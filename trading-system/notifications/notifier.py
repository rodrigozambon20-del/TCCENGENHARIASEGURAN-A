"""
Notifier — envia eventos do bot para o dono, por dois canais independentes:

1. TELEGRAM DIRETO (recomendado, zero infraestrutura):
   defina TELEGRAM_BOT_TOKEN e TELEGRAM_CHAT_ID no .env e pronto —
   placar, trades e alertas chegam no celular via Bot API oficial.

2. n8n WEBHOOK (opcional, para roteamento avançado):
   defina N8N_WEBHOOK_URL e o fluxo n8n_workflow.json roteia por
   severidade para Telegram/WhatsApp/Google Sheets.

Os dois podem ficar ativos ao mesmo tempo. Resiliência: fila local com
reenvio; notificação NUNCA trava o pipeline de trading (timeout curto,
fire-and-forget).
"""
from __future__ import annotations

import asyncio
import logging

import aiohttp

from core.config import AppConfig
from core.event_bus import EventBus
from core.events import ExecutionReport, Notification, Topic

log = logging.getLogger("notifier")

SEND_TIMEOUT = 8
MAX_QUEUE = 500
LEVEL_EMOJI = {"info": "ℹ️", "warning": "⚠️", "critical": "🚨"}


class Notifier:
    def __init__(self, bus: EventBus, config: AppConfig) -> None:
        self.bus = bus
        self.webhook_url = config.n8n_webhook_url
        self.tg_token = config.telegram_bot_token
        self.tg_chat = config.telegram_chat_id
        self.outbox: asyncio.Queue = asyncio.Queue(maxsize=MAX_QUEUE)

    @property
    def telegram_enabled(self) -> bool:
        return bool(self.tg_token and self.tg_chat)

    async def run(self) -> None:
        if not (self.webhook_url or self.telegram_enabled):
            log.warning("Nenhum canal de notificação configurado "
                        "(TELEGRAM_BOT_TOKEN/CHAT_ID ou N8N_WEBHOOK_URL) — "
                        "notificações desativadas")
            return
        channels = []
        if self.telegram_enabled:
            channels.append("Telegram")
        if self.webhook_url:
            channels.append("n8n")
        log.info("Notificações ativas via: %s", ", ".join(channels))

        inbox = self.bus.subscribe(Topic.NOTIFICATION, Topic.EXECUTION_REPORT,
                                   Topic.KILL_SWITCH)
        asyncio.create_task(self._sender_loop())

        async for topic, event in self.bus.stream(inbox):
            if topic == Topic.NOTIFICATION:
                payload = self._from_notification(event)
                # Telegram: só resultado de trade (to_telegram) ou crítico.
                payload["_telegram"] = bool(getattr(event, "to_telegram", False)
                                            or event.level == "critical")
                self._enqueue(payload)
            elif topic == Topic.EXECUTION_REPORT:
                payload = self._from_execution(event)
                payload["_telegram"] = False   # entradas não vão ao Telegram
                self._enqueue(payload)
            elif topic == Topic.KILL_SWITCH:
                self._enqueue({"level": "critical", "title": "🛑 KILL SWITCH",
                               "body": event.reason, "source": event.source,
                               "_telegram": True})

    def _enqueue(self, payload: dict) -> None:
        try:
            self.outbox.put_nowait(payload)
        except asyncio.QueueFull:
            self.outbox.get_nowait()  # descarta a mais antiga
            self.outbox.put_nowait(payload)

    # ------------------------------------------------------------------ #
    # Loop de envio (com retry) para todos os canais ativos              #
    # ------------------------------------------------------------------ #
    async def _sender_loop(self) -> None:
        timeout = aiohttp.ClientTimeout(total=SEND_TIMEOUT)
        async with aiohttp.ClientSession(timeout=timeout) as session:
            while True:
                payload = await self.outbox.get()
                # Telegram só recebe o que estiver marcado (resultado de trade
                # / crítico). O n8n, se configurado, recebe tudo.
                if self.telegram_enabled and payload.get("_telegram"):
                    await self._post_retry(session, self._telegram_request(payload))
                if self.webhook_url:
                    await self._post_retry(session,
                                           {"url": self.webhook_url, "json": payload})

    def _telegram_request(self, payload: dict) -> dict:
        emoji = LEVEL_EMOJI.get(payload.get("level", "info"), "")
        text = f"{emoji} {payload.get('title', '')}\n{payload.get('body', '')}".strip()
        return {
            "url": f"https://api.telegram.org/bot{self.tg_token}/sendMessage",
            "json": {"chat_id": self.tg_chat, "text": text[:4000],
                     "disable_notification": payload.get("level") == "info"},
        }

    @staticmethod
    async def _post_retry(session: aiohttp.ClientSession, req: dict) -> None:
        for attempt in range(3):
            try:
                async with session.post(req["url"], json=req["json"]) as r:
                    if r.status < 300:
                        return
                    body = await r.text()
                    log.warning("Canal respondeu %s: %.200s (tentativa %d)",
                                r.status, body, attempt + 1)
            except Exception as exc:
                log.warning("Envio falhou: %s (tentativa %d)", exc, attempt + 1)
            await asyncio.sleep(2 ** attempt)

    # ------------------------------------------------------------------ #
    # Formatação                                                         #
    # ------------------------------------------------------------------ #
    @staticmethod
    def _from_notification(n: Notification) -> dict:
        return {"level": n.level, "title": n.title, "body": n.body,
                "source": n.source, "ts": n.ts,
                **({"data": n.payload} if n.payload else {})}

    @staticmethod
    def _from_execution(r: ExecutionReport) -> dict:
        emoji = "✅" if r.status == "filled" else "⚠️"
        return {
            "level": "info" if r.status == "filled" else "warning",
            "title": f"{emoji} Execução {r.symbol}: {r.status}",
            "body": (f"{r.side} {r.filled_qty} @ {r.fill_price:.2f} | "
                     f"slippage {r.slippage_bps:.1f} bps | taxa {r.fees:.4f}"
                     + (f"\nErro: {r.error}" if r.error else "")),
            "source": r.source, "ts": r.ts,
        }
