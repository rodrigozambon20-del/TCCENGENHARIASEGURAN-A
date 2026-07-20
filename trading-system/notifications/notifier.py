"""
Notifier — ponte entre o Event Bus e o n8n (Telegram/WhatsApp/e-mail).

Arquitetura de integração com n8n:

  Bot ──HTTP POST──> n8n Webhook Node ──Switch(level)──┬─> Telegram Node
                                                        ├─> WhatsApp (Twilio/Evolution)
                                                        ├─> Gmail / Google Tasks
                                                        └─> Google Sheets (log de trades)

O bot só conhece UMA URL (o webhook do n8n). Todo o roteamento,
formatação por canal, deduplicação e escalonamento (ex: só acordar o
dono de madrugada se level == critical) fica no n8n — trocar de canal
não exige tocar no código do bot.

Resiliência: fila local com reenvio; notificação NUNCA pode travar o
pipeline de trading (fire-and-forget com timeout curto).
"""
from __future__ import annotations

import asyncio
import logging

import aiohttp

from core.event_bus import EventBus
from core.events import (ExecutionReport, KillSwitch, Notification, Topic)

log = logging.getLogger("notifier")

SEND_TIMEOUT = 5
MAX_QUEUE = 500


class Notifier:
    def __init__(self, bus: EventBus, webhook_url: str) -> None:
        self.bus = bus
        self.webhook_url = webhook_url
        self.outbox: asyncio.Queue = asyncio.Queue(maxsize=MAX_QUEUE)

    async def run(self) -> None:
        if not self.webhook_url:
            log.warning("N8N_WEBHOOK_URL não configurada — notificações desativadas")
            return
        inbox = self.bus.subscribe(Topic.NOTIFICATION, Topic.EXECUTION_REPORT,
                                   Topic.KILL_SWITCH, Topic.HEARTBEAT)
        asyncio.create_task(self._sender_loop())
        heartbeat_agg: dict[str, float] = {}

        async for topic, event in self.bus.stream(inbox):
            if topic == Topic.NOTIFICATION:
                self._enqueue(self._from_notification(event))
            elif topic == Topic.EXECUTION_REPORT:
                self._enqueue(self._from_execution(event))
            elif topic == Topic.KILL_SWITCH:
                self._enqueue({"level": "critical", "title": "🛑 KILL SWITCH",
                               "body": event.reason, "source": event.source})
            elif topic == Topic.HEARTBEAT:
                heartbeat_agg[event.source] = event.ts
                # Health-check é puxado pelo n8n via Schedule Trigger + HTTP
                # Request no endpoint /health do bot; aqui só agregamos.

    def _enqueue(self, payload: dict) -> None:
        try:
            self.outbox.put_nowait(payload)
        except asyncio.QueueFull:
            self.outbox.get_nowait()  # descarta a mais antiga
            self.outbox.put_nowait(payload)

    async def _sender_loop(self) -> None:
        async with aiohttp.ClientSession() as session:
            while True:
                payload = await self.outbox.get()
                for attempt in range(3):
                    try:
                        async with session.post(
                                self.webhook_url, json=payload,
                                timeout=aiohttp.ClientTimeout(total=SEND_TIMEOUT)) as r:
                            if r.status < 300:
                                break
                            log.warning("n8n respondeu %s (tentativa %d)",
                                        r.status, attempt + 1)
                    except Exception as exc:
                        log.warning("Envio ao n8n falhou: %s (tentativa %d)",
                                    exc, attempt + 1)
                    await asyncio.sleep(2 ** attempt)

    @staticmethod
    def _from_notification(n: Notification) -> dict:
        return {"level": n.level, "title": n.title, "body": n.body,
                "source": n.source, "ts": n.ts, **({"data": n.payload} if n.payload else {})}

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
