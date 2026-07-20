"""
Gerenciador de WebSockets da Binance com reconexão resiliente.

Regras de resiliência:
1. Reconexão automática com backoff exponencial + jitter (evita
   thundering herd quando a Binance reinicia endpoints).
2. Watchdog de silêncio: se nenhuma mensagem chegar em N segundos,
   a conexão é considerada morta e derrubada à força (o TCP pode
   ficar "zumbi" sem erro explícito).
3. A Binance encerra streams a cada 24h — a reconexão é tratada como
   evento NORMAL, não como falha.
4. Se o stream ficar indisponível por mais de `max_outage_seconds`,
   publica DEFENSIVE_MODE: operar cego é inaceitável.
5. Ao reconectar, o consumidor deve RESSINCRONIZAR estado via REST
   (snapshot do book / ordens abertas) antes de confiar no stream.
"""
from __future__ import annotations

import asyncio
import json
import logging
import random
import time
from typing import Awaitable, Callable

import websockets

from core.event_bus import EventBus
from core.events import DefensiveMode, Topic

log = logging.getLogger("ws_manager")

BINANCE_WS = "wss://stream.binance.com:9443/stream"
BINANCE_TESTNET_WS = "wss://stream.testnet.binance.vision/stream"

SILENCE_TIMEOUT = 60          # s sem mensagens => conexão zumbi
MAX_BACKOFF = 60.0
MAX_OUTAGE_BEFORE_DEFENSIVE = 120


class WebSocketManager:
    def __init__(self, bus: EventBus, streams: list[str],
                 on_message: Callable[[dict], Awaitable[None]],
                 on_resync: Callable[[], Awaitable[None]] | None = None,
                 testnet: bool = True) -> None:
        self.bus = bus
        self.streams = streams          # ex: ["btcusdt@kline_1m", "btcusdt@depth20@100ms"]
        self.on_message = on_message
        self.on_resync = on_resync      # ressincronização REST pós-reconexão
        self.url = (BINANCE_TESTNET_WS if testnet else BINANCE_WS) \
            + "?streams=" + "/".join(streams)
        self._last_msg_at = time.time()
        self._outage_started: float | None = None

    async def run_forever(self) -> None:
        backoff = 1.0
        while True:
            try:
                async with websockets.connect(self.url, ping_interval=20,
                                              ping_timeout=10,
                                              close_timeout=5) as ws:
                    log.info("WebSocket conectado (%d streams)", len(self.streams))
                    backoff = 1.0
                    self._outage_started = None
                    if self.on_resync:
                        await self.on_resync()   # snapshot REST antes de confiar no stream
                    watchdog = asyncio.create_task(self._silence_watchdog(ws))
                    try:
                        async for raw in ws:
                            self._last_msg_at = time.time()
                            try:
                                await self.on_message(json.loads(raw))
                            except Exception:
                                log.exception("Handler de mensagem falhou (stream segue)")
                    finally:
                        watchdog.cancel()
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                log.warning("WebSocket caiu (%s). Reconectando em %.1fs",
                            exc, backoff)
                await self._track_outage()
                # Backoff exponencial com jitter
                await asyncio.sleep(backoff + random.uniform(0, backoff / 2))
                backoff = min(backoff * 2, MAX_BACKOFF)

    async def _silence_watchdog(self, ws) -> None:
        while True:
            await asyncio.sleep(SILENCE_TIMEOUT / 2)
            if time.time() - self._last_msg_at > SILENCE_TIMEOUT:
                log.warning("Sem mensagens há %ds — derrubando conexão zumbi",
                            SILENCE_TIMEOUT)
                await ws.close()
                return

    async def _track_outage(self) -> None:
        now = time.time()
        if self._outage_started is None:
            self._outage_started = now
        elif now - self._outage_started > MAX_OUTAGE_BEFORE_DEFENSIVE:
            log.critical("Stream fora há %ds — ativando modo defensivo",
                         int(now - self._outage_started))
            await self.bus.publish(Topic.DEFENSIVE_MODE, DefensiveMode(
                source="ws_manager", enabled=True,
                reason="Perda prolongada de dados de mercado (WebSocket)",
                cooldown_seconds=600))
            self._outage_started = now  # evita spam do evento
