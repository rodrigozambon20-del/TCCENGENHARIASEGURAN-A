"""
Classe base de todos os agentes.

Cada agente é uma corrotina de longa duração com:
- loop de consumo de eventos do bus (reativo) e/ou loop periódico (proativo);
- heartbeat para monitoramento de saúde (consumido pelo Notifier/n8n);
- tratamento de exceções que NUNCA derruba o processo: um agente que
  falha reinicia com backoff, e falhas repetidas disparam DEFENSIVE_MODE.
"""
from __future__ import annotations

import asyncio
import logging
import time
from abc import ABC, abstractmethod

from core.config import AppConfig
from core.event_bus import EventBus
from core.events import BaseEvent, DefensiveMode, Topic
from core.state import BotState

MAX_CRASHES_BEFORE_DEFENSIVE = 3
HEARTBEAT_INTERVAL = 30


class BaseAgent(ABC):
    name: str = "base"

    def __init__(self, bus: EventBus, state: BotState, config: AppConfig) -> None:
        self.bus = bus
        self.state = state
        self.config = config
        self.log = logging.getLogger(self.name)
        self._crash_count = 0

    @abstractmethod
    async def run(self) -> None:
        """Loop principal do agente (implementado por cada especialista)."""

    async def start(self) -> None:
        """Executa o agente com supervisão: reinício com backoff exponencial."""
        asyncio.create_task(self._heartbeat_loop())
        backoff = 1.0
        while True:
            try:
                await self.run()
                return  # saída limpa
            except asyncio.CancelledError:
                raise
            except Exception:
                self._crash_count += 1
                self.log.exception("Agente %s falhou (%d/%d). Reiniciando em %.0fs",
                                   self.name, self._crash_count,
                                   MAX_CRASHES_BEFORE_DEFENSIVE, backoff)
                if self._crash_count >= MAX_CRASHES_BEFORE_DEFENSIVE:
                    await self.bus.publish(Topic.DEFENSIVE_MODE, DefensiveMode(
                        source=self.name, enabled=True,
                        reason=f"Agente {self.name} instável ({self._crash_count} falhas)"))
                await asyncio.sleep(backoff)
                backoff = min(backoff * 2, 60)

    async def _heartbeat_loop(self) -> None:
        while True:
            hb = BaseEvent(source=self.name)
            await self.bus.publish(Topic.HEARTBEAT, hb)
            await asyncio.sleep(HEARTBEAT_INTERVAL)

    async def emit(self, topic: Topic, event: BaseEvent) -> None:
        event.source = self.name
        await self.bus.publish(topic, event)
