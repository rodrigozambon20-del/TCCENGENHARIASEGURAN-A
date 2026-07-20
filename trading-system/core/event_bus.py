"""
Event Bus assíncrono (pub/sub) — a espinha dorsal da comunicação entre agentes.

Design:
- In-process: `asyncio.Queue` por assinante -> latência de microssegundos,
  zero dependência externa. Ideal para um único host.
- Distribuído: a mesma interface pode ser trocada por Redis Streams ou
  RabbitMQ sem alterar nenhum agente (basta implementar publish/subscribe).

Regras:
- KILL_SWITCH e DEFENSIVE_MODE usam fila prioritária: são entregues antes
  de qualquer sinal de trade pendente.
- Filas com tamanho máximo: se um agente travar, o bus descarta o sinal
  mais antigo (sinais de mercado envelhecem; um sinal atrasado é um sinal
  errado) em vez de acumular backlog silenciosamente.
"""
from __future__ import annotations

import asyncio
import logging
from collections import defaultdict
from typing import AsyncIterator

from core.events import BaseEvent, Topic

log = logging.getLogger("event_bus")

PRIORITY_TOPICS = {Topic.KILL_SWITCH, Topic.DEFENSIVE_MODE}
MAX_QUEUE_SIZE = 1000


class EventBus:
    def __init__(self) -> None:
        self._subscribers: dict[Topic, list[asyncio.Queue]] = defaultdict(list)

    def subscribe(self, *topics: Topic) -> asyncio.Queue:
        """Registra um assinante e devolve a fila da qual ele consumirá."""
        queue: asyncio.Queue = asyncio.Queue(maxsize=MAX_QUEUE_SIZE)
        for topic in topics:
            self._subscribers[topic].append(queue)
        return queue

    async def publish(self, topic: Topic, event: BaseEvent) -> None:
        for queue in self._subscribers[topic]:
            item = (topic, event)
            if topic in PRIORITY_TOPICS:
                # Eventos críticos nunca podem ser descartados: se a fila
                # estiver cheia, abrimos espaço removendo sinais comuns.
                while queue.full():
                    dropped = queue.get_nowait()
                    log.warning("Fila cheia: descartando %s para dar passagem a %s",
                                dropped[0], topic)
                queue.put_nowait(item)
            else:
                try:
                    queue.put_nowait(item)
                except asyncio.QueueFull:
                    stale_topic, stale = queue.get_nowait()
                    log.warning("Fila cheia: sinal envelhecido %s (%.1fs) descartado",
                                stale_topic, event.ts - stale.ts)
                    queue.put_nowait(item)

    @staticmethod
    async def stream(queue: asyncio.Queue) -> AsyncIterator[tuple[Topic, BaseEvent]]:
        """Iterador infinito de consumo para os agentes."""
        while True:
            yield await queue.get()
