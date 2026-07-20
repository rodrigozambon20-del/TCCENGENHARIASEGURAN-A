"""
Entrypoint — sobe o ecossistema completo de agentes.

Ordem de inicialização (importa!):
1. Config + estado global
2. Cliente de exchange (valida credenciais ANTES de subir agentes)
3. Event Bus
4. Notifier (para que falhas de boot já sejam notificadas)
5. Risk Manager (o guardião sobe antes de quem propõe trades)
6. Demais agentes

Desligamento gracioso: SIGINT/SIGTERM cancela ordens abertas antes de
encerrar (posições ficam protegidas pelos stops OCO na exchange).
"""
from __future__ import annotations

import asyncio
import logging
import signal

from agents.execution_agent import ExecutionAgent
from agents.orchestrator_agent import OrchestratorAgent
from agents.quant_agent import QuantAgent
from agents.risk_agent import RiskAgent
from agents.sentiment_agent import SentimentAgent
from core.config import AppConfig
from core.event_bus import EventBus
from core.state import BotState
from exchange.binance_client import BinanceClient
from notifications.notifier import Notifier

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s [%(name)s] %(message)s",
)
log = logging.getLogger("main")


async def main() -> None:
    config = AppConfig.load()
    if not config.binance_api_key or not config.binance_api_secret:
        raise SystemExit("Defina BINANCE_API_KEY e BINANCE_API_SECRET no ambiente (.env)")

    exchange = BinanceClient(config.binance_api_key, config.binance_api_secret,
                             testnet=config.testnet)
    # Fail-fast: valida credenciais e conectividade antes de qualquer agente
    equity = await exchange.fetch_total_equity()
    log.info("Conectado à Binance (%s). Equity inicial: %.2f USDT",
             "TESTNET" if config.testnet else "PRODUÇÃO", equity or 0.0)

    bus = EventBus()
    state = BotState(initial_equity=equity or config.initial_equity)

    notifier = Notifier(bus, config.n8n_webhook_url)
    agents = [
        RiskAgent(bus, state, config, exchange),        # guardião primeiro
        ExecutionAgent(bus, state, config, exchange),
        QuantAgent(bus, state, config, exchange),
        SentimentAgent(bus, state, config),
        OrchestratorAgent(bus, state, config),
    ]

    tasks = [asyncio.create_task(notifier.run(), name="notifier")]
    tasks += [asyncio.create_task(a.start(), name=a.name) for a in agents]
    tasks.append(asyncio.create_task(_daily_roll(state), name="daily_roll"))

    stop = asyncio.Event()
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGINT, signal.SIGTERM):
        loop.add_signal_handler(sig, stop.set)

    log.info("Sistema no ar com %d agentes. Ctrl+C para desligar.", len(agents))
    await stop.wait()

    log.info("Desligamento gracioso: cancelando ordens abertas...")
    for symbol in config.symbols:
        try:
            await exchange.cancel_all_orders(symbol)
        except Exception as exc:
            log.error("Falha ao cancelar ordens de %s: %s", symbol, exc)
    for t in tasks:
        t.cancel()
    await asyncio.gather(*tasks, return_exceptions=True)
    await exchange.close()
    log.info("Encerrado. Posições abertas seguem protegidas por OCO na exchange.")


async def _daily_roll(state: BotState) -> None:
    """Vira o marcador de drawdown diário à meia-noite UTC e semanal na segunda."""
    import datetime as dt
    while True:
        now = dt.datetime.now(dt.timezone.utc)
        tomorrow = (now + dt.timedelta(days=1)).replace(hour=0, minute=0,
                                                        second=0, microsecond=0)
        await asyncio.sleep((tomorrow - now).total_seconds())
        await state.roll_daily()
        if dt.datetime.now(dt.timezone.utc).weekday() == 0:
            await state.roll_weekly()


if __name__ == "__main__":
    asyncio.run(main())
