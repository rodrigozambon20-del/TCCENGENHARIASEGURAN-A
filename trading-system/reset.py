"""
Reset de conta: cancela TODAS as ordens abertas e fecha TODAS as posições
(a mercado) da conta configurada. Útil para limpar acúmulos de testnet.

Uso:
    python reset.py                         # usa config/config.yaml (spot)
    python reset.py config/config.futures.yaml   # futuros
"""
from __future__ import annotations

import asyncio
import logging
import sys

from core.config import AppConfig
from exchange.binance_client import BinanceClient

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger("reset")


async def main() -> None:
    path = sys.argv[1] if len(sys.argv) > 1 else "config/config.yaml"
    cfg = AppConfig.load(path)
    if not cfg.binance_api_key:
        raise SystemExit("Defina as chaves no .env/.env.futures")

    ex = BinanceClient(cfg.binance_api_key, cfg.binance_api_secret,
                       testnet=cfg.testnet, market_type=cfg.market_type,
                       leverage=cfg.leverage)
    try:
        log.info("Cancelando ordens abertas...")
        for symbol in cfg.symbols:
            try:
                await ex.cancel_all_orders(symbol)
                log.info("  ordens de %s canceladas", symbol)
            except Exception as exc:
                log.warning("  %s: %s", symbol, exc)

        if cfg.market_type == "futures":
            log.info("Fechando posições abertas...")
            positions = await ex.fetch_open_positions()
            if not positions:
                log.info("  nenhuma posição aberta")
            for p in positions:
                sym = p["symbol"]
                side = p.get("side") or ("long" if float(p["contracts"]) > 0 else "short")
                qty = abs(float(p["contracts"]))
                try:
                    await ex.close_position_market(sym, side, qty)
                    log.info("  posição %s %s %.4f FECHADA", sym, side, qty)
                except Exception as exc:
                    log.warning("  falha ao fechar %s: %s", sym, exc)
        log.info("Reset concluído. Conta limpa.")
    finally:
        await ex.close()


if __name__ == "__main__":
    asyncio.run(main())
