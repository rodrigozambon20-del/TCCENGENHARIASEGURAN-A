"""
Cliente Binance via CCXT com política de segurança máxima.

REGRAS DE SEGURANÇA (não negociáveis):
1. API key criada com permissões MÍNIMAS: "Enable Reading" +
   "Enable Spot & Margin Trading". SAQUE (withdraw) SEMPRE DESABILITADO.
2. Restrição de IP na Binance: a key só funciona a partir do IP do servidor.
3. Segredos apenas em variáveis de ambiente / secret manager. Nunca em
   código, YAML ou git.
4. `enableRateLimit=True` no ccxt + retry com backoff exponencial e
   respeito a HTTP 429/418 (ban temporário da Binance).
5. Diferenciar erros RECUPERÁVEIS (rede, timeout, rate limit -> retry)
   de erros FATAIS (fundos insuficientes, key inválida -> não retentar,
   alertar).
6. `testnet=True` por padrão — produção exige mudança explícita de config.
7. Stops de proteção são criados NA EXCHANGE (stop-limit/OCO), não apenas
   na memória do bot: se o processo morrer, a posição continua protegida.
"""
from __future__ import annotations

import asyncio
import logging
import os

import ccxt.async_support as ccxt

from core.state import Position

log = logging.getLogger("binance")

RETRYABLE = (ccxt.NetworkError, ccxt.ExchangeNotAvailable,
             ccxt.RequestTimeout, ccxt.DDoSProtection, ccxt.RateLimitExceeded)
FATAL = (ccxt.AuthenticationError, ccxt.PermissionDenied,
         ccxt.InsufficientFunds, ccxt.InvalidOrder)

MAX_RETRIES = 4
BASE_BACKOFF = 2.0  # 2s, 4s, 8s, 16s


class BinanceClient:
    def __init__(self, api_key: str, api_secret: str, testnet: bool = True,
                 quote_currency: str = "USDT", market_type: str = "spot",
                 leverage: int = 1) -> None:
        self.quote = quote_currency
        self.market_type = market_type       # "spot" | "futures"
        self.leverage = leverage
        default_type = "future" if market_type == "futures" else "spot"
        options = {
            "defaultType": default_type,
            "adjustForTimeDifference": True,  # evita erro -1021 (timestamp)
        }
        if market_type == "futures":
            # Carrega só os mercados USDT-M — boot mais rápido, menos falhas
            options["fetchMarkets"] = ["linear"]
        self.exchange = ccxt.binance({
            "apiKey": api_key,
            "secret": api_secret,
            "enableRateLimit": True,          # throttle automático do ccxt
            "options": options,
        })
        if testnet:
            if market_type == "futures":
                # A "Demo Trading" nova da Binance (onde a chave é criada em
                # demo.binance.com) usa o host demo-fapi.binance.com — NÃO o
                # antigo testnet.binancefuture.com. ccxt descontinuou o sandbox
                # de futuros, então montamos as URLs de futuros à mão.
                demo_host = os.environ.get("BINANCE_FUTURES_DEMO_HOST",
                                           "https://demo-fapi.binance.com")
                for k, v in self.exchange.urls.get("test", {}).items():
                    if k.startswith("fapi"):
                        # preserva o sufixo /fapi/vN, troca só o host
                        suffix = v.split("testnet.binancefuture.com", 1)[-1]
                        self.exchange.urls["api"][k] = demo_host + suffix
                # fetch_currencies usa um endpoint spot de produção que a
                # testnet de futuros não tem — desativa (não é usado no trade).
                self.exchange.has["fetchCurrencies"] = False
            else:
                self.exchange.set_sandbox_mode(True)
            log.warning("MODO TESTNET ATIVO — nenhuma ordem real será enviada")
        log.warning("Mercado: %s%s", market_type.upper(),
                    f" | alavancagem {leverage}x" if market_type == "futures" else "")

    async def setup_futures(self, symbols: list[str]) -> None:
        """Configura alavancagem e margem ISOLADA (mais segura) por símbolo."""
        if self.market_type != "futures":
            return
        for symbol in symbols:
            try:
                await self._call(self.exchange.set_margin_mode, "isolated", symbol)
            except Exception as exc:
                log.info("Margem isolada %s: %s (pode já estar configurada)",
                         symbol, exc)
            try:
                await self._call(self.exchange.set_leverage, self.leverage, symbol)
                log.info("%s: alavancagem %dx, margem isolada", symbol, self.leverage)
            except Exception as exc:
                log.warning("Falha ao configurar alavancagem de %s: %s", symbol, exc)

    async def close(self) -> None:
        await self.exchange.close()

    # ------------------------------------------------------------------ #
    # Wrapper com retry/backoff e classificação de erros                 #
    # ------------------------------------------------------------------ #
    async def _call(self, fn, *args, **kwargs):
        delay = BASE_BACKOFF
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                return await fn(*args, **kwargs)
            except FATAL as exc:
                # Erro fatal: retentar seria inútil ou perigoso (ordem duplicada)
                log.error("Erro FATAL em %s: %s — sem retry", fn.__name__, exc)
                raise
            except RETRYABLE as exc:
                if attempt == MAX_RETRIES:
                    log.error("%s falhou após %d tentativas: %s",
                              fn.__name__, MAX_RETRIES, exc)
                    raise
                log.warning("%s: erro recuperável (%s), retry %d/%d em %.0fs",
                            fn.__name__, type(exc).__name__, attempt,
                            MAX_RETRIES, delay)
                await asyncio.sleep(delay)
                delay *= 2

    # ------------------------------------------------------------------ #
    # Dados de mercado                                                   #
    # ------------------------------------------------------------------ #
    async def fetch_ohlcv(self, symbol: str, timeframe: str, limit: int = 200):
        return await self._call(self.exchange.fetch_ohlcv, symbol,
                                timeframe=timeframe, limit=limit)

    async def fetch_order_book(self, symbol: str, depth: int = 20):
        return await self._call(self.exchange.fetch_order_book, symbol, depth)

    async def fetch_last_price(self, symbol: str) -> float:
        ticker = await self._call(self.exchange.fetch_ticker, symbol)
        return float(ticker["last"])

    async def fetch_total_equity(self) -> float | None:
        """Equity total da conta convertida para a moeda de cotação.

        Usa fetch_tickers (1 chamada para todos os preços) em vez de um
        request por ativo — inicialização e watchdog muito mais rápidos.
        """
        balance = await self._call(self.exchange.fetch_balance)
        total = float(balance["total"].get(self.quote, 0.0))
        others = {a: q for a, q in balance["total"].items()
                  if a != self.quote and q}
        if not others:
            return total
        tickers = await self._call(self.exchange.fetch_tickers)
        for asset, qty in others.items():
            ticker = tickers.get(f"{asset}/{self.quote}")
            if ticker and ticker.get("last"):
                total += qty * float(ticker["last"])
        return total

    # ------------------------------------------------------------------ #
    # Ordens                                                             #
    # ------------------------------------------------------------------ #
    async def create_limit_order(self, symbol: str, side: str, qty: float,
                                 price: float, post_only: bool = True):
        qty = float(self.exchange.amount_to_precision(symbol, qty))
        price = float(self.exchange.price_to_precision(symbol, price))
        params = {"timeInForce": "GTX"} if post_only else {}  # GTX = post-only
        return await self._call(self.exchange.create_order, symbol, "limit",
                                side, qty, price, params)

    async def create_market_order(self, symbol: str, side: str, qty: float):
        qty = float(self.exchange.amount_to_precision(symbol, qty))
        return await self._call(self.exchange.create_order, symbol, "market",
                                side, qty)

    async def fetch_order(self, order_id: str, symbol: str):
        return await self._call(self.exchange.fetch_order, order_id, symbol)

    async def cancel_order(self, order_id: str, symbol: str):
        try:
            return await self._call(self.exchange.cancel_order, order_id, symbol)
        except ccxt.OrderNotFound:
            return None  # já preenchida/cancelada — ok

    async def cancel_all_orders(self, symbol: str):
        try:
            return await self._call(self.exchange.cancel_all_orders, symbol)
        except (ccxt.InvalidOrder, ccxt.OrderNotFound):
            return None  # Binance -2011: não havia ordens abertas — nada a fazer

    # ------------------------------------------------------------------ #
    # Proteção na exchange (stop + take profit)                          #
    # ------------------------------------------------------------------ #
    async def place_protective_orders(self, symbol: str, pos: Position):
        """
        Proteção NA EXCHANGE (sobrevive a quedas do bot):
        - spot: ordem OCO (stop-limit + take-profit) que se auto-cancela;
        - futures: STOP_MARKET + TAKE_PROFIT_MARKET reduce-only (fecham a
          posição no gatilho, cada um cobrindo um lado).
        """
        side = "sell" if pos.side == "long" else "buy"
        qty = float(self.exchange.amount_to_precision(symbol, pos.qty))
        tp = float(self.exchange.price_to_precision(symbol, pos.take_profit))
        sl = float(self.exchange.price_to_precision(symbol, pos.stop_loss))

        if self.market_type == "futures":
            # Stop de perda: fecha a posição a mercado no gatilho
            await self._call(self.exchange.create_order, symbol, "STOP_MARKET",
                             side, qty, None,
                             {"stopPrice": sl, "reduceOnly": True})
            # Alvo de lucro: idem, do outro lado
            return await self._call(self.exchange.create_order, symbol,
                                    "TAKE_PROFIT_MARKET", side, qty, None,
                                    {"stopPrice": tp, "reduceOnly": True})

        # spot: OCO stop-limit ligeiramente dentro do gatilho p/ garantir fill
        sl_limit = float(self.exchange.price_to_precision(
            symbol, sl * (0.998 if side == "sell" else 1.002)))
        return await self._call(
            self.exchange.create_order, symbol, "limit", side, qty, tp,
            {"stopPrice": sl, "stopLimitPrice": sl_limit, "type": "oco"})

    async def replace_stop_order(self, symbol: str, pos: Position):
        """Trailing stop: cancela a proteção antiga e recria com stop novo."""
        await self.cancel_all_orders(symbol)
        return await self.place_protective_orders(symbol, pos)
