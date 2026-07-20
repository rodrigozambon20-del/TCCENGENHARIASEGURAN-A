"""
AGENTE 2 — QUANT E ANÁLISE TÉCNICA (Quantitative Analytics)

Responsabilidades:
1. Varredura contínua de múltiplos timeframes (1m, 5m, 15m, 1h, 1d) via
   OHLCV do ccxt + WebSocket de klines para o candle corrente.
2. Indicadores: KAMA (média adaptativa de Kaufman), RSI, Bandas de
   Bollinger, ATR (para stops) e desequilíbrio do livro de ordens.
3. Suporte/resistência por fractais de pivô + confirmação de volume.
4. Fusão dos timeframes em um score único ponderado (timeframes maiores
   pesam mais — tendência manda, timeframe curto só refina a entrada).

O agente NUNCA decide operar: apenas publica QuantSignal com evidência
matemática. Decisão é do Orquestrador; aprovação é do Risco.
"""
from __future__ import annotations

import asyncio

import numpy as np

from agents.base import BaseAgent
from core.events import Direction, QuantSignal, Topic

SCAN_INTERVAL = 15  # segundos entre varreduras completas


class QuantAgent(BaseAgent):
    name = "quant_agent"

    def __init__(self, bus, state, config, market_data) -> None:
        super().__init__(bus, state, config)
        self.md = market_data  # exchange.binance_client.BinanceClient

    async def run(self) -> None:
        while True:
            for symbol in self.config.symbols:
                try:
                    signal = await self.analyze_symbol(symbol)
                    if signal and abs(signal.score) >= self.config.quant.min_score_to_signal:
                        await self.emit(Topic.QUANT_SIGNAL, signal)
                        self.log.info("%s: %s score=%.2f votos=%s", symbol,
                                      signal.direction.value, signal.score,
                                      signal.timeframe_votes)
                except Exception as exc:
                    self.log.warning("Análise de %s falhou: %s", symbol, exc)
            await asyncio.sleep(SCAN_INTERVAL)

    # ------------------------------------------------------------------ #
    # Pipeline de análise                                                #
    # ------------------------------------------------------------------ #
    async def analyze_symbol(self, symbol: str) -> QuantSignal | None:
        cfg = self.config.quant
        votes: dict[str, float] = {}
        closes_by_tf: dict[str, np.ndarray] = {}
        atr_1h = None
        last_close = None
        support = resistance = None

        for tf in cfg.timeframes:
            ohlcv = await self.md.fetch_ohlcv(symbol, tf, limit=200)
            if len(ohlcv) < 50:
                continue
            closes = np.array([c[4] for c in ohlcv], dtype=float)
            highs = np.array([c[2] for c in ohlcv], dtype=float)
            lows = np.array([c[3] for c in ohlcv], dtype=float)
            last_close = closes[-1]
            closes_by_tf[tf] = closes

            votes[tf] = self.score_timeframe(closes, cfg)
            if tf == "1h":
                atr_1h = self.atr(highs, lows, closes, cfg.atr_period)
                support, resistance = self.pivot_levels(highs, lows, last_close)

        if not votes or last_close is None:
            return None

        # Order flow: desequilíbrio bid/ask nos topos do livro
        imbalance = await self.orderbook_imbalance(symbol)

        # ------------------------------------------------------------------
        # MODO PULLBACK (vencedor do laboratório de backtest 2024-2026):
        # só compra o RECUO dentro de tendência de ALTA do gráfico diário.
        # Condições: preço > SMA(50) diária E RSI(1h) <= 45 E preço <= média
        # das Bandas de Bollinger 1h. Saída fica por conta do trailing stop.
        # ------------------------------------------------------------------
        if cfg.mode == "pullback":
            daily = closes_by_tf.get("1d")
            hourly = closes_by_tf.get("1h")
            if daily is None or hourly is None or len(daily) < cfg.trend_sma_days:
                return None
            price = float(hourly[-1])
            sma_daily = float(daily[-cfg.trend_sma_days:].mean())
            rsi_1h = self.rsi(hourly, cfg.rsi_period)
            bb_mid_1h = float(hourly[-cfg.bollinger_period:].mean())

            uptrend = price > sma_daily
            in_pullback = rsi_1h <= cfg.pullback_rsi and price <= bb_mid_1h
            if not (uptrend and in_pullback):
                return None
            self.log.info("%s: PULLBACK detectado (preço %.2f > SMA%dd %.2f, "
                          "RSI1h %.1f, abaixo da média BB)", symbol, price,
                          cfg.trend_sma_days, sma_daily, rsi_1h)
            return QuantSignal(
                symbol=symbol, direction=Direction.LONG, strategy="pullback",
                score=0.8, timeframe_votes=votes, support=support,
                resistance=resistance, atr=atr_1h, last_price=price,
                orderflow_imbalance=float(imbalance),
            )

        # Modo VOTES (original): média ponderada dos timeframes + order flow
        total_w = sum(cfg.timeframe_weights.get(tf, 0.1) for tf in votes)
        score = sum(v * cfg.timeframe_weights.get(tf, 0.1) for tf, v in votes.items()) / total_w
        score = 0.8 * score + 0.2 * imbalance

        direction = (Direction.LONG if score > 0
                     else Direction.SHORT if score < 0
                     else Direction.FLAT)
        return QuantSignal(
            symbol=symbol, direction=direction, score=float(score),
            timeframe_votes=votes, support=support, resistance=resistance,
            atr=atr_1h, last_price=float(last_close),
            orderflow_imbalance=float(imbalance),
        )

    # ------------------------------------------------------------------ #
    # Indicadores                                                        #
    # ------------------------------------------------------------------ #
    def score_timeframe(self, closes: np.ndarray, cfg) -> float:
        """Combina KAMA + RSI + Bollinger num voto -1..+1 para o timeframe."""
        price = closes[-1]
        kama = self.kama(closes, cfg.kama_period)
        rsi = self.rsi(closes, cfg.rsi_period)
        bb_mid = closes[-cfg.bollinger_period:].mean()
        bb_std = closes[-cfg.bollinger_period:].std()
        bb_upper = bb_mid + cfg.bollinger_std * bb_std
        bb_lower = bb_mid - cfg.bollinger_std * bb_std

        vote = 0.0
        # Tendência: preço acima/abaixo da média adaptativa
        vote += 0.4 if price > kama else -0.4
        # Momentum: RSI (evita comprar sobrecomprado / vender sobrevendido)
        if rsi < 30:
            vote += 0.3
        elif rsi > 70:
            vote -= 0.3
        # Reversão à média nas bandas
        if price <= bb_lower:
            vote += 0.3
        elif price >= bb_upper:
            vote -= 0.3
        return float(np.clip(vote, -1.0, 1.0))

    @staticmethod
    def kama(closes: np.ndarray, period: int, fast: int = 2, slow: int = 30) -> float:
        """Média Móvel Adaptativa de Kaufman: acelera em tendência, trava no ruído."""
        change = abs(closes[-1] - closes[-period - 1])
        volatility = np.abs(np.diff(closes[-period - 1:])).sum() or 1e-12
        er = change / volatility  # efficiency ratio
        sc = (er * (2 / (fast + 1) - 2 / (slow + 1)) + 2 / (slow + 1)) ** 2
        kama = closes[-period]
        for price in closes[-period + 1:]:
            kama += sc * (price - kama)
        return float(kama)

    @staticmethod
    def rsi(closes: np.ndarray, period: int) -> float:
        deltas = np.diff(closes[-(period + 1):])
        gains = deltas[deltas > 0].sum() / period
        losses = -deltas[deltas < 0].sum() / period
        if losses == 0:
            return 100.0
        rs = gains / losses
        return float(100 - 100 / (1 + rs))

    @staticmethod
    def atr(highs: np.ndarray, lows: np.ndarray, closes: np.ndarray, period: int) -> float:
        tr = np.maximum(highs[1:] - lows[1:],
                        np.maximum(abs(highs[1:] - closes[:-1]),
                                   abs(lows[1:] - closes[:-1])))
        return float(tr[-period:].mean())

    @staticmethod
    def pivot_levels(highs: np.ndarray, lows: np.ndarray,
                     price: float, lookback: int = 100) -> tuple[float, float]:
        """Suporte = maior fundo de pivô abaixo do preço; resistência = menor topo acima."""
        pivots_high, pivots_low = [], []
        h, l = highs[-lookback:], lows[-lookback:]
        for i in range(2, len(h) - 2):
            if h[i] == max(h[i - 2:i + 3]):
                pivots_high.append(h[i])
            if l[i] == min(l[i - 2:i + 3]):
                pivots_low.append(l[i])
        support = max((p for p in pivots_low if p < price), default=price * 0.97)
        resistance = min((p for p in pivots_high if p > price), default=price * 1.03)
        return float(support), float(resistance)

    async def orderbook_imbalance(self, symbol: str, depth: int = 20) -> float:
        """(-1..+1) Pressão líquida no livro: >0 = compradores dominam."""
        try:
            book = await self.md.fetch_order_book(symbol, depth)
            bid_vol = sum(b[1] for b in book["bids"][:depth])
            ask_vol = sum(a[1] for a in book["asks"][:depth])
            total = bid_vol + ask_vol
            return (bid_vol - ask_vol) / total if total else 0.0
        except Exception:
            return 0.0
