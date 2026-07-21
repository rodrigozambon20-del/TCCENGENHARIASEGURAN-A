"""
AGENTE 1 — ANÁLISE MACRO E SENTIMENTO (News & Social Sentiment)

Fontes: 100% GRATUITAS e internacionais (ver datasources/news_connectors.py):
RSS de CoinTelegraph/CoinDesk/Decrypt/Bitcoin Magazine, Google News,
Reddit (JSON público) e Fear & Greed Index. CryptoPanic opcional (token grátis).

Motor NLP: VADER com léxico estendido para o vocabulário cripto/financeiro
("hack", "ETF approval", "depeg"...). Roda 100% local, custo zero e latência
de milissegundos. Fontes que já entregam score numérico (Fear & Greed)
pulam o NLP.

Ponderação: score de cada manchete x credibilidade da fonte x frescor
(meia-vida de 30 min). GATILHO RÁPIDO: manchete com palavra catastrófica
publica DEFENSIVE_MODE imediatamente, antes de qualquer agregação.
"""
from __future__ import annotations

import asyncio
import math
import time

from agents.base import BaseAgent
from core.events import DefensiveMode, Sentiment, SentimentSignal, Topic
from datasources.news_connectors import NewsAggregator, NewsItem

# Palavras-gatilho de evento catastrófico (checadas ANTES do NLP,
# para latência mínima na resposta defensiva)
CATASTROPHIC_KEYWORDS = (
    "hack", "hacked", "exploit", "bankruptcy", "insolvency", "insolvent",
    "sec lawsuit", "banned", "depeg", "halted withdrawals", "rug pull",
    "flash crash", "liquidation cascade",
)

# Termos que provam que a notícia é sobre o MERCADO/infraestrutura cripto —
# e não sobre um assunto qualquer que só cita bitcoin de passagem (ex.:
# "resgate em bitcoin" num caso policial). Um evento só é tratado como
# catastrófico se combinar uma palavra de perigo COM um destes termos.
CRYPTO_MARKET_TERMS = (
    "exchange", "binance", "coinbase", "kraken", "okx", "bybit", "bitfinex",
    "stablecoin", "tether", "usdt", "usdc", "etf", "sec", "cftc",
    "defi", "protocol", "bridge", "custodian", "custody", "wallet provider",
    "crypto firm", "crypto exchange", "trading platform", "token", "blockchain",
)

# Léxico adicional p/ o VADER: vocabulário do mercado cripto que o léxico
# padrão (redes sociais genéricas) não conhece. Escala VADER: -4..+4.
CRYPTO_LEXICON = {
    "hack": -3.5, "hacked": -3.5, "exploit": -3.0, "stolen": -3.0,
    "bankruptcy": -3.5, "insolvent": -3.5, "lawsuit": -2.0, "sued": -2.0,
    "ban": -2.5, "banned": -2.5, "crackdown": -2.0, "depeg": -3.0,
    "liquidation": -2.0, "liquidations": -2.0, "selloff": -2.0,
    "sell-off": -2.0, "plunge": -2.5, "plunges": -2.5, "crash": -3.0,
    "dump": -2.0, "fud": -1.5, "bearish": -2.0, "correction": -1.0,
    "rally": 2.0, "rallies": 2.0, "surge": 2.5, "surges": 2.5,
    "soar": 2.5, "soars": 2.5, "breakout": 2.0, "bullish": 2.0,
    "adoption": 1.5, "approval": 2.0, "approved": 2.0, "etf": 0.5,
    "halving": 1.0, "institutional": 1.0, "accumulation": 1.5,
    "all-time high": 3.0, "ath": 2.0, "moon": 1.5, "pump": 1.5,
    "upgrade": 1.0, "partnership": 1.0,
}


def _build_vader():
    """VADER com léxico cripto; fallback simples se a lib não estiver instalada."""
    try:
        from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
        analyzer = SentimentIntensityAnalyzer()
        analyzer.lexicon.update(CRYPTO_LEXICON)
        return analyzer
    except ImportError:
        return None


class SentimentAgent(BaseAgent):
    name = "sentiment_agent"

    def __init__(self, bus, state, config) -> None:
        super().__init__(bus, state, config)
        self.aggregator = NewsAggregator()
        self.vader = _build_vader()
        if self.vader is None:
            self.log.warning("vaderSentiment não instalado — usando léxico simples")

    async def run(self) -> None:
        cfg = self.config.sentiment
        while True:
            items = await self.aggregator.fetch_all()
            if items:
                await self.process_batch(items)
            await asyncio.sleep(cfg.poll_interval_seconds)

    # ------------------------------------------------------------------ #
    # NLP e agregação                                                    #
    # ------------------------------------------------------------------ #
    async def process_batch(self, items: list[NewsItem]) -> None:
        cfg = self.config.sentiment
        now = time.time()
        weighted_sum = 0.0
        weight_total = 0.0
        headlines: list[str] = []

        for item in items:
            # 1) Fast-path catastrófico: keyword match antes de qualquer NLP
            lowered = item.headline.lower()
            # Só é catastrófico se houver perigo E relevância de mercado cripto.
            # Evita alarme falso com notícias que apenas citam "bitcoin"/"hack"
            # fora do contexto do mercado (ex.: crime político, ransomware geral).
            has_danger = any(k in lowered for k in CATASTROPHIC_KEYWORDS)
            is_market_relevant = any(t in lowered for t in CRYPTO_MARKET_TERMS)
            if has_danger and is_market_relevant:
                await self.trigger_defensive(item)

            # 2) Score: pronto (Fear & Greed) ou via NLP local
            raw_score = (item.precomputed_score
                         if item.precomputed_score is not None
                         else self.classify(item.headline))

            # 3) Peso = credibilidade da fonte x decaimento temporal (meia-vida 30 min)
            source_w = cfg.source_weights.get(item.source, 0.2)
            age_min = max(0.0, (now - item.published_at) / 60)
            freshness = math.exp(-age_min / 30 * math.log(2))
            w = source_w * freshness

            weighted_sum += raw_score * w
            weight_total += w
            headlines.append(f"[{item.source}] {item.headline}")

        if weight_total == 0:
            return

        score = weighted_sum / weight_total
        confidence = min(1.0, weight_total / 3.0)  # mais fontes concordando => mais confiança
        sentiment = (Sentiment.BULLISH if score > 0.2
                     else Sentiment.BEARISH if score < -0.2
                     else Sentiment.NEUTRAL)

        signal = SentimentSignal(
            sentiment=sentiment, score=score, confidence=confidence,
            headlines=headlines[:10],
            is_catastrophic=(score <= cfg.catastrophic_score_threshold
                             and confidence >= cfg.catastrophic_confidence_threshold),
        )
        if signal.is_catastrophic:
            await self.emit(Topic.DEFENSIVE_MODE, DefensiveMode(
                enabled=True, reason=f"Sentimento catastrófico agregado ({score:.2f})"))
        await self.emit(Topic.SENTIMENT_SIGNAL, signal)
        self.log.info("Sentimento: %s score=%.2f conf=%.2f (%d manchetes novas)",
                      sentiment.value, score, confidence, len(items))

    def classify(self, headline: str) -> float:
        """Score -1..+1 da manchete. VADER local (rápido e gratuito)."""
        if self.vader is not None:
            return float(self.vader.polarity_scores(headline)["compound"])
        # Fallback sem dependências: soma do léxico cripto normalizada
        lowered = headline.lower()
        raw = sum(v for k, v in CRYPTO_LEXICON.items() if k in lowered)
        return max(-1.0, min(1.0, raw / 4.0))

    async def trigger_defensive(self, item: NewsItem) -> None:
        """Resposta rápida: <1s entre detecção e bloqueio de novas entradas."""
        self.log.critical("EVENTO CATASTRÓFICO DETECTADO: %s", item.headline)
        await self.emit(Topic.DEFENSIVE_MODE, DefensiveMode(
            enabled=True,
            reason=f"Notícia catastrófica [{item.source}]: {item.headline}",
            cooldown_seconds=3600,
        ))
