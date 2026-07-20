"""
AGENTE 1 — ANÁLISE MACRO E SENTIMENTO (News & Social Sentiment)

Responsabilidades:
1. Coletar manchetes de portais financeiros (RSS/APIs: CoinTelegraph,
   CoinDesk, agregadores como CryptoPanic; Bloomberg/Reuters via API paga)
   e redes sociais (X/Twitter API v2, Reddit via PRAW).
2. Classificar sentimento com NLP. Duas opções de motor, em cascata:
   a) Modelo local rápido (VADER/FinBERT) — baixa latência, custo zero;
   b) LLM (Claude) para manchetes ambíguas ou de alto impacto — melhor
      compreensão de contexto macro ("Fed segura juros" != "hack de $500M").
3. Ponderar o score pela credibilidade da fonte (Bloomberg pesa mais que
   um tweet) e pelo frescor da notícia (decaimento exponencial).
4. GATILHO DE RESPOSTA RÁPIDA: notícia catastrófica (hack de exchange,
   ação regulatória, colapso de stablecoin) publica DEFENSIVE_MODE
   imediatamente, sem esperar o ciclo do orquestrador.
"""
from __future__ import annotations

import asyncio
import math
import time
from dataclasses import dataclass

from agents.base import BaseAgent
from core.events import DefensiveMode, Sentiment, SentimentSignal, Topic

# Palavras-gatilho de evento catastrófico (verificadas ANTES do NLP,
# para latência mínima na resposta defensiva)
CATASTROPHIC_KEYWORDS = (
    "hack", "hacked", "exploit", "bankruptcy", "insolvency", "insolvent",
    "sec lawsuit", "banned", "depeg", "halted withdrawals", "rug pull",
    "flash crash", "liquidation cascade",
)


@dataclass
class NewsItem:
    source: str          # "bloomberg", "twitter", ...
    headline: str
    published_at: float


class SentimentAgent(BaseAgent):
    name = "sentiment_agent"

    async def run(self) -> None:
        cfg = self.config.sentiment
        while True:
            items = await self.fetch_news()
            if items:
                await self.process_batch(items)
            await asyncio.sleep(cfg.poll_interval_seconds)

    # ------------------------------------------------------------------ #
    # Coleta                                                             #
    # ------------------------------------------------------------------ #
    async def fetch_news(self) -> list[NewsItem]:
        """
        TEMPLATE: integre aqui seus conectores reais.

        - RSS gratuito:  feedparser em https://cointelegraph.com/rss
        - CryptoPanic:   GET https://cryptopanic.com/api/v1/posts/?auth_token=...
        - X/Twitter v2:  filtered stream com regras ("BTC", "Binance", "SEC")
        - Reddit:        PRAW em r/CryptoCurrency (hot + new)

        Todos os conectores devem rodar com timeout e falhar de forma
        isolada (uma fonte fora do ar não pode travar o agente).
        """
        results: list[NewsItem] = []
        # for connector in self.connectors:
        #     try:
        #         results += await asyncio.wait_for(connector.fetch(), timeout=10)
        #     except Exception as exc:
        #         self.log.warning("Fonte %s indisponível: %s", connector.name, exc)
        return results

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
            if any(k in lowered for k in CATASTROPHIC_KEYWORDS):
                await self.trigger_defensive(item)

            # 2) Score NLP (-1..+1)
            raw_score = await self.classify(item.headline)

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
        self.log.info("Sentimento: %s score=%.2f conf=%.2f (%d manchetes)",
                      sentiment.value, score, confidence, len(items))

    async def classify(self, headline: str) -> float:
        """
        TEMPLATE do motor NLP em cascata.

        Camada 1 (sempre): modelo local
            from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
            return SentimentIntensityAnalyzer().polarity_scores(headline)["compound"]

        Camada 2 (só p/ manchetes de alto impacto, |score| ambíguo):
            resposta = await anthropic_client.messages.create(
                model="claude-sonnet-5",
                max_tokens=10,
                messages=[{"role": "user", "content":
                    f"Classifique o impacto desta manchete no preço do Bitcoin "
                    f"nas próximas horas. Responda só um número entre -1 e 1: {headline}"}])
            return float(resposta.content[0].text)
        """
        return 0.0

    async def trigger_defensive(self, item: NewsItem) -> None:
        """Resposta rápida: <1s entre detecção e bloqueio de novas entradas."""
        self.log.critical("EVENTO CATASTRÓFICO DETECTADO: %s", item.headline)
        await self.emit(Topic.DEFENSIVE_MODE, DefensiveMode(
            enabled=True,
            reason=f"Notícia catastrófica [{item.source}]: {item.headline}",
            cooldown_seconds=3600,
        ))
