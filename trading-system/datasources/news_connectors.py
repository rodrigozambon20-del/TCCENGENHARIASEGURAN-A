"""
Conectores de notícias e sentimento — SOMENTE FONTES GRATUITAS.

Fontes internacionais sem custo e sem chave de API:
- CoinTelegraph, CoinDesk, Decrypt, Bitcoin Magazine  -> RSS público
- Google News (busca "bitcoin OR crypto", últimas 24h) -> RSS público
- Reddit r/CryptoCurrency e r/Bitcoin                  -> endpoint JSON público
- Fear & Greed Index (alternative.me)                  -> API pública

Opcional (grátis com cadastro, só ativa se CRYPTOPANIC_TOKEN existir):
- CryptoPanic (agregador de manchetes com voto da comunidade)

Regras de resiliência:
- Cada fonte roda isolada com timeout: uma fonte fora do ar não derruba
  a coleta das demais.
- Deduplicação por hash da manchete (janela de 24h): a mesma notícia
  replicada em vários portais não vira sinal duplicado.
"""
from __future__ import annotations

import asyncio
import calendar
import hashlib
import logging
import os
import time
from dataclasses import dataclass
from typing import Optional

import aiohttp
import feedparser

log = logging.getLogger("news")

RSS_FEEDS = {
    "cointelegraph": "https://cointelegraph.com/rss",
    "coindesk": "https://www.coindesk.com/arc/outboundfeeds/rss/",
    "decrypt": "https://decrypt.co/feed",
    # bitcoinmagazine removido: certificado SSL expirado (erro recorrente).
    "google_news": ("https://news.google.com/rss/search"
                    "?q=bitcoin+OR+cryptocurrency+when:1d&hl=en-US&gl=US&ceid=US:en"),
}
# Reddit desativado: o endpoint público passou a bloquear (HTTP 403) e
# poluía o log. As fontes RSS + Fear & Greed já dão sinal suficiente.
REDDIT_SUBS = ()
REDDIT_URL = "https://www.reddit.com/r/{sub}/hot.json?limit=25"
FEAR_GREED_URL = "https://api.alternative.me/fng/?limit=1"
CRYPTOPANIC_URL = ("https://cryptopanic.com/api/v1/posts/"
                   "?auth_token={token}&public=true&kind=news")

TIMEOUT = aiohttp.ClientTimeout(total=12)
HEADERS = {"User-Agent": "Mozilla/5.0 (research; crypto-sentiment-bot/1.0)"}
DEDUP_WINDOW = 24 * 3600
MAX_ITEM_AGE = 6 * 3600      # ignora notícia com mais de 6h (não é mais "notícia")
MIN_REDDIT_UPVOTES = 50      # filtra ruído de posts irrelevantes


@dataclass
class NewsItem:
    source: str                       # chave usada nos pesos do config
    headline: str
    published_at: float
    url: str = ""
    precomputed_score: Optional[float] = None  # fontes que já entregam score (-1..1)


class NewsAggregator:
    """Coleta paralela de todas as fontes gratuitas, com dedup e isolamento."""

    def __init__(self) -> None:
        self._seen: dict[str, float] = {}   # hash da manchete -> ts da 1ª vez
        self._cryptopanic_token = os.environ.get("CRYPTOPANIC_TOKEN", "")

    async def fetch_all(self) -> list[NewsItem]:
        tasks: list[asyncio.Task] = []
        async with aiohttp.ClientSession(timeout=TIMEOUT, headers=HEADERS) as session:
            for source, url in RSS_FEEDS.items():
                tasks.append(asyncio.create_task(self._fetch_rss(session, source, url)))
            for sub in REDDIT_SUBS:
                tasks.append(asyncio.create_task(self._fetch_reddit(session, sub)))
            tasks.append(asyncio.create_task(self._fetch_fear_greed(session)))
            if self._cryptopanic_token:
                tasks.append(asyncio.create_task(self._fetch_cryptopanic(session)))
            results = await asyncio.gather(*tasks, return_exceptions=True)

        items: list[NewsItem] = []
        for result in results:
            if isinstance(result, Exception):
                # Nível debug: fonte fora do ar é rotina, não erro (as demais
                # cobrem). Não polui a tela com "erros" que não são erros.
                log.debug("Fonte indisponível (ignorada neste ciclo): %s", result)
            else:
                items.extend(result)
        return self._dedup(items)

    # ------------------------------------------------------------------ #
    # RSS (CoinTelegraph, CoinDesk, Decrypt, Bitcoin Magazine, G.News)   #
    # ------------------------------------------------------------------ #
    async def _fetch_rss(self, session: aiohttp.ClientSession,
                         source: str, url: str) -> list[NewsItem]:
        async with session.get(url) as resp:
            resp.raise_for_status()
            raw = await resp.read()
        # feedparser é síncrono -> roda em thread p/ não travar o event loop
        feed = await asyncio.to_thread(feedparser.parse, raw)
        now = time.time()
        items = []
        for entry in feed.entries[:30]:
            parsed = entry.get("published_parsed") or entry.get("updated_parsed")
            ts = calendar.timegm(parsed) if parsed else now
            if now - ts > MAX_ITEM_AGE:
                continue
            title = (entry.get("title") or "").strip()
            if title:
                items.append(NewsItem(source=source, headline=title,
                                      published_at=ts, url=entry.get("link", "")))
        log.debug("%s: %d manchetes frescas", source, len(items))
        return items

    # ------------------------------------------------------------------ #
    # Reddit (JSON público, sem chave)                                   #
    # ------------------------------------------------------------------ #
    async def _fetch_reddit(self, session: aiohttp.ClientSession,
                            sub: str) -> list[NewsItem]:
        async with session.get(REDDIT_URL.format(sub=sub)) as resp:
            resp.raise_for_status()
            data = await resp.json()
        now = time.time()
        items = []
        for child in data.get("data", {}).get("children", []):
            post = child.get("data", {})
            if post.get("stickied") or post.get("ups", 0) < MIN_REDDIT_UPVOTES:
                continue
            ts = float(post.get("created_utc", now))
            if now - ts > MAX_ITEM_AGE:
                continue
            items.append(NewsItem(source="reddit",
                                  headline=(post.get("title") or "").strip(),
                                  published_at=ts,
                                  url="https://reddit.com" + post.get("permalink", "")))
        return items

    # ------------------------------------------------------------------ #
    # Fear & Greed Index (score pronto, sem NLP)                         #
    # ------------------------------------------------------------------ #
    async def _fetch_fear_greed(self, session: aiohttp.ClientSession) -> list[NewsItem]:
        async with session.get(FEAR_GREED_URL) as resp:
            resp.raise_for_status()
            data = await resp.json()
        entry = data["data"][0]
        value = int(entry["value"])                    # 0 = medo extremo, 100 = ganância
        score = (value - 50) / 50.0                    # -1..+1
        return [NewsItem(
            source="fear_greed",
            headline=f"Fear & Greed Index: {value} ({entry['value_classification']})",
            published_at=time.time(),
            precomputed_score=score,
        )]

    # ------------------------------------------------------------------ #
    # CryptoPanic (opcional — token gratuito)                            #
    # ------------------------------------------------------------------ #
    async def _fetch_cryptopanic(self, session: aiohttp.ClientSession) -> list[NewsItem]:
        url = CRYPTOPANIC_URL.format(token=self._cryptopanic_token)
        async with session.get(url) as resp:
            resp.raise_for_status()
            data = await resp.json()
        now = time.time()
        items = []
        for post in data.get("results", [])[:30]:
            title = (post.get("title") or "").strip()
            if title:
                items.append(NewsItem(source="cryptopanic", headline=title,
                                      published_at=now, url=post.get("url", "")))
        return items

    # ------------------------------------------------------------------ #
    # Deduplicação (24h)                                                 #
    # ------------------------------------------------------------------ #
    def _dedup(self, items: list[NewsItem]) -> list[NewsItem]:
        now = time.time()
        self._seen = {h: ts for h, ts in self._seen.items()
                      if now - ts < DEDUP_WINDOW}
        fresh = []
        for item in items:
            key = hashlib.sha1(item.headline.lower().encode()).hexdigest()
            if key in self._seen:
                continue
            self._seen[key] = now
            fresh.append(item)
        return fresh
