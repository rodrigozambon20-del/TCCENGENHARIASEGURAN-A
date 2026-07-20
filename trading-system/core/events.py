"""
Eventos tipados que trafegam no Event Bus.

Cada agente se comunica EXCLUSIVAMENTE por meio destes eventos — nenhum
agente chama outro diretamente. Isso garante desacoplamento total:
qualquer agente pode ser substituído, escalado ou movido para outro
processo (Redis/RabbitMQ) sem alterar os demais.
"""
from __future__ import annotations

import time
import uuid
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class Topic(str, Enum):
    """Tópicos do barramento de mensagens (pub/sub)."""
    SENTIMENT_SIGNAL = "sentiment.signal"        # Sentimento -> Orquestrador
    QUANT_SIGNAL = "quant.signal"                # Quant -> Orquestrador
    TRADE_PROPOSAL = "trade.proposal"            # Orquestrador -> Risco
    RISK_VERDICT = "risk.verdict"                # Risco -> Orquestrador
    EXECUTION_ORDER = "execution.order"          # Orquestrador -> Execução
    EXECUTION_REPORT = "execution.report"        # Execução -> todos
    KILL_SWITCH = "system.kill_switch"           # Qualquer agente -> todos (prioridade máxima)
    DEFENSIVE_MODE = "system.defensive_mode"     # Sentimento/Risco -> todos
    HEARTBEAT = "system.heartbeat"               # Saúde dos agentes
    NOTIFICATION = "system.notification"         # -> Notifier (n8n/Telegram)


class Direction(str, Enum):
    LONG = "long"
    SHORT = "short"
    FLAT = "flat"


class Sentiment(str, Enum):
    BULLISH = "bullish"
    BEARISH = "bearish"
    NEUTRAL = "neutral"


@dataclass
class BaseEvent:
    event_id: str = field(default_factory=lambda: uuid.uuid4().hex, init=False)
    ts: float = field(default_factory=time.time, init=False)
    source: str = "unknown"


@dataclass
class SentimentSignal(BaseEvent):
    """Emitido pelo Agente de Sentimento."""
    symbol: str = "BTC/USDT"
    sentiment: Sentiment = Sentiment.NEUTRAL
    score: float = 0.0            # -1.0 (bearish extremo) .. +1.0 (bullish extremo)
    confidence: float = 0.0       # 0..1, ponderado pela relevância das fontes
    headlines: list[str] = field(default_factory=list)
    is_catastrophic: bool = False # gatilho de resposta rápida


@dataclass
class QuantSignal(BaseEvent):
    """Emitido pelo Agente Quant após varredura multi-timeframe."""
    symbol: str = "BTC/USDT"
    direction: Direction = Direction.FLAT
    strategy: str = "votes"       # "votes" | "pullback" (define stops no orquestrador)
    score: float = 0.0            # -1..+1 confluência técnica
    timeframe_votes: dict[str, float] = field(default_factory=dict)  # {"1m": 0.3, "1h": 0.8, ...}
    support: Optional[float] = None
    resistance: Optional[float] = None
    atr: Optional[float] = None   # volatilidade p/ dimensionar stops
    last_price: Optional[float] = None
    orderflow_imbalance: float = 0.0  # -1..+1 (pressão vendedora/compradora no book)


@dataclass
class TradeProposal(BaseEvent):
    """Proposta do Orquestrador. NUNCA vai direto para a exchange."""
    symbol: str = "BTC/USDT"
    direction: Direction = Direction.LONG
    entry_price: float = 0.0
    stop_loss: float = 0.0
    take_profit: float = 0.0
    conviction: float = 0.0       # 0..1 fusão sentimento+quant
    rationale: str = ""


@dataclass
class RiskVerdict(BaseEvent):
    """Veredito do Risk Manager — palavra final antes da execução."""
    proposal_id: str = ""
    approved: bool = False
    reason: str = ""
    position_size: float = 0.0    # quantidade na moeda base, já dimensionada
    stop_loss: float = 0.0        # stops do Risco SOBRESCREVEM os da proposta
    take_profit: float = 0.0
    trailing_stop_pct: Optional[float] = None


@dataclass
class ExecutionOrder(BaseEvent):
    """Comando final aprovado, enviado ao Execution Bot."""
    proposal: Optional[TradeProposal] = None
    verdict: Optional[RiskVerdict] = None


@dataclass
class ExecutionReport(BaseEvent):
    """Resultado real da execução na exchange."""
    symbol: str = ""
    order_id: str = ""
    status: str = ""              # filled | partial | canceled | rejected | error
    side: str = ""
    requested_price: float = 0.0
    fill_price: float = 0.0
    filled_qty: float = 0.0
    slippage_bps: float = 0.0
    fees: float = 0.0
    error: Optional[str] = None


@dataclass
class KillSwitch(BaseEvent):
    """Para TUDO. Emitido em drawdown máximo, notícia catastrófica ou falha grave."""
    reason: str = ""
    close_positions: bool = False


@dataclass
class DefensiveMode(BaseEvent):
    """Bloqueia NOVAS entradas, mantém gestão das posições abertas."""
    enabled: bool = True
    reason: str = ""
    cooldown_seconds: int = 3600


@dataclass
class Notification(BaseEvent):
    """Mensagem para o pipeline n8n -> Telegram/WhatsApp."""
    level: str = "info"           # info | warning | critical
    title: str = ""
    body: str = ""
    payload: dict = field(default_factory=dict)
