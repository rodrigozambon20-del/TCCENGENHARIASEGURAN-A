"""
Carregamento e validação de configuração.

Segredos (API keys) vêm SOMENTE de variáveis de ambiente (.env) — nunca
do YAML e nunca commitados no repositório. O YAML carrega apenas
parâmetros de estratégia e risco.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

import yaml


@dataclass
class RiskConfig:
    max_risk_per_trade_pct: float = 1.0      # % do capital arriscado por trade (1-2%)
    max_position_pct: float = 10.0           # exposição máxima por posição
    max_open_positions: int = 3
    max_daily_drawdown_pct: float = 3.0      # circuito diário
    max_weekly_drawdown_pct: float = 8.0     # circuito semanal
    max_consecutive_losses: int = 4
    max_trades_per_day: int = 10
    default_trailing_stop_pct: float = 1.5
    min_risk_reward_ratio: float = 1.5


@dataclass
class QuantConfig:
    timeframes: list[str] = field(default_factory=lambda: ["1m", "5m", "15m", "1h", "1d"])
    timeframe_weights: dict[str, float] = field(
        default_factory=lambda: {"1m": 0.10, "5m": 0.15, "15m": 0.20, "1h": 0.30, "1d": 0.25})
    rsi_period: int = 14
    bollinger_period: int = 20
    bollinger_std: float = 2.0
    kama_period: int = 10                    # média móvel adaptativa (Kaufman)
    atr_period: int = 14
    min_score_to_signal: float = 0.35


@dataclass
class SentimentConfig:
    poll_interval_seconds: int = 60
    source_weights: dict[str, float] = field(
        default_factory=lambda: {
            "coindesk": 0.9, "cointelegraph": 0.8, "decrypt": 0.7,
            "bitcoinmagazine": 0.6, "google_news": 0.6, "cryptopanic": 0.6,
            "fear_greed": 0.8, "reddit": 0.3,
        })
    catastrophic_score_threshold: float = -0.75
    catastrophic_confidence_threshold: float = 0.6


@dataclass
class ExecutionConfig:
    prefer_limit_orders: bool = True
    limit_offset_bps: float = 2.0            # oferta 2 bps dentro do spread (maker)
    limit_timeout_seconds: int = 20          # após isso, cancela e reavalia
    max_slippage_bps: float = 15.0           # aborta market order acima disso
    order_poll_interval: float = 1.0


@dataclass
class AppConfig:
    symbols: list[str] = field(default_factory=lambda: ["BTC/USDT", "ETH/USDT"])
    initial_equity: float = 10_000.0
    testnet: bool = True                     # SEMPRE começar em testnet
    risk: RiskConfig = field(default_factory=RiskConfig)
    quant: QuantConfig = field(default_factory=QuantConfig)
    sentiment: SentimentConfig = field(default_factory=SentimentConfig)
    execution: ExecutionConfig = field(default_factory=ExecutionConfig)
    # Segredos — somente via ambiente:
    binance_api_key: str = ""
    binance_api_secret: str = ""
    n8n_webhook_url: str = ""

    @classmethod
    def load(cls, path: str | Path = "config/config.yaml") -> "AppConfig":
        cfg = cls()
        p = Path(path)
        if p.exists():
            raw = yaml.safe_load(p.read_text()) or {}
            cfg.symbols = raw.get("symbols", cfg.symbols)
            cfg.initial_equity = raw.get("initial_equity", cfg.initial_equity)
            cfg.testnet = raw.get("testnet", cfg.testnet)
            for section, target in (("risk", cfg.risk), ("quant", cfg.quant),
                                    ("sentiment", cfg.sentiment), ("execution", cfg.execution)):
                for k, v in (raw.get(section) or {}).items():
                    if hasattr(target, k):
                        setattr(target, k, v)
        cfg.binance_api_key = os.environ.get("BINANCE_API_KEY", "")
        cfg.binance_api_secret = os.environ.get("BINANCE_API_SECRET", "")
        cfg.n8n_webhook_url = os.environ.get("N8N_WEBHOOK_URL", "")
        return cfg
