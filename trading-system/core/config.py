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

try:
    from dotenv import load_dotenv
    load_dotenv()  # carrega o arquivo .env da pasta atual para o ambiente
except ImportError:
    pass  # sem python-dotenv, exige variáveis já exportadas no shell


@dataclass
class RiskConfig:
    max_risk_per_trade_pct: float = 2.0      # % do capital arriscado por trade (1-2%)
    max_position_pct: float = 15.0           # exposição máxima por posição
    max_open_positions: int = 4
    max_daily_drawdown_pct: float = 3.0      # circuito diário
    max_weekly_drawdown_pct: float = 8.0     # circuito semanal
    max_consecutive_losses: int = 4
    max_trades_per_day: int = 10
    default_trailing_stop_pct: float = 1.5
    min_risk_reward_ratio: float = 1.5
    # Envelope diário em DÓLARES (0 = desligado):
    max_daily_loss_usd: float = 0.0    # perdeu isso no dia -> para até amanhã
    # Trava de lucro do dia (trailing no resultado diário): quando o lucro do
    # dia atinge 'trigger', ativa um PISO 'gap' abaixo do pico; o piso sobe
    # junto com o lucro. Se o resultado do dia recuar até o piso, fecha e para
    # (protegendo o ganho). Continua operando enquanto o lucro subir.
    daily_profit_lock_trigger_usd: float = 0.0
    daily_profit_lock_gap_usd: float = 0.0


@dataclass
class QuantConfig:
    # "pullback" = recuo na tendência (vencedora do backtest 2024-2026);
    # "votes" = confluência multi-indicador original
    mode: str = "pullback"
    side: str = "long"                # "long" | "short" | "both" (both/short exige futures)
    trend_sma_days: int = 50          # pullback: só compra acima da SMA(n) diária
    pullback_rsi: float = 45.0        # pullback: RSI 1h máximo p/ considerar recuo
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
    symbols: list[str] = field(
        default_factory=lambda: ["BTC/USDT", "ETH/USDT", "BNB/USDT", "SOL/USDT"])
    initial_equity: float = 10_000.0
    testnet: bool = True                     # SEMPRE começar em testnet
    market_type: str = "spot"                # "spot" | "futures"
    leverage: int = 1                        # só usado em futures (2-3 recomendado)
    # Capital de operação simulado (0 = usa o saldo real da conta). Ex.: na
    # testnet o saldo é $5000, mas com trading_capital_usd=2000 o bot
    # dimensiona as ordens como se a conta tivesse só $2000 — espelha o real.
    trading_capital_usd: float = 0.0
    risk: RiskConfig = field(default_factory=RiskConfig)
    quant: QuantConfig = field(default_factory=QuantConfig)
    sentiment: SentimentConfig = field(default_factory=SentimentConfig)
    execution: ExecutionConfig = field(default_factory=ExecutionConfig)
    # Segredos — somente via ambiente:
    binance_api_key: str = ""
    binance_api_secret: str = ""
    n8n_webhook_url: str = ""
    telegram_bot_token: str = ""
    telegram_chat_id: str = ""

    @classmethod
    def load(cls, path: str | Path = "config/config.yaml") -> "AppConfig":
        cfg = cls()
        # Carrega o arquivo de ambiente certo diretamente: .env.futures para
        # a config de futuros, .env para o spot. Robusto e independente de
        # como o processo foi iniciado (não depende do .bat injetar variáveis).
        try:
            from dotenv import load_dotenv
            env_file = ".env.futures" if "futures" in str(path) else ".env"
            if Path(env_file).exists():
                load_dotenv(env_file, override=True)
        except ImportError:
            pass
        p = Path(path)
        if p.exists():
            raw = yaml.safe_load(p.read_text()) or {}
            cfg.symbols = raw.get("symbols", cfg.symbols)
            cfg.initial_equity = raw.get("initial_equity", cfg.initial_equity)
            cfg.testnet = raw.get("testnet", cfg.testnet)
            cfg.market_type = raw.get("market_type", cfg.market_type)
            cfg.leverage = raw.get("leverage", cfg.leverage)
            cfg.trading_capital_usd = raw.get("trading_capital_usd", cfg.trading_capital_usd)
            for section, target in (("risk", cfg.risk), ("quant", cfg.quant),
                                    ("sentiment", cfg.sentiment), ("execution", cfg.execution)):
                for k, v in (raw.get(section) or {}).items():
                    if hasattr(target, k):
                        setattr(target, k, v)
        cfg.binance_api_key = os.environ.get("BINANCE_API_KEY", "")
        cfg.binance_api_secret = os.environ.get("BINANCE_API_SECRET", "")
        cfg.n8n_webhook_url = os.environ.get("N8N_WEBHOOK_URL", "")
        cfg.telegram_bot_token = os.environ.get("TELEGRAM_BOT_TOKEN", "")
        cfg.telegram_chat_id = os.environ.get("TELEGRAM_CHAT_ID", "")
        return cfg
