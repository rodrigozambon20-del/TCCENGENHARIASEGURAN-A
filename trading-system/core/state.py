"""
Estado global compartilhado do bot (single source of truth).

Somente o Risk Manager escreve nos campos de risco; os demais agentes
apenas leem. O acesso é protegido por lock para evitar condições de corrida
entre corrotinas.
"""
from __future__ import annotations

import asyncio
import time
from dataclasses import dataclass, field


@dataclass
class Position:
    symbol: str
    side: str                 # long | short
    qty: float
    entry_price: float
    stop_loss: float
    take_profit: float
    trailing_stop_pct: float | None = None
    highest_price: float = 0.0    # p/ trailing em long
    lowest_price: float = 1e18    # p/ trailing em short
    opened_at: float = field(default_factory=time.time)


class BotState:
    def __init__(self, initial_equity: float) -> None:
        self._lock = asyncio.Lock()
        self.initial_equity = initial_equity
        self.equity = initial_equity
        self.equity_high_watermark = initial_equity
        self.day_start_equity = initial_equity
        self.week_start_equity = initial_equity
        self.open_positions: dict[str, Position] = {}
        self.trading_halted = False          # kill switch ativado
        self.defensive_mode_until: float = 0.0
        self.consecutive_losses = 0
        self.trades_today = 0

    @property
    def daily_drawdown_pct(self) -> float:
        if self.day_start_equity <= 0:
            return 0.0
        return max(0.0, (self.day_start_equity - self.equity) / self.day_start_equity * 100)

    @property
    def weekly_drawdown_pct(self) -> float:
        if self.week_start_equity <= 0:
            return 0.0
        return max(0.0, (self.week_start_equity - self.equity) / self.week_start_equity * 100)

    @property
    def max_drawdown_from_peak_pct(self) -> float:
        if self.equity_high_watermark <= 0:
            return 0.0
        return max(0.0, (self.equity_high_watermark - self.equity) / self.equity_high_watermark * 100)

    @property
    def in_defensive_mode(self) -> bool:
        return time.time() < self.defensive_mode_until

    async def update_equity(self, new_equity: float) -> None:
        async with self._lock:
            self.equity = new_equity
            self.equity_high_watermark = max(self.equity_high_watermark, new_equity)

    async def roll_daily(self) -> None:
        async with self._lock:
            self.day_start_equity = self.equity
            self.trades_today = 0

    async def roll_weekly(self) -> None:
        async with self._lock:
            self.week_start_equity = self.equity
