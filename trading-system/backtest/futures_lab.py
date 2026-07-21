"""
LABORATÓRIO DE FUTUROS — a mesma estratégia pullback, agora com alavancagem.

Diferente do backtest spot, aqui modelamos os TRÊS custos reais que a
alavancagem traz (e que quase todo "robô milagroso" esconde):

1. MULTIPLICAÇÃO SIMÉTRICA: alavancagem L multiplica o lucro E o prejuízo
   de cada operação por L. Não existe multiplicar só o ganho.

2. FUNDING RATE: futuros perpétuos cobram uma taxa a cada 8h de quem está
   posicionado (padrão ~0,01%/8h sobre o NOCIONAL, não sobre a margem).
   Segurar posição alavancada custa dinheiro o tempo todo.

3. LIQUIDAÇÃO: se o preço andar contra ~ (100/L - margem de manutenção)%,
   a corretora ZERA a posição e você perde toda a margem daquele trade —
   antes mesmo do seu stop. Ex.: 10x liquida com ~9,5% contra; 3x com ~33%.

Sizing: o risco por trade é multiplicado pela alavancagem (é isso que
"usar alavancagem para multiplicar" significa na prática). A margem
exigida permanece constante; o que muda é o tamanho da aposta.

Uso:
    python backtest/futures_lab.py --symbol BTC/USDT --tf 4h --days 730
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from strategy_backtrader import load_data  # noqa: E402

BAR_HOURS = {"1m": 1 / 60, "5m": 5 / 60, "15m": 0.25, "30m": 0.5,
             "1h": 1, "2h": 2, "4h": 4, "1d": 24}


def _rsi(close: np.ndarray, period: int = 14) -> np.ndarray:
    delta = np.diff(close, prepend=close[0])
    gain = np.where(delta > 0, delta, 0.0)
    loss = np.where(delta < 0, -delta, 0.0)
    avg_gain = pd.Series(gain).rolling(period).mean().values
    avg_loss = pd.Series(loss).rolling(period).mean().values
    rs = np.divide(avg_gain, avg_loss, out=np.full_like(avg_gain, np.inf),
                   where=avg_loss > 0)
    return 100 - 100 / (1 + rs)


def simulate(df: pd.DataFrame, *, leverage: float, side: str, tf: str,
             trail_pct: float = 8.0, trend_days: int = 50,
             pullback_rsi: float = 45.0, risk_pct: float = 1.0,
             funding_8h: float = 0.0001, taker_fee: float = 0.0004,
             maint_margin: float = 0.005) -> dict:
    close = df["close"].values.astype(float)
    high = df["high"].values.astype(float)
    low = df["low"].values.astype(float)
    n = len(close)

    bar_h = BAR_HOURS.get(tf, 1)
    trend_bars = int(trend_days * 24 / bar_h)
    sma = pd.Series(close).rolling(trend_bars).mean().values
    bb_mid = pd.Series(close).rolling(20).mean().values
    rsi = _rsi(close, 14)

    equity = 10_000.0
    peak = equity
    max_dd = 0.0
    pos = None
    trades = wins = liquidations = 0
    # adverse move fraction que dispara liquidação (isolada, aprox.)
    liq_frac = max(1e-9, 1.0 / leverage - maint_margin)

    for i in range(trend_bars + 1, n):
        price = close[i]

        if pos is not None:
            # --- custo de funding enquanto a posição está aberta ---
            equity -= pos["notional"] * funding_8h * (bar_h / 8.0)

            exit_price = None
            liquidated = False
            if pos["side"] == "long":
                pos["peak"] = max(pos["peak"], high[i])
                pos["trail"] = max(pos["trail"], pos["peak"] * (1 - trail_pct / 100))
                liq_price = pos["entry"] * (1 - liq_frac)
                if low[i] <= liq_price:              # liquidação antes do stop
                    exit_price, liquidated = liq_price, True
                elif low[i] <= pos["trail"]:         # stop móvel
                    exit_price = pos["trail"]
            else:  # short
                pos["trough"] = min(pos["trough"], low[i])
                pos["trail"] = min(pos["trail"], pos["trough"] * (1 + trail_pct / 100))
                liq_price = pos["entry"] * (1 + liq_frac)
                if high[i] >= liq_price:
                    exit_price, liquidated = liq_price, True
                elif high[i] >= pos["trail"]:
                    exit_price = pos["trail"]

            if exit_price is not None:
                move = (exit_price - pos["entry"]) / pos["entry"]
                if pos["side"] == "short":
                    move = -move
                pnl = pos["notional"] * move
                equity += pnl - pos["notional"] * taker_fee  # taxa de saída
                if liquidated:
                    liquidations += 1
                if pnl > 0:
                    wins += 1
                trades += 1
                pos = None

        elif not np.isnan(sma[i]) and not np.isnan(bb_mid[i]):
            uptrend = price >= sma[i]
            downtrend = price < sma[i]
            long_ok = (side in ("long", "both") and uptrend
                       and rsi[i] <= pullback_rsi and price <= bb_mid[i])
            short_ok = (side in ("short", "both") and downtrend
                        and rsi[i] >= (100 - pullback_rsi) and price >= bb_mid[i])

            if long_ok or short_ok:
                # risco efetivo = risco base x alavancagem (a "multiplicação")
                risk_budget = equity * (risk_pct * leverage) / 100
                notional = risk_budget / (trail_pct / 100)
                if notional / leverage <= equity:    # margem cabe na conta
                    equity -= notional * taker_fee    # taxa de entrada
                    pos = {"side": "long" if long_ok else "short",
                           "entry": price, "notional": notional,
                           "peak": price, "trough": price,
                           "trail": (price * (1 - trail_pct / 100) if long_ok
                                     else price * (1 + trail_pct / 100))}

        if equity <= 0:                               # conta zerada
            equity = 0.0
            peak = max(peak, equity)
            max_dd = 100.0
            break
        peak = max(peak, equity)
        max_dd = max(max_dd, (peak - equity) / peak * 100)

    return {
        "ret_pct": (equity / 10_000 - 1) * 100,
        "max_dd": max_dd,
        "trades": trades,
        "win_rate": (wins / trades * 100) if trades else 0.0,
        "liquidations": liquidations,
        "final": equity,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--symbol", default="BTC/USDT")
    ap.add_argument("--tf", default="4h")
    ap.add_argument("--days", type=int, default=730)
    ap.add_argument("--side", default="long", choices=["long", "short", "both"])
    args = ap.parse_args()

    df = load_data(args.symbol, args.tf, args.days)

    print()
    print("=" * 82)
    print(f"  FUTUROS — {args.symbol} {args.tf}, {args.days} dias | "
          f"lado: {args.side} | capital inicial 10.000")
    print("=" * 82)
    print(f"  {'Alavancagem':<14} {'Retorno':>10} {'MaxDD':>8} "
          f"{'Trades':>7} {'Acerto':>7} {'Liquidações':>12}")
    print("-" * 82)
    for lev in (1, 2, 3, 5, 10, 20):
        r = simulate(df, leverage=lev, side=args.side, tf=args.tf)
        flag = "  <-- CONTA ZERADA" if r["final"] <= 0 else ""
        print(f"  {str(lev) + 'x':<14} {r['ret_pct']:>+9.2f}% {r['max_dd']:>7.2f}% "
              f"{r['trades']:>7d} {r['win_rate']:>6.1f}% {r['liquidations']:>12d}{flag}")
    print("=" * 82)
    print("  Modelo inclui: funding 0,01%/8h, taxa taker 0,04%, liquidação isolada.")
    print("  Resultado passado NÃO garante futuro. Alavancagem pode zerar a conta.")


if __name__ == "__main__":
    main()
