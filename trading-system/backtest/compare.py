"""
Laboratório de estratégias: roda a MESMA base de dados por várias
variações da estratégia e imprime um ranking comparativo.

Uso:
    python backtest/compare.py --symbol BTC/USDT --tf 1h --days 365
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import backtrader as bt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from strategy_backtrader import MultiAgentMirrorStrategy, load_data  # noqa: E402

# Barras por dia em cada timeframe (p/ converter "SMA de N dias" em barras)
BARS_PER_DAY = {"1m": 1440, "5m": 288, "15m": 96, "30m": 48,
                "1h": 24, "2h": 12, "4h": 6, "1d": 1}

# Cada variação = nome + parâmetros sobrescritos ("trend_days" é convertido)
VARIATIONS: dict[str, dict] = {
    "A. Original (score 0.35)": {},
    "B. Score alto (0.60)": {"min_vote": 0.60},
    "C. Tendência diária (SMA50)": {"trend_days": 50},
    "D. Filtro volatilidade (ATR<=1.2%)": {"max_atr_pct": 1.2},
    "E. Tendência + score alto": {"trend_days": 50, "min_vote": 0.60},
    "F. Completa (tend + score + vol + RR 2)": {
        "trend_days": 50, "min_vote": 0.60, "max_atr_pct": 1.2, "min_rr": 2.0},
    "G. Pullback na tendência (RSI<45, RR2)": {
        "mode": "pullback", "trend_days": 50, "min_rr": 2.0},
    "H. Pullback conservador (RSI<40, RR2)": {
        "mode": "pullback", "trend_days": 50, "min_rr": 2.0, "pullback_rsi": 40.0},
    "I. Pullback + trailing 8%": {
        "mode": "pullback", "trend_days": 50, "exit_mode": "trail", "trail_pct": 8.0},
    "J. Pullback + trailing 5%": {
        "mode": "pullback", "trend_days": 50, "exit_mode": "trail", "trail_pct": 5.0},
}


def run_config(df, overrides: dict, timeframe: str) -> dict:
    overrides = dict(overrides)
    trend_days = overrides.pop("trend_days", 0)
    if trend_days:
        overrides["trend_bars"] = trend_days * BARS_PER_DAY.get(timeframe, 24)
    cerebro = bt.Cerebro()
    cerebro.adddata(bt.feeds.PandasData(dataname=df))
    cerebro.addstrategy(MultiAgentMirrorStrategy, **overrides)
    cerebro.broker.setcash(10_000)
    cerebro.broker.setcommission(commission=0.001)
    cerebro.broker.set_slippage_perc(0.0005)
    cerebro.addanalyzer(bt.analyzers.DrawDown, _name="dd")
    cerebro.addanalyzer(bt.analyzers.TradeAnalyzer, _name="trades")
    strat = cerebro.run()[0]

    trades = strat.analyzers.trades.get_analysis()
    total = trades.get("total", {}).get("closed", 0)
    won = trades.get("won", {}).get("total", 0)
    final = cerebro.broker.getvalue()
    return {
        "ret_pct": (final / 10_000 - 1) * 100,
        "max_dd": strat.analyzers.dd.get_analysis().max.drawdown,
        "trades": total,
        "win_rate": (won / total * 100) if total else 0.0,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--symbol", default="BTC/USDT")
    ap.add_argument("--tf", default="1h")
    ap.add_argument("--days", type=int, default=365)
    args = ap.parse_args()

    df = load_data(args.symbol, args.tf, args.days)

    results = {}
    for name, overrides in VARIATIONS.items():
        print(f"Testando: {name} ...")
        results[name] = run_config(df, overrides, args.tf)

    print()
    print("=" * 78)
    print(f"  COMPARATIVO — {args.symbol} {args.tf}, {args.days} dias "
          f"(capital inicial 10.000)")
    print("=" * 78)
    print(f"  {'Variação':<40} {'Retorno':>9} {'MaxDD':>7} {'Trades':>7} {'Acerto':>7}")
    print("-" * 78)
    ranked = sorted(results.items(), key=lambda kv: kv[1]["ret_pct"], reverse=True)
    for name, r in ranked:
        print(f"  {name:<40} {r['ret_pct']:>+8.2f}% {r['max_dd']:>6.2f}% "
              f"{r['trades']:>7d} {r['win_rate']:>6.1f}%")
    print("=" * 78)
    print("  Critério de aprovação: retorno positivo E drawdown máximo < 15%.")
    print("  Lembre: resultado passado não garante resultado futuro.")


if __name__ == "__main__":
    main()
