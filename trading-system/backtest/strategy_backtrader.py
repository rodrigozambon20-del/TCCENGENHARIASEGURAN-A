"""
Template de backtest com Backtrader replicando as MESMAS regras do bot:
- Sinal: KAMA/SMA + RSI + Bollinger (espelho do QuantAgent)
- Risco: 1% por trade, stop por ATR, trailing stop (espelho do RiskAgent)

Regra de ouro: NENHUMA regra vai para produção sem antes ser aprovada
aqui e depois em paper trading (testnet). Backtest, forward test e
produção devem compartilhar os mesmos parâmetros do config.yaml.

Uso:
    python backtest/strategy_backtrader.py --symbol BTC/USDT --tf 1h --days 365
"""
from __future__ import annotations

import argparse
import datetime as dt

import backtrader as bt


class MultiAgentMirrorStrategy(bt.Strategy):
    params = dict(
        risk_per_trade_pct=1.0,
        atr_period=14,
        atr_stop_mult=1.5,
        rsi_period=14,
        bb_period=20,
        bb_dev=2.0,
        kama_period=10,
        trailing_pct=1.5,
        min_rr=1.5,
        max_daily_dd_pct=3.0,
    )

    def __init__(self) -> None:
        self.kama = bt.ind.AdaptiveMovingAverage(period=self.p.kama_period)
        self.rsi = bt.ind.RSI(period=self.p.rsi_period)
        self.bb = bt.ind.BollingerBands(period=self.p.bb_period,
                                        devfactor=self.p.bb_dev)
        self.atr = bt.ind.ATR(period=self.p.atr_period)
        self.order = None
        self.day_start_value = None
        self.current_day = None
        self.halted_today = False

    # -------------------- espelho do RiskAgent -------------------- #
    def circuit_breaker_hit(self) -> bool:
        """Replica o drawdown diário: se estourar, para de operar no dia."""
        today = self.data.datetime.date(0)
        if today != self.current_day:
            self.current_day = today
            self.day_start_value = self.broker.getvalue()
            self.halted_today = False
        dd = (self.day_start_value - self.broker.getvalue()) / self.day_start_value * 100
        if dd >= self.p.max_daily_dd_pct:
            self.halted_today = True
        return self.halted_today

    def position_size(self, entry: float, stop: float) -> float:
        risk_capital = self.broker.getvalue() * self.p.risk_per_trade_pct / 100
        risk_per_unit = abs(entry - stop)
        return risk_capital / risk_per_unit if risk_per_unit else 0.0

    # -------------------- espelho do QuantAgent -------------------- #
    def quant_vote(self) -> float:
        price = self.data.close[0]
        vote = 0.4 if price > self.kama[0] else -0.4
        if self.rsi[0] < 30:
            vote += 0.3
        elif self.rsi[0] > 70:
            vote -= 0.3
        if price <= self.bb.lines.bot[0]:
            vote += 0.3
        elif price >= self.bb.lines.top[0]:
            vote -= 0.3
        return vote

    def next(self) -> None:
        if self.order or self.circuit_breaker_hit():
            return

        price = self.data.close[0]
        vote = self.quant_vote()

        if not self.position:
            if vote >= 0.35:  # min_score_to_signal do config
                stop = price - self.p.atr_stop_mult * self.atr[0]
                target = price + self.p.min_rr * (price - stop)
                size = self.position_size(price, stop)
                if size > 0:
                    # bracket = entrada + stop + alvo, como o OCO em produção
                    self.order = self.buy_bracket(
                        size=size, price=price,
                        stopprice=stop, limitprice=target,
                        exectype=bt.Order.Market)[0]
        # (Trailing stop pode ser adicionado com bt.Order.StopTrail
        #  usando trailpercent=self.p.trailing_pct / 100)

    def notify_order(self, order) -> None:
        if order.status in (order.Completed, order.Canceled,
                            order.Margin, order.Rejected):
            self.order = None


def run(symbol: str, timeframe: str, days: int) -> None:
    cerebro = bt.Cerebro()
    cerebro.addstrategy(MultiAgentMirrorStrategy)
    cerebro.broker.setcash(10_000)
    cerebro.broker.setcommission(commission=0.001)  # 0.1% taker Binance
    cerebro.broker.set_slippage_perc(0.0005)        # slippage realista

    # Dados: baixe OHLCV via ccxt e converta para PandasData
    #   import ccxt, pandas as pd
    #   ex = ccxt.binance()
    #   since = ex.milliseconds() - days * 86_400_000
    #   rows = ex.fetch_ohlcv(symbol, timeframe, since=since, limit=1000)
    #   df = pd.DataFrame(rows, columns=["ts","open","high","low","close","volume"])
    #   df["ts"] = pd.to_datetime(df["ts"], unit="ms"); df.set_index("ts", inplace=True)
    #   cerebro.adddata(bt.feeds.PandasData(dataname=df))

    cerebro.addanalyzer(bt.analyzers.SharpeRatio, _name="sharpe")
    cerebro.addanalyzer(bt.analyzers.DrawDown, _name="dd")
    cerebro.addanalyzer(bt.analyzers.TradeAnalyzer, _name="trades")

    results = cerebro.run()
    strat = results[0]
    print("Sharpe:", strat.analyzers.sharpe.get_analysis())
    print("Max Drawdown:", strat.analyzers.dd.get_analysis().max.drawdown, "%")
    print("Valor final:", cerebro.broker.getvalue())


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--symbol", default="BTC/USDT")
    ap.add_argument("--tf", default="1h")
    ap.add_argument("--days", type=int, default=365)
    args = ap.parse_args()
    run(args.symbol, args.tf, args.days)
