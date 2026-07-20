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
        # --- Filtros opcionais (laboratório de melhoria) ---
        min_vote=0.35,       # score mínimo de confluência para entrar
        trend_bars=0,        # >0: só compra acima da SMA(n barras) — use n = dias
                             #     x barras/dia p/ emular a média DIÁRIA
        max_atr_pct=0.0,     # >0: não entra se ATR% > limite (volatilidade extrema)
        mode="votes",        # "votes" (original) | "pullback" (recuo na tendência)
        pullback_rsi=45.0,   # modo pullback: RSI máximo para considerar "recuo"
        exit_mode="bracket", # "bracket" (stop+alvo fixos) | "trail" (stop móvel)
        trail_pct=8.0,       # modo trail: distância % do stop móvel
    )

    def __init__(self) -> None:
        self.kama = bt.ind.AdaptiveMovingAverage(period=self.p.kama_period)
        self.rsi = bt.ind.RSI(period=self.p.rsi_period)
        self.bb = bt.ind.BollingerBands(period=self.p.bb_period,
                                        devfactor=self.p.bb_dev)
        self.atr = bt.ind.ATR(period=self.p.atr_period)
        self.daily_sma = None
        if self.p.trend_bars:
            self.daily_sma = bt.ind.SMA(self.data.close, period=self.p.trend_bars)
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
            # Filtro 1: tendência do gráfico diário — só compra em mercado de alta
            if self.daily_sma is not None and price < self.daily_sma[0]:
                return
            # Filtro 2: volatilidade extrema — fora do mercado em pânico
            if self.p.max_atr_pct and (self.atr[0] / price * 100) > self.p.max_atr_pct:
                return
            # Sinal de entrada conforme o modo
            if self.p.mode == "pullback":
                # Recuo dentro da tendência: RSI respirou E preço voltou à média
                should_enter = (self.rsi[0] <= self.p.pullback_rsi
                                and price <= self.bb.lines.mid[0])
            else:
                # Filtro 3 (modo votes): exigência de confluência mínima
                should_enter = vote >= self.p.min_vote
            if should_enter:
                if self.p.exit_mode == "trail":
                    # Saída por stop móvel: risco inicial = trail_pct
                    stop = price * (1 - self.p.trail_pct / 100)
                    size = self.position_size(price, stop)
                    if size > 0:
                        self.order = self.buy(size=size)
                else:
                    stop = price - self.p.atr_stop_mult * self.atr[0]
                    target = price + self.p.min_rr * (price - stop)
                    size = self.position_size(price, stop)
                    if size > 0:
                        # bracket = entrada + stop + alvo, como o OCO em produção
                        self.order = self.buy_bracket(
                            size=size, price=price,
                            stopprice=stop, limitprice=target,
                            exectype=bt.Order.Market)[0]

    def notify_order(self, order) -> None:
        if order.status == order.Completed and order.isbuy() \
                and self.p.exit_mode == "trail":
            # Arma o stop móvel logo após o fill da entrada
            self.sell(size=order.executed.size, exectype=bt.Order.StopTrail,
                      trailpercent=self.p.trail_pct / 100)
        if order.status in (order.Completed, order.Canceled,
                            order.Margin, order.Rejected):
            self.order = None


def load_data(symbol: str, timeframe: str, days: int):
    """Baixa OHLCV real do espelho público de dados da Binance.

    data-api.binance.vision: sem chave de API, sem bloqueio regional e
    sem os 17 MB de metadados do exchangeInfo — só os candles.
    """
    import time as _time

    import pandas as pd
    import requests

    url = "https://data-api.binance.vision/api/v3/klines"
    market = symbol.replace("/", "")
    since = int(_time.time() * 1000) - days * 86_400_000
    rows: list = []
    print(f"Baixando {days} dias de {symbol} ({timeframe}) da Binance...")
    session = requests.Session()
    while True:
        resp = session.get(url, params={"symbol": market, "interval": timeframe,
                                        "startTime": since, "limit": 1000},
                           timeout=30)
        resp.raise_for_status()
        batch = resp.json()
        if not batch:
            break
        rows += [[c[0], float(c[1]), float(c[2]), float(c[3]),
                  float(c[4]), float(c[5])] for c in batch]
        since = batch[-1][0] + 1
        print(f"  {len(rows)} candles...", end="\r")
        if len(batch) < 1000:
            break
        _time.sleep(0.15)  # cortesia com o rate limit público
    df = pd.DataFrame(rows, columns=["ts", "open", "high", "low", "close", "volume"])
    df["ts"] = pd.to_datetime(df["ts"], unit="ms")
    df.set_index("ts", inplace=True)
    print(f"\n{len(df)} candles carregados "
          f"({df.index[0]:%d/%m/%Y} a {df.index[-1]:%d/%m/%Y})")
    return df


def run(symbol: str, timeframe: str, days: int) -> None:
    cerebro = bt.Cerebro()
    cerebro.addstrategy(MultiAgentMirrorStrategy)
    cerebro.broker.setcash(10_000)
    cerebro.broker.setcommission(commission=0.001)  # 0.1% taker Binance
    cerebro.broker.set_slippage_perc(0.0005)        # slippage realista

    df = load_data(symbol, timeframe, days)
    cerebro.adddata(bt.feeds.PandasData(dataname=df))

    cerebro.addanalyzer(bt.analyzers.SharpeRatio, _name="sharpe",
                        timeframe=bt.TimeFrame.Days, riskfreerate=0.0)
    cerebro.addanalyzer(bt.analyzers.DrawDown, _name="dd")
    cerebro.addanalyzer(bt.analyzers.TradeAnalyzer, _name="trades")

    print("Rodando backtest...")
    strat = cerebro.run()[0]

    final_value = cerebro.broker.getvalue()
    ret_pct = (final_value / 10_000 - 1) * 100
    dd = strat.analyzers.dd.get_analysis()
    trades = strat.analyzers.trades.get_analysis()
    sharpe = strat.analyzers.sharpe.get_analysis().get("sharperatio")

    total = trades.get("total", {}).get("closed", 0)
    won = trades.get("won", {}).get("total", 0)
    lost = trades.get("lost", {}).get("total", 0)
    win_rate = won / total * 100 if total else 0.0
    avg_win = trades.get("won", {}).get("pnl", {}).get("average", 0.0)
    avg_loss = trades.get("lost", {}).get("pnl", {}).get("average", 0.0)

    print()
    print("=" * 56)
    print(f"  RESULTADO DO BACKTEST — {symbol} {timeframe}, {days} dias")
    print("=" * 56)
    print(f"  Capital inicial:     10,000.00 USDT")
    print(f"  Capital final:       {final_value:,.2f} USDT")
    print(f"  Retorno:             {ret_pct:+.2f}%")
    print(f"  Drawdown máximo:     {dd.max.drawdown:.2f}%")
    print(f"  Sharpe (diário):     {sharpe if sharpe is not None else 'n/d'}")
    print(f"  Trades fechados:     {total} ({won} ganhos / {lost} perdas)")
    print(f"  Taxa de acerto:      {win_rate:.1f}%")
    print(f"  Ganho médio:         {avg_win:+.2f} USDT")
    print(f"  Perda média:         {avg_loss:+.2f} USDT")
    print("=" * 56)
    print("  Lembre: resultado passado não garante resultado futuro.")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--symbol", default="BTC/USDT")
    ap.add_argument("--tf", default="1h")
    ap.add_argument("--days", type=int, default=365)
    args = ap.parse_args()
    run(args.symbol, args.tf, args.days)
