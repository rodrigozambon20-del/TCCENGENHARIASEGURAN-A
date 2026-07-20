"""
AGENTE 5 — ORQUESTRADOR / DIRETOR (Chief Executive Agent)

Fluxo de decisão (pipeline de aprovação):

  SentimentSignal ─┐
                   ├─> fusão (conviction) ─> TradeProposal ─> RiskAgent
  QuantSignal ─────┘                                             │
                                                        RiskVerdict (veto?)
                                                                 │ aprovado
                                                          ExecutionOrder
                                                                 │
                                                          ExecutionAgent

Regras de fusão:
- O Quant dá a DIREÇÃO e os NÍVEIS (entrada, stop por ATR, alvo em S/R).
- O Sentimento atua como filtro/booster: sentimento contrário forte
  cancela a entrada; sentimento alinhado aumenta a convicção.
- Sinais têm validade (TTL): sentimento de 10 min atrás não conta.
- O Orquestrador nunca calcula tamanho de posição — isso é do Risco.
"""
from __future__ import annotations

import time

from agents.base import BaseAgent
from core.events import (Direction, ExecutionOrder, Notification, QuantSignal,
                         RiskVerdict, Sentiment, SentimentSignal, Topic,
                         TradeProposal)

SENTIMENT_TTL = 600      # s — validade de um sinal de sentimento
QUANT_TTL = 120          # s — validade de um sinal técnico
MIN_CONVICTION = 0.5     # convicção mínima para propor trade


class OrchestratorAgent(BaseAgent):
    name = "orchestrator"

    def __init__(self, bus, state, config) -> None:
        super().__init__(bus, state, config)
        self.latest_sentiment: SentimentSignal | None = None
        self.latest_quant: dict[str, QuantSignal] = {}
        self.pending_proposals: dict[str, TradeProposal] = {}

    async def run(self) -> None:
        inbox = self.bus.subscribe(
            Topic.SENTIMENT_SIGNAL, Topic.QUANT_SIGNAL,
            Topic.RISK_VERDICT, Topic.KILL_SWITCH,
        )
        async for topic, event in self.bus.stream(inbox):
            if topic == Topic.KILL_SWITCH:
                self.pending_proposals.clear()
                self.log.critical("Kill switch recebido — pipeline de decisão limpo")
            elif topic == Topic.SENTIMENT_SIGNAL:
                self.latest_sentiment = event
            elif topic == Topic.QUANT_SIGNAL:
                self.latest_quant[event.symbol] = event
                await self.maybe_propose(event.symbol)
            elif topic == Topic.RISK_VERDICT:
                await self.on_verdict(event)

    # ------------------------------------------------------------------ #
    # Fusão de sinais -> proposta                                        #
    # ------------------------------------------------------------------ #
    async def maybe_propose(self, symbol: str) -> None:
        if self.state.trading_halted or self.state.in_defensive_mode:
            return

        quant = self.latest_quant.get(symbol)
        now = time.time()
        if not quant or now - quant.ts > QUANT_TTL or quant.direction == Direction.FLAT:
            return

        sent = self.latest_sentiment
        sent_score = 0.0
        if sent and now - sent.ts <= SENTIMENT_TTL:
            sent_score = sent.score * sent.confidence
            # Filtro de veto por sentimento contrário forte
            if quant.direction == Direction.LONG and sent.sentiment == Sentiment.BEARISH \
               and sent.confidence > 0.5:
                self.log.info("%s: entrada LONG vetada por sentimento bearish", symbol)
                return
            if quant.direction == Direction.SHORT and sent.sentiment == Sentiment.BULLISH \
               and sent.confidence > 0.5:
                return

        # Convicção = 70% técnica + 30% sentimento (alinhado soma, contrário subtrai)
        direction_mult = 1.0 if quant.direction == Direction.LONG else -1.0
        conviction = 0.7 * abs(quant.score) + 0.3 * max(0.0, sent_score * direction_mult)
        if conviction < MIN_CONVICTION:
            return

        price = quant.last_price
        if quant.strategy == "pullback":
            # Estratégia vencedora do backtest: risco inicial = trailing %,
            # saída real fica por conta do stop móvel do Risk Manager
            trail = self.config.risk.default_trailing_stop_pct
            stop = price * (1 - trail / 100)
            target = price * (1 + 2 * trail / 100)   # R:R formal de 2:1
        elif quant.direction == Direction.LONG:
            # Modo votes: stop = 1.5x ATR; alvo no S/R na direção do trade
            atr = quant.atr or price * 0.01
            stop = max(price - 1.5 * atr, (quant.support or 0) * 0.999)
            target = quant.resistance or price + 3 * atr
        else:
            atr = quant.atr or price * 0.01
            stop = min(price + 1.5 * atr, (quant.resistance or 1e18) * 1.001)
            target = quant.support or price - 3 * atr

        proposal = TradeProposal(
            symbol=symbol, direction=quant.direction,
            entry_price=price, stop_loss=stop, take_profit=target,
            conviction=conviction,
            rationale=(f"Quant {quant.score:+.2f} {quant.timeframe_votes} | "
                       f"sentimento {sent_score:+.2f} | "
                       f"orderflow {quant.orderflow_imbalance:+.2f}"),
        )
        self.pending_proposals[proposal.event_id] = proposal
        self.log.info("Proposta %s %s @ %.2f (SL %.2f / TP %.2f, conv %.2f) -> Risco",
                      symbol, quant.direction.value, price, stop, target, conviction)
        await self.emit(Topic.TRADE_PROPOSAL, proposal)

    # ------------------------------------------------------------------ #
    # Resposta do Risk Manager                                           #
    # ------------------------------------------------------------------ #
    async def on_verdict(self, verdict: RiskVerdict) -> None:
        proposal = self.pending_proposals.pop(verdict.proposal_id, None)
        if proposal is None:
            return
        if not verdict.approved:
            self.log.info("Proposta %s VETADA pelo Risco: %s",
                          proposal.symbol, verdict.reason)
            return
        self.log.info("Proposta %s APROVADA (%s) -> Execução",
                      proposal.symbol, verdict.reason)
        await self.emit(Topic.EXECUTION_ORDER,
                        ExecutionOrder(proposal=proposal, verdict=verdict))
        await self.emit(Topic.NOTIFICATION, Notification(
            level="info", title=f"📈 Nova operação: {proposal.symbol}",
            body=f"{proposal.direction.value.upper()} {verdict.position_size} @ "
                 f"{proposal.entry_price:.2f} | SL {verdict.stop_loss:.2f} | "
                 f"TP {verdict.take_profit:.2f}\n{proposal.rationale}",
            payload={"conviction": proposal.conviction}))
