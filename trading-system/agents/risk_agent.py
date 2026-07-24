"""
AGENTE 3 — GERENCIAMENTO DE RISCO (Risk Manager) — O MAIS IMPORTANTE

Princípio: NENHUMA ordem chega à exchange sem passar por aqui.
O Risk Manager tem PODER DE VETO ABSOLUTO e seus stops SOBRESCREVEM
qualquer valor sugerido pelo Orquestrador. Ele é deliberadamente
"burro e teimoso": não conhece a tese do trade, só conhece os números.

Camadas de proteção (todas precisam passar):
1. Kill switch / modo defensivo ativos?           -> VETO
2. Drawdown diário/semanal acima do limite?       -> VETO + KILL SWITCH
3. Perdas consecutivas / limite de trades do dia? -> VETO
4. Stop loss válido e risco/retorno >= mínimo?    -> VETO se não
5. Position sizing: risco por trade limitado a 1-2% do capital
6. Exposição total e nº de posições simultâneas dentro do teto

Além do veto, roda dois loops contínuos:
- Trailing stop: sobe o stop conforme o preço avança (nunca desce).
- Vigia de drawdown: recalcula equity a cada ciclo e aciona o
  circuit breaker sem depender de nenhum outro agente.
"""
from __future__ import annotations

import asyncio
import time

from agents.base import BaseAgent
from core.events import (Direction, ExecutionReport, KillSwitch, Notification,
                         RiskVerdict, Topic, TradeProposal)


class RiskAgent(BaseAgent):
    name = "risk_agent"

    def __init__(self, bus, state, config, market_data) -> None:
        super().__init__(bus, state, config)
        self.md = market_data

    async def run(self) -> None:
        inbox = self.bus.subscribe(Topic.TRADE_PROPOSAL, Topic.EXECUTION_REPORT,
                                   Topic.DEFENSIVE_MODE)
        asyncio.create_task(self.drawdown_watchdog())
        asyncio.create_task(self.trailing_stop_loop())
        asyncio.create_task(self.position_close_monitor())

        async for topic, event in self.bus.stream(inbox):
            if topic == Topic.TRADE_PROPOSAL:
                verdict = await self.evaluate(event)
                await self.emit(Topic.RISK_VERDICT, verdict)
            elif topic == Topic.EXECUTION_REPORT:
                await self.on_execution_report(event)
            elif topic == Topic.DEFENSIVE_MODE:
                self.state.defensive_mode_until = time.time() + event.cooldown_seconds
                self.log.warning("MODO DEFENSIVO por %ds: %s",
                                 event.cooldown_seconds, event.reason)

    # ------------------------------------------------------------------ #
    # Avaliação de propostas (compliance)                                #
    # ------------------------------------------------------------------ #
    async def evaluate(self, p: TradeProposal) -> RiskVerdict:
        cfg = self.config.risk
        veto = lambda reason: RiskVerdict(proposal_id=p.event_id, approved=False,
                                          reason=reason)

        # Camada 1 — estado do sistema
        if self.state.trading_halted:
            return veto("Kill switch ativo")
        if self.state.in_defensive_mode:
            return veto("Modo defensivo ativo — novas entradas bloqueadas")

        # Camada 2 — envelope diário em dólares (meta e limite de perda)
        if cfg.max_daily_loss_usd > 0 and \
                self.state.daily_pnl_usd <= -cfg.max_daily_loss_usd:
            await self.halt(f"Perda diária de ${abs(self.state.daily_pnl_usd):.2f} "
                            f">= limite ${cfg.max_daily_loss_usd:.2f} — parando até amanhã")
            return veto("Limite de perda diária atingido")
        if self.state.daily_profit_locked or (
                cfg.daily_profit_target_usd > 0 and
                self.state.daily_pnl_usd >= cfg.daily_profit_target_usd):
            self.state.daily_profit_locked = True
            return veto(f"Meta diária de ${cfg.daily_profit_target_usd:.2f} atingida "
                        f"— ganho travado, sem novas entradas hoje")

        # Camada 2b — circuit breakers de drawdown (percentual, rede de segurança)
        if self.state.daily_drawdown_pct >= cfg.max_daily_drawdown_pct:
            await self.halt(f"Drawdown diário {self.state.daily_drawdown_pct:.2f}% "
                            f">= limite {cfg.max_daily_drawdown_pct}%")
            return veto("Drawdown diário máximo atingido")
        if self.state.weekly_drawdown_pct >= cfg.max_weekly_drawdown_pct:
            await self.halt(f"Drawdown semanal {self.state.weekly_drawdown_pct:.2f}% "
                            f">= limite {cfg.max_weekly_drawdown_pct}%")
            return veto("Drawdown semanal máximo atingido")

        # Camada 3 — disciplina operacional
        if self.state.consecutive_losses >= cfg.max_consecutive_losses:
            return veto(f"{self.state.consecutive_losses} perdas consecutivas — pausa forçada")
        if self.state.trades_today >= cfg.max_trades_per_day:
            return veto("Limite de trades do dia atingido (overtrading)")
        if len(self.state.open_positions) >= cfg.max_open_positions:
            return veto("Número máximo de posições simultâneas atingido")
        if p.symbol in self.state.open_positions:
            return veto(f"Já existe posição aberta em {p.symbol}")

        # Camada 4 — sanidade do trade
        if p.entry_price <= 0 or p.stop_loss <= 0:
            return veto("Proposta sem preço de entrada ou stop loss válidos")
        risk_per_unit = abs(p.entry_price - p.stop_loss)
        if risk_per_unit == 0:
            return veto("Stop loss igual ao preço de entrada")
        is_long = p.direction == Direction.LONG
        if (is_long and p.stop_loss >= p.entry_price) or \
           (not is_long and p.stop_loss <= p.entry_price):
            return veto("Stop loss do lado errado da entrada")
        reward = abs(p.take_profit - p.entry_price)
        rr = reward / risk_per_unit
        if rr < cfg.min_risk_reward_ratio:
            return veto(f"Risco/retorno {rr:.2f} abaixo do mínimo {cfg.min_risk_reward_ratio}")

        # Camada 5 — POSITION SIZING (o coração da preservação de capital)
        # risco_monetário = equity * max_risk_pct  ==>  qty = risco / dist_do_stop
        risk_capital = self.state.equity * (cfg.max_risk_per_trade_pct / 100.0)
        qty = risk_capital / risk_per_unit

        # Camada 6 — teto de exposição nominal por posição
        max_notional = self.state.equity * (cfg.max_position_pct / 100.0)
        if qty * p.entry_price > max_notional:
            qty = max_notional / p.entry_price
            self.log.info("Qty reduzida pelo teto de exposição (%.1f%% do capital)",
                          cfg.max_position_pct)

        return RiskVerdict(
            proposal_id=p.event_id, approved=True,
            reason=f"Aprovado: risco {cfg.max_risk_per_trade_pct}% "
                   f"(${risk_capital:.2f}), R:R {rr:.2f}",
            position_size=round(qty, 8),
            stop_loss=p.stop_loss,           # imutável após aprovação
            take_profit=p.take_profit,
            trailing_stop_pct=cfg.default_trailing_stop_pct,
        )

    # ------------------------------------------------------------------ #
    # Circuit breaker                                                    #
    # ------------------------------------------------------------------ #
    async def halt(self, reason: str, close_positions: bool = False) -> None:
        self.state.trading_halted = True
        self.log.critical("KILL SWITCH: %s", reason)
        await self.emit(Topic.KILL_SWITCH, KillSwitch(reason=reason,
                                                      close_positions=close_positions))
        await self.emit(Topic.NOTIFICATION, Notification(
            level="critical", title="🛑 KILL SWITCH ATIVADO", body=reason,
            payload={"equity": self.state.equity,
                     "daily_dd": self.state.daily_drawdown_pct,
                     "weekly_dd": self.state.weekly_drawdown_pct}))

    STATUS_REPORT_EVERY = 60  # ciclos de 30s => placar a cada 30 min

    def _placar(self) -> tuple[str, str]:
        """Monta (emoji, texto) do placar: saldo, resultado do dia, posições."""
        s = self.state
        pnl_day = s.equity - s.day_start_equity
        pnl_pct = (pnl_day / s.day_start_equity * 100) if s.day_start_equity else 0.0
        emoji = "🟢" if pnl_day >= 0 else "🔴"
        positions = (", ".join(f"{sym} {p.side} @ {p.entry_price:.2f}"
                               for sym, p in s.open_positions.items())
                     or "nenhuma")
        text = (f"Saldo: {s.equity:,.2f} USDT\n"
                f"Resultado do dia: {pnl_day:+,.2f} USDT ({pnl_pct:+.2f}%)\n"
                f"Posições abertas: {positions}\n"
                f"Trades hoje: {s.trades_today}")
        return emoji, text

    async def send_placar(self, header: str | None = None) -> None:
        """Envia o placar ao Telegram (e loga no console)."""
        emoji, text = self._placar()
        self.log.info("%s PLACAR | %s", emoji, text.replace("\n", " | "))
        await self.emit(Topic.NOTIFICATION, Notification(
            level="info", to_telegram=True,
            title=header or f"{emoji} Placar do bot", body=text))

    async def report_status(self) -> None:
        await self.send_placar()

    async def position_close_monitor(self) -> None:
        """Detecta quando uma posição fecha na exchange (stop/alvo/trailing)
        e envia ao Telegram SÓ o essencial: ganhou/perdeu, quanto e em qual
        moeda."""
        if self.md.market_type != "futures":
            return  # (spot fecha via on_execution_report)
        while True:
            try:
                open_syms = {str(p["symbol"]).split(":")[0]
                             for p in await self.md.fetch_open_positions()}
                for sym, pos in list(self.state.open_positions.items()):
                    if sym not in open_syms:
                        await self.on_position_closed(sym, pos)
            except Exception as exc:
                self.log.warning("Monitor de fechamento: %s", exc)
            await asyncio.sleep(12)

    async def on_position_closed(self, sym: str, pos) -> None:
        try:
            exit_price = await self.md.fetch_last_price(sym)
        except Exception:
            exit_price = pos.stop_loss
        pnl = (exit_price - pos.entry_price) * pos.qty
        if pos.side == "short":
            pnl = -pnl
        win = pnl >= 0
        self.state.consecutive_losses = 0 if win else self.state.consecutive_losses + 1
        self.state.open_positions.pop(sym, None)
        emoji = "🟢" if win else "🔴"
        verb = "GANHEI" if win else "PERDI"
        self.log.info("%s POSIÇÃO FECHADA %s: %+.2f USDT", emoji, sym, pnl)
        # Placar com o resultado da ordem no cabeçalho (fechamento é "ordem")
        moeda = "compra" if pos.side == "long" else "venda"
        await self.send_placar(
            header=f"{emoji} {verb} {abs(pnl):.2f} USDT em {sym} ({moeda})")

    async def drawdown_watchdog(self) -> None:
        """Vigia independente: mesmo sem novas propostas, o drawdown é checado."""
        cfg = self.config.risk
        cycles = 0
        while True:
            try:
                equity = await self.md.fetch_total_equity()
                if equity is not None:
                    await self.state.update_equity(equity)
                if not self.state.trading_halted:
                    # Envelope diário em dólares
                    if cfg.max_daily_loss_usd > 0 and \
                            self.state.daily_pnl_usd <= -cfg.max_daily_loss_usd:
                        await self.halt(
                            f"Perda diária de ${abs(self.state.daily_pnl_usd):.2f} "
                            f"atingiu o limite de ${cfg.max_daily_loss_usd:.2f}",
                            close_positions=True)
                    elif (cfg.daily_profit_target_usd > 0 and not
                          self.state.daily_profit_locked and
                          self.state.daily_pnl_usd >= cfg.daily_profit_target_usd):
                        self.state.daily_profit_locked = True
                        self.log.info("🎯 META DIÁRIA de $%.2f atingida — ganho "
                                      "travado, sem novas entradas hoje",
                                      cfg.daily_profit_target_usd)
                        await self.send_placar(
                            header=f"🎯 Meta diária de ${cfg.daily_profit_target_usd:.0f} "
                                   f"atingida! Parando por hoje.")
                    elif self.state.daily_drawdown_pct >= cfg.max_daily_drawdown_pct:
                        await self.halt("Watchdog: drawdown diário máximo",
                                        close_positions=True)
                    elif self.state.weekly_drawdown_pct >= cfg.max_weekly_drawdown_pct:
                        await self.halt("Watchdog: drawdown semanal máximo",
                                        close_positions=True)
                if cycles % self.STATUS_REPORT_EVERY == 0:
                    await self.report_status()
            except Exception as exc:
                self.log.warning("Watchdog: falha ao ler equity: %s", exc)
            cycles += 1
            await asyncio.sleep(30)

    # ------------------------------------------------------------------ #
    # Trailing stop (stop móvel) — nunca afrouxa, só aperta              #
    # ------------------------------------------------------------------ #
    # Só recria o stop na exchange se ele melhorar mais que isto (evita
    # recriar ordem a cada tick — principal fonte de acúmulo de -4045).
    TRAIL_MIN_STEP = 0.004  # 0,4%

    async def trailing_stop_loop(self) -> None:
        while True:
            for symbol, pos in list(self.state.open_positions.items()):
                if not pos.trailing_stop_pct:
                    continue
                try:
                    price = await self.md.fetch_last_price(symbol)
                    if pos.side == "long":
                        pos.highest_price = max(pos.highest_price, price)
                        candidate = pos.highest_price * (1 - pos.trailing_stop_pct / 100)
                        if candidate > pos.stop_loss * (1 + self.TRAIL_MIN_STEP):
                            self.log.info("%s trailing stop: %.2f -> %.2f",
                                          symbol, pos.stop_loss, candidate)
                            pos.stop_loss = candidate
                            await self.md.replace_stop_order(symbol, pos)
                    else:
                        pos.lowest_price = min(pos.lowest_price, price)
                        candidate = pos.lowest_price * (1 + pos.trailing_stop_pct / 100)
                        if candidate < pos.stop_loss * (1 - self.TRAIL_MIN_STEP):
                            pos.stop_loss = candidate
                            await self.md.replace_stop_order(symbol, pos)
                except Exception as exc:
                    self.log.warning("Trailing de %s falhou: %s", symbol, exc)
            await asyncio.sleep(10)

    # ------------------------------------------------------------------ #
    # Contabilidade pós-execução                                         #
    # ------------------------------------------------------------------ #
    async def on_execution_report(self, report: ExecutionReport) -> None:
        if report.status != "filled":
            return
        self.state.trades_today += 1
        pos = self.state.open_positions.get(report.symbol)
        is_close = pos and ((pos.side == "long" and report.side == "sell") or
                            (pos.side == "short" and report.side == "buy"))
        if is_close:
            pnl = (report.fill_price - pos.entry_price) * report.filled_qty
            if pos.side == "short":
                pnl = -pnl
            self.state.consecutive_losses = (self.state.consecutive_losses + 1
                                             if pnl < 0 else 0)
            del self.state.open_positions[report.symbol]
        else:
            # Entrada (abertura): manda o placar após a ordem.
            side = "compra" if report.side == "buy" else "venda"
            await self.send_placar(
                header=f"📥 Ordem de {side} em {report.symbol} @ "
                       f"{report.fill_price:.2f}")
