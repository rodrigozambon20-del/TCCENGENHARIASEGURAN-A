"""
AGENTE 4 — EXECUÇÃO E LATÊNCIA (Execution Bot)

Estratégia de execução inteligente:
1. Preferir ordem LIMITADA post-only levemente dentro do spread
   (paga taxa maker, menor; slippage zero).
2. Se não preencher em `limit_timeout_seconds`, cancelar e reavaliar:
   - preço ainda dentro da tolerância de slippage -> MARKET (garante fill);
   - preço fugiu além de `max_slippage_bps` -> ABORTAR e reportar.
3. Após o fill de entrada, registrar a posição e delegar stops ao
   Risk Manager (que mantém trailing e OCO stop/target na exchange).
4. Todo resultado (fill, parcial, abortado, erro) vira ExecutionReport
   no bus — o sistema inteiro fica auditável.

Kill switch: ao receber, cancela TODAS as ordens abertas e, se
solicitado, fecha posições a mercado.
"""
from __future__ import annotations

import asyncio

from agents.base import BaseAgent
from core.events import (Direction, ExecutionOrder, ExecutionReport,
                         Notification, Topic)
from core.state import Position


class ExecutionAgent(BaseAgent):
    name = "execution_agent"

    def __init__(self, bus, state, config, exchange) -> None:
        super().__init__(bus, state, config)
        self.ex = exchange  # exchange.binance_client.BinanceClient

    async def run(self) -> None:
        inbox = self.bus.subscribe(Topic.EXECUTION_ORDER, Topic.KILL_SWITCH)
        async for topic, event in self.bus.stream(inbox):
            if topic == Topic.KILL_SWITCH:
                await self.emergency_stop(close_positions=event.close_positions)
            elif topic == Topic.EXECUTION_ORDER:
                # Revalida o estado NO MOMENTO da execução (pode ter mudado
                # entre a aprovação do risco e a chegada aqui)
                if self.state.trading_halted:
                    self.log.warning("Ordem descartada: kill switch ativou após aprovação")
                    continue
                await self.execute(event)

    # ------------------------------------------------------------------ #
    # Execução inteligente                                               #
    # ------------------------------------------------------------------ #
    async def execute(self, order: ExecutionOrder) -> None:
        cfg = self.config.execution
        p, v = order.proposal, order.verdict
        side = "buy" if p.direction == Direction.LONG else "sell"

        report = ExecutionReport(symbol=p.symbol, side=side,
                                 requested_price=p.entry_price)
        try:
            filled = None
            if cfg.prefer_limit_orders:
                filled = await self.try_limit_fill(p.symbol, side,
                                                   v.position_size, p.entry_price)
            if filled is None:
                # Fallback market — mas só se o slippage estimado for aceitável
                last = await self.ex.fetch_last_price(p.symbol)
                slippage_bps = abs(last - p.entry_price) / p.entry_price * 10_000
                if slippage_bps > cfg.max_slippage_bps:
                    report.status = "rejected"
                    report.error = (f"Slippage estimado {slippage_bps:.1f} bps > "
                                    f"máximo {cfg.max_slippage_bps} bps — abortado")
                    self.log.warning(report.error)
                    await self.emit(Topic.EXECUTION_REPORT, report)
                    return
                filled = await self.ex.create_market_order(p.symbol, side,
                                                           v.position_size)

            report.order_id = str(filled.get("id", ""))
            report.fill_price = float(filled.get("average") or filled.get("price")
                                      or p.entry_price)
            report.filled_qty = float(filled.get("filled") or v.position_size)
            report.fees = float((filled.get("fee") or {}).get("cost") or 0.0)
            report.slippage_bps = (abs(report.fill_price - p.entry_price)
                                   / p.entry_price * 10_000)
            report.status = "filled"

            # Registra posição e arma proteção OCO (stop + alvo) na exchange:
            # a proteção mora na Binance, não só no bot — se o bot cair,
            # o stop continua lá.
            pos = Position(symbol=p.symbol, side=p.direction.value,
                           qty=report.filled_qty, entry_price=report.fill_price,
                           stop_loss=v.stop_loss, take_profit=v.take_profit,
                           trailing_stop_pct=v.trailing_stop_pct,
                           highest_price=report.fill_price,
                           lowest_price=report.fill_price)
            self.state.open_positions[p.symbol] = pos
            await self.ex.place_protective_orders(p.symbol, pos)

            self.log.info("EXECUTADO %s %s %.6f @ %.2f (slippage %.1f bps)",
                          side, p.symbol, report.filled_qty, report.fill_price,
                          report.slippage_bps)
        except Exception as exc:
            report.status = "error"
            report.error = str(exc)
            self.log.exception("Falha na execução de %s", p.symbol)
            await self.emit(Topic.NOTIFICATION, Notification(
                level="critical", title=f"⚠️ Falha de execução {p.symbol}",
                body=str(exc)))

        await self.emit(Topic.EXECUTION_REPORT, report)

    async def try_limit_fill(self, symbol: str, side: str, qty: float,
                             ref_price: float) -> dict | None:
        """Ordem limitada post-only com timeout; devolve fill ou None."""
        cfg = self.config.execution
        book = await self.ex.fetch_order_book(symbol, 5)
        best_bid, best_ask = book["bids"][0][0], book["asks"][0][0]
        offset = ref_price * cfg.limit_offset_bps / 10_000
        limit_price = (best_bid + offset) if side == "buy" else (best_ask - offset)

        order = await self.ex.create_limit_order(symbol, side, qty, limit_price,
                                                 post_only=True)
        deadline = asyncio.get_event_loop().time() + cfg.limit_timeout_seconds
        while asyncio.get_event_loop().time() < deadline:
            await asyncio.sleep(cfg.order_poll_interval)
            status = await self.ex.fetch_order(order["id"], symbol)
            if status["status"] in ("closed", "filled"):
                return status
            if status["status"] in ("canceled", "rejected", "expired"):
                return None
        # Timeout: cancela o resto e devolve parcial se houver
        await self.ex.cancel_order(order["id"], symbol)
        status = await self.ex.fetch_order(order["id"], symbol)
        return status if float(status.get("filled") or 0) > 0 else None

    # ------------------------------------------------------------------ #
    # Parada de emergência                                               #
    # ------------------------------------------------------------------ #
    async def emergency_stop(self, close_positions: bool) -> None:
        self.log.critical("PARADA DE EMERGÊNCIA (fechar posições=%s)", close_positions)
        for symbol in self.config.symbols:
            try:
                await self.ex.cancel_all_orders(symbol)
            except Exception as exc:
                self.log.error("Falha ao cancelar ordens de %s: %s", symbol, exc)
        if close_positions:
            for symbol, pos in list(self.state.open_positions.items()):
                try:
                    side = "sell" if pos.side == "long" else "buy"
                    await self.ex.create_market_order(symbol, side, pos.qty)
                    del self.state.open_positions[symbol]
                    self.log.critical("Posição %s fechada a mercado", symbol)
                except Exception as exc:
                    self.log.error("FALHA AO FECHAR %s: %s — INTERVENÇÃO MANUAL",
                                   symbol, exc)
                    await self.emit(Topic.NOTIFICATION, Notification(
                        level="critical",
                        title=f"🚨 INTERVENÇÃO MANUAL NECESSÁRIA: {symbol}",
                        body=f"Não consegui fechar a posição: {exc}"))
