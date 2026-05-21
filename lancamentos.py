"""Geração de lançamentos contábeis automáticos a partir de regras
aplicadas sobre os pendentes do OFX (tarifas, IOF, juros, etc)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Any

from parser_xlsx import Transacao


@dataclass
class LancamentoContabil:
    data: date                   # data de pagamento (vem do OFX)
    historico: str               # texto do histórico contábil (vem da regra)
    valor: Decimal               # valor do pagamento (do OFX)
    banco: str                   # banco identificado no OFX
    memo_original: str           # memo bruto do OFX, pra referência
    padrao_match: str            # qual padrão da regra casou
    transacao_origem: Transacao | None = None  # ref à Transacao do OFX original


def _matches(memo: str, padrao: str) -> bool:
    """Checa se ``padrao`` aparece como substring case-insensitive em ``memo``."""
    if not padrao:
        return False
    return padrao.strip().upper() in (memo or "").upper()


def gerar_lancamentos_contabeis(
    pendentes_ofx: list[Transacao],
    regras: list[dict[str, Any]],
) -> list[LancamentoContabil]:
    """Pra cada pendente do OFX, busca a primeira regra cujo padrão case
    com o memo. Gera um LancamentoContabil com data, valor, banco do OFX
    + histórico (e regra) da configuração.

    Não modifica a lista original de pendentes — chamador decide o que fazer.
    """
    lancamentos: list[LancamentoContabil] = []
    for t in pendentes_ofx:
        memo = t.descricao or ""
        for regra in regras:
            padrao = (regra.get("padrao") or "").strip()
            if not padrao:
                continue
            if _matches(memo, padrao):
                lancamentos.append(LancamentoContabil(
                    data=t.data,
                    historico=(regra.get("historico") or "").strip(),
                    valor=t.valor,
                    banco=t.extras.get("banco", "") or "",
                    memo_original=memo,
                    padrao_match=padrao,
                    transacao_origem=t,
                ))
                break  # primeira regra que casa "ganha"
    return lancamentos
