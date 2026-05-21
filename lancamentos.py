"""Geração de lançamentos contábeis automáticos a partir de regras.

Duas origens:
- ``pendentes_ofx`` (Transacao) → regras tipo ``memo`` (tarifas, IOF, juros).
- ``pares_sem_dominio`` (Par) → regras tipo ``fornecedor`` (CNPJ ou nome).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import TYPE_CHECKING, Any

from parser_xlsx import Transacao

if TYPE_CHECKING:
    from matcher import Par


@dataclass
class LancamentoContabil:
    data: date                                # data de pagamento (sempre do OFX = data de compensação no banco)
    historico: str                            # vem da regra
    valor: Decimal                            # do OFX
    banco: str                                # do OFX
    memo_original: str                        # memo bruto do OFX (referência)
    padrao_match: str                         # padrão da regra
    conta: str = ""                           # código contábil da regra
    tipo_regra: str = "memo"                  # "memo" ou "fornecedor"
    fornecedor: str = ""                      # nome do fornecedor (quando aplicável)
    cnpj: str = ""                            # CNPJ (quando aplicável)
    transacao_origem: Transacao | None = None # ref ao OFX original
    par_origem: "Par | None" = None           # ref ao par P×O (quando origem=fornecedor)


def _matches(texto: str, padrao: str) -> bool:
    if not padrao:
        return False
    return padrao.strip().upper() in (texto or "").upper()


def _tipo_regra(regra: dict[str, Any]) -> str:
    """Retorna o tipo da regra (default 'memo' para retrocompatibilidade)."""
    return (regra.get("tipo") or "memo").strip().lower()


def _gerar_de_pendentes_ofx(
    pendentes_ofx: list[Transacao],
    regras_memo: list[dict[str, Any]],
) -> list[LancamentoContabil]:
    """Gera lançamentos a partir dos pendentes do OFX casando memo."""
    lancamentos: list[LancamentoContabil] = []
    for t in pendentes_ofx:
        memo = t.descricao or ""
        for regra in regras_memo:
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
                    conta=(regra.get("conta") or "").strip(),
                    tipo_regra="memo",
                    transacao_origem=t,
                ))
                break
    return lancamentos


def _gerar_de_pares_sem_dominio(
    pares_sem_dominio: list["Par"],
    regras_fornecedor: list[dict[str, Any]],
) -> list[LancamentoContabil]:
    """Gera lançamentos a partir dos pares P×OFX que faltam no Domínio,
    casando contra CNPJ ou nome do fornecedor da planilha."""
    lancamentos: list[LancamentoContabil] = []
    for par in pares_sem_dominio:
        cnpj = (par.planilha.extras.get("cnpj") or "").strip()
        fornecedor = (par.planilha.extras.get("fornecedor") or "").strip()
        campo_busca = f"{cnpj} {fornecedor}"
        for regra in regras_fornecedor:
            padrao = (regra.get("padrao") or "").strip()
            if not padrao:
                continue
            if _matches(campo_busca, padrao):
                # Data de pagamento: sempre a do OFX (data de compensação
                # efetiva no banco, é o que vai no lançamento contábil).
                lancamentos.append(LancamentoContabil(
                    data=par.ofx.data,
                    historico=(regra.get("historico") or "").strip(),
                    valor=par.planilha.valor,
                    banco=par.ofx.extras.get("banco", "") or "",
                    memo_original=par.ofx.descricao or "",
                    padrao_match=padrao,
                    conta=(regra.get("conta") or "").strip(),
                    tipo_regra="fornecedor",
                    fornecedor=fornecedor,
                    cnpj=cnpj,
                    transacao_origem=par.ofx,
                    par_origem=par,
                ))
                break
    return lancamentos


def gerar_lancamentos_contabeis(
    pendentes_ofx: list[Transacao],
    regras: list[dict[str, Any]],
    pares_sem_dominio: list["Par"] | None = None,
) -> list[LancamentoContabil]:
    """Pipeline completo: aplica regras tipo ``memo`` aos pendentes do OFX e
    regras tipo ``fornecedor`` aos pares conciliados que faltam no Domínio.

    Retrocompatibilidade: chamadas antigas passando só os 2 primeiros args
    continuam funcionando (pares_sem_dominio default []).
    """
    pares_sem_dominio = pares_sem_dominio or []
    regras_memo = [r for r in regras if _tipo_regra(r) == "memo"]
    regras_fornecedor = [r for r in regras if _tipo_regra(r) == "fornecedor"]
    return (
        _gerar_de_pendentes_ofx(pendentes_ofx, regras_memo)
        + _gerar_de_pares_sem_dominio(pares_sem_dominio, regras_fornecedor)
    )
