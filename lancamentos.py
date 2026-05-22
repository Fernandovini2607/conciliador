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
    """Gera lançamentos a partir dos pendentes do OFX casando o padrão da
    regra contra (memo + documento) da transação.

    Se a regra tem ``banco`` preenchido, exige que esse texto também esteja
    presente no nome do banco da transação OFX (case-insensitive). Útil
    quando bancos diferentes têm memos parecidos (ex.: "TARIFA DE
    MANUTENÇÃO") mas devem cair em contas contábeis diferentes.
    """
    lancamentos: list[LancamentoContabil] = []
    for t in pendentes_ofx:
        memo = t.descricao or ""
        documento = t.extras.get("documento", "") or ""
        banco_t = (t.extras.get("banco", "") or "").strip()
        # Padrão bate substring em memo OU documento (concatenados)
        campo_busca = f"{memo} {documento}"
        for regra in regras_memo:
            padrao = (regra.get("padrao") or "").strip()
            if not padrao:
                continue
            if not _matches(campo_busca, padrao):
                continue
            # Filtro adicional opcional: banco da regra precisa bater
            banco_regra = (regra.get("banco") or "").strip()
            if banco_regra and not _matches(banco_t, banco_regra):
                continue
            lancamentos.append(LancamentoContabil(
                data=t.data,
                historico=(regra.get("historico") or "").strip(),
                valor=t.valor,
                banco=banco_t,
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
    casando contra CNPJ, nome do fornecedor OU histórico da planilha."""
    lancamentos: list[LancamentoContabil] = []
    for par in pares_sem_dominio:
        cnpj = (par.planilha.extras.get("cnpj") or "").strip()
        fornecedor = (par.planilha.extras.get("fornecedor") or "").strip()
        historico = (par.planilha.extras.get("historico") or "").strip()
        # Padrão bate substring em CNPJ, fornecedor OU histórico (concatenados)
        campo_busca = f"{cnpj} {fornecedor} {historico}"
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


def _gerar_de_pendentes_planilha(
    pendentes_planilha: list[Transacao],
    regras_fornecedor: list[dict[str, Any]],
) -> list[LancamentoContabil]:
    """Aplica regras tipo ``fornecedor`` aos pendentes da PLANILHA que
    não casaram com nenhum OFX. Como não há banco, o lançamento sai com
    banco='Caixa geral'."""
    lancamentos: list[LancamentoContabil] = []
    for t in pendentes_planilha:
        cnpj = (t.extras.get("cnpj") or "").strip()
        fornecedor = (t.extras.get("fornecedor") or "").strip()
        historico = (t.extras.get("historico") or "").strip()
        campo_busca = f"{cnpj} {fornecedor} {historico}"
        for regra in regras_fornecedor:
            padrao = (regra.get("padrao") or "").strip()
            if not padrao:
                continue
            if _matches(campo_busca, padrao):
                lancamentos.append(LancamentoContabil(
                    # Data: prioriza data de pagamento da planilha, senão vencimento
                    data=t.data_pagamento or t.data,
                    historico=(regra.get("historico") or "").strip(),
                    valor=t.valor,
                    banco="Caixa geral",  # sem OFX correspondente
                    memo_original="",
                    padrao_match=padrao,
                    conta=(regra.get("conta") or "").strip(),
                    tipo_regra="fornecedor_planilha",
                    fornecedor=fornecedor,
                    cnpj=cnpj,
                    transacao_origem=t,
                    par_origem=None,
                ))
                break
    return lancamentos


def gerar_lancamentos_contabeis(
    pendentes_ofx: list[Transacao],
    regras: list[dict[str, Any]],
    pares_sem_dominio: list["Par"] | None = None,
    pendentes_planilha: list[Transacao] | None = None,
) -> list[LancamentoContabil]:
    """Pipeline completo: aplica regras a TRÊS fontes:
    - ``pendentes_ofx`` → regras ``memo``
    - ``pares_sem_dominio`` → regras ``fornecedor`` (com OFX casado)
    - ``pendentes_planilha`` → regras ``fornecedor`` (sem OFX, vira Caixa geral)

    Retrocompatibilidade: chamadas antigas passando só os 2 ou 3 primeiros
    args continuam funcionando.
    """
    pares_sem_dominio = pares_sem_dominio or []
    pendentes_planilha = pendentes_planilha or []
    regras_memo = [r for r in regras if _tipo_regra(r) == "memo"]
    regras_fornecedor = [r for r in regras if _tipo_regra(r) == "fornecedor"]
    return (
        _gerar_de_pendentes_ofx(pendentes_ofx, regras_memo)
        + _gerar_de_pares_sem_dominio(pares_sem_dominio, regras_fornecedor)
        + _gerar_de_pendentes_planilha(pendentes_planilha, regras_fornecedor)
    )
