from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from parser_xlsx import Transacao


@dataclass
class Par:
    """Par de transações conciliadas (auto ou manual) ou sugerido.

    ``dominio`` é preenchido na segunda fase (Conciliados × Domínio) quando
    existe um lançamento equivalente no Domínio (data_venc + valor + NF).
    """
    planilha: Transacao
    ofx: Transacao
    tipo: str = "auto"  # "auto", "manual", "sugestao"
    diff_dias: int = 0
    diff_valor: Decimal = Decimal("0")
    dominio: Transacao | None = None


@dataclass
class Resultado:
    conciliados: list[Par] = field(default_factory=list)
    pendentes_planilha: list[Transacao] = field(default_factory=list)
    pendentes_ofx: list[Transacao] = field(default_factory=list)
    sugestoes: list[Par] = field(default_factory=list)


TOLERANCIA_DIAS = 2
TOLERANCIA_VALOR = Decimal("10.00")


def _quantizar(v: Decimal) -> Decimal:
    return v.quantize(Decimal("0.01"))


def _chave(t: Transacao) -> tuple[date, Decimal]:
    """Chave padrão usando t.data — vencimento (planilha/Domínio) ou
    data da transação (OFX)."""
    return (t.data, _quantizar(t.valor))


def _data_match_ofx(t: Transacao) -> date:
    """Data a usar quando comparando com OFX. Para a planilha usa
    data_pagamento (se mapeada); para OFX/Domínio usa t.data."""
    return t.data_pagamento if t.data_pagamento is not None else t.data


def _chave_ofx(t: Transacao) -> tuple[date, Decimal]:
    return (_data_match_ofx(t), _quantizar(t.valor))


def conciliar_automatico(
    planilha: list[Transacao],
    ofx: list[Transacao],
) -> tuple[list[Par], list[Transacao], list[Transacao]]:
    """Match exato por (data, valor) — onde "data" da planilha é a
    data_pagamento (se mapeada) e do OFX é a data de compensação.
    Devolve pares + pendentes de cada lado."""
    indice_ofx: dict[tuple[date, Decimal], list[Transacao]] = defaultdict(list)
    for t in ofx:
        indice_ofx[_chave_ofx(t)].append(t)

    pares: list[Par] = []
    pendentes_p: list[Transacao] = []

    for t in planilha:
        candidatos = indice_ofx.get(_chave_ofx(t))
        if candidatos:
            par = candidatos.pop(0)
            pares.append(Par(planilha=t, ofx=par, tipo="auto"))
        else:
            pendentes_p.append(t)

    pendentes_o: list[Transacao] = []
    for restantes in indice_ofx.values():
        pendentes_o.extend(restantes)

    return pares, pendentes_p, pendentes_o


def gerar_sugestoes(
    pendentes_planilha: list[Transacao],
    pendentes_ofx: list[Transacao],
    dias_tol: int = TOLERANCIA_DIAS,
    valor_tol: Decimal = TOLERANCIA_VALOR,
) -> list[Par]:
    """Pares onde data difere até `dias_tol` E valor difere até `valor_tol`.

    Filtra pelos dois critérios simultaneamente — o match exato já foi feito
    antes, então os pares aqui sempre têm diff_dias>0 ou diff_valor>0.
    Ordena pelo "mais próximo" (soma ponderada de diferenças).
    """
    sugestoes: list[Par] = []
    for tp in pendentes_planilha:
        data_p = _data_match_ofx(tp)
        for to in pendentes_ofx:
            diff_dias = abs((data_p - to.data).days)
            diff_valor = abs(_quantizar(tp.valor) - _quantizar(to.valor))
            if diff_dias <= dias_tol and diff_valor <= valor_tol:
                sugestoes.append(Par(
                    planilha=tp,
                    ofx=to,
                    tipo="sugestao",
                    diff_dias=diff_dias,
                    diff_valor=diff_valor,
                ))
    sugestoes.sort(key=lambda p: (p.diff_dias, p.diff_valor))
    return sugestoes


def conciliar_completo(
    planilha: list[Transacao],
    ofx: list[Transacao],
    dias_tol: int = TOLERANCIA_DIAS,
    valor_tol: Decimal = TOLERANCIA_VALOR,
) -> Resultado:
    pares, pendentes_p, pendentes_o = conciliar_automatico(planilha, ofx)
    sugestoes = gerar_sugestoes(pendentes_p, pendentes_o, dias_tol, valor_tol)
    return Resultado(
        conciliados=pares,
        pendentes_planilha=pendentes_p,
        pendentes_ofx=pendentes_o,
        sugestoes=sugestoes,
    )


def diferenca(p: Transacao, o: Transacao) -> tuple[int, Decimal]:
    """Diferença entre planilha e OFX — usa data_pagamento da planilha
    quando disponível (mesma lógica do match)."""
    data_p = _data_match_ofx(p)
    return (
        abs((data_p - o.data).days),
        abs(_quantizar(p.valor) - _quantizar(o.valor)),
    )
