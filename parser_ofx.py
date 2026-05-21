from decimal import Decimal
from pathlib import Path

from ofxparse import OfxParser

from parser_xlsx import Transacao


def _identificar_banco(conta, fallback: str) -> str:
    """Tenta identificar o banco a partir dos campos do OFX. Cai no nome do
    arquivo quando o OFX não traz a informação."""
    # Tenta institution.organization (mais comum em OFXs brasileiros)
    try:
        if conta.institution and conta.institution.organization:
            return str(conta.institution.organization).strip()
    except AttributeError:
        pass
    # Tenta routing_number / bank_id
    for attr in ("routing_number", "bank_id", "branch_id"):
        try:
            v = getattr(conta, attr, None)
            if v:
                return f"{attr}={v}"
        except AttributeError:
            continue
    # Tenta account_id
    try:
        if conta.account_id:
            return f"conta {conta.account_id}"
    except AttributeError:
        pass
    return fallback


def ler_ofx(caminho: str | Path) -> tuple[list[Transacao], int]:
    """Lê o OFX retornando apenas pagamentos (valores negativos), convertidos
    para positivo para casarem com a planilha. Recebimentos são ignorados.

    Adiciona em ``extras`` o banco identificado e o nome do arquivo de origem.
    Retorna ``(pagamentos, n_recebimentos_ignorados)``.
    """
    arquivo = Path(caminho).stem  # nome do arquivo sem extensão
    with open(caminho, "rb") as f:
        ofx = OfxParser.parse(f)

    transacoes: list[Transacao] = []
    ignorados = 0
    for conta in ofx.accounts:
        banco = _identificar_banco(conta, fallback=arquivo)
        for t in conta.statement.transactions:
            valor = Decimal(str(t.amount))
            if valor >= 0:
                ignorados += 1
                continue
            data = t.date.date() if hasattr(t.date, "date") else t.date
            descricao = (t.memo or t.payee or "").strip()
            transacoes.append(Transacao(
                data=data, valor=-valor, descricao=descricao, origem="ofx",
                extras={"banco": banco, "arquivo": arquivo},
            ))
    return transacoes, ignorados
