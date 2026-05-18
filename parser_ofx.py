from decimal import Decimal
from pathlib import Path

from ofxparse import OfxParser

from parser_xlsx import Transacao


def ler_ofx(caminho: str | Path) -> tuple[list[Transacao], int]:
    """Lê o OFX retornando apenas pagamentos (valores negativos), convertidos
    para positivo para casarem com a planilha. Recebimentos são ignorados.

    Retorna ``(pagamentos, n_recebimentos_ignorados)``.
    """
    with open(caminho, "rb") as f:
        ofx = OfxParser.parse(f)

    transacoes: list[Transacao] = []
    ignorados = 0
    for conta in ofx.accounts:
        for t in conta.statement.transactions:
            valor = Decimal(str(t.amount))
            if valor >= 0:
                ignorados += 1
                continue
            data = t.date.date() if hasattr(t.date, "date") else t.date
            descricao = (t.memo or t.payee or "").strip()
            transacoes.append(Transacao(
                data=data, valor=-valor, descricao=descricao, origem="ofx",
            ))
    return transacoes, ignorados
