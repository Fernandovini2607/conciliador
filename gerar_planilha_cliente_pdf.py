"""Gera PLANILHA_CLIENTE.pdf — versao simples, sem detalhes tecnicos,
pro cliente final montar a planilha de controle dele.
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

FONT_DIR = Path("C:/Windows/Fonts")
pdfmetrics.registerFont(TTFont("Arial", str(FONT_DIR / "arial.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Bold", str(FONT_DIR / "arialbd.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Italic", str(FONT_DIR / "ariali.ttf")))
pdfmetrics.registerFont(TTFont("Arial-BoldItalic", str(FONT_DIR / "arialbi.ttf")))
pdfmetrics.registerFontFamily(
    "Arial", normal="Arial", bold="Arial-Bold",
    italic="Arial-Italic", boldItalic="Arial-BoldItalic",
)

styles = getSampleStyleSheet()

TITULO = ParagraphStyle(
    "Titulo", parent=styles["Title"], fontName="Arial-Bold",
    fontSize=22, leading=26, spaceAfter=4,
    textColor=colors.HexColor("#1f3a68"),
)
SUBTITULO = ParagraphStyle(
    "Subtitulo", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=11, leading=15, spaceAfter=20,
    textColor=colors.HexColor("#555555"), alignment=1,
)
H1 = ParagraphStyle(
    "H1", parent=styles["Heading1"], fontName="Arial-Bold",
    fontSize=15, leading=19, spaceBefore=18, spaceAfter=6,
    textColor=colors.HexColor("#1f3a68"),
)
TEXTO = ParagraphStyle(
    "Texto", parent=styles["Normal"], fontName="Arial",
    fontSize=11, leading=15, spaceAfter=8,
)
DESTAQUE = ParagraphStyle(
    "Destaque", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=10.5, leading=14, spaceAfter=8,
    textColor=colors.HexColor("#1f3a68"),
    leftIndent=10, rightIndent=10,
    borderPadding=8, borderWidth=0,
    backColor=colors.HexColor("#eef2fa"),
)


def bullets(itens: list[str]) -> ListFlowable:
    return ListFlowable(
        [ListItem(Paragraph(t, TEXTO), leftIndent=14) for t in itens],
        bulletType="bullet", start="•", leftIndent=14, bulletFontName="Arial",
    )


def tabela_colunas() -> Table:
    cabecalho = ["Coluna", "Precisa ter?", "O que preencher"]
    linhas = [
        ["Data de vencimento", "Sim",
         "Dia em que o pagamento vence"],
        ["Valor", "Sim",
         "Valor em reais do pagamento"],
        ["Data de pagamento", "Recomendado",
         "Dia em que o dinheiro realmente saiu do banco"],
        ["Nº da nota fiscal", "Recomendado",
         "Número da NF associada"],
        ["CNPJ do fornecedor", "Recomendado",
         "CNPJ de quem recebeu o pagamento"],
        ["Nome do fornecedor", "Recomendado",
         "Razão social ou nome fantasia"],
        ["Histórico", "Opcional",
         "Descrição / observação do pagamento"],
        ["Tipo / Categoria", "Opcional",
         "Classificação da despesa (ex: aluguel, luz, material)"],
    ]
    dados = [cabecalho] + [
        [Paragraph(f"<b>{c}</b>", TEXTO), p, Paragraph(oq, TEXTO)]
        for c, p, oq in linhas
    ]
    t = Table(dados, colWidths=[4.5 * cm, 3 * cm, 9 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 11),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def tabela_exemplo() -> Table:
    cabecalho = [
        "Vencimento", "Pagamento", "Valor", "NF",
        "CNPJ", "Fornecedor", "Histórico", "Tipo",
    ]
    linhas = [
        ["05/10/2026", "05/10/2026", "1.234,56", "12345",
         "12.345.678/0001-90", "ACME LTDA",
         "Comissão outubro", "Serviço"],
        ["10/10/2026", "11/10/2026", "500,00", "67890",
         "98.765.432/0001-10", "FORNECEDOR X",
         "Material escritório", "Material"],
        ["15/10/2026", "15/10/2026", "2.789,30", "11111",
         "11.222.333/0001-44", "ENERGIA SA",
         "Conta de luz", "Utilidade"],
    ]
    dados = [cabecalho] + linhas
    t = Table(dados, colWidths=[
        1.9 * cm, 1.9 * cm, 1.6 * cm, 1.2 * cm,
        2.6 * cm, 2.5 * cm, 2.8 * cm, 1.6 * cm,
    ])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("ALIGN", (3, 0), (3, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Arial", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawRightString(20 * cm, 1.2 * cm, f"Página {doc.page}")
    canvas.drawString(
        2 * cm, 1.2 * cm,
        "Janco Assessoria Contábil — Planilha de controle",
    )
    canvas.restoreState()


def construir() -> list:
    flow: list = []
    flow.append(Paragraph("Planilha de controle de pagamentos", TITULO))
    flow.append(Paragraph(
        "Como montar a sua planilha para enviar à contabilidade",
        SUBTITULO,
    ))

    flow.append(Paragraph(
        "Esta é a planilha de controle dos pagamentos da sua empresa. "
        "Nós usamos ela para conferir com o extrato bancário e com os "
        "lançamentos do sistema contábil, garantindo que tudo está "
        "batendo certinho.",
        TEXTO,
    ))
    flow.append(Paragraph(
        "Para funcionar bem, a planilha precisa ter algumas colunas. "
        "Veja abaixo quais são:",
        TEXTO,
    ))

    flow.append(Paragraph("Colunas da planilha", H1))
    flow.append(tabela_colunas())
    flow.append(Spacer(1, 10))

    flow.append(Paragraph(
        "<b>Quanto mais colunas preenchidas, melhor.</b> Com mais "
        "informação, o sistema consegue bater automaticamente mais "
        "pagamentos e você recebe os relatórios mais rápido, com "
        "menos ajustes manuais.",
        DESTAQUE,
    ))

    flow.append(Paragraph("Exemplo de como deve ficar", H1))
    flow.append(Paragraph(
        "Veja um exemplo de planilha preenchida (os valores são "
        "fictícios, só para ilustrar o formato):",
        TEXTO,
    ))
    flow.append(tabela_exemplo())
    flow.append(Spacer(1, 10))

    flow.append(Paragraph("Como preencher cada coluna", H1))
    flow.append(bullets([
        "<b>Datas</b> no formato <b>dd/mm/aaaa</b> (ex: 05/10/2026) ou "
        "formatadas como data no Excel.",
        "<b>Valor</b> em reais, com vírgula para centavos "
        "(ex: 1.234,56). Não precisa colocar o símbolo R$.",
        "<b>Nº da nota fiscal</b>: só o número, sem letras ou pontos.",
        "<b>CNPJ</b>: pode preencher com ou sem máscara "
        "(12.345.678/0001-90 ou 12345678000190 — os dois funcionam).",
        "<b>Fornecedor</b>: nome da empresa que recebeu o pagamento. "
        "A razão social é o ideal, mas o nome fantasia também vale.",
    ]))

    flow.append(Paragraph("Dicas importantes", H1))
    flow.append(bullets([
        "<b>A primeira linha da planilha é o cabeçalho</b> — ponha os "
        "nomes das colunas na linha 1 e os pagamentos a partir da "
        "linha 2.",
        "<b>Uma planilha por empresa.</b> Se a sua empresa tem matriz "
        "e filiais, faça uma planilha separada para cada uma.",
        "<b>Evite células mescladas</b> (aquelas que juntam várias "
        "colunas numa só) nas linhas de pagamento. Pode usar mescla "
        "só no título acima do cabeçalho, se quiser.",
        "<b>Valor em uma coluna só.</b> Se você hoje separa pagamento "
        "em colunas de \"débito\" e \"crédito\", prefira juntar numa "
        "coluna só \"Valor\".",
        "<b>Linhas em branco</b> no meio da planilha não atrapalham, "
        "podem ficar.",
        "<b>Salve como .xlsx</b> (formato padrão do Excel).",
    ]))

    flow.append(Paragraph("E se faltar alguma informação?", H1))
    flow.append(Paragraph(
        "Tudo bem — a planilha funciona mesmo que você só tenha "
        "<b>vencimento e valor</b>. As outras colunas são para melhorar "
        "a conferência automática. Se você já tem uma planilha pronta "
        "com nomes de coluna um pouco diferentes dos nomes aqui "
        "(ex: \"Val Pgto\" em vez de \"Valor\"), também não tem problema: "
        "o sistema deixa a gente ajustar isso na importação. O "
        "importante é ter as informações.",
        TEXTO,
    ))

    flow.append(Spacer(1, 14))
    flow.append(Paragraph(
        "Qualquer dúvida sobre como montar a planilha, pode falar com "
        "a gente — a Janco está à disposição.",
        DESTAQUE,
    ))

    return flow


def main() -> None:
    saida = Path(__file__).parent / "PLANILHA_CLIENTE.pdf"
    doc = SimpleDocTemplate(
        str(saida),
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=1.6 * cm, bottomMargin=2 * cm,
        title="Planilha de controle — como montar",
        author="Janco Assessoria Contábil",
    )
    doc.build(construir(), onFirstPage=_rodape, onLaterPages=_rodape)
    print(f"PDF gerado: {saida}")


if __name__ == "__main__":
    main()
