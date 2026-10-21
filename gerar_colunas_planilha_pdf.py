"""Gera COLUNAS_PLANILHA.pdf a partir de COLUNAS_PLANILHA.md.

Documento curto pra entregar ao cliente/operador com as colunas que
a planilha de controle precisa ter pra ser usada no Conciliador.
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
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Fontes Arial do Windows
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
    fontSize=22, leading=26, spaceAfter=4, textColor=colors.HexColor("#1f3a68"),
)
SUBTITULO = ParagraphStyle(
    "Subtitulo", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=11, leading=15, spaceAfter=18, textColor=colors.HexColor("#555555"),
    alignment=1,
)
H1 = ParagraphStyle(
    "H1", parent=styles["Heading1"], fontName="Arial-Bold",
    fontSize=15, leading=19, spaceBefore=16, spaceAfter=6,
    textColor=colors.HexColor("#1f3a68"),
)
H2 = ParagraphStyle(
    "H2", parent=styles["Heading2"], fontName="Arial-Bold",
    fontSize=12, leading=16, spaceBefore=10, spaceAfter=4,
    textColor=colors.HexColor("#333333"),
)
TEXTO = ParagraphStyle(
    "Texto", parent=styles["Normal"], fontName="Arial",
    fontSize=10, leading=14, spaceAfter=6,
)
NOTA = ParagraphStyle(
    "Nota", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=9.5, leading=13, spaceAfter=6,
    textColor=colors.HexColor("#1f3a68"),
    leftIndent=10, rightIndent=10,
    borderPadding=6, borderWidth=0,
    backColor=colors.HexColor("#eef2fa"),
)


def bullets(itens: list[str]) -> ListFlowable:
    return ListFlowable(
        [ListItem(Paragraph(t, TEXTO), leftIndent=14) for t in itens],
        bulletType="bullet", start="•", leftIndent=14, bulletFontName="Arial",
    )


def tabela_resumo() -> Table:
    cabecalho = ["Campo", "Obrigatório?", "Para que serve"]
    linhas = [
        ["Data de vencimento", "✅ Sim",
         "Chave de match com OFX e Domínio"],
        ["Valor", "✅ Sim",
         "Chave de match (valor pago/devido)"],
        ["Data de pagamento", "Opcional",
         "Match preciso com OFX (dia que o dinheiro saiu do banco)"],
        ["Nº da nota fiscal (NF)", "Opcional",
         "Match forte com o Domínio (triple-match)"],
        ["CNPJ do fornecedor", "Opcional",
         "Match por fornecedor quando NF/data divergem"],
        ["Nome do fornecedor", "Opcional",
         "Match por nome quando CNPJ falta"],
        ["Histórico contábil", "Opcional",
         "Histórico padrão dos lançamentos contábeis"],
        ["Tipo / Categoria", "Opcional",
         "Auxílio de classificação por fornecedor"],
        ["Data de emissão", "Opcional",
         "Data da nota fiscal (só exibição em relatórios)"],
    ]
    dados = [cabecalho] + [
        [c, ob, Paragraph(ps, TEXTO)] for c, ob, ps in linhas
    ]
    t = Table(dados, colWidths=[4.6 * cm, 2.6 * cm, 9.2 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def tabela_exemplo_completo() -> Table:
    cabecalho = [
        "Vencimento", "Pagamento", "Valor", "NF", "CNPJ",
        "Fornecedor", "Histórico", "Tipo",
    ]
    linhas = [
        ["05/10/2026", "05/10/2026", "1.234,56", "12345",
         "12.345.678/0001-90", "ACME LTDA", "Comissão outubro", "Serviço"],
        ["10/10/2026", "11/10/2026", "500,00", "67890",
         "98.765.432/0001-10", "FORNECEDOR X", "Material escritório",
         "Material"],
        ["15/10/2026", "", "2.789,30", "11111",
         "11.222.333/0001-44", "ENERGIA SA", "Conta de luz", "Utilidade"],
    ]
    dados = [cabecalho] + linhas
    t = Table(dados, colWidths=[
        1.8 * cm, 1.8 * cm, 1.6 * cm, 1.2 * cm, 2.6 * cm,
        2.6 * cm, 2.8 * cm, 1.8 * cm,
    ])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 7.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("ALIGN", (3, 0), (3, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def tabela_exemplo_minimo() -> Table:
    cabecalho = ["Vencimento", "Valor"]
    linhas = [
        ["05/10/2026", "1.234,56"],
        ["10/10/2026", "500,00"],
        ["15/10/2026", "2.789,30"],
    ]
    dados = [cabecalho] + linhas
    t = Table(dados, colWidths=[4 * cm, 3 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 9.5),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Arial", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawRightString(
        20 * cm, 1.2 * cm,
        f"Página {doc.page}",
    )
    canvas.drawString(
        2 * cm, 1.2 * cm,
        "Conciliador — Colunas da planilha de controle",
    )
    canvas.restoreState()


def construir() -> list:
    flow: list = []
    flow.append(Paragraph(
        "Colunas da planilha de controle", TITULO,
    ))
    flow.append(Paragraph(
        "O que a planilha .xlsx precisa ter pra ser importada "
        "no Conciliador", SUBTITULO,
    ))

    flow.append(Paragraph(
        "O Conciliador aceita qualquer planilha Excel (.xlsx), "
        "desde que ela tenha pelo menos <b>duas colunas obrigatórias</b> "
        "— data de vencimento e valor. As demais são opcionais, mas "
        "quanto mais informação a planilha trouxer, mais pagamentos "
        "casam automaticamente com o OFX bancário e com o Domínio.",
        TEXTO,
    ))
    flow.append(Spacer(1, 6))

    flow.append(Paragraph("Resumo das colunas", H1))
    flow.append(tabela_resumo())

    # Obrigatórias
    flow.append(Paragraph("Colunas obrigatórias", H1))

    flow.append(Paragraph("1. Data de vencimento", H2))
    flow.append(Paragraph(
        "Data em que o pagamento deveria ocorrer (ou efetivamente "
        "ocorreu, se a planilha só traz uma data). O sistema aceita "
        "qualquer um destes nomes de coluna no cabeçalho:",
        TEXTO,
    ))
    flow.append(bullets([
        "Data / Date / Dt / Dia",
        "Data lançamento / Data do lançamento",
        "Data movimento / Data mov",
        "<b>Vencimento</b> / Data vencimento / Data venc / Venc",
        "Data pagamento / Data pgto / Data operação",
        "Data crédito / Data débito / Data compensação",
    ]))
    flow.append(Paragraph(
        "<b>Formato aceito:</b> dd/mm/aaaa, dd-mm-aaaa, aaaa-mm-dd, "
        "dd.mm.aaaa ou célula já formatada como data no Excel.",
        TEXTO,
    ))

    flow.append(Paragraph("2. Valor", H2))
    flow.append(Paragraph(
        "O valor em reais do pagamento. Nomes de coluna aceitos:",
        TEXTO,
    ))
    flow.append(bullets([
        "<b>Valor</b> / Vlr / Vl / Amount / Value",
        "Valor (R$) / Valor R$ / R$",
        "Valor lançamento / Valor parcela / Valor da parcela",
        "Total / Montante / Quantia",
        "Crédito / Débito / Entrada / Saída",
    ]))
    flow.append(Paragraph(
        "<b>Formato aceito:</b> número puro (1234.56 ou 1234,56), "
        "com separador de milhar (1.234,56 ou 1,234.56), entre "
        "parênteses para negativo ((123,45) vira -123,45), ou com "
        "símbolo R$ (ignorado automaticamente).",
        TEXTO,
    ))

    flow.append(PageBreak())

    # Opcionais
    flow.append(Paragraph("Colunas opcionais (recomendadas)", H1))
    flow.append(Paragraph(
        "Essas colunas não são obrigatórias, mas <b>enriquecem a "
        "conciliação</b> — o sistema consegue casar automaticamente "
        "muito mais pagamentos quando elas estão presentes.",
        TEXTO,
    ))

    flow.append(Paragraph("3. Data de pagamento", H2))
    flow.append(Paragraph(
        "Dia em que o dinheiro efetivamente saiu do banco. Muito "
        "recomendada quando a planilha controla vencimento e "
        "pagamento em colunas separadas — permite o match exato "
        "com o OFX bancário.",
        TEXTO,
    ))
    flow.append(bullets([
        "Data pagamento / Data pgto / Data de pagamento",
        "Pagamento / Pago em",
        "Data baixa / Data quitação / Data liquidação",
    ]))

    flow.append(Paragraph("4. Nº da nota fiscal (NF)", H2))
    flow.append(Paragraph(
        "Essencial para o <b>triple-match</b> com o Domínio "
        "(data + valor + NF). Sem a NF, o sistema cai pros matches "
        "aproximados, menos precisos.",
        TEXTO,
    ))
    flow.append(bullets([
        "NF / N NF / Nº NF / Número NF",
        "Número da NF / Num NF / N Nota / Nº Nota",
        "Nota / Nota fiscal / Número nota fiscal",
        "Número / Numero documento / Doc",
    ]))

    flow.append(Paragraph("5. CNPJ do fornecedor", H2))
    flow.append(bullets([
        "CNPJ / CGC / CNPJ/CPF",
        "CNPJ fornecedor / CNPJ do fornecedor",
        "Documento fornecedor / Doc fornecedor",
    ]))
    flow.append(Paragraph(
        "Com ou sem máscara (12.345.678/0001-90 ou 12345678000190).",
        TEXTO,
    ))

    flow.append(Paragraph("6. Nome do fornecedor / razão social", H2))
    flow.append(bullets([
        "Fornecedor / Nome fornecedor / Nome do fornecedor",
        "Razão social / Razão / Razao",
        "Beneficiário / Favorecido / Credor",
        "Pagamento a / Para / Destinatário",
    ]))

    flow.append(Paragraph("7. Histórico contábil", H2))
    flow.append(Paragraph(
        "Texto descritivo que vira o histórico padrão do lançamento "
        "contábil gerado pelas regras de fornecedor.",
        TEXTO,
    ))
    flow.append(bullets([
        "Histórico / Hist / Histórico contábil",
        "Descrição / Complemento / Obs / Observação",
        "Detalhes / Narrativa",
    ]))

    flow.append(Paragraph("8. Tipo / Categoria", H2))
    flow.append(Paragraph(
        "Classificação da despesa. Usado como critério auxiliar "
        "no match por fornecedor quando mais de um pagamento tem "
        "o mesmo valor na mesma data.",
        TEXTO,
    ))
    flow.append(bullets([
        "Tipo / Tipo de despesa / Tipo lançamento",
        "Classificação / Categoria / Grupo",
        "Natureza / Natureza da despesa",
    ]))

    flow.append(Paragraph("9. Data de emissão da nota", H2))
    flow.append(bullets([
        "Data emissão / Emissão",
        "Data NF / Data da nota / Data documento",
    ]))

    flow.append(PageBreak())

    # Exemplos
    flow.append(Paragraph("Exemplo de planilha mínima", H1))
    flow.append(Paragraph(
        "Essa é a planilha mais simples que funciona — só com as "
        "duas colunas obrigatórias:",
        TEXTO,
    ))
    flow.append(tabela_exemplo_minimo())
    flow.append(Spacer(1, 6))
    flow.append(Paragraph(
        "A conciliação vai depender só de data + valor. Se dois "
        "pagamentos têm o mesmo valor na mesma data, o sistema pode "
        "pedir confirmação manual.",
        NOTA,
    ))

    flow.append(Paragraph("Exemplo de planilha completa (recomendada)", H1))
    flow.append(Paragraph(
        "Com todas as colunas, o Conciliador consegue aplicar as "
        "6 fases de match com o Domínio e maximiza o número de "
        "pagamentos conciliados automaticamente:",
        TEXTO,
    ))
    flow.append(tabela_exemplo_completo())
    flow.append(Spacer(1, 10))

    # Boas práticas
    flow.append(Paragraph("Boas práticas", H1))
    flow.append(bullets([
        "<b>Primeira linha deve ser o cabeçalho</b> com os nomes das "
        "colunas. O sistema identifica cada coluna pelo nome.",
        "<b>Não misture pagamentos de empresas diferentes</b> na mesma "
        "planilha. Em grupo empresarial, cada filial tem a sua.",
        "<b>Evite células mescladas (merge)</b> nas linhas de dados — "
        "bagunçam a leitura. Pode haver mescla no título acima do "
        "cabeçalho.",
        "<b>Formate como data</b> as células de data (não como texto). "
        "Se estiver como texto, use dd/mm/aaaa.",
        "<b>Valor em uma única coluna.</b> Se a planilha tem "
        "Débito/Crédito em colunas separadas, consolide antes de "
        "importar (ou use só uma delas se os pagamentos forem sempre "
        "do mesmo lado).",
        "<b>Linhas em branco</b> no meio da planilha são ignoradas.",
        "<b>PDFs de comprovantes</b> (Sicoob/Bradesco e PIX Sicoob) "
        "podem complementar ou substituir a planilha — o sistema "
        "extrai data, valor, NF, CNPJ e fornecedor direto do PDF.",
    ]))

    flow.append(Paragraph(
        "Se o sistema não achar as colunas automaticamente", H1,
    ))
    flow.append(Paragraph(
        "O Conciliador tenta identificar as colunas pelo nome do "
        "cabeçalho. Se alguma coluna estiver com nome não-padrão "
        "(ex.: <i>Val Pgto</i> no lugar de <i>Valor</i>), ao importar "
        "a planilha o sistema abre um diálogo <b>Editar colunas</b> "
        "onde o operador associa manualmente cada campo à coluna "
        "correta da planilha. O mapeamento é salvo <b>por empresa</b>, "
        "então só precisa ser feito uma vez.",
        TEXTO,
    ))

    return flow


def main() -> None:
    saida = Path(__file__).parent / "COLUNAS_PLANILHA.pdf"
    doc = SimpleDocTemplate(
        str(saida),
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=1.6 * cm, bottomMargin=2 * cm,
        title="Colunas da planilha de controle — Conciliador",
        author="Janco Assessoria Contábil",
    )
    doc.build(construir(), onFirstPage=_rodape, onLaterPages=_rodape)
    print(f"PDF gerado: {saida}")


if __name__ == "__main__":
    main()
