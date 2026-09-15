"""Gera MANUAL_USO.pdf a partir do conteúdo do MANUAL_USO.md.

PDF simples para operadores — sem detalhes técnicos, foco em uso.
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
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
    fontSize=24, leading=28, spaceAfter=4, textColor=colors.HexColor("#1f3a68"),
)
SUBTITULO = ParagraphStyle(
    "Subtitulo", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=12, leading=16, spaceAfter=20, textColor=colors.HexColor("#555555"),
    alignment=1,
)
PASSO = ParagraphStyle(
    "Passo", parent=styles["Heading1"], fontName="Arial-Bold",
    fontSize=16, leading=20, spaceBefore=18, spaceAfter=4,
    textColor=colors.HexColor("#1f3a68"),
)
H2 = ParagraphStyle(
    "H2", parent=styles["Heading2"], fontName="Arial-Bold",
    fontSize=13, leading=16, spaceBefore=12, spaceAfter=4,
    textColor=colors.HexColor("#2a4d7d"),
)
H3 = ParagraphStyle(
    "H3", parent=styles["Heading3"], fontName="Arial-Bold",
    fontSize=11, leading=14, spaceBefore=8, spaceAfter=3,
    textColor=colors.HexColor("#333333"),
)
TEXTO = ParagraphStyle(
    "Texto", parent=styles["Normal"], fontName="Arial",
    fontSize=10, leading=14, alignment=TA_LEFT, spaceAfter=4,
)
TEXTO_BULLET = ParagraphStyle(
    "Bullet", parent=TEXTO, leftIndent=14, bulletIndent=0, spaceAfter=2,
)
NOTA = ParagraphStyle(
    "Nota", parent=TEXTO, fontName="Arial-Italic",
    textColor=colors.HexColor("#666666"), leftIndent=14, borderPadding=4,
)
CODIGO = ParagraphStyle(
    "Codigo", parent=styles["Normal"], fontName="Courier",
    fontSize=9, leading=12, leftIndent=14, spaceAfter=6,
    textColor=colors.HexColor("#333333"), backColor=colors.HexColor("#f0f2f5"),
    borderPadding=6,
)


def bullets(itens):
    return ListFlowable(
        [ListItem(Paragraph(t, TEXTO_BULLET), leftIndent=14) for t in itens],
        bulletType="bullet", start="•", bulletColor=colors.HexColor("#1f3a68"),
        leftIndent=8, bulletFontSize=10,
    )


def tabela_simples(cabecalho, linhas, larguras=None):
    """Cria tabela azul padrão."""
    dados = [cabecalho] + [
        [Paragraph(str(c), TEXTO) if not str(c).startswith("<b>") else Paragraph(c, TEXTO)
         for c in linha]
        for linha in linhas
    ]
    t = Table(dados, colWidths=larguras)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def caixa_dica(texto):
    """Caixa colorida com texto de dica/nota."""
    t = Table([[Paragraph(f"<b>💡 Dica:</b> {texto}", TEXTO)]],
              colWidths=[16 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fff8e1")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#f0c243")),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return t


def _rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Arial", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawString(2 * cm, 1.2 * cm, "Manual do Conciliador — uso diário")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Página {doc.page}")
    canvas.restoreState()


def construir():
    flow = []

    # ============ Capa
    flow.append(Spacer(1, 3 * cm))
    flow.append(Paragraph("Manual do Conciliador", TITULO))
    flow.append(Paragraph(
        "Passo a passo para uso diário — Janco Assessoria Contábil",
        SUBTITULO,
    ))
    flow.append(Spacer(1, 2 * cm))
    flow.append(Paragraph(
        "Este manual explica como usar o sistema Conciliador no dia a dia.<br/>"
        "Não é sobre <b>instalar</b> — é sobre <b>usar</b> depois de instalado.",
        TEXTO,
    ))
    flow.append(Spacer(1, 0.5 * cm))
    flow.append(caixa_dica(
        "Se o sistema ainda não está instalado na sua máquina, "
        "fale com o administrador antes de continuar."
    ))
    flow.append(PageBreak())

    # ============ Passo 1
    flow.append(Paragraph("Passo 1 — Entrar no sistema", PASSO))
    flow.append(Paragraph("Como abrir o app e fazer login pela primeira vez:", TEXTO))
    flow.append(bullets([
        "Duplo-clique no atalho <b>Conciliador</b> na área de trabalho.",
        "Digite seu <b>usuário</b> e <b>senha</b> (o administrador te forneceu).",
        "Clique em <b>Entrar</b> (ou aperte Enter).",
    ]))
    flow.append(Spacer(1, 4))
    flow.append(caixa_dica(
        "Primeira vez usando? Clique em <b>Minha senha</b> (barra do topo) "
        "e troque a senha inicial pela sua senha pessoal."
    ))

    # ============ Passo 2
    flow.append(Paragraph("Passo 2 — Selecionar a empresa", PASSO))
    flow.append(bullets([
        "Na barra do topo, clique em <b>Selecionar empresa</b>.",
        "Escolha a empresa da lista e clique em <b>OK</b>.",
        "Se aparecer a pergunta <i>“Trocar de empresa? Os dados carregados "
        "serão descartados”</i>, clique em <b>Sim</b>. É normal — cada "
        "empresa trabalha com dados próprios.",
    ]))
    flow.append(Spacer(1, 4))
    flow.append(caixa_dica(
        "A empresa que você escolher fica salva na sua conta. Da próxima vez "
        "que abrir o app, vai voltar direto pra ela."
    ))

    # ============ Passo 3
    flow.append(Paragraph("Passo 3 — Carregar os dados da empresa", PASSO))
    flow.append(Paragraph(
        "Você vai trabalhar com <b>até 3 fontes de dados</b> ao mesmo tempo:",
        TEXTO,
    ))

    flow.append(Paragraph("3.1 Planilha de contas a pagar (opcional)", H3))
    flow.append(bullets([
        "Clique em <b>Abrir planilha (.xlsx)</b> e escolha o arquivo.",
        "Se for a primeira vez com essa empresa, marque qual coluna é o quê:",
        "&nbsp;&nbsp;<b>Vencimento</b> (obrigatório)",
        "&nbsp;&nbsp;<b>Valor</b> (obrigatório)",
        "&nbsp;&nbsp;Outros: Pagamento, Fornecedor, CNPJ, Nº NF, Histórico, Tipo",
        "&nbsp;&nbsp;Se a planilha não tiver algum campo, deixe <b>(deixar vazia)</b>.",
        "Clique em <b>Confirmar</b>.",
    ]))
    flow.append(caixa_dica(
        "Da próxima vez que abrir uma planilha dessa mesma empresa, o "
        "sistema já lembra o mapeamento — vai direto sem perguntar."
    ))

    flow.append(Paragraph(
        "3.1.b Comprovantes PDF de pagamento (alternativa)", H3,
    ))
    flow.append(Paragraph(
        "Se a empresa não tem planilha de controle mas você tem "
        "comprovantes de boleto em PDF baixados do internet banking:",
        TEXTO,
    ))
    flow.append(bullets([
        "Clique em <b>Importar comprovantes PDF</b>.",
        "Selecione um ou mais arquivos (Ctrl+clique).",
        "<b>Bancos suportados</b>: Sicoob e Bradesco.",
        "Aguarde o processamento — rápido em PDFs pequenos, alguns "
        "segundos em PDFs de 100+ páginas.",
        "Os comprovantes viram uma \"planilha virtual\" — mesmas colunas, "
        "mesmo fluxo de conciliação.",
    ]))
    flow.append(caixa_dica(
        "Pode combinar planilha .xlsx + comprovantes PDF. O sistema "
        "detecta duplicatas e importa cada lançamento uma vez só. O "
        "popup avisa quantos entraram e quantos foram ignorados. "
        "Comprovantes PDF trazem CNPJ + nome, o que ajuda muito a "
        "conciliação com o Domínio."
    ))

    flow.append(Paragraph("3.2 Extrato bancário OFX", H3))
    flow.append(bullets([
        "Clique em <b>Importar OFX</b>.",
        "Escolha <b>um ou mais arquivos .OFX</b> (Ctrl+clique pra selecionar "
        "vários bancos de uma vez).",
        "Confirme.",
    ]))

    flow.append(Paragraph("3.3 Domínio Contábil", H3))
    flow.append(bullets([
        "Clique em <b>Carregar pagamentos</b>.",
        "Aguarde ele buscar as parcelas do Domínio da empresa.",
    ]))
    flow.append(caixa_dica(
        "<b>Empresa matriz com filiais?</b> O sistema detecta "
        "automaticamente pelo <b>CNPJ raiz</b> (primeiros 8 dígitos) e "
        "traz as parcelas de <b>todas as empresas do grupo</b> — útil "
        "quando a matriz paga boletos emitidos contra as filiais.<br/>"
        "• Aparece uma mensagem listando quais empresas foram incluídas.<br/>"
        "• A aba <b>Domínio dados</b> ganha a coluna <b>Empresa (código)</b> "
        "mostrando de onde veio cada parcela.<br/>"
        "• O título da aba fica <b>'Domínio dados (N | X empresas)'</b>.<br/>"
        "• O <b>plano de contas continua sendo o da matriz</b> — as "
        "parcelas das filiais são lançadas nas contas da matriz."
    ))
    flow.append(Spacer(1, 6))

    flow.append(Paragraph("3.4 Plano de contas (só na primeira vez do dia)", H3))
    flow.append(bullets([
        "Clique em <b>Carregar plano contas</b>.",
        "Ele traz todas as contas <b>analíticas</b> da empresa "
        "(pra usar nos lançamentos contábeis).",
    ]))

    # ============ Passo 4
    flow.append(PageBreak())
    flow.append(Paragraph("Passo 4 — Conciliar planilha × OFX", PASSO))
    flow.append(Paragraph(
        "Depois de carregar planilha + OFX:", TEXTO,
    ))
    flow.append(bullets([
        "Clique no botão <b>Conciliar</b> (barra do meio).",
        "O sistema tenta casar cada linha da planilha com uma linha do OFX "
        "por <b>data + valor</b>.",
    ]))
    flow.append(Paragraph("O que acontece:", H3))
    flow.append(tabela_simples(
        ["Onde vai", "O que é"],
        [
            ["<b>Aba Conciliados</b>", "Casou tudo certinho "
             "(verde = automático, azul = manual)"],
            ["<b>Aba Pendentes</b>", "O que sobrou dos dois lados (não casou)"],
            ["<b>Aba Sugestões</b>", "Pares com pequena diferença "
             "(até 2 dias e até R$ 10)"],
        ],
        larguras=[5 * cm, 11 * cm],
    ))

    # ============ Passo 5
    flow.append(Paragraph("Passo 5 — Revisar as sugestões", PASSO))
    flow.append(Paragraph(
        "Se tiver algo na aba <b>Sugestões</b>:", TEXTO,
    ))
    flow.append(bullets([
        "Vá pra aba <b>Sugestões</b>.",
        "Marque as linhas que fazem sentido casar (você reconhece os pagamentos).",
        "Clique em <b>Aceitar selecionadas</b> (pra várias) ou "
        "<b>Selecionar tudo</b> + Aceitar (se todas estão OK).",
        "Elas vão pra aba <b>Conciliados</b>.",
    ]))

    # ============ Passo 6
    flow.append(Paragraph("Passo 6 — Tratar os pendentes", PASSO))
    flow.append(Paragraph(
        "A aba <b>Pendentes</b> tem 2 blocos:", TEXTO,
    ))

    flow.append(Paragraph("Bloco de cima: Só na planilha", H3))
    flow.append(Paragraph(
        "Linhas que estavam na planilha mas não bateram com OFX. Podem ser:",
        TEXTO,
    ))
    flow.append(bullets([
        "Pagamento em dinheiro / Caixa geral (não passou pelo banco)",
        "Duplicidade / erro",
        "Ainda não foi pago",
    ]))

    flow.append(Paragraph("Bloco de baixo: Só no OFX", H3))
    flow.append(Paragraph(
        "Linhas do extrato que não bateram com a planilha. Normalmente são:",
        TEXTO,
    ))
    flow.append(bullets([
        "Tarifas do banco (manutenção, IOF, TED)",
        "Juros / rendimentos",
        "Pagamentos não previstos na planilha",
    ]))

    flow.append(Paragraph("Como resolver linhas do OFX (tarifas)", H2))

    flow.append(Paragraph(
        "<b>Se é uma tarifa que se repete todo mês</b> (ex: “TARIFA PACOTE SERVICOS”):",
        TEXTO,
    ))
    flow.append(bullets([
        "Selecione a linha.",
        "Clique em <b>Criar regra (memo)</b>.",
        "Preencha:",
        "&nbsp;&nbsp;<b>Padrão</b>: o texto que identifica (ex: TARIFA PACOTE)",
        "&nbsp;&nbsp;<b>Banco</b> (opcional): nome do banco (evita confusão "
        "com tarifas de nome parecido de outros bancos)",
        "&nbsp;&nbsp;<b>Histórico contábil</b>: como quer que apareça",
        "&nbsp;&nbsp;<b>Conta contábil</b>: escolha da lista",
        "Clique OK. Todas as tarifas iguais viram lançamento automático "
        "agora e nas próximas conciliações.",
    ]))

    flow.append(Paragraph(
        "<b>Se é um caso único, não recorrente:</b>", TEXTO,
    ))
    flow.append(bullets([
        "Selecione a linha.",
        "Clique em <b>Lançamento manual</b>.",
        "Preencha histórico + conta contábil e OK.",
    ]))

    flow.append(Paragraph("Como resolver linhas da planilha", H2))

    flow.append(Paragraph(
        "<b>Pagamento que se repete (ex: aluguel, condomínio):</b>", TEXTO,
    ))
    flow.append(bullets([
        "Selecione a linha.",
        "Clique em <b>Criar regra (fornecedor)</b>.",
        "O padrão já vem preenchido com Tipo/CNPJ/nome. Confira, ajuste o "
        "histórico e escolha a conta contábil.",
        "OK. Casos iguais viram lançamento automático.",
    ]))

    flow.append(Paragraph("<b>Caso único:</b>", TEXTO))
    flow.append(bullets([
        "Selecione a linha → <b>Lançamento manual</b> → preencher e OK.",
    ]))
    flow.append(Spacer(1, 4))
    flow.append(caixa_dica(
        "Pagamentos da planilha sem OFX correspondente entram como "
        "<b>Caixa geral</b> (sem banco associado)."
    ))

    # ============ Passo 7
    flow.append(PageBreak())
    flow.append(Paragraph("Passo 7 — Comparar com o Domínio", PASSO))
    flow.append(Paragraph(
        "Depois de conciliar planilha × OFX:", TEXTO,
    ))
    flow.append(bullets([
        "Clique em <b>Comparar com Domínio</b> (barra do meio).",
        "Vai pra aba <b>Comparação</b> automaticamente.",
    ]))
    flow.append(Paragraph("Na aba Comparação você vê 6 cores:", H3))
    flow.append(tabela_simples(
        ["Cor", "Significado", "O que fazer"],
        [
            ["🟢 <b>Verde (OK)</b>", "Já bateu no Domínio, tudo certo",
             "Nada — está fechado"],
            ["🟡 <b>Amarelo (Falta dom)</b>",
             "Conciliado no banco, falta lançar no Domínio",
             "Criar lançamento ou regra"],
            ["🔵 <b>Azul (Caixa OK)</b>", "Caixa geral já no Domínio",
             "Nada — está fechado"],
            ["⚪ <b>Cinza (Caixa falta)</b>",
             "Caixa geral, falta lançar",
             "Criar lançamento ou regra"],
            ["🩵 <b>Ciano (OFX OK)</b>", "Extrato bateu no Domínio",
             "Nada — está fechado"],
            ["🟠 <b>Laranja (OFX falta)</b>",
             "Extrato sem lançamento no Domínio",
             "Criar lançamento ou regra"],
        ],
        larguras=[4 * cm, 6 * cm, 6 * cm],
    ))

    flow.append(Paragraph("Ações na aba Comparação", H3))
    flow.append(bullets([
        "<b>Editar dados</b>: se a planilha tem erros (NF errada, valor "
        "trocado) que impediram o match com o Domínio, você corrige aqui "
        "e o sistema tenta o match de novo.",
        "<b>Lançar manualmente</b>: cria lançamento contábil pra essa linha.",
        "<b>Criar regra de fornecedor</b>: cria uma regra que também "
        "vai pegar outras linhas similares.",
        "<b>Exportar pendências</b>: gera Excel com todas as linhas <b>falta</b> "
        "(amarelas, cinzas, laranjas) — útil pra revisar ou mandar pra alguém.",
    ]))

    # ============ Passo 8
    flow.append(Paragraph("Passo 8 — Revisar os lançamentos contábeis", PASSO))
    flow.append(Paragraph(
        "A aba <b>Lançamentos contábeis</b> mostra tudo que vai virar "
        "lançamento no Domínio:", TEXTO,
    ))
    flow.append(bullets([
        "Cada linha tem: data, banco, valor, conta, histórico, regra que gerou.",
        "<b>Editar lançamento</b>: muda data/valor/conta/histórico de uma "
        "linha específica sem afetar a regra.",
        "<b>Excluir lançamento</b>: remove essa linha. A transação volta "
        "pra aba Pendentes.",
        "<b>Exportar para Excel</b>: gera .xlsx com todos os lançamentos + "
        "linha TOTAL no fim.",
    ]))
    flow.append(Spacer(1, 4))
    flow.append(caixa_dica(
        "Essa aba é o <b>produto final</b> do seu trabalho. É o que você "
        "vai lançar no Domínio (manualmente ou via importação, "
        "dependendo do processo do escritório)."
    ))

    # ============ Passo 9
    flow.append(Paragraph("Passo 9 — Exportar relatórios (Excel)", PASSO))
    flow.append(Paragraph(
        "Você pode exportar em várias abas pra revisar ou entregar pra alguém:",
        TEXTO,
    ))
    flow.append(bullets([
        "<b>Aba Pendentes</b> → botão <i>Exportar para Excel</i>: 2 abas "
        "(planilha e OFX pendentes).",
        "<b>Aba Conciliados × Domínio</b> → tudo que já está fechado "
        "com o Domínio.",
        "<b>Aba Lançamentos contábeis</b> → os lançamentos que vão pro "
        "Domínio, com totalização.",
        "<b>Aba Comparação</b> → botão <i>Exportar pendências</i>: só o "
        "que falta (amarelo + cinza + laranja).",
    ]))
    flow.append(Paragraph(
        "Os arquivos são salvos com nome sugerido:", TEXTO,
    ))
    flow.append(Paragraph(
        "&lt;tipo&gt;_&lt;código_empresa&gt;_&lt;data&gt;.xlsx<br/>"
        "Exemplo: lancamentos_55_2026-09-14.xlsx",
        CODIGO,
    ))

    # ============ Fluxo típico
    flow.append(PageBreak())
    flow.append(Paragraph("Fluxo típico do dia a dia", PASSO))
    flow.append(Paragraph(
        "Resumo do que você faz a cada empresa/período:", TEXTO,
    ))
    flow.append(Paragraph(
        "1. Abre o app → login<br/>"
        "2. Seleciona empresa<br/>"
        "3. Carrega planilha + OFX + Domínio<br/>"
        "4. Clica em Conciliar<br/>"
        "5. Aceita sugestões que fazem sentido<br/>"
        "6. Trata os Pendentes (cria regras / lançamentos)<br/>"
        "7. Clica em Comparar com Domínio<br/>"
        "8. Trata linhas amarelas / cinzas / laranjas<br/>"
        "9. Vai em Lançamentos contábeis → confere<br/>"
        "10. Exporta pra Excel se precisar<br/>"
        "11. Fecha o app",
        CODIGO,
    ))
    flow.append(Paragraph(
        "Da próxima vez que abrir para a <b>mesma empresa</b>, o sistema já lembra:",
        TEXTO,
    ))
    flow.append(bullets([
        "O mapeamento das colunas da planilha",
        "Todas as regras que você criou",
        "A empresa que estava selecionada",
    ]))
    flow.append(Spacer(1, 4))
    flow.append(caixa_dica(
        "As regras cadastradas por qualquer usuário ficam disponíveis pra "
        "todos que trabalham nessa empresa — não precisa cadastrar de novo "
        "quando é outro operador."
    ))

    # ============ Uso compartilhado
    # Nova seção — Limpar OFX/planilha preservando trabalho
    flow.append(Paragraph(
        "Trocar OFX ou planilha no meio do trabalho", PASSO,
    ))
    flow.append(Paragraph(
        "Se você percebeu que importou o OFX errado (ou a planilha "
        "errada) e já fez conciliação em cima:", TEXTO,
    ))
    flow.append(Paragraph("Limpar OFX", H3))
    flow.append(bullets([
        "Fecha só o OFX. As <b>conciliações que você já fez ficam "
        "preservadas</b>.",
        "A planilha (ou os comprovantes PDF) continuam.",
        "Ao importar o novo OFX e Conciliar de novo, o sistema <b>não "
        "vai refazer</b> as linhas já conciliadas — só as pendentes "
        "tentam casar com o novo OFX.",
    ]))
    flow.append(Paragraph("Limpar planilha", H3))
    flow.append(bullets([
        "Simétrico: fecha só a planilha, OFX continua, conciliações "
        "preservadas.",
    ]))
    flow.append(caixa_dica(
        "Utilíssimo pra corrigir um OFX incompleto sem perder 20 minutos "
        "de classificação já feita. O popup te avisa quantas conciliações "
        "e lançamentos serão preservados antes de aplicar."
    ))

    flow.append(Paragraph("Uso compartilhado (multi-usuário)", PASSO))
    flow.append(bullets([
        "Cada operador tem seu <b>próprio usuário e senha</b>.",
        "Cada operador pode estar em uma <b>empresa diferente ao mesmo tempo</b> "
        "(Aline na empresa 55, Carlos na 82) — não interfere um no outro.",
        "<b>Regras e mapeamentos</b> são da EMPRESA (compartilhados). Se "
        "você criar uma regra de tarifa da empresa 55, seu colega vê a "
        "mesma regra quando trabalhar na 55.",
        "<b>Dados temporários</b> (planilha carregada, OFX importado) são só na "
        "sua tela — quando você fecha o app, some. Se precisar continuar "
        "depois, exporte pra Excel.",
    ]))

    # ============ Perguntas comuns
    flow.append(PageBreak())
    flow.append(Paragraph("Perguntas comuns", PASSO))

    perguntas = [
        ("Errei uma regra, como corrijo?",
         "Botão <b>Configurar taxas</b> (barra do meio) → seleciona a "
         "regra → Editar ou Excluir."),
        ("Meu colega criou uma regra errada, o que faço?",
         "Igual acima. Regras são compartilhadas — qualquer um pode "
         "ajustar/apagar."),
        ("Quero trocar de empresa no meio do trabalho, perco tudo?",
         "Sim: planilha/OFX carregados são descartados (o sistema avisa "
         "antes). Exporte pra Excel se quiser guardar."),
        ("Tem como \"salvar\" o trabalho pra continuar depois?",
         "Só via export pra Excel. As <b>regras</b> e <b>mapeamentos</b> "
         "ficam salvos automaticamente no servidor — só os dados carregados "
         "(planilha/OFX) são temporários."),
        ("Como sei que uma regra automática pegou minha linha?",
         "Vá na aba <b>Lançamentos contábeis</b> — se a linha está lá, "
         "virou lançamento (a coluna “Regra” mostra qual padrão casou)."),
        ("A conta contábil errada foi escolhida numa regra que já pegou "
         "várias linhas. Como corrijo tudo?",
         "Vai em <b>Configurar taxas</b> → edita a regra corrigindo a "
         "conta. Mas o sistema <b>NÃO</b> reaplica automaticamente nas "
         "linhas que já viraram lançamento. Pra corrigir cada uma: "
         "vai em <b>Lançamentos contábeis</b> → seleciona a linha errada "
         "→ <b>Editar lançamento</b> → muda a conta manualmente."),
        ("Um lançamento automático saiu errado, mas a regra está certa. "
         "E agora?",
         "Vai em <b>Lançamentos contábeis</b> → seleciona → <b>Editar</b> "
         "(muda o que precisar) OU <b>Excluir</b> (a transação volta pra "
         "Pendentes)."),
        ("Como troco minha senha?",
         "Botão <b>Minha senha</b> (barra do topo) → coloca senha atual + nova."),
        ("Esqueci minha senha, como reseto?",
         "Fale com um admin. Só admin pode resetar senhas de outros "
         "usuários em <b>Gerenciar usuários</b>."),
    ]
    for pergunta, resposta in perguntas:
        flow.append(Paragraph(f"<b>{pergunta}</b>", H3))
        flow.append(Paragraph(resposta, TEXTO))
        flow.append(Spacer(1, 2))

    # ============ Problemas comuns
    flow.append(Paragraph("Problemas comuns", PASSO))

    problemas = [
        ("“Auto-conexão do Domínio falhou…” em vermelho no topo",
         "O sistema não conseguiu conectar no Domínio Contábil. "
         "Verifique se o Domínio está aberto/rodando na sua máquina. "
         "Se persistir, fale com o admin."),
        ("Tela de login não aparece / app não abre",
         "Verifique se o servidor de banco está acessível. Falar com o admin."),
        ("Após aceitar sugestões, algumas voltam pra Pendentes",
         "Verifique se as datas e valores realmente batem. Se batem mas o "
         "sistema não aceita, exporte a Comparação pra planilha e mande "
         "pro admin analisar."),
        ("Clico em “Comparar com Domínio” e não abre a aba",
         "Confirma que você clicou em <b>Carregar pagamentos</b> antes "
         "(o Domínio precisa estar carregado)."),
    ]
    for problema, solucao in problemas:
        flow.append(Paragraph(f"<b>{problema}</b>", H3))
        flow.append(Paragraph(solucao, TEXTO))
        flow.append(Spacer(1, 2))

    # ============ Suporte
    flow.append(Paragraph("Suporte", PASSO))
    flow.append(Paragraph(
        "Problema no sistema? Fale com o admin informando:", TEXTO,
    ))
    flow.append(bullets([
        "O que você estava fazendo (qual passo do manual).",
        "Qual mensagem apareceu (<b>screenshot ajuda muito</b>).",
        "Qual empresa estava trabalhando.",
    ]))

    return flow


def main():
    saida = Path(__file__).parent / "MANUAL_USO.pdf"
    doc = SimpleDocTemplate(
        str(saida),
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=1.8 * cm,
        title="Manual do Conciliador — Uso Diário",
        author="Conciliador OFX × Planilha × Domínio",
    )
    doc.build(construir(), onFirstPage=_rodape, onLaterPages=_rodape)
    print(f"PDF gerado: {saida}")


if __name__ == "__main__":
    main()
