"""Gera relatorio_sistema.pdf com a documentação das funcionalidades implementadas."""

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

# Fontes Arial do Windows — têm suporte amplo a Unicode (Δ, →, ≤, ×).
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
    fontSize=20, leading=24, spaceAfter=6, textColor=colors.HexColor("#1f3a68"),
)
SUBTITULO = ParagraphStyle(
    "Subtitulo", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=11, leading=14, spaceAfter=18, textColor=colors.HexColor("#555555"),
)
H1 = ParagraphStyle(
    "H1", parent=styles["Heading1"], fontName="Arial-Bold",
    fontSize=15, leading=18, spaceBefore=14, spaceAfter=6,
    textColor=colors.HexColor("#1f3a68"),
)
H2 = ParagraphStyle(
    "H2", parent=styles["Heading2"], fontName="Arial-Bold",
    fontSize=12, leading=15, spaceBefore=10, spaceAfter=4,
    textColor=colors.HexColor("#2a4d7d"),
)
TEXTO = ParagraphStyle(
    "Texto", parent=styles["Normal"], fontName="Arial",
    fontSize=10, leading=14, alignment=TA_LEFT, spaceAfter=4,
)
TEXTO_BULLET = ParagraphStyle(
    "Bullet", parent=TEXTO, leftIndent=12, bulletIndent=0, spaceAfter=2,
)
TEXTO_CODIGO = ParagraphStyle(
    "Codigo", parent=styles["Normal"], fontName="Courier",
    fontSize=8.5, leading=11, leftIndent=12, spaceAfter=4,
    textColor=colors.HexColor("#333333"),
)


def bullets(itens: list[str]) -> ListFlowable:
    return ListFlowable(
        [ListItem(Paragraph(t, TEXTO_BULLET), leftIndent=14) for t in itens],
        bulletType="bullet", start="•", bulletColor=colors.HexColor("#1f3a68"),
        leftIndent=8, bulletFontSize=10,
    )


def tabela_abas() -> Table:
    cabecalho = ["#", "Aba", "Conteúdo"]
    linhas = [
        ["0", "Planilha", "Dados crus da planilha (linha, vencimento, "
         "pagamento, emissão, valor, NF, CNPJ, fornecedor, histórico). "
         "Botão Editar lançamento selecionado."],
        ["1", "OFX", "Pagamentos do extrato (data compensação, banco, valor, memo). "
         "Aceita múltiplos OFX de bancos diferentes."],
        ["2", "Domínio dados", "Parcelas vindas do Domínio com status "
         "(Aberto / Parcial / Paga, com cores)."],
        ["3", "Conciliados", "Todos os pares Planilha × OFX casados (verde=auto, "
         "azul=manual). Coluna Origem mostra o banco do OFX."],
        ["4", "Pendentes", "Layout vertical (planilha em cima, OFX embaixo). "
         "4 botões: Lançamento manual e Criar regra para cada lado."],
        ["5", "Sugestões", "Pares aproximados P × OFX (Δ ≤ 2d e Δ ≤ R$ 10). "
         "Seleção múltipla + 'Selecionar tudo'."],
        ["6", "Conciliados × Domínio", "Pares triple-matched (P × OFX × Domínio) "
         "+ pendentes da planilha (Caixa geral) que casaram com o Domínio."],
        ["7", "Comparação", "4 status com cores: ok (verde), falta dom (amarelo), "
         "Caixa ok (azul), Caixa falta (cinza). 3 botões de ação."],
        ["8", "Lançamentos contábeis", "Saídas geradas por regras (memo / fornecedor / "
         "fornecedor_planilha) ou manualmente (5 tipos). Botões Editar e Excluir."],
        ["9", "Plano de contas", "Plano da empresa carregado do Domínio (ctcontas) "
         "FILTRADO só por contas analíticas."],
    ]
    dados = [cabecalho] + [
        [c, ab, Paragraph(co, TEXTO)] for c, ab, co in linhas
    ]
    t = Table(dados, colWidths=[0.7 * cm, 3.4 * cm, 12.3 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("FONTNAME", (0, 1), (-1, -1), "Arial"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def tabela_tipos_lancamento() -> Table:
    cabecalho = ["tipo_regra", "Origem", "Banco", "Notas"]
    linhas = [
        ["memo", "Regra automática casa contra memo+documento do OFX",
         "Do OFX",
         "Tarifas, IOF, juros. Campo Banco opcional na regra para evitar "
         "falso positivo entre bancos diferentes."],
        ["fornecedor", "Regra automática casa contra par P×OFX sem Domínio",
         "Do OFX (banco do par)",
         "Casa contra CNPJ, fornecedor OU histórico da planilha."],
        ["fornecedor_planilha", "Regra automática casa contra pendente da planilha sem OFX",
         "Caixa geral",
         "Mesma regra de fornecedor, mas aplicada quando não existe par com OFX."],
        ["manual", "Lançamento manual a partir de par P×OFX",
         "Do OFX (banco do par)",
         "Botão 'Lançar manualmente' na aba Comparação."],
        ["manual_ofx", "Lançamento manual a partir de pendente OFX",
         "Do OFX",
         "Botão 'Lançamento manual' na aba Pendentes (lado OFX)."],
        ["manual_planilha", "Lançamento manual a partir de pendente planilha",
         "Caixa geral",
         "Botão 'Lançamento manual' na aba Pendentes (lado Planilha)."],
    ]
    dados = [cabecalho] + [
        [Paragraph(f"<b>{t}</b>", TEXTO), Paragraph(o, TEXTO),
         Paragraph(b, TEXTO), Paragraph(n, TEXTO)]
        for t, o, b, n in linhas
    ]
    t = Table(dados, colWidths=[3.0 * cm, 4.6 * cm, 3.0 * cm, 5.8 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a68")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Arial-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#f4f6fa"), colors.white]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Arial", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawString(2 * cm, 1.2 * cm, "Conciliador OFX × Planilha × Domínio")
    canvas.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Página {doc.page}")
    canvas.restoreState()


def construir() -> list:
    flow = []

    flow.append(Paragraph("Relatório do Sistema", TITULO))
    flow.append(Paragraph(
        "Conciliador OFX × Planilha × Domínio — funcionalidades implementadas",
        SUBTITULO,
    ))

    # ============================ 1
    flow.append(Paragraph("1. Visão geral", H1))
    flow.append(Paragraph(
        "Aplicativo desktop em Python/Tkinter para conciliação em 3 níveis "
        "e classificação contábil automática. Suporta inclusive lançamentos "
        "de Caixa geral (entradas na planilha sem contraparte no extrato "
        "bancário).",
        TEXTO,
    ))
    flow.append(bullets([
        "<b>Planilha (.xlsx)</b> com contas a pagar — só Vencimento e Valor "
        "são obrigatórios; demais campos (pagamento, emissão, NF, CNPJ, "
        "fornecedor, histórico) são opcionais.",
        "<b>Extrato bancário (OFX)</b>: aceita <b>múltiplos arquivos</b> de "
        "bancos diferentes; só pagamentos (negativos) são considerados.",
        "<b>Sistema Domínio (Escrita Fiscal + Contábil)</b> via ODBC "
        "(read-only): puxa parcelas, status (Aberto/Parcial/Paga) e o plano "
        "de contas da empresa (filtrado por contas analíticas).",
        "<b>Lançamentos contábeis automáticos</b>: pendentes do OFX, pares "
        "que faltam no Domínio E pendentes da planilha sem OFX podem ser "
        "classificados por regras (salvas por empresa) ou manualmente.",
        "<b>Mapeamento de colunas por empresa</b>: a primeira vez que você "
        "mapeia as colunas da planilha de uma empresa fica salvo — próximas "
        "planilhas dessa empresa não precisam re-mapear (a menos que o "
        "cabeçalho tenha mudado).",
    ]))

    # ============================ 2
    flow.append(Paragraph("2. Leitura de dados", H1))

    flow.append(Paragraph("Planilha Excel (.xlsx) — parser_xlsx.py", H2))
    flow.append(bullets([
        "Detecção automática da linha do cabeçalho (busca nas primeiras 15).",
        "Auto-detecção dos campos por nome (dezenas de aliases pra Data, "
        "Vencimento, Pagamento, Emissão, Valor, Nº NF, CNPJ, Fornecedor, "
        "Histórico).",
        "Fallback por conteúdo (datas → Data, números → Valor, textos "
        "longos → Fornecedor).",
        "<b>Apenas Vencimento e Valor são obrigatórios</b> — os demais são "
        "opcionais. O diálogo de mapeamento tem opção <i>(deixar vazia)</i> "
        "em cada combo para colunas que a planilha não tem.",
        "Diálogo de mapeamento com preview ao vivo (10 linhas), marcação "
        "vermelha em linhas inválidas e labels diferenciadas (opcional "
        "em cinza, obrigatório em preto).",
        "Conversão BR de valores (R$ 1.234,56, parênteses pra negativo).",
    ]))

    flow.append(Paragraph("Extrato OFX — parser_ofx.py", H2))
    flow.append(bullets([
        "Leitura via <i>ofxparse</i>.",
        "<b>Aceita múltiplos arquivos</b> num único Importar (Ctrl+clique "
        "no dialog) — soma todas as transações em <i>transacoes_ofx</i>.",
        "Identifica o <b>banco</b> de cada lançamento via OFX "
        "(institution.organization, routing_number, account_id) com "
        "fallback ao nome do arquivo.",
        "Extrai o número do <b>documento</b> (CHECKNUM ou FITID) — usado "
        "no match de regras tipo memo.",
        "Filtra automaticamente só pagamentos (valores negativos), "
        "invertendo o sinal para casar com a planilha.",
    ]))

    flow.append(Paragraph("Sistema Domínio (ODBC) — parser_dominio.py", H2))
    flow.append(bullets([
        "Padrão Janco: <i>data/dominio_config.json</i> guarda credenciais; "
        "<i>connect_dominio(readonly=True)</i> como context manager.",
        "<b>Duas fontes independentes</b> (<i>dominio_fonte_pagamentos</i> e "
        "<i>dominio_fonte_plano_contas</i>) — cada uma com seu SQL e mapeamento.",
        "Modelos: <i>Transacao</i> (pagamentos) e <i>ContaContabil</i> "
        "(plano de contas).",
        "<i>listar_empresas(conn)</i> lê <i>bethadba.geempre</i>; "
        "<i>listar_tabelas</i> filtra schema <i>bethadba</i> por default.",
        "<i>extrair_pagamentos</i> detecta colunas opcionais <i>status_parcela</i> "
        "e <i>valor_pago</i> vindas do SQL (status calculado por sub-query em "
        "<i>efentradaspag</i>).",
        "<i>extrair_plano_contas</i> <b>filtra automaticamente só contas "
        "analíticas</b> (tipo = \"A\"). Reconhece a coluna de tipo pelo "
        "nome (TIPO_CTA, CTRG_CTA, TIPO_CTC, ANAL_CTA, etc.) ou por "
        "mapeamento explícito.",
        "Injeção automática de <i>CODI_EMP = ?</i> quando o SQL tem <i>?</i>.",
    ]))

    flow.append(Paragraph("SQL recomendado para o plano de contas", H2))
    flow.append(Paragraph(
        "Cole isso no modo Query SQL manual da Fonte de plano de contas:",
        TEXTO,
    ))
    flow.append(Paragraph(
        "SELECT CLAS_CTA, NOME_CTA, TIPO_CTA<br/>"
        "FROM bethadba.ctcontas<br/>"
        "WHERE CODI_EMP = ? AND TIPO_CTA = 'A'<br/>"
        "ORDER BY CLAS_CTA",
        TEXTO_CODIGO,
    ))

    # ============================ 3
    flow.append(Paragraph("3. Modelo de dados", H1))
    flow.append(Paragraph("Campos da planilha (8 — apenas 2 obrigatórios)", H2))
    flow.append(bullets([
        "<b>Data vencimento</b> — chave do match (obrigatório)",
        "<b>Valor</b> — chave do match (obrigatório)",
        "Data pagamento — opcional, usada quando disponível",
        "Data emissão — opcional",
        "Nº NF — opcional, usado no match com Domínio quando disponível",
        "CNPJ fornecedor — opcional, usado em regras fornecedor",
        "Nome do fornecedor — opcional, usado em regras fornecedor",
        "Histórico — opcional, usado em regras fornecedor (case-insensitive substring)",
    ]))
    flow.append(Paragraph("Campos do Domínio (pagamentos): 6 obrigatórios + 2 opcionais", TEXTO))
    flow.append(bullets([
        "Obrigatórios: Data vencimento, Valor, Data emissão, NF, CNPJ, Fornecedor",
        "Opcionais (vindos do SELECT): <i>status_parcela</i> e <i>valor_pago</i>",
    ]))
    flow.append(Paragraph("Campos do plano de contas: 2 obrigatórios + 1 opcional", TEXTO))
    flow.append(bullets([
        "Obrigatórios: <b>Código</b> (CLAS_CTA) e <b>Descrição</b> (NOME_CTA)",
        "Opcional mas <b>recomendado</b>: <b>Tipo</b> (TIPO_CTA — A/S analítica/sintética) "
        "— filtra para mostrar só contas analíticas",
    ]))

    # ============================ 4
    flow.append(PageBreak())
    flow.append(Paragraph("4. Lógica de conciliação", H1))

    flow.append(Paragraph("Fase 1 — Planilha × OFX (matcher.py)", H2))
    flow.append(bullets([
        "<b>conciliar_automatico</b> em DUAS sub-fases:",
        "<b>Sub-fase 1.1 — match exato</b>: por (data_pagamento ou "
        "vencimento, valor). Da planilha usa <i>data_pagamento</i> se "
        "mapeada; do OFX usa <i>data</i> (compensação).",
        "<b>Sub-fase 1.2 — match aproximado pequeno</b>: nos restantes, "
        "Δdias ≤ 2 E Δvalor ≤ R$ 0,10. Esses caem direto em Conciliados "
        "(visíveis na coluna Diferenças) — não viram sugestão.",
        "<b>gerar_sugestoes</b>: pares com Δdias ≤ 2 e Δvalor ≤ R$ 10,00 "
        "(mais permissivo) — ficam na aba Sugestões com seleção múltipla.",
        "<b>Conciliação manual</b>: selecionar 1 em cada lado dos Pendentes.",
    ]))

    flow.append(Paragraph("Fase 2 — Comparação com Domínio (main.py)", H2))
    flow.append(bullets([
        "<i>_filtrar_conciliados_por_dominio</i> roda em DUAS fontes:",
        "<b>Pares P×OFX</b> (têm prioridade) e <b>pendentes da planilha</b> "
        "(Caixa geral, sem OFX casado).",
        "<b>Sub-fase 2.1 — match exato</b>: por (data_vencimento + valor + Nº NF).",
        "<b>Sub-fase 2.2 — match aproximado 2-de-3</b>: para os restantes, "
        "pelo menos 2 de 3 critérios (CNPJ, data_vencimento, valor) iguais. "
        "Diferença registrada em <i>diff_dias_dominio</i> / "
        "<i>diff_valor_dominio</i> e mostrada na coluna Δ Domínio.",
        "Cada Transação do Domínio só pode casar com 1 item — pares têm "
        "prioridade sobre pendentes em ambas as sub-fases.",
        "Pares triple-matched → aba <b>Conciliados × Domínio</b>.",
        "Pendentes da planilha que casaram → também aparecem em "
        "<b>Conciliados × Domínio</b> com origem \"Caixa geral\".",
        "Pares que não bateram (amarelos) e pendentes que não bateram "
        "(cinzas) ficam na aba <b>Comparação</b> — alvos de classificação.",
    ]))

    # ============================ 5
    flow.append(Paragraph("5. Lançamentos contábeis", H1))
    flow.append(Paragraph(
        "Existem 6 tipos de lançamento contábil no app, cobrindo todas as "
        "combinações possíveis de origem (pendente / par) e modo (regra "
        "automática / manual avulso):", TEXTO,
    ))
    flow.append(tabela_tipos_lancamento())
    flow.append(Spacer(1, 6))

    flow.append(Paragraph("Regras automáticas — configuração", H2))
    flow.append(bullets([
        "Botão <b>\"Configurar taxas\"</b> abre diálogo com lista editável "
        "de regras. Dois botões separados: <i>+ Regra por memo</i> e "
        "<i>+ Regra por fornecedor</i>.",
        "Cada regra tem: tipo, padrão, histórico contábil, conta contábil. "
        "Regras tipo <i>memo</i> têm também o campo opcional <b>Banco</b> "
        "(se preenchido, regra só dispara em transações do banco específico).",
        "O campo <b>Conta contábil</b> é um <b>Combobox autocomplete</b> "
        "carregando o plano da empresa filtrado por analíticas.",
        "Regras tipo <i>fornecedor</i> casam contra <b>CNPJ + nome + "
        "histórico</b> da planilha (substring case-insensitive). Atendem "
        "tanto pares P×OFX quanto pendentes-só-planilha (Caixa geral).",
        "Regras são <b>salvas por empresa</b> em "
        "<i>cfg[\"regras_taxas_por_empresa\"][codi_emp]</i>. Trocar de "
        "empresa muda automaticamente o conjunto ativo.",
        "<b>Atalhos contextuais na aba Pendentes</b>: 2 botões em cada lado "
        "(Lançamento manual + Criar regra) pré-populam dados a partir do "
        "pendente selecionado.",
        "<b>Atalho na aba Comparação</b>: <i>\"Criar regra de fornecedor\"</i> "
        "funciona tanto pra par amarelo quanto pra Caixa cinza.",
    ]))

    flow.append(Paragraph("Edição e exclusão de lançamentos", H2))
    flow.append(bullets([
        "Aba <b>Lançamentos contábeis</b> tem botões <b>Editar lançamento</b> "
        "e <b>Excluir lançamento</b>.",
        "<b>Editar</b>: abre diálogo com data, valor, banco, conta e "
        "histórico editáveis. Se o lançamento veio de uma regra automática, "
        "ele é \"promovido a manual\" — a versão editada persiste e a "
        "transação origem fica marcada como ignorada para não regerar.",
        "<b>Excluir</b>: remove o lançamento (manual) ou marca a transação "
        "origem como ignorada (automático). A transação volta para a aba "
        "Pendentes (do lado correspondente).",
        "Estado de ignorados fica em <i>self.lancamentos_ignorados: set[int]</i> "
        "e é aplicado a cada recálculo.",
    ]))

    flow.append(Paragraph("Edição de dados do par e da Transacao", H2))
    flow.append(bullets([
        "Aba <b>Planilha</b>: botão <i>Editar lançamento selecionado</i> abre "
        "diálogo com os 8 campos da Transacao (vencimento, pagamento, "
        "emissão, valor, NF, CNPJ, fornecedor, histórico). Vencimento e "
        "Valor são obrigatórios. Após salvar, os resultados de conciliação "
        "são limpos — usuário precisa rodar Conciliar de novo.",
        "Aba <b>Comparação</b>: botão <i>Editar dados</i> para pares amarelos "
        "(P×OFX sem Domínio). Após salvar, o app re-tenta o match e o par "
        "migra automaticamente para <i>Conciliados × Domínio</i> se passar a "
        "bater.",
    ]))

    # ============================ 6
    flow.append(PageBreak())
    flow.append(Paragraph("6. Interface gráfica (Tkinter)", H1))

    flow.append(Paragraph("Barra de ações (4 blocos no topo)", H2))
    flow.append(bullets([
        "<b>1. Domínio:</b> Conectar | Selecionar empresa | Fonte: pagamentos | "
        "Fonte: plano contas | Carregar pagamentos | Carregar plano contas | label",
        "<b>2. Planilha:</b> Abrir planilha (.xlsx) | Editar colunas | "
        "Limpar planilha | label",
        "<b>3. OFX:</b> Importar OFX (multi-select) | Limpar OFX | label",
        "<b>4. Ações:</b> Conciliar | Comparar com Domínio | Configurar taxas | resumo",
    ]))
    flow.append(Paragraph(
        "Os botões do Domínio ficam <b>em cima</b> porque define a empresa "
        "e o plano de contas. Trocar de empresa <b>limpa automaticamente</b> "
        "planilha, OFX, pagamentos Domínio e plano de contas (com "
        "confirmação) — força importar dados específicos da nova empresa. "
        "As regras de taxas e o mapeamento de colunas continuam salvos por "
        "empresa.",
        TEXTO,
    ))

    flow.append(Paragraph("Filtros das abas de dados", H2))
    flow.append(bullets([
        "<b>Busca global</b> (campo \"Buscar:\") filtra em todas as colunas "
        "em tempo real.",
        "<b>Status</b> (na Domínio dados): dropdown Todos / Aberto / "
        "Parcial / Paga.",
        "<b>Filtros estilo Excel por coluna</b>: clica no cabeçalho \"▾\" → "
        "popup com checkboxes dos valores únicos + busca interna + marcar/"
        "desmarcar tudo. Cabeçalho mostra \"▼ ★\" com filtro ativo.",
        "Todos combinam com AND; botão \"Limpar\" zera tudo.",
    ]))

    flow.append(Paragraph("Cores na aba Comparação", H2))
    flow.append(bullets([
        "🟢 <b>ok</b> (verde claro) — Conciliado P×OFX e no Domínio",
        "🟡 <b>falta_dominio</b> (amarelo) — Conciliado P×OFX, falta no Domínio",
        "🔵 <b>caixa_ok</b> (azul claro) — Caixa geral no Domínio",
        "⚪ <b>caixa_falta</b> (cinza) — Caixa geral falta no Domínio",
    ]))

    flow.append(Paragraph("Cores nas abas Domínio dados e Conciliados × Domínio", H2))
    flow.append(bullets([
        "🟢 <b>Aberto</b> — verde (ainda pode ser tratado / lançado)",
        "🔵 <b>Parcial</b> — azul claro (pagamento parcial)",
        "⚪ <b>Paga</b> — cinza (já liquidada no Domínio)",
    ]))

    flow.append(Paragraph("Abas do Notebook (10 abas em 3 blocos)", H2))
    flow.append(tabela_abas())
    flow.append(Spacer(1, 8))

    flow.append(Paragraph("Layout da aba Pendentes", H2))
    flow.append(bullets([
        "<b>Layout vertical</b>: planilha em cima (com colunas Vencimento, "
        "Pagamento, Valor, Nº NF, Fornecedor, Histórico), OFX embaixo "
        "(com colunas Data pagamento, Banco, Documento, Valor, Memo OFX).",
        "Cada bloco tem 2 botões abaixo da tabela: <b>Lançamento manual</b> "
        "e <b>Criar regra</b>.",
        "Botão <b>Conciliar selecionadas</b> ancorado no rodapé absoluto.",
    ]))

    # ============================ 7
    flow.append(Paragraph("7. Persistência e configuração", H1))
    flow.append(bullets([
        "<b>config.json</b> (raiz, gitignored): preferências do app.",
        "<b>data/dominio_config.json</b> (gitignored): credenciais ODBC.",
        "Migrações automáticas no startup: <i>cfg[\"dominio\"]</i> antigo → "
        "estrutura nova; <i>cfg[\"regras_taxas\"]</i> global → por empresa; "
        "<i>cfg[\"dominio_fonte\"]</i> → <i>dominio_fonte_pagamentos</i>.",
    ]))

    flow.append(Paragraph("Conteúdo do config.json", H2))
    flow.append(bullets([
        "<b>dominio_empresa</b>: {codi_emp, razao, cnpj} da empresa ativa.",
        "<b>dominio_fonte_pagamentos</b>: {modo, sql ou tabela, mapeamento, where}.",
        "<b>dominio_fonte_plano_contas</b>: idem, para o plano.",
        "<b>regras_taxas_por_empresa</b>: dict {codi_emp: [regras…]}.",
        "<b>mapeamentos_planilha_por_empresa</b>: dict {codi_emp: "
        "{campo: nome_coluna}}. Mapeamento salvo POR NOME DE COLUNA — "
        "resiste a reordenação de colunas entre planilhas.",
    ]))

    flow.append(Paragraph("Estrutura típica de regras_taxas_por_empresa", H2))
    flow.append(Paragraph(
        '{ "55": [ '
        '{ "tipo": "memo", "padrao": "TAR PACOTE", "banco": "Banco do Brasil", '
        '"historico": "Tarifa BB", "conta": "4.2.1.001" }, '
        '{ "tipo": "fornecedor", "padrao": "07358761", "historico": "Compras Gerdau", '
        '"conta": "1.1.3.005" } '
        '] }',
        TEXTO_CODIGO,
    ))

    flow.append(Paragraph("Estrutura típica de mapeamentos_planilha_por_empresa", H2))
    flow.append(Paragraph(
        '{ "55": { '
        '"data": "Vencimento", "data_pagamento": "Data Pgto", '
        '"valor": "Valor", "fornecedor": "Fornecedor", '
        '"historico": "Histórico" '
        '} }',
        TEXTO_CODIGO,
    ))

    # ============================ 8
    flow.append(Paragraph("8. Infraestrutura", H1))
    flow.append(bullets([
        "<b>requirements.txt:</b> openpyxl, ofxparse, pyodbc, reportlab.",
        "<b>.gitignore:</b> protege config.json, data/dominio_config.json, "
        ".venv/, __pycache__/, planilhas, OFX e CSV.",
        "<b>iniciar.bat:</b> duplo-clique no Explorer → faz git pull + "
        "ativa venv + abre o app, sem precisar abrir terminal.",
        "<b>gerar_pdf.py:</b> regenera este relatório.",
        "<b>Repositório:</b> github.com/Fernandovini2607/conciliador (privado).",
    ]))

    # ============================ 9
    flow.append(Paragraph("9. Convenções herdadas do projeto Janco", H1))
    flow.append(bullets([
        "Tabelas do Domínio Escrita Fiscal e Contábil vivem no schema "
        "<b>bethadba</b> (sempre prefixar).",
        "Conexão ODBC é sempre <b>read-only</b> — escrita exige RPA "
        "(fora do escopo deste app).",
        "<b>Gotcha SQL Anywhere</b>: DISTINCT + ORDER BY exige aliases "
        "do SELECT (não nomes qualificados — erro -854).",
        "Empresas compartilham a mesma base — separação por "
        "<b>CODI_EMP</b> em cada tabela; o app injeta o filtro "
        "automaticamente quando o SQL tem <i>?</i>.",
    ]))

    # ============================ 10
    flow.append(Paragraph("10. Tabelas do Domínio utilizadas", H1))
    flow.append(bullets([
        "<b>bethadba.efentradas</b> — cabeçalho de NF de entrada.",
        "<b>bethadba.efentradaspar</b> — parcelas/duplicatas (vcto, valor, número da parcela).",
        "<b>bethadba.efentradaspag</b> — pagamentos registrados (pgto_entp, "
        "vpag_entp) usado pra calcular status agregado.",
        "<b>bethadba.effornece</b> — cadastro de fornecedores "
        "(codi_for, nome_for, cgce_for).",
        "<b>bethadba.geempre</b> — cadastro de empresas (codi_emp + razão + CNPJ).",
        "<b>bethadba.ctcontas</b> — plano de contas (CLAS_CTA, NOME_CTA, "
        "TIPO_CTA) com filtro por CODI_EMP e TIPO_CTA = 'A' (só analíticas).",
    ]))

    # ============================ 11
    flow.append(Paragraph("11. O que ainda não existe", H1))
    flow.append(bullets([
        "Tolerâncias configuráveis via UI (hoje fixas em código).",
        "Exportação dos lançamentos contábeis para CSV/Excel.",
        "Tratamento de planilhas com colunas separadas Crédito/Débito.",
        "Match por similaridade textual (fuzzy match) em descrição/fornecedor.",
        "Histórico de conciliações entre execuções (auditoria).",
        "Escrita de volta no Domínio (exigiria RPA via pywinauto).",
        "Edição dos dados do lado OFX (só o lado planilha é editável hoje).",
        "Exportação do plano de contas para arquivo (atualmente fica só em memória).",
    ]))

    return flow


def main() -> None:
    saida = Path(__file__).parent / "relatorio_sistema.pdf"
    doc = SimpleDocTemplate(
        str(saida),
        pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=1.8 * cm,
        title="Relatório do Sistema — Conciliador",
        author="Conciliador OFX × Planilha × Domínio",
    )
    doc.build(construir(), onFirstPage=_rodape, onLaterPages=_rodape)
    print(f"PDF gerado: {saida}")


if __name__ == "__main__":
    main()
