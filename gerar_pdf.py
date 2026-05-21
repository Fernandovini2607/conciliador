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
         "pagamento, emissão, valor, NF, CNPJ, fornecedor)."],
        ["1", "OFX", "Pagamentos do extrato (data compensação, banco, valor, memo). "
         "Aceita múltiplos OFX de bancos diferentes."],
        ["2", "Domínio dados", "Parcelas vindas do Domínio com status "
         "(Aberto / Parcial / Paga, com cores)."],
        ["3", "Conciliados", "Todos os pares Planilha × OFX casados (verde=auto, "
         "azul=manual)."],
        ["4", "Pendentes", "Só na planilha / só no OFX. Botão pra conciliar manual "
         "e pra criar lançamento padrão (atalho)."],
        ["5", "Sugestões", "Pares aproximados P × OFX (Δ ≤ 2d e Δ ≤ R$ 10)."],
        ["6", "Conciliados × Domínio", "Pares triple-matched (P × OFX × Domínio) "
         "com status do Domínio em cores."],
        ["7", "Comparação", "Conciliados ok (verde) + falta no Domínio (amarelo). "
         "3 botões: Editar dados, Lançar manualmente, Criar regra de fornecedor."],
        ["8", "Lançamentos contábeis", "Saídas geradas automaticamente por regras "
         "(memo / fornecedor) ou manualmente. Mostra data, banco, valor, conta, "
         "histórico, memo, regra."],
        ["9", "Plano de contas", "Plano de contas da empresa carregado do Domínio "
         "(ctcontas). Coluna código / descrição / tipo, com busca."],
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
        "e classificação contábil automática:",
        TEXTO,
    ))
    flow.append(bullets([
        "<b>Planilha (.xlsx)</b> com contas a pagar (vencimento, pagamento, "
        "valor, NF, CNPJ, fornecedor, emissão).",
        "<b>Extrato bancário (OFX)</b>: aceita <b>múltiplos arquivos</b> de "
        "bancos diferentes; só pagamentos (negativos) são considerados.",
        "<b>Sistema Domínio (Escrita Fiscal + Contábil)</b> via ODBC "
        "(read-only): puxa parcelas, status (Aberto/Parcial/Paga) e o plano "
        "de contas da empresa.",
        "<b>Lançamentos contábeis automáticos</b>: pendentes do OFX e pares "
        "que faltam no Domínio podem ser classificados por regras "
        "(salvas por empresa) ou manualmente.",
    ]))

    # ============================ 2
    flow.append(Paragraph("2. Leitura de dados", H1))

    flow.append(Paragraph("Planilha Excel (.xlsx) — parser_xlsx.py", H2))
    flow.append(bullets([
        "Detecção automática da linha do cabeçalho (busca nas primeiras 15).",
        "Auto-detecção dos 7 campos obrigatórios por nome (dezenas de "
        "aliases pra Data vencimento, Data pagamento, Data emissão, Valor, "
        "Nº NF, CNPJ, Fornecedor).",
        "Fallback por conteúdo (datas → Data, números → Valor, textos "
        "longos → Fornecedor).",
        "Diálogo de mapeamento com preview ao vivo (10 linhas) e marcação "
        "vermelha em linhas inválidas.",
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
        "<i>extrair_plano_contas</i> com mapeamento {codigo, descricao, tipo?}.",
        "Injeção automática de <i>CODI_EMP = ?</i> quando o SQL tem <i>?</i>.",
    ]))

    # ============================ 3
    flow.append(Paragraph("3. Modelo de dados", H1))
    flow.append(Paragraph("Campos obrigatórios da planilha (7):", TEXTO))
    flow.append(bullets([
        "<b>Data vencimento</b> — chave do match Conciliados × Domínio",
        "<b>Data pagamento</b> — chave do match Planilha × OFX",
        "<b>Data emissão</b>",
        "<b>Valor</b> — da parcela",
        "<b>Nº NF</b>",
        "<b>CNPJ fornecedor</b>",
        "<b>Nome do fornecedor</b>",
    ]))
    flow.append(Paragraph("Campos do Domínio (pagamentos): 6 obrigatórios + 2 opcionais", TEXTO))
    flow.append(bullets([
        "Obrigatórios: Data vencimento (vcto_entp), Valor, Data emissão, NF, CNPJ, Fornecedor",
        "Opcionais (vindos do SELECT): <i>status_parcela</i> e <i>valor_pago</i>",
    ]))
    flow.append(Paragraph("Campos do plano de contas: 2 obrigatórios + 1 opcional", TEXTO))
    flow.append(bullets([
        "Obrigatórios: <b>Código</b> (codi_cta) e <b>Descrição</b> (nome_cta)",
        "Opcional: <b>Tipo</b> (tipo_cta — A/S analítica/sintética)",
    ]))

    # ============================ 4
    flow.append(PageBreak())
    flow.append(Paragraph("4. Lógica de conciliação", H1))

    flow.append(Paragraph("Fase 1 — Planilha × OFX (matcher.py)", H2))
    flow.append(bullets([
        "<b>conciliar_automatico</b>: match exato por (data_pagamento, valor). "
        "Da planilha usa <i>data_pagamento</i>; do OFX usa <i>data</i> "
        "(compensação no banco).",
        "<b>gerar_sugestoes</b>: pares com Δdias ≤ 2 e Δvalor ≤ R$ 10,00.",
        "<b>Conciliação manual</b>: selecionar 1 em cada lado dos Pendentes.",
        "<b>Atalho \"Criar lançamento padrão\"</b> na aba Pendentes: pega "
        "um pendente OFX e cria uma regra de memo direto do contexto.",
    ]))

    flow.append(Paragraph("Fase 2 — Conciliados × Domínio (main.py)", H2))
    flow.append(bullets([
        "<i>_filtrar_conciliados_por_dominio</i>: cada par P×OFX casa com o "
        "Domínio por (data_vencimento + valor + Nº NF).",
        "Pares que batem aparecem na aba <b>Conciliados × Domínio</b>.",
        "Pares que não batem aparecem na aba <b>Comparação</b> (amarelos) — "
        "esses são alvo de classificação contábil.",
    ]))

    # ============================ 5
    flow.append(Paragraph("5. Lançamentos contábeis automáticos", H1))
    flow.append(Paragraph(
        "O app classifica automaticamente os lançamentos que não vão pra "
        "Conciliados × Domínio em 3 origens, todas com data = data de "
        "compensação do OFX:", TEXTO,
    ))
    flow.append(bullets([
        "<b>Regras de memo</b>: para cada pendente do OFX (tarifa, IOF, "
        "juros), bate substring contra o memo do banco. Tarifas recorrentes "
        "saem dos Pendentes automaticamente.",
        "<b>Regras de fornecedor</b>: para cada par \"Conciliado, falta no "
        "Domínio\" (amarelo na Comparação), bate substring contra CNPJ ou "
        "nome do fornecedor da planilha. Pares classificados saem da "
        "Comparação automaticamente.",
        "<b>Lançamento manual avulso</b>: na aba Comparação, selecionar um "
        "par amarelo e clicar \"Lançar manualmente\" — pede só histórico + "
        "conta, sem criar regra. Útil pra lançamentos únicos.",
    ]))

    flow.append(Paragraph("Configuração de regras", H2))
    flow.append(bullets([
        "Botão <b>\"Configurar taxas\"</b> abre diálogo com lista editável "
        "de regras. Dois botões separados: <i>+ Regra por memo</i> e "
        "<i>+ Regra por fornecedor</i>.",
        "Cada regra tem: tipo, padrão, histórico contábil, conta contábil.",
        "O campo <b>Conta contábil</b> é um <b>Combobox autocomplete</b> "
        "carregando o plano da empresa (filtro substring por código ou "
        "descrição; só o código é salvo).",
        "Regras são <b>salvas por empresa</b> em "
        "<i>cfg[\"regras_taxas_por_empresa\"][codi_emp]</i>. Trocar de "
        "empresa muda automaticamente o conjunto ativo.",
        "<b>Atalho contextual</b>: na aba Comparação, botão <i>\"Criar regra "
        "de fornecedor\"</i> pré-popula com o CNPJ do par selecionado.",
    ]))

    flow.append(Paragraph("Edição de dados do par", H2))
    flow.append(bullets([
        "Botão <b>\"Editar dados\"</b> na aba Comparação abre um diálogo "
        "permitindo corrigir os 6 campos da planilha (data vencimento, "
        "emissão, valor, NF, CNPJ, fornecedor).",
        "Após salvar, o app <b>re-tenta o match com o Domínio</b>. Se passar "
        "a bater, o par migra automaticamente pra <i>Conciliados × Domínio</i>.",
        "Útil quando o motivo do par não conciliar com o Domínio é dado "
        "preenchido errado (NF com 1 dígito a menos, valor digitado errado, etc.).",
    ]))

    # ============================ 6
    flow.append(PageBreak())
    flow.append(Paragraph("6. Interface gráfica (Tkinter)", H1))

    flow.append(Paragraph("Barra de ações (4 linhas no topo)", H2))
    flow.append(bullets([
        "<b>Planilha:</b> Abrir planilha (.xlsx) | Editar colunas | label",
        "<b>OFX:</b> Importar OFX (multi-select) | label",
        "<b>Domínio:</b> Conectar | Selecionar empresa | Fonte: pagamentos | "
        "Fonte: plano contas | Carregar pagamentos | Carregar plano contas | label",
        "<b>Ações:</b> Conciliar | Comparar com Domínio | Configurar taxas | resumo",
    ]))

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

    flow.append(Paragraph("Cores nas abas Domínio dados e Conciliados × Domínio", H2))
    flow.append(bullets([
        "🟢 <b>Aberto</b> — verde (ainda pode ser tratado / lançado)",
        "🔵 <b>Parcial</b> — azul claro (pagamento parcial)",
        "⚪ <b>Paga</b> — cinza (já liquidada no Domínio)",
    ]))

    flow.append(Paragraph("Abas do Notebook (10 abas em 3 blocos)", H2))
    flow.append(tabela_abas())
    flow.append(Spacer(1, 8))

    # ============================ 7
    flow.append(Paragraph("7. Persistência e configuração", H1))
    flow.append(bullets([
        "<b>config.json</b> (raiz, gitignored): preferências do app — "
        "mapeamento de colunas da planilha, empresa selecionada, fontes "
        "(pagamentos e plano de contas), <b>regras por empresa</b>.",
        "<b>data/dominio_config.json</b> (gitignored): credenciais ODBC.",
        "Migrações automáticas: <i>cfg[\"dominio\"]</i> antigo → estrutura nova; "
        "<i>cfg[\"regras_taxas\"]</i> global → por empresa; "
        "<i>cfg[\"dominio_fonte\"]</i> → <i>dominio_fonte_pagamentos</i>.",
    ]))

    flow.append(Paragraph("Estrutura típica de cfg[\"regras_taxas_por_empresa\"]", H2))
    flow.append(Paragraph(
        '{ "55": [ '
        '{ "tipo": "memo", "padrao": "TAR PACOTE", "historico": "Tarifa", "conta": "4.2.1.001" }, '
        '{ "tipo": "fornecedor", "padrao": "07358761", "historico": "Compras Gerdau", "conta": "1.1.3.005" } '
        '] }',
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
        "<b>bethadba.ctcontas</b> — plano de contas (codi_cta, nome_cta, tipo_cta) "
        "com filtro por codi_emp.",
    ]))

    # ============================ 11
    flow.append(Paragraph("11. O que ainda não existe", H1))
    flow.append(bullets([
        "Tolerância configurável via UI (hoje fixa em 2 dias / R$ 10,00).",
        "Exportação dos lançamentos contábeis para CSV/Excel.",
        "Tratamento de planilhas com colunas separadas Crédito/Débito.",
        "Match por similaridade textual (descrição/fornecedor).",
        "Histórico de conciliações entre execuções.",
        "Escrita de volta no Domínio (exigiria RPA via pywinauto).",
        "Edição dos dados do lado OFX (só o lado planilha é editável hoje).",
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
