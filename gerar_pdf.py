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
    cabecalho = ["#", "Aba", "Conteúdo", "Ações / Notas"]
    linhas = [
        ["0", "Planilha", "Dados crus da planilha (linha original, vencimento, "
         "pagamento, emissão, valor, NF, CNPJ, fornecedor).",
         "Filtros Excel-style por coluna"],
        ["1", "OFX", "Pagamentos do extrato bancário (data de compensação, "
         "valor, memo do banco).", "Filtros Excel-style por coluna"],
        ["2", "Domínio dados", "Parcelas vindas do Domínio (vencimento, "
         "emissão, valor parcela, valor pago, status, NF, CNPJ, fornecedor).",
         "Filtros + dropdown Status"],
        ["3", "Conciliados", "TODOS os pares Planilha × OFX (verde = auto, "
         "azul = manual). Mostra Δdias / Δvalor.",
         "Desfazer conciliação manual"],
        ["4", "Pendentes", "Dois lados: \"Só na planilha\" (com NF e fornecedor) "
         "e \"Só no OFX\" (com memo).", "Conciliar selecionadas →"],
        ["5", "Sugestões", "Pares aproximados P × OFX (Δ ≤ 2d e Δ ≤ R$ 10).",
         "Aceitar como conciliação"],
        ["6", "Conciliados × Domínio", "APENAS pares triple-matched: P × OFX "
         "que também batem no Domínio (vencimento + valor + NF). Coluna Status "
         "colorida (Aberto/Parcial/Paga).", "—"],
        ["7", "Comparação", "Visão tripla: 🟢 em tudo, 🟡 falta na contabilidade, "
         "🔴 só no Domínio (sem conciliação).", "—"],
    ]
    dados = [cabecalho] + [
        [c, ab, Paragraph(co, TEXTO), Paragraph(ac, TEXTO)] for c, ab, co, ac in linhas
    ]
    t = Table(dados, colWidths=[0.7 * cm, 3.2 * cm, 7.6 * cm, 4.9 * cm])
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
        "Aplicativo desktop em Python/Tkinter para conciliação tripla:",
        TEXTO,
    ))
    flow.append(bullets([
        "<b>Planilha (.xlsx)</b> com contas a pagar (vencimento, pagamento, "
        "valor da parcela, NF, fornecedor, CNPJ, emissão).",
        "<b>Extrato bancário (OFX)</b> com pagamentos efetivados — só valores "
        "negativos (saídas) são considerados.",
        "<b>Sistema Domínio (Escrita Fiscal)</b> via ODBC (read-only) — puxa "
        "parcelas da tabela <i>bethadba.efentradaspar</i> com JOIN em "
        "<i>efentradas</i> e <i>effornece</i>, com status calculado a partir "
        "de <i>efentradaspag</i> (Aberto / Parcial / Paga).",
    ]))
    flow.append(Paragraph(
        "Fluxo principal: carregar as 3 fontes → casar Planilha × OFX (match "
        "exato + sugestões com tolerância) → cruzar o resultado com o Domínio "
        "por data de vencimento + valor + número da NF (segunda fase).",
        TEXTO,
    ))

    # ============================ 2
    flow.append(Paragraph("2. Leitura de dados", H1))

    flow.append(Paragraph("Planilha Excel (.xlsx) — parser_xlsx.py", H2))
    flow.append(bullets([
        "<b>Detecção automática da linha do cabeçalho</b> (procura nas primeiras "
        "15 linhas, pula linhas em branco, títulos, logos).",
        "<b>Auto-detecção dos 7 campos obrigatórios por nome</b>, com dezenas de "
        "aliases: <i>Data vencimento</i>, <i>Data pagamento</i> (data pgto / "
        "pago em / data baixa / data quitação / data liquidação / data "
        "compensação…), <i>Data emissão</i>, <i>Valor</i>, <i>Nº NF</i>, "
        "<i>CNPJ</i>, <i>Fornecedor</i>.",
        "<b>Fallback por conteúdo</b>: quando o nome não bate, classifica a "
        "coluna pelo tipo das células — mais datas → Data, mais números → "
        "Valor, textos mais longos → Fornecedor.",
        "<b>Diálogo de mapeamento manual</b> com preview ao vivo das primeiras "
        "10 linhas; linhas em vermelho destacam quando alguma data ou valor "
        "não pôde ser convertido.",
        "<b>Conversão robusta de valores</b> no formato brasileiro: R$ 1.234,56, "
        "vírgula decimal, parênteses para negativo.",
        "<b>Múltiplos formatos de data</b>: dd/mm/aaaa, aaaa-mm-dd, dd-mm-aaaa, "
        "dd/mm/aa, dd.mm.aaaa e células nativas de data do Excel.",
    ]))

    flow.append(Paragraph("Extrato OFX — parser_ofx.py", H2))
    flow.append(bullets([
        "Leitura via <i>ofxparse</i>.",
        "<b>Filtra automaticamente apenas pagamentos</b> (valores negativos), "
        "ignorando recebimentos — o app mostra a contagem de ignorados.",
        "Inverte o sinal para casar com a planilha (que tem valores positivos).",
    ]))

    flow.append(Paragraph("Sistema Domínio (ODBC) — parser_dominio.py", H2))
    flow.append(bullets([
        "<b>Padrão Janco</b>: credenciais em <i>data/dominio_config.json</i> "
        "(gitignored), separadas do <i>config.json</i> de preferências.",
        "<b>connect_dominio(readonly=True)</b> como context manager para scripts; "
        "<b>open_connection()</b> para conexões longas mantidas pela GUI.",
        "<b>load_odbc_config / save_odbc_config</b> para persistir e ler "
        "credenciais.",
        "<b>listar_empresas(conn)</b> lê <i>bethadba.geempre</i> com fallback "
        "entre variantes de nomes de coluna (RAZA_EMP / RAZAO_EMP / NOME_EMP).",
        "<b>listar_tabelas</b> filtra schema <i>bethadba</i> por padrão; "
        "checkbox no diálogo libera ver todos os schemas.",
        "<b>amostra</b> usa fallback TOP / FIRST / LIMIT para compatibilidade "
        "entre dialetos do SQL Anywhere.",
        "<b>extrair_pagamentos</b> aceita dois modos (tabela+WHERE ou SQL livre) "
        "e injeta automaticamente <i>CODI_EMP = ?</i> como parâmetro pyodbc. "
        "Detecta colunas opcionais <i>status_parcela</i> e <i>valor_pago</i> "
        "vindas do SQL e armazena em <i>Transacao.extras</i>.",
        "<b>Gotchas documentados</b>: DISTINCT + ORDER BY exige aliases do SELECT "
        "(não nomes qualificados — erro -854 no SA).",
    ]))

    # ============================ 3
    flow.append(Paragraph("3. Modelo de dados", H1))
    flow.append(Paragraph("Campos obrigatórios da planilha (7):", TEXTO))
    flow.append(bullets([
        "<b>Data vencimento</b> — usada no match com Domínio (vcto_entp)",
        "<b>Data pagamento</b> — usada no match com OFX (data de compensação)",
        "<b>Data emissão</b> — data do documento fiscal",
        "<b>Valor da parcela</b>",
        "<b>Número da NF</b>",
        "<b>CNPJ do fornecedor</b>",
        "<b>Nome do fornecedor</b>",
    ]))
    flow.append(Paragraph("Campos obrigatórios do Domínio (6 — sem data_pagamento):", TEXTO))
    flow.append(bullets([
        "Data vencimento (vcto_entp), Valor, Data emissão, Nº NF, CNPJ, Fornecedor",
    ]))
    flow.append(Paragraph(
        "O <b>OFX</b> é diferente: tem apenas <i>data de pagamento, valor</i> "
        "e <i>memo</i> (do banco). O memo aparece nas abas onde o OFX está "
        "envolvido.",
        TEXTO,
    ))

    flow.append(Paragraph("Query padrão usada no Domínio", H2))
    flow.append(Paragraph(
        "JOIN das 3 tabelas + sub-query agregando pagamentos para derivar "
        "status (Aberto/Parcial/Paga), parametrizado por empresa:",
        TEXTO,
    ))
    flow.append(Paragraph(
        "SELECT par.vcto_entp AS data_vencimento, par.vlor_entp AS valor_parcela, "
        "par.parc_entp AS numero_parcela, ent.ddoc_ent AS data_emissao, "
        "ent.nume_ent AS numero_nf, forn.cgce_for AS cnpj_fornecedor, "
        "forn.nome_for AS nome_fornecedor, "
        "COALESCE(pag.total_pago, 0) AS valor_pago, "
        "CASE WHEN COALESCE(pag.total_pago,0)=0 THEN 'Aberto' "
        "WHEN COALESCE(pag.total_pago,0)&gt;=par.vlor_entp THEN 'Paga' "
        "ELSE 'Parcial' END AS status_parcela "
        "FROM bethadba.efentradaspar par "
        "JOIN bethadba.efentradas ent ON ent.codi_emp=par.codi_emp AND "
        "ent.codi_ent=par.codi_ent "
        "JOIN bethadba.effornece forn ON forn.codi_emp=ent.codi_emp AND "
        "forn.codi_for=ent.codi_for "
        "LEFT JOIN (SELECT codi_emp,codi_ent,parc_entp,SUM(vpag_entp) "
        "AS total_pago FROM bethadba.efentradaspag WHERE pgto_entp IS NOT NULL "
        "GROUP BY codi_emp,codi_ent,parc_entp) pag "
        "ON pag.codi_emp=par.codi_emp AND pag.codi_ent=par.codi_ent "
        "AND pag.parc_entp=par.parc_entp "
        "WHERE par.codi_emp=?",
        TEXTO_CODIGO,
    ))

    # ============================ 4
    flow.append(PageBreak())
    flow.append(Paragraph("4. Lógica de conciliação", H1))

    flow.append(Paragraph("Fase 1 — Planilha × OFX (matcher.py)", H2))
    flow.append(bullets([
        "<b>conciliar_automatico</b>: match exato por (data_pagamento, valor) "
        "quantizado para 2 casas. Da planilha usa <i>data_pagamento</i> "
        "(quando mapeada — recomendado), com fallback para <i>t.data</i> "
        "(vencimento) se não houver.",
        "<b>gerar_sugestoes</b>: pares de pendentes onde a diferença "
        "<i>|Δdias| ≤ 2</i> e <i>|Δvalor| ≤ R$ 10,00</i>. Ordenado pelos "
        "mais próximos.",
        "<b>diferenca(p, o)</b>: calcula Δdias e Δvalor entre planilha e OFX "
        "(usando data_pagamento).",
        "<b>Conciliação manual</b>: selecionar 1 em cada lado dos Pendentes e "
        "clicar \"Conciliar selecionadas →\".",
    ]))

    flow.append(Paragraph("Fase 2 — Pares × Domínio (main.py)", H2))
    flow.append(bullets([
        "<b>_filtrar_conciliados_por_dominio</b>: para cada par P × OFX, busca "
        "uma transação do Domínio que case por (<i>data_vencimento</i> + "
        "<i>valor</i> + <i>NF</i>). Quando encontra, marca <i>par.dominio</i>. "
        "Cada transação do Domínio só pode casar com 1 par.",
        "Disparado automaticamente após: Conciliar, Carregar pagamentos do "
        "Domínio (se já havia conciliação), aceitar sugestão, conciliação "
        "manual, e botão Comparar com Domínio.",
        "Aba <b>Conciliados × Domínio</b> mostra apenas pares onde "
        "<i>par.dominio</i> está preenchido.",
        "Aba <b>Comparação</b> mostra a visão tripla (verde / amarelo / "
        "vermelho) incluindo pagamentos do Domínio sem conciliação.",
    ]))

    # ============================ 5
    flow.append(Paragraph("5. Interface gráfica (Tkinter) — main.py", H1))

    flow.append(Paragraph("Barra de ações (4 linhas no topo)", H2))
    flow.append(bullets([
        "<b>Linha 1 — Planilha:</b> Abrir planilha (.xlsx) | Editar colunas | label",
        "<b>Linha 2 — OFX:</b> Importar OFX | label",
        "<b>Linha 3 — Domínio:</b> Conectar Domínio | Selecionar empresa | "
        "Configurar fonte | Carregar pagamentos | label de status",
        "<b>Linha 4 — Ações:</b> Conciliar | Comparar com Domínio | resumo",
    ]))

    flow.append(Paragraph("Diálogo de mapeamento da planilha — DialogoMapeamento", H2))
    flow.append(bullets([
        "7 dropdowns (todos obrigatórios) pré-selecionados com a sugestão "
        "automática.",
        "<b>Preview ao vivo</b> mostra as primeiras 10 linhas com as 7 colunas "
        "mapeadas — atualiza a cada mudança de dropdown.",
        "<b>Linhas em vermelho</b> quando vencimento, pagamento, emissão ou "
        "valor não puderam ser convertidos.",
        "Status com contagem N/M linhas válidas.",
    ]))

    flow.append(Paragraph("Diálogos do Domínio — dialogos_dominio.py", H2))
    flow.append(bullets([
        "<b>DialogoConexao:</b> DSN + usuário + senha; testa a conexão "
        "(read-only) e persiste em <i>data/dominio_config.json</i>.",
        "<b>DialogoSelecionarEmpresa:</b> lista buscável de "
        "<i>bethadba.geempre</i> (código + razão social + CNPJ); duplo-clique "
        "seleciona.",
        "<b>DialogoFonte:</b> alterna entre <i>Tabela + WHERE</i> e <i>Query "
        "SQL manual</i>; lista tabelas do schema <i>bethadba</i> (com opção de "
        "ver outros schemas); preview até 100 linhas; mapeia os 6 campos.",
        "No modo SQL, quando há <i>?</i> na query, o preview e a extração "
        "passam automaticamente o CODI_EMP da empresa selecionada como "
        "parâmetro pyodbc.",
    ]))

    flow.append(Paragraph("Filtros das abas de dados", H2))
    flow.append(bullets([
        "<b>Busca global</b>: campo \"Buscar:\" filtra linhas que contenham o "
        "texto em qualquer coluna visível (case-insensitive, em tempo real).",
        "<b>Status (só Domínio dados)</b>: dropdown \"Todos / Aberto / Parcial "
        "/ Paga\" filtra pelo status calculado da parcela.",
        "<b>Filtros estilo Excel por coluna</b>: clica no cabeçalho de "
        "qualquer coluna → abre popup com checkboxes dos valores únicos "
        "daquela coluna. Tem busca interna, marcar/desmarcar tudo, OK/Cancelar. "
        "Cabeçalho mostra \"▾\" sem filtro e \"▼ ★\" com filtro ativo.",
        "Todos os filtros combinam com <b>AND</b>. Botão \"Limpar\" zera tudo "
        "de uma vez (global + por coluna + status).",
    ]))

    flow.append(Paragraph("Abas do Notebook (8 abas em 2 blocos)", H2))
    flow.append(tabela_abas())
    flow.append(Spacer(1, 8))

    flow.append(Paragraph("Comportamentos transversais", H2))
    flow.append(bullets([
        "Resumo no topo com contagem por categoria; títulos das abas mostram "
        "número entre parênteses (ex.: <i>Conciliados (142)</i>).",
        "Carregar nova planilha/OFX <b>limpa automaticamente</b> os resultados "
        "anteriores de conciliação.",
        "Aviso quando o mapeamento produz <b>0 lançamentos</b>.",
        "Aviso quando <b>Carregar pagamentos</b> sem empresa selecionada — "
        "evita puxar dados de todas as empresas misturadas.",
        "Conciliação <b>manual</b> entre pendentes; <b>aceitar sugestão</b>; "
        "<b>desfazer conciliação manual</b> devolve o par para pendentes e "
        "recalcula sugestões.",
    ]))

    # ============================ 6
    flow.append(PageBreak())
    flow.append(Paragraph("6. Persistência e configuração", H1))
    flow.append(bullets([
        "<b>config.json</b> (raiz do projeto, gitignored): preferências do app "
        "— mapeamento de colunas da planilha, empresa selecionada do Domínio, "
        "fonte do Domínio (modo, tabela ou SQL, WHERE, mapeamento das 6 "
        "colunas).",
        "<b>data/dominio_config.json</b> (gitignored): credenciais ODBC — "
        "DSN, usuário, senha em texto.",
        "<b>Migração automática</b>: configs antigas com <i>cfg[\"dominio\"]</i> "
        "misturado são divididas automaticamente na primeira execução do app "
        "atualizado.",
    ]))

    # ============================ 7
    flow.append(Paragraph("7. Infraestrutura", H1))
    flow.append(bullets([
        "<b>requirements.txt:</b> openpyxl, ofxparse, pyodbc, reportlab.",
        "<b>.gitignore:</b> protege <i>config.json</i>, "
        "<i>data/dominio_config.json</i>, <i>.venv/</i>, <i>__pycache__/</i>, "
        "planilhas, OFX e CSV.",
        "<b>README.md:</b> guia completo de instalação, configuração do DSN "
        "ODBC para o Domínio (SQL Anywhere 17), fluxo de uso, convenções do "
        "projeto Janco.",
        "<b>iniciar.bat:</b> clique-duplo no Explorer → faz <i>git pull</i> + "
        "ativa venv + abre o app, sem precisar abrir terminal.",
        "<b>gerar_pdf.py:</b> script para regenerar este relatório.",
        "<b>Repositório:</b> github.com/Fernandovini2607/conciliador (privado).",
    ]))

    # ============================ 8
    flow.append(Paragraph("8. Convenções herdadas do projeto Janco", H1))
    flow.append(bullets([
        "Toda tabela do Domínio Escrita Fiscal vive no schema <b>bethadba</b> "
        "(sempre prefixar).",
        "Conexão ODBC é sempre <b>read-only</b> (o usuário ODBC do Domínio "
        "não tem GRANT de UPDATE; qualquer escrita exige RPA — fora do escopo "
        "deste app).",
        "<b>Gotcha SQL Anywhere</b>: DISTINCT + ORDER BY exige usar os aliases "
        "do SELECT, não nomes qualificados (erro -854 com <i>p.CODI_GRU</i>; "
        "use <i>grupo</i> se aliasou).",
        "Empresas vivem na mesma base; separação é por <b>CODI_EMP</b> em cada "
        "tabela. O app injeta esse filtro automaticamente quando há <i>?</i> "
        "no SQL.",
    ]))

    # ============================ 9
    flow.append(Paragraph("9. Tabelas do Domínio utilizadas", H1))
    flow.append(bullets([
        "<b>bethadba.efentradas</b> — cabeçalho de NF de entrada (codi_emp, "
        "codi_ent, codi_for, nume_ent, ddoc_ent, vcon_ent).",
        "<b>bethadba.efentradaspar</b> — parcelas/duplicatas geradas no "
        "import do XML (vcto_entp, vlor_entp, parc_entp).",
        "<b>bethadba.efentradaspag</b> — pagamentos registrados (pgto_entp, "
        "vpag_entp). Usada para calcular o status agregado.",
        "<b>bethadba.effornece</b> — cadastro de fornecedores (codi_for, "
        "nome_for, cgce_for).",
        "<b>bethadba.geempre</b> — cadastro de empresas (codi_emp + razão + "
        "CNPJ). Usada para a lista da seleção de empresa.",
    ]))

    # ============================ 10
    flow.append(Paragraph("10. O que ainda não existe", H1))
    flow.append(bullets([
        "Tolerância configurável via UI (hoje fixa em 2 dias / R$ 10,00).",
        "Exportação do resultado para Excel/PDF/CSV (com cores).",
        "Tratamento de planilhas com colunas separadas de Crédito/Débito.",
        "Match por similaridade textual (descrição ou fornecedor).",
        "Suporte a múltiplas contas no mesmo OFX.",
        "Histórico de conciliações entre execuções.",
        "Conciliação manual diretamente na aba do Domínio.",
        "Escrita de volta no Domínio (exigiria RPA via pywinauto).",
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
