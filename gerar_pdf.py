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


# ---------------------------------------------------- Tabelas auxiliares

def tabela_abas() -> Table:
    cabecalho = ["#", "Aba", "Conteúdo"]
    linhas = [
        ["0", "Planilha", "Dados crus da planilha .xlsx OU dos comprovantes "
         "PDF (Sicoob/Bradesco). Botões Abrir planilha e Importar "
         "comprovantes PDF. Deduplica automaticamente entre xlsx+PDF."],
        ["1", "OFX", "Pagamentos do extrato (data, banco, valor, memo, "
         "documento). Multi-arquivo. Após conciliação, ganha colunas "
         "Fornecedor (via PDF) + CNPJ (via PDF) + fundo azul nas linhas "
         "enriquecidas."],
        ["2", "Domínio dados", "Parcelas do Domínio com status (Aberto/Parcial/Paga)."],
        ["3", "Conciliados", "Pares Planilha × OFX casados. Coluna 'Origem' "
         "mostra o banco do OFX."],
        ["4", "Pendentes", "Layout vertical (planilha em cima, OFX embaixo). "
         "4 botões: Lançamento manual e Criar regra para cada lado. "
         "Botão Exportar Excel."],
        ["5", "Sugestões", "Pares aproximados (Δ ≤ 2d e Δ ≤ R$ 10). Seleção múltipla."],
        ["6", "Conciliados × Domínio", "Pares triple-matched + Caixa geral + OFX "
         "(sem planilha) que casaram no Domínio. Botão Exportar Excel."],
        ["7", "Comparação", "6 status coloridos (ok, falta dom, caixa ok/falta, "
         "OFX ok/falta). Legenda + 4 botões (Editar dados, Lançar manual, "
         "Criar regra, Exportar pendências)."],
        ["8", "Lançamentos contábeis", "Saídas geradas por regras ou manualmente "
         "(6 tipos). Botões Editar, Excluir, Exportar Excel."],
        ["9", "Plano de contas", "Plano da empresa carregado do Domínio, "
         "FILTRADO só por contas analíticas (tipo A)."],
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
        ["memo",
         "Regra automática casa contra memo + documento do OFX",
         "Do OFX",
         "Tarifas, IOF, juros. Regra tem campo opcional Banco para evitar "
         "falso positivo entre bancos diferentes."],
        ["fornecedor",
         "Regra automática casa contra par P × OFX sem Domínio",
         "Do OFX (banco do par)",
         "Casa contra CNPJ, fornecedor, histórico OU tipo da planilha."],
        ["fornecedor_planilha",
         "Regra automática casa contra pendente da planilha (sem OFX)",
         "Caixa geral",
         "Mesma regra fornecedor, aplicada ao pendente-só-planilha."],
        ["manual",
         "Lançamento manual a partir de par P × OFX (aba Comparação)",
         "Do OFX",
         "Botão 'Lançar manualmente' — não cria regra."],
        ["manual_ofx",
         "Lançamento manual a partir de pendente OFX",
         "Do OFX",
         "Botão na aba Pendentes lado OFX."],
        ["manual_planilha",
         "Lançamento manual a partir de pendente da planilha",
         "Caixa geral",
         "Botão na aba Pendentes lado Planilha."],
    ]
    dados = [cabecalho] + [
        [Paragraph(f"<b>{t}</b>", TEXTO), Paragraph(o, TEXTO),
         Paragraph(b, TEXTO), Paragraph(n, TEXTO)]
        for t, o, b, n in linhas
    ]
    t = Table(dados, colWidths=[3.2 * cm, 4.6 * cm, 3.0 * cm, 5.6 * cm])
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


def tabela_schema_banco() -> Table:
    cabecalho = ["Tabela", "Conteúdo", "Colunas principais"]
    linhas = [
        ["app_config",
         "Configurações globais (fontes SQL do Domínio, chaves/valores JSON)",
         "chave (PK), valor (JSON), atualizado_em"],
        ["regra_taxa",
         "Regras de classificação contábil, uma linha por regra",
         "id (PK), codi_emp, tipo (memo/fornecedor), padrao, historico, "
         "conta, banco, ordem"],
        ["mapeamento_planilha",
         "Mapeamento coluna→campo por empresa",
         "id (PK), codi_emp, campo, nome_coluna (UNIQUE codi_emp+campo)"],
        ["usuario",
         "Usuários do app com senha (PBKDF2)",
         "id (PK), username (UNIQUE), nome, email, senha_hash, ativo, "
         "admin, empresa_ativa (JSON), ultimo_login, criado_em"],
    ]
    dados = [cabecalho] + [
        [Paragraph(f"<b>{t}</b>", TEXTO), Paragraph(c, TEXTO), Paragraph(co, TEXTO)]
        for t, c, co in linhas
    ]
    t = Table(dados, colWidths=[3.8 * cm, 5.4 * cm, 7.2 * cm])
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
        "Conciliador OFX × Planilha × Domínio — funcionalidades implementadas "
        "(Janco Assessoria Contábil)",
        SUBTITULO,
    ))

    # ============================ 1
    flow.append(Paragraph("1. Visão geral", H1))
    flow.append(Paragraph(
        "Aplicativo desktop em Python/Tkinter para conciliação em 3 níveis "
        "e classificação contábil automática. Multi-usuário (rede) com "
        "autenticação individual, dados persistidos em MariaDB compartilhado.",
        TEXTO,
    ))
    flow.append(bullets([
        "<b>Planilha (.xlsx)</b> com contas a pagar — só Vencimento e Valor "
        "são obrigatórios; demais campos (pagamento, emissão, NF, CNPJ, "
        "fornecedor, histórico, tipo) são opcionais.",
        "<b>Extrato bancário (OFX)</b>: aceita <b>múltiplos arquivos</b> de "
        "bancos diferentes; identifica o banco de cada transação.",
        "<b>Sistema Domínio (Escrita Fiscal + Contábil)</b> via ODBC "
        "(read-only): puxa parcelas, status (Aberto/Parcial/Paga) e o plano "
        "de contas da empresa (filtrado só por contas analíticas).",
        "<b>Lançamentos contábeis automáticos</b>: pendentes do OFX, pares "
        "que faltam no Domínio E pendentes da planilha (Caixa geral) podem "
        "ser classificados por regras salvas por empresa ou manualmente.",
        "<b>Comparação OFX × Domínio direta</b>: mesmo sem planilha, dá pra "
        "conferir se os pagamentos do extrato bateram com o Domínio.",
        "<b>Multi-usuário com login individual</b>: cada operador tem seu "
        "próprio username/senha; empresa ativa é individual; regras e "
        "mapeamentos são compartilhados.",
    ]))

    # ============================ 2
    flow.append(Paragraph("2. Arquitetura", H1))

    flow.append(Paragraph("Stack tecnológico", H2))
    flow.append(bullets([
        "<b>Python 3.11+</b> + <b>Tkinter</b> (UI desktop nativa)",
        "<b>MariaDB 10.5+</b> como banco de configurações compartilhado "
        "(via PyMySQL, pure Python — sem compilação C)",
        "<b>SQL Anywhere (Domínio Contábil)</b> via <b>pyodbc</b> (read-only)",
        "<b>openpyxl</b> para .xlsx (leitura e escrita), <b>ofxparse</b> para OFX",
        "<b>reportlab</b> para geração do PDF de documentação",
        "<b>PyInstaller</b> para empacotar em .exe distribuível (Windows)",
    ]))

    flow.append(Paragraph("Modelo de deployment", H2))
    flow.append(bullets([
        "<b>Um servidor MariaDB</b> na rede local (ex: 10.0.1.47:3307) "
        "guardando todas as configurações, regras, mapeamentos e usuários.",
        "<b>Máquinas cliente</b> rodam o Conciliador.exe empacotado, "
        "conectando no MariaDB via rede.",
        "<b>Domínio Contábil</b> continua sendo lido diretamente via ODBC "
        "em cada máquina (não é feito cache no MariaDB).",
        "<b>Distribuição</b>: ZIP com .exe + pasta data/ (credenciais do "
        "MariaDB já preenchidas). Cada operador só configura o "
        "<i>dominio_config.json</i> com o próprio login do Domínio.",
    ]))

    # ============================ 3
    flow.append(Paragraph("3. Autenticação e usuários", H1))

    flow.append(Paragraph("Login individual", H2))
    flow.append(bullets([
        "Tela de login aparece antes da UI principal, forçada ao topo "
        "(-topmost + focus_force pra não ficar atrás de outras janelas).",
        "Autenticação por username + senha, validados contra a tabela "
        "<i>usuario</i> do MariaDB.",
        "Senhas armazenadas com <b>PBKDF2-HMAC-SHA256</b>, 200.000 iterações, "
        "salt aleatório de 16 bytes. Formato: <i>iter$salt$hash</i>. "
        "Comparação em tempo constante (hmac.compare_digest).",
        "Após login: <i>usuario.ultimo_login</i> atualizado; "
        "<i>config.set_usuario_atual(id)</i> registra pra config saber "
        "de qual usuário é a <i>empresa_ativa</i>.",
    ]))

    flow.append(Paragraph("Perfis", H2))
    flow.append(bullets([
        "<b>admin</b>: gerencia (criar/editar/desativar) outros usuários, "
        "vê botões <i>Fonte: pagamentos</i>, <i>Fonte: plano contas</i>, "
        "<i>Gerenciar usuários</i>.",
        "<b>operador</b>: uso normal do app + trocar a própria senha. "
        "Botões de configuração escondidos.",
    ]))

    flow.append(Paragraph("Empresa ativa por usuário", H2))
    flow.append(bullets([
        "Antes: <i>dominio_empresa</i> era global no config.json.",
        "Agora: coluna <i>usuario.empresa_ativa</i> (JSON) — cada login "
        "tem sua própria empresa selecionada.",
        "<b>Regras e mapeamentos continuam por empresa</b> (compartilhados) "
        "— A muda regra da empresa 55, B vê o mesmo estado.",
        "<b>Fontes SQL do Domínio</b> continuam globais (mesma config pra "
        "todos, faz sentido — SQL de acesso ao Domínio é do sistema).",
    ]))

    flow.append(Paragraph("Barra do topo", H2))
    flow.append(Paragraph(
        "Todo usuário vê 'Usuário logado — perfil' + botões 'Minha senha' e "
        "'Trocar usuário'. Admin vê também 'Gerenciar usuários'.",
        TEXTO,
    ))

    # ============================ 4
    flow.append(PageBreak())
    flow.append(Paragraph("4. Leitura de dados", H1))

    flow.append(Paragraph("Planilha Excel (.xlsx) — parser_xlsx.py", H2))
    flow.append(bullets([
        "Detecção automática da linha do cabeçalho (busca nas primeiras 15).",
        "Auto-detecção dos campos por nome, com dezenas de aliases pra cada "
        "(Vencimento, Pagamento, Emissão, Valor, NF, CNPJ, Fornecedor, "
        "Histórico, Tipo).",
        "Fallback por conteúdo (datas → Data, números → Valor, textos "
        "longos → Fornecedor).",
        "<b>Apenas Vencimento e Valor são obrigatórios</b>. No diálogo de "
        "mapeamento cada combo tem opção <i>(deixar vazia)</i>.",
        "Preview ao vivo (10 linhas) com marcação vermelha em linhas inválidas.",
        "Conversão BR de valores (R$ 1.234,56, parênteses pra negativo).",
    ]))

    flow.append(Paragraph("Extrato OFX — parser_ofx.py", H2))
    flow.append(bullets([
        "Leitura via <i>ofxparse</i>. Aceita <b>múltiplos arquivos</b> num "
        "único Importar (Ctrl+clique).",
        "Identifica o <b>banco</b> de cada lançamento via OFX "
        "(institution.organization, routing_number, account_id) com "
        "fallback ao nome do arquivo.",
        "Extrai o número do <b>documento</b> (CHECKNUM ou FITID) — usado "
        "em regras tipo memo.",
        "Filtra automaticamente só pagamentos (valores negativos), invertendo "
        "o sinal pra casar com a planilha.",
    ]))

    flow.append(Paragraph("Comprovantes PDF de pagamento — parser_pdf.py", H2))
    flow.append(bullets([
        "Extração <b>nativa</b> de PDFs (via <i>pdfplumber</i>) de "
        "comprovantes de pagamento de boleto. Não requer OCR — funciona "
        "só em PDFs digitais (baixados de internet banking).",
        "<b>Bancos suportados hoje</b>: Sicoob (SISBR) e Bradesco NET "
        "Empresa. Arquitetura permite adicionar novos bancos com "
        "detecção automática por marcadores fortes no texto.",
        "<b>Campos extraídos</b>: valor pago, data pagamento, data "
        "vencimento, beneficiário (nome e CNPJ), nº documento, banco.",
        "<b>Multi-comprovante</b>: um único PDF pode ter dezenas ou "
        "centenas de comprovantes empilhados — o parser divide por "
        "cabeçalhos e processa cada um separadamente.",
        "<b>Streaming pra PDFs grandes</b>: performance calibrada pra "
        "arquivos de 500+ páginas (~50 ms/pág). UI mostra progresso.",
        "<b>Deduplicação cruzada</b>: se a planilha já tem esse lançamento, "
        "detecta e ignora (ver seção 5).",
    ]))

    flow.append(Paragraph("Sistema Domínio (ODBC) — parser_dominio.py", H2))
    flow.append(bullets([
        "<b>Auto-conexão no startup</b>: usa credenciais salvas em "
        "<i>data/dominio_config.json</i>. Silencioso se falhar — usuário "
        "pode reconectar manualmente.",
        "<b>Duas fontes independentes</b> (<i>pagamentos</i> e "
        "<i>plano_contas</i>) — cada uma com seu SQL e mapeamento próprios.",
        "<i>extrair_pagamentos</i> detecta colunas opcionais "
        "<i>status_parcela</i> e <i>valor_pago</i>.",
        "<i>extrair_plano_contas</i> <b>filtra automaticamente só contas "
        "analíticas</b> (tipo A). Reconhece a coluna de tipo pelo nome "
        "(TIPO_CTA, CTRG_CTA, etc.) ou mapeamento explícito.",
        "Injeção automática de <i>CODI_EMP = ?</i> quando o SQL tem <i>?</i>.",
        "<i>listar_filiais(conn, cnpj_matriz)</i> descobre empresas do "
        "mesmo grupo via <b>CNPJ raiz</b> (primeiros 8 dígitos): útil "
        "quando a matriz paga boletos das filiais.",
        "<i>encontrar_matriz(conn, cnpj)</i> identifica a matriz do "
        "grupo pelo sufixo <b>/0001-XX</b> do CNPJ (regra Receita "
        "Federal). Fallback: menor <i>codi_emp</i> do grupo.",
    ]))

    flow.append(Paragraph("SQL recomendado — plano de contas", H2))
    flow.append(Paragraph(
        "SELECT CLAS_CTA, NOME_CTA, TIPO_CTA<br/>"
        "FROM bethadba.ctcontas<br/>"
        "WHERE CODI_EMP = ? AND TIPO_CTA = 'A'<br/>"
        "ORDER BY CLAS_CTA",
        TEXTO_CODIGO,
    ))

    # ============================ 5
    flow.append(Paragraph("5. Lógica de conciliação", H1))

    flow.append(Paragraph("Nível 1 — Planilha × OFX (matcher.py)", H2))
    flow.append(bullets([
        "<b>conciliar_automatico</b> em DUAS sub-fases:",
        "<b>1.1 Match exato</b> por (data_pagamento se mapeada senão "
        "vencimento, valor). Cada Transacao do OFX só casa com 1 planilha.",
        "<b>1.2 Match aproximado pequeno</b>: nos restantes, Δdias ≤ 2 E "
        "Δvalor ≤ R$ 0,10. Caem em Conciliados (com diferença visível).",
        "<b>gerar_sugestoes</b>: Δdias ≤ 2 e Δvalor ≤ R$ 10 — ficam na aba "
        "Sugestões com seleção múltipla e 'Aceitar tudo'.",
    ]))

    flow.append(Paragraph(
        "Deduplicação Planilha × Comprovantes PDF", H2,
    ))
    flow.append(bullets([
        "Antes da conciliação, se o usuário importar planilha .xlsx E "
        "comprovantes PDF, o app detecta lançamentos que aparecem nos "
        "dois pra evitar duplicidade.",
        "<b>Regra de match</b>: valor exato + data (vencimento OU pagamento) "
        "coincidem + identidade (CNPJ igual OU nome bate por substring "
        "normalizada, sem sufixos societários).",
        "Quando detecta duplicata: mantém a linha da planilha e usa o "
        "PDF só pra <b>enriquecer</b> campos faltantes (ex: se a planilha "
        "não tem CNPJ, herda do PDF).",
        "Nunca considera duplicata se ambos lados estão sem CNPJ e sem "
        "nome (evita colar tarifas iguais em datas iguais como se fossem "
        "uma só).",
        "Popup pós-importação lista quantos novos entraram e quantos "
        "duplicados foram ignorados, pra visibilidade.",
    ]))

    flow.append(Paragraph("Nível 2 — Comparação com Domínio (3 fases)", H2))
    flow.append(bullets([
        "<i>_filtrar_conciliados_por_dominio</i> processa <b>3 fontes</b> "
        "(pares P×OFX, pendentes planilha, pendentes OFX) em <b>3 fases</b> "
        "hierárquicas. Cada Transacao do Domínio só casa com 1 item.",
        "<b>Fase 1 — Exata</b>: data_vencimento + valor + NF iguais.",
        "<b>Fase 2 — 2 de 3</b>: pelo menos 2 dentre (CNPJ, data, valor) "
        "iguais. O campo restante pode divergir; a diferença aparece na "
        "coluna Δ Domínio.",
        "<b>Fase 3 — Fornecedor + Valor</b>: valor exato E (CNPJ bate "
        "OU nome bate por substring normalizada). Data usada apenas como "
        "desempate — NÃO precisa bater. Útil quando o Domínio tem a "
        "mesma parcela mas com vencimento renegociado/prorrogado.",
        "Prioridade nas 3 fases: pares > pendentes planilha > pendentes OFX.",
    ]))

    flow.append(Paragraph("Comparação OFX × Domínio direta (sem planilha)", H2))
    flow.append(bullets([
        "Cenário útil quando o operador só tem OFX e Domínio (não recebeu "
        "planilha de contas a pagar).",
        "Botão 'Conciliar' habilita com planilha OU OFX (não exige os dois).",
        "Match OFX × Domínio segue as mesmas 3 fases.",
        "Pendentes OFX que casarem com Domínio somem da aba Pendentes e "
        "aparecem em Conciliados × Domínio com 'Origem = banco do OFX'.",
    ]))

    flow.append(Paragraph(
        "Enriquecimento cruzado — PDF → OFX → Domínio", H2,
    ))
    flow.append(bullets([
        "Após conciliar, se um par tem <i>planilha.origem = 'pdf'</i>, "
        "o app copia beneficiário/CNPJ/nº doc do PDF para o extras do OFX.",
        "Efeito visual: aba OFX (dados crus) mostra colunas Fornecedor e "
        "CNPJ populadas + linhas com fundo azul claro. Aba Pendentes lado "
        "OFX também mostra dados enriquecidos, útil pra criar regras memo.",
        "Efeito funcional: regras memo automaticamente usam o fornecedor "
        "enriquecido como padrão de match; lançamentos manuais herdam "
        "CNPJ/fornecedor sem digitação.",
        "<b>Aba Conciliados × Domínio</b>: dados priorizam sempre o Domínio "
        "> planilha/PDF > OFX. Se o CNPJ da planilha está errado ou vazio, "
        "puxa o do Domínio (que é a fonte contábil confiável).",
    ]))

    flow.append(Paragraph("Preservação seletiva ao limpar", H2))
    flow.append(bullets([
        "Ao clicar em <b>Limpar planilha</b>: mantém intactos pares "
        "conciliados, matches com Domínio, lançamentos contábeis E "
        "pendentes OFX que ainda estão carregados.",
        "Ao clicar em <b>Limpar OFX</b>: idem, preservando pendentes "
        "da planilha.",
        "<b>Reconciliação seguinte não duplica</b>: linhas já pareadas "
        "não são reoferecidas — só o que sobrou tenta casar com o novo "
        "OFX/planilha importado.",
        "Trocar de empresa continua fazendo <b>reset total</b> (dados "
        "são específicos por empresa).",
    ]))

    flow.append(Paragraph(
        "Grupo matriz + filiais — carregamento consolidado", H2,
    ))
    flow.append(bullets([
        "<b>Problema real</b>: matriz frequentemente paga boletos "
        "emitidos contra filiais. Se o app só puxa parcelas do "
        "<i>codi_emp</i> da matriz, as notas ficam órfãs (amarelo — "
        "'falta no Domínio') sem motivo.",
        "<b>Detecção automática</b>: no clique de 'Carregar pagamentos', "
        "<i>listar_filiais</i> compara o CNPJ raiz (8 primeiros dígitos) "
        "da matriz com todas as empresas cadastradas no Domínio e "
        "retorna as que casam.",
        "<b>Carregamento</b>: chama <i>extrair_pagamentos</i> uma vez "
        "por <i>codi_emp</i> encontrado. Cada Transacao recebe "
        "<i>extras['codi_emp_origem']</i> e <i>extras['razao_empresa']</i>, "
        "então dá pra identificar de qual filial veio a parcela.",
        "<b>Interface</b>: aba <i>Domínio dados</i> ganha coluna "
        "<b>'Empresa (código)'</b>. Título da aba mostra "
        "'<i>Domínio dados (N | X empresas)</i>' quando grupo detectado. "
        "Ao carregar, aparece dialog listando as empresas incluídas.",
        "<b>Fallback</b>: sem CNPJ da matriz, sem raiz válida, ou grupo "
        "com só 1 empresa, comportamento é idêntico ao anterior.",
        "<b>Plano de contas</b> continua vindo <b>só da matriz</b>, "
        "como definido pela regra contábil. Se o operador seleciona uma "
        "filial no combo de empresa, o app usa <i>encontrar_matriz</i> "
        "pra achar o <i>codi_emp</i> da matriz (sufixo /0001) e chama "
        "<i>extrair_plano_contas</i> com esse código — o messagebox "
        "informa qual matriz foi usada.",
        "<b>Regras de taxa por empresa</b> continuam separadas: cada "
        "empresa do grupo tem suas próprias regras salvas em "
        "<i>regra_taxa.codi_emp</i>. Isso é intencional — filiais têm "
        "contas bancárias diferentes, então as regras que apontam pra "
        "essas contas ficam por empresa.",
    ]))

    # ============================ 6
    flow.append(PageBreak())
    flow.append(Paragraph("6. Lançamentos contábeis", H1))
    flow.append(Paragraph(
        "6 tipos de lançamento cobrindo todas as combinações "
        "origem (par/pendente) × modo (regra/manual):", TEXTO,
    ))
    flow.append(tabela_tipos_lancamento())
    flow.append(Spacer(1, 6))

    flow.append(Paragraph("Regras automáticas — configuração", H2))
    flow.append(bullets([
        "Botão <b>Configurar taxas</b> abre diálogo com lista editável.",
        "2 botões: <i>+ Regra por memo</i> e <i>+ Regra por fornecedor</i>.",
        "Regras <b>memo</b> podem ter campo opcional <i>Banco</i> — se "
        "preenchido, regra só dispara em transações daquele banco (evita "
        "falso positivo entre bancos com memos parecidos).",
        "Regras <b>fornecedor</b> casam contra CNPJ + nome + histórico + "
        "TIPO da planilha (substring case-insensitive).",
        "Campo Conta é Combobox autocomplete carregando plano de contas "
        "filtrado por analíticas.",
        "Regras salvas por empresa em <i>regra_taxa.codi_emp</i>.",
        "<b>Atalhos contextuais na aba Pendentes</b>: pré-populam a regra "
        "com dados da linha selecionada.",
        "<b>Histórico contábil é opcional</b>: quando o campo Histórico "
        "da regra fica em branco, o lançamento gerado herda o texto bruto "
        "da origem — <i>memo do OFX</i> para regras <b>memo</b>, "
        "<i>histórico da planilha</i> para regras <b>fornecedor</b>. "
        "Implementado em <i>lancamentos._historico_final</i>. Útil quando "
        "cada linha tem descritivo próprio e não faz sentido padronizar.",
    ]))

    flow.append(Paragraph("Editar e excluir lançamentos", H2))
    flow.append(bullets([
        "Aba <b>Lançamentos contábeis</b> tem botões <b>Editar lançamento</b> "
        "e <b>Excluir lançamento</b>.",
        "<b>Editar</b>: diálogo com data/valor/banco/conta/histórico editáveis. "
        "Se lançamento veio de regra, é 'promovido a manual' — a versão "
        "editada persiste, transação origem é marcada como ignorada.",
        "<b>Excluir</b>: manual remove; automático marca origem como ignorada. "
        "Transação volta pra aba Pendentes.",
    ]))

    # ============================ 7
    flow.append(Paragraph("7. Aba Comparação — 6 cores", H1))
    flow.append(Paragraph(
        "Cabeçalho tem legenda com 6 chips coloridos indicando o significado "
        "de cada cor:", TEXTO,
    ))
    flow.append(bullets([
        "🟢 <b>OK</b> (verde) — Conciliado P × OFX e no Domínio",
        "🟡 <b>Falta dom</b> (amarelo) — Conciliado P × OFX, falta no Domínio",
        "🔵 <b>Caixa OK</b> (azul claro) — Pendente planilha (Caixa geral) no Domínio",
        "⚪ <b>Caixa falta</b> (cinza) — Pendente planilha, falta no Domínio",
        "🩵 <b>OFX OK</b> (ciano) — Pendente OFX (sem planilha) no Domínio",
        "🟠 <b>OFX falta</b> (laranja) — Pendente OFX, falta no Domínio",
    ]))
    flow.append(Paragraph(
        "Título da aba mostra contagem por grupo. Ações (Editar dados, "
        "Lançar manualmente, Criar regra, Exportar pendências) funcionam "
        "em qualquer linha 'falta'.",
        TEXTO,
    ))

    # ============================ 8
    flow.append(Paragraph("8. Interface gráfica", H1))

    flow.append(Paragraph("Barras superiores (4 blocos)", H2))
    flow.append(bullets([
        "<b>Usuário</b>: logado, perfil, botões Minha senha / Gerenciar / Trocar.",
        "<b>Domínio</b>: Conectar, Selecionar empresa, Fonte pagamentos "
        "(admin), Fonte plano contas (admin), Carregar pagamentos, Carregar "
        "plano contas.",
        "<b>Planilha</b>: Abrir, Editar colunas, Limpar planilha.",
        "<b>OFX</b>: Importar, Limpar OFX.",
        "<b>Ações</b>: Conciliar, Comparar com Domínio, Configurar taxas.",
    ]))

    flow.append(Paragraph("Filtros das abas de dados", H2))
    flow.append(bullets([
        "<b>Busca global</b> (Buscar:) filtra em tempo real.",
        "<b>Status</b> (Domínio dados): dropdown Todos/Aberto/Parcial/Paga.",
        "<b>Filtros estilo Excel por coluna</b>: clique no cabeçalho ▾ → "
        "popup com checkboxes + busca + marcar/desmarcar tudo.",
        "Cabeçalho mostra ▼ ★ em colunas com filtro ativo.",
    ]))

    flow.append(Paragraph("Abas do Notebook (10 abas)", H2))
    flow.append(tabela_abas())

    # ============================ 9
    flow.append(PageBreak())
    flow.append(Paragraph("9. Persistência (MariaDB)", H1))

    flow.append(Paragraph("Banco compartilhado", H2))
    flow.append(bullets([
        "Todas as configurações do sistema, regras, mapeamentos e usuários "
        "ficam em MariaDB compartilhado na rede local.",
        "Cada máquina cliente lê/escreve diretamente no banco — não há "
        "backend intermediário.",
        "Camada de acesso: <i>db.py</i> (PyMySQL + context manager) e "
        "<i>config.py</i> (API carregar/salvar compatível com o config.json "
        "antigo, agora sobre MariaDB).",
    ]))

    flow.append(Paragraph("Schema (4 tabelas)", H2))
    flow.append(tabela_schema_banco())
    flow.append(Spacer(1, 6))

    flow.append(Paragraph("Setup e migração", H2))
    flow.append(bullets([
        "<i>setup_db.py</i>: script interativo — pede credenciais MariaDB, "
        "cria banco + tabelas, cadastra 1º admin, migra <i>config.json</i> "
        "antigo se houver.",
        "Idempotente (CREATE IF NOT EXISTS); pode rodar de novo sem risco.",
        "Bootstrap: se não existe nenhum usuário, força criação do 1º admin "
        "no setup.",
    ]))

    flow.append(Paragraph("Credenciais", H2))
    flow.append(bullets([
        "<i>data/db_config.json</i>: host, porta, user, senha, database.",
        "<i>data/dominio_config.json</i>: DSN, usuário, senha do Domínio "
        "(individual por operador — cada consulta ODBC fica rastreada no "
        "Domínio).",
        "Ambos são gitignored — nunca vão pro repositório.",
    ]))

    # ============================ 10
    flow.append(Paragraph("10. Exportação para Excel (.xlsx)", H1))
    flow.append(Paragraph(
        "3 abas do app têm botão 'Exportar para Excel' — geram .xlsx com "
        "estilo consistente (cabeçalho azul, freeze pane na linha 1, cores "
        "de status preservadas):", TEXTO,
    ))
    flow.append(bullets([
        "<b>Aba Pendentes</b>: gera .xlsx com 2 abas (Pendentes Planilha + "
        "Pendentes OFX) — colunas específicas de cada lado.",
        "<b>Aba Conciliados × Domínio</b>: pares triple-matched + Caixa "
        "geral + OFX (todos os matchados). Coluna Origem indica banco ou "
        "'Caixa geral'.",
        "<b>Aba Lançamentos contábeis</b>: todos os lançamentos com "
        "tipo_regra legível + linha TOTAL com fórmula =SUM() no fim.",
        "<b>Aba Comparação</b>: botão Exportar pendências separa em 3 abas "
        "por tipo (Amarelos P × OFX / Cinzas Caixa geral / Laranjas OFX-só). "
        "Cores das linhas preservadas.",
    ]))

    # ============================ 11
    flow.append(Paragraph("11. Distribuição (build .exe)", H1))
    flow.append(Paragraph(
        "Empacotamento via PyInstaller (spec <i>Conciliador.spec</i>). Modo "
        "onedir — gera pasta <i>dist/Conciliador/</i> com:", TEXTO,
    ))
    flow.append(bullets([
        "<b>Conciliador.exe</b> — entry point (6 MB)",
        "<b>_internal/</b> — libs Python + DLLs (pyodbc, tkinter, MariaDB "
        "driver, reportlab, etc.)",
        "<b>data/db_config.json</b> — apontando pro servidor MariaDB (já "
        "preenchido, distribuído junto).",
        "<b>data/dominio_config.json.EXEMPLO</b> — template pro operador "
        "preencher com seu login individual do Domínio.",
        "<b>LEIA-ME.txt</b> — instruções básicas.",
        "Zipado (~25 MB) e distribuído por pen drive, rede ou email.",
    ]))
    flow.append(Paragraph("Detecção runtime do modo empacotado", H2))
    flow.append(Paragraph(
        "<i>db.py</i> e <i>parser_dominio.py</i> têm helper <i>_base_dir()</i> "
        "que detecta <i>sys.frozen</i> (PyInstaller) e usa o diretório do "
        "executável em vez de <i>__file__.parent</i>. Assim as credenciais "
        "ficam ao lado do .exe (editáveis) e não enterradas em _internal.",
        TEXTO,
    ))

    # ============================ 12
    flow.append(Paragraph("12. Diagnóstico ODBC (testar_dominio.py)", H1))
    flow.append(bullets([
        "Script standalone que testa a conexão ODBC com o Domínio em 6 "
        "passos, identificando exatamente onde trava:",
        "1. pyodbc carregado? 2. dominio_config.json existe e é JSON válido? "
        "3. DSN listada no ODBC 64-bit? 4. Conexão abre? 5. SELECT roda? "
        "6. Tem permissão em bethadba?",
        "Detecta erros específicos: IM014 (incompatibilidade 32/64 bit), "
        "28000 (credenciais inválidas), 08001 (servidor Domínio fora).",
        "Empacotável como .exe standalone pra distribuir aos operadores "
        "quando derem problema.",
    ]))

    # ============================ 13
    flow.append(PageBreak())
    flow.append(Paragraph("13. Convenções herdadas do projeto Janco", H1))
    flow.append(bullets([
        "Tabelas do Domínio Escrita Fiscal e Contábil vivem no schema "
        "<b>bethadba</b> (sempre prefixar).",
        "Conexão ODBC é sempre <b>read-only</b> — escrita exigiria RPA "
        "(fora do escopo deste app).",
        "<b>Gotcha SQL Anywhere</b>: DISTINCT + ORDER BY exige aliases "
        "do SELECT (não nomes qualificados — erro -854).",
        "Empresas compartilham a mesma base — separação por <b>CODI_EMP</b> "
        "em cada tabela; o app injeta o filtro automaticamente quando o "
        "SQL tem <i>?</i>.",
    ]))

    # ============================ 14
    flow.append(Paragraph("14. Tabelas do Domínio utilizadas", H1))
    flow.append(bullets([
        "<b>bethadba.efentradas</b> — cabeçalho de NF de entrada.",
        "<b>bethadba.efentradaspar</b> — parcelas/duplicatas (vcto, valor, "
        "número da parcela).",
        "<b>bethadba.efentradaspag</b> — pagamentos registrados (usado pra "
        "calcular status agregado da parcela).",
        "<b>bethadba.effornece</b> — cadastro de fornecedores.",
        "<b>bethadba.geempre</b> — cadastro de empresas.",
        "<b>bethadba.ctcontas</b> — plano de contas (CLAS_CTA, NOME_CTA, "
        "TIPO_CTA) filtrado por CODI_EMP e TIPO_CTA = 'A'.",
    ]))

    # ============================ 15
    flow.append(Paragraph("15. Infraestrutura e código-fonte", H1))
    flow.append(bullets([
        "<b>requirements.txt</b>: openpyxl, ofxparse, pyodbc, reportlab, "
        "PyMySQL.",
        "<b>.gitignore</b>: protege credenciais (db_config, dominio_config), "
        "config.json legado, backups .sql, dist/, build/, .exe, planilhas, "
        "OFX, CSV.",
        "<b>iniciar.bat</b>: git pull + venv + pip install + roda setup_db "
        "se primeira execução + abre app.",
        "<b>Conciliador.spec</b>: config do PyInstaller com hidden imports "
        "declarados e pasta data/ incluída.",
        "<b>gerar_pdf.py</b>: regenera este relatório.",
        "<b>Repositório</b>: github.com/Fernandovini2607/conciliador (privado).",
    ]))

    # ============================ 16
    flow.append(Paragraph("16. Documentação incluída no projeto", H1))
    flow.append(bullets([
        "<b>INSTALACAO_CLIENTE.md</b> — método Git (para desenvolvedores): "
        "Python + Git + venv + pip install. ~30 min.",
        "<b>INSTALACAO_CLIENTE_EXE.md</b> — método ZIP (para operadores): "
        "extrair, criar dominio_config.json, atalho. ~10 min. Recomendado.",
        "<b>relatorio_sistema.pdf</b> — este documento.",
    ]))

    # ============================ 17
    flow.append(Paragraph("17. O que ainda não existe", H1))
    flow.append(bullets([
        "Tolerâncias configuráveis via UI (hoje fixas em código).",
        "Histórico de conciliações entre execuções (auditoria).",
        "Match por similaridade textual (fuzzy match) em fornecedor.",
        "Escrita de volta no Domínio (exigiria RPA via pywinauto).",
        "Edição dos dados do lado OFX (só a planilha é editável hoje).",
        "Tela de configuração de credenciais Domínio no primeiro login "
        "(hoje é manual editando JSON).",
        "Recuperação de senha por email.",
        "Auditoria de ações (quem fez o quê e quando).",
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
