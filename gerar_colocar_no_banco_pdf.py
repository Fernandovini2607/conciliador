"""Gera COLOCAR_SISTEMA_NO_BANCO.pdf a partir do markdown.

Passo a passo simples pra passar pra outra pessoa que vai colocar
outro sistema no banco de dados.
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
    Paragraph,
    SimpleDocTemplate,
    Spacer,
)

FONT_DIR = Path("C:/Windows/Fonts")
pdfmetrics.registerFont(TTFont("Arial", str(FONT_DIR / "arial.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Bold", str(FONT_DIR / "arialbd.ttf")))
pdfmetrics.registerFont(TTFont("Arial-Italic", str(FONT_DIR / "ariali.ttf")))

styles = getSampleStyleSheet()

TITULO = ParagraphStyle(
    "Titulo", parent=styles["Title"], fontName="Arial-Bold",
    fontSize=22, leading=26, spaceAfter=6,
    textColor=colors.HexColor("#1f3a68"),
)
SUBTITULO = ParagraphStyle(
    "Subtitulo", parent=styles["Normal"], fontName="Arial-Italic",
    fontSize=11, leading=15, spaceAfter=18,
    textColor=colors.HexColor("#555"), alignment=1,
)
PASSO = ParagraphStyle(
    "Passo", parent=styles["Heading1"], fontName="Arial-Bold",
    fontSize=15, leading=19, spaceBefore=16, spaceAfter=4,
    textColor=colors.HexColor("#1f3a68"),
)
H2 = ParagraphStyle(
    "H2", parent=styles["Heading2"], fontName="Arial-Bold",
    fontSize=12, leading=16, spaceBefore=10, spaceAfter=3,
    textColor=colors.HexColor("#2a4d7d"),
)
TEXTO = ParagraphStyle(
    "Texto", parent=styles["Normal"], fontName="Arial",
    fontSize=10, leading=14, alignment=TA_LEFT, spaceAfter=4,
)
CODIGO = ParagraphStyle(
    "Codigo", parent=styles["Normal"], fontName="Courier",
    fontSize=9, leading=12, leftIndent=14, spaceAfter=6,
    textColor=colors.HexColor("#222"),
    backColor=colors.HexColor("#f0f2f5"), borderPadding=6,
)
NOTA = ParagraphStyle(
    "Nota", parent=TEXTO, fontName="Arial-Italic",
    textColor=colors.HexColor("#555"), leftIndent=14,
    borderPadding=4,
)


def bullets(itens):
    return ListFlowable(
        [ListItem(Paragraph(x, TEXTO), leftIndent=14) for x in itens],
        bulletType="bullet", leftIndent=14,
    )


def gerar():
    saida = Path(__file__).parent / "COLOCAR_SISTEMA_NO_BANCO.pdf"
    doc = SimpleDocTemplate(
        str(saida), pagesize=A4,
        leftMargin=1.8 * cm, rightMargin=1.8 * cm,
        topMargin=1.8 * cm, bottomMargin=1.8 * cm,
        title="Como colocar um sistema no banco de dados",
        author="Janco Assessoria Contábil",
    )
    f = []

    # ============ Capa
    f.append(Paragraph("Como colocar um sistema no banco de dados", TITULO))
    f.append(Paragraph(
        "Passo a passo simples, sem jargão, pra quem nunca fez.<br/>"
        "Serve pra qualquer sistema Python que hoje guarda dados em<br/>"
        "arquivos locais (JSON, CSV, SQLite) e precisa passar pra um<br/>"
        "banco compartilhado entre várias máquinas.",
        SUBTITULO,
    ))
    f.append(Paragraph("<b>Estimativa:</b> 2 a 4 horas na primeira vez.", TEXTO))
    f.append(Spacer(1, 10))
    f.append(Paragraph("<b>O que você vai fazer:</b>", TEXTO))
    f.append(bullets([
        "Escolher o servidor",
        "Instalar o banco (MariaDB)",
        "Criar o usuário e o banco vazio",
        "Liberar acesso pela rede",
        "Criar as tabelas do sistema",
        "Migrar os dados que já existiam nos arquivos locais",
        "Testar de outra máquina",
        "Configurar cada máquina cliente",
        "Fazer backup automático",
    ]))

    # ============ Passo 1
    f.append(Paragraph("Passo 1 — Escolher o servidor", PASSO))
    f.append(Paragraph(
        "<b>O que é:</b> uma máquina que fica sempre ligada, na rede do "
        "escritório, e vai guardar os dados que todo mundo compartilha.",
        TEXTO,
    ))
    f.append(Paragraph("<b>Pode ser:</b>", TEXTO))
    f.append(bullets([
        "Um computador antigo que fica só pra isso (mais barato).",
        "Um NAS (Synology, QNAP) — quase todos têm MariaDB pronto.",
        "Uma máquina virtual num servidor Windows Server.",
        "Um serviço na nuvem (AWS RDS, DigitalOcean, etc) — mais caro, "
        "mas não precisa cuidar de hardware.",
    ]))
    f.append(Paragraph("<b>Anotar antes de começar:</b>", TEXTO))
    f.append(bullets([
        "<b>IP interno do servidor</b> (ex: 10.0.1.47). Descubra com "
        "<i>ipconfig</i> no cmd da máquina servidor.",
        "<b>Sistema operacional</b> (Windows, Linux).",
    ]))

    # ============ Passo 2
    f.append(Paragraph("Passo 2 — Instalar o banco (MariaDB)", PASSO))
    f.append(Paragraph(
        "<b>Por que MariaDB:</b> é grátis e sem limite, todos os operadores "
        "conseguem acessar ao mesmo tempo, é rápido pra dados de escritório "
        "(até milhões de linhas sem suar).", TEXTO,
    ))
    f.append(Paragraph("<b>2.1 Baixar</b>", H2))
    f.append(Paragraph(
        "Site oficial: <b>mariadb.org/download</b>. Windows: escolha o "
        ".msi de 64 bits. Linux: geralmente já vem no gerenciador de "
        "pacotes (<i>apt install mariadb-server</i> no Ubuntu).", TEXTO,
    ))
    f.append(Paragraph("<b>2.2 Instalar (Windows)</b>", H2))
    f.append(bullets([
        "Duplo-clique no instalador. Aceite os termos → Next.",
        "Deixe o caminho padrão.",
        "Na tela <i>Default instance properties</i>: crie uma <b>senha "
        "forte pro root</b> e anote em lugar seguro.",
        "<b>NÃO marque</b> 'Enable access from remote machines' ainda "
        "— vamos configurar isso do jeito seguro no Passo 4.",
        "Porta: pode deixar 3306, ou trocar pra 3307 se já tiver outro "
        "banco na 3306.",
        "Próxima → Install → Finish.",
    ]))
    f.append(Paragraph("<b>2.3 Testar</b>", H2))
    f.append(Paragraph("Abra o cmd na máquina servidor:", TEXTO))
    f.append(Paragraph("mysql -u root -p", CODIGO))
    f.append(Paragraph(
        "Digite a senha do root. Se aparecer <i>MariaDB [(none)]></i>, "
        "deu certo. Digite <i>exit</i> pra sair.", TEXTO,
    ))

    # ============ Passo 3
    f.append(Paragraph("Passo 3 — Criar o usuário e o banco vazio", PASSO))
    f.append(Paragraph(
        "<b>Por que:</b> o usuário root é só pra administrar. O sistema "
        "vai usar um usuário próprio, com acesso só ao banco dele — mais "
        "seguro.", TEXTO,
    ))
    f.append(Paragraph("<b>3.1 Entrar no MariaDB</b>", H2))
    f.append(Paragraph("mysql -u root -p", CODIGO))
    f.append(Paragraph("<b>3.2 Rodar os comandos (uma linha por vez)</b>", H2))
    f.append(Paragraph(
        "Copia e cola. Troca <i>NOMEDOSISTEMA</i> pelo nome curto do seu "
        "sistema (ex: conciliador, financeiro, estoque). Escolha uma "
        "senha forte pra <i>SENHAFORTE</i> — anote.", TEXTO,
    ))
    f.append(Paragraph(
        "CREATE DATABASE NOMEDOSISTEMA CHARACTER SET utf8mb4<br/>"
        "&nbsp;&nbsp;COLLATE utf8mb4_unicode_ci;<br/>"
        "<br/>"
        "CREATE USER 'NOMEDOSISTEMA_master'@'%' IDENTIFIED BY<br/>"
        "&nbsp;&nbsp;'SENHAFORTE';<br/>"
        "<br/>"
        "GRANT ALL PRIVILEGES ON NOMEDOSISTEMA.* TO<br/>"
        "&nbsp;&nbsp;'NOMEDOSISTEMA_master'@'%';<br/>"
        "<br/>"
        "FLUSH PRIVILEGES;<br/>"
        "EXIT;",
        CODIGO,
    ))
    f.append(Paragraph("<b>O que cada linha faz:</b>", TEXTO))
    f.append(bullets([
        "Linha 1: cria o banco vazio.",
        "Linha 2: cria o usuário. O <i>'%'</i> significa 'pode conectar "
        "de qualquer máquina da rede'.",
        "Linha 3: dá permissão total pra esse usuário só no banco dele.",
        "Linha 4: aplica as mudanças imediatamente.",
        "Linha 5: sai.",
    ]))

    # ============ Passo 4
    f.append(Paragraph("Passo 4 — Liberar acesso pela rede", PASSO))
    f.append(Paragraph(
        "Sem isso, só a máquina servidor consegue conectar. Vamos abrir "
        "pra rede local.", TEXTO,
    ))
    f.append(Paragraph("<b>4.1 Editar a configuração do MariaDB</b>", H2))
    f.append(Paragraph(
        "<b>Windows</b>: abre com o Bloco de Notas o arquivo:", TEXTO,
    ))
    f.append(Paragraph("C:\\Program Files\\MariaDB 11.x\\data\\my.ini", CODIGO))
    f.append(Paragraph("<b>Linux:</b>", TEXTO))
    f.append(Paragraph(
        "sudo nano /etc/mysql/mariadb.conf.d/50-server.cnf", CODIGO,
    ))
    f.append(Paragraph(
        "Procure a linha que começa com <b>bind-address</b>. Deve estar tipo:",
        TEXTO,
    ))
    f.append(Paragraph("bind-address = 127.0.0.1", CODIGO))
    f.append(Paragraph("Troque pra:", TEXTO))
    f.append(Paragraph("bind-address = 0.0.0.0", CODIGO))
    f.append(Paragraph(
        "Isso diz 'aceita conexão de qualquer IP', não só da própria "
        "máquina. Salva e fecha.", TEXTO,
    ))
    f.append(Paragraph("<b>4.2 Reiniciar o MariaDB</b>", H2))
    f.append(Paragraph(
        "<b>Windows:</b> abre 'Serviços' (Win+R → services.msc), procura "
        "'MariaDB', botão direito → <b>Reiniciar</b>.<br/><br/>"
        "<b>Linux:</b>", TEXTO,
    ))
    f.append(Paragraph("sudo systemctl restart mariadb", CODIGO))
    f.append(Paragraph("<b>4.3 Liberar a porta no Firewall</b>", H2))
    f.append(Paragraph("<b>Windows:</b>", TEXTO))
    f.append(bullets([
        "Abre 'Windows Defender Firewall com Segurança Avançada'.",
        "<b>Regras de Entrada</b> → <b>Nova Regra</b>.",
        "Tipo: <b>Porta</b> → TCP → Porta específica: <b>3306</b> "
        "(ou 3307 se trocou).",
        "Ação: <b>Permitir a conexão</b>.",
        "Perfil: marque <b>Domínio + Privada</b> (rede local). "
        "Desmarque Pública (segurança).",
        "Nome: <b>MariaDB</b>. Concluir.",
    ]))
    f.append(Paragraph("<b>Linux (com ufw):</b>", TEXTO))
    f.append(Paragraph("sudo ufw allow 3306/tcp", CODIGO))

    # ============ Passo 5
    f.append(Paragraph("Passo 5 — Criar as tabelas do sistema", PASSO))
    f.append(Paragraph(
        "Cada sistema tem uma estrutura de tabelas diferente. Você precisa "
        "escrever um script Python que cria essas tabelas.", TEXTO,
    ))
    f.append(Paragraph("<b>5.1 Modelo pronto (adaptar)</b>", H2))
    f.append(Paragraph(
        "No Conciliador, esse script é o <i>setup_db.py</i>. Use-o como "
        "referência — só troque o SQL das tabelas pelas suas. Estrutura "
        "básica do script:", TEXTO,
    ))
    f.append(Paragraph(
        "import pymysql<br/><br/>"
        "CONFIG = {<br/>"
        "&nbsp;&nbsp;'host': '10.0.1.47',<br/>"
        "&nbsp;&nbsp;'port': 3306,<br/>"
        "&nbsp;&nbsp;'user': 'NOMEDOSISTEMA_master',<br/>"
        "&nbsp;&nbsp;'password': 'SENHAFORTE',<br/>"
        "&nbsp;&nbsp;'database': 'NOMEDOSISTEMA',<br/>"
        "&nbsp;&nbsp;'charset': 'utf8mb4',<br/>"
        "}<br/><br/>"
        "TABELAS = [<br/>"
        "&nbsp;&nbsp;\"\"\"<br/>"
        "&nbsp;&nbsp;CREATE TABLE IF NOT EXISTS usuarios (<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;id INT AUTO_INCREMENT PRIMARY KEY,<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;username VARCHAR(50) UNIQUE NOT NULL,<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;senha_hash VARCHAR(200) NOT NULL<br/>"
        "&nbsp;&nbsp;) ENGINE=InnoDB<br/>"
        "&nbsp;&nbsp;\"\"\",<br/>"
        "]<br/><br/>"
        "conn = pymysql.connect(**CONFIG)<br/>"
        "with conn.cursor() as cur:<br/>"
        "&nbsp;&nbsp;for sql in TABELAS:<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;cur.execute(sql)<br/>"
        "conn.commit()<br/>"
        "conn.close()",
        CODIGO,
    ))
    f.append(Paragraph("<b>5.2 Rodar o script</b>", H2))
    f.append(Paragraph(
        "Numa máquina que já tem Python + acesso ao servidor:", TEXTO,
    ))
    f.append(Paragraph(
        "pip install pymysql<br/>python setup_db.py", CODIGO,
    ))
    f.append(Paragraph(
        "<b>Dica:</b> o <i>CREATE TABLE IF NOT EXISTS</i> só cria se não "
        "existir. Pode rodar o script quantas vezes quiser sem quebrar "
        "dados existentes.", NOTA,
    ))

    # ============ Passo 6
    f.append(Paragraph(
        "Passo 6 — Migrar os dados que já existiam", PASSO,
    ))
    f.append(Paragraph(
        "Se o sistema já rodava com arquivos locais (JSON, CSV), a gente "
        "precisa mover esses dados pro banco.", TEXTO,
    ))
    f.append(Paragraph("<b>Estratégia:</b>", TEXTO))
    f.append(bullets([
        "Abrir o arquivo local (JSON, por exemplo).",
        "Pra cada registro, executar um INSERT no banco.",
        "<b>Deixar o arquivo antigo como backup</b> (não apagar).",
    ]))
    f.append(Paragraph("<b>Exemplo de script:</b>", TEXTO))
    f.append(Paragraph(
        "import json<br/>"
        "import pymysql<br/><br/>"
        "with open('config.json', 'r', encoding='utf-8') as f:<br/>"
        "&nbsp;&nbsp;dados_antigos = json.load(f)<br/><br/>"
        "conn = pymysql.connect(**CONFIG)<br/>"
        "with conn.cursor() as cur:<br/>"
        "&nbsp;&nbsp;for chave, valor in dados_antigos.items():<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;cur.execute(<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'INSERT INTO app_config"
        " (chave, valor) VALUES (%s, %s)',<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(chave, json.dumps(valor)),<br/>"
        "&nbsp;&nbsp;&nbsp;&nbsp;)<br/>"
        "conn.commit()<br/>conn.close()",
        CODIGO,
    ))
    f.append(Paragraph("<b>Backup antes:</b>", H2))
    f.append(Paragraph("Antes de rodar, faz uma cópia dos dados antigos:", TEXTO))
    f.append(Paragraph(
        "copy C:\\SistemaAntigo\\data\\config.json<br/>"
        "&nbsp;&nbsp;C:\\SistemaAntigo\\data\\config.json.bak",
        CODIGO,
    ))

    # ============ Passo 7
    f.append(Paragraph("Passo 7 — Testar de outra máquina", PASSO))
    f.append(Paragraph(
        "Vá pra uma máquina <b>que não é o servidor</b> e tem Python "
        "instalado.", TEXTO,
    ))
    f.append(Paragraph("<b>7.1 Instalar o driver</b>", H2))
    f.append(Paragraph("pip install pymysql", CODIGO))
    f.append(Paragraph("<b>7.2 Rodar um teste</b>", H2))
    f.append(Paragraph("Criar um arquivo teste.py:", TEXTO))
    f.append(Paragraph(
        "import pymysql<br/><br/>"
        "conn = pymysql.connect(<br/>"
        "&nbsp;&nbsp;host='10.0.1.47',<br/>"
        "&nbsp;&nbsp;port=3306,<br/>"
        "&nbsp;&nbsp;user='NOMEDOSISTEMA_master',<br/>"
        "&nbsp;&nbsp;password='SENHAFORTE',<br/>"
        "&nbsp;&nbsp;database='NOMEDOSISTEMA',<br/>"
        ")<br/>"
        "with conn.cursor() as cur:<br/>"
        "&nbsp;&nbsp;cur.execute('SELECT VERSION()')<br/>"
        "&nbsp;&nbsp;print('Conectado:', cur.fetchone()[0])<br/>"
        "conn.close()",
        CODIGO,
    ))
    f.append(Paragraph(
        "Se aparecer <i>Conectado: 10.x.x-MariaDB</i>, <b>funcionou</b>.",
        TEXTO,
    ))
    f.append(Paragraph("<b>7.3 Se der erro</b>", H2))
    f.append(bullets([
        "<b>'Can't connect to server'</b>: firewall bloqueando. "
        "Volta no passo 4.3.",
        "<b>'Access denied for user'</b>: senha errada ou usuário criado "
        "com 'localhost' em vez de '%'. Volta no passo 3.2 e recria com '%'.",
        "<b>'Unknown database'</b>: o nome do banco no CONFIG está "
        "diferente do que criou. Confere.",
    ]))

    # ============ Passo 8
    f.append(Paragraph("Passo 8 — Configurar cada máquina cliente", PASSO))
    f.append(Paragraph(
        "Em cada computador que vai usar o sistema:", TEXTO,
    ))
    f.append(Paragraph("<b>8.1 Instalar Python e o sistema</b>", H2))
    f.append(Paragraph(
        "Do jeito normal do sistema (git clone, pip install, etc). "
        "Ver o INSTALACAO_CLIENTE.md como exemplo.", TEXTO,
    ))
    f.append(Paragraph("<b>8.2 Colocar as credenciais</b>", H2))
    f.append(Paragraph(
        "Cria o arquivo de configuração (no Conciliador é "
        "data/db_config.json):", TEXTO,
    ))
    f.append(Paragraph(
        "{<br/>"
        "&nbsp;&nbsp;\"host\": \"10.0.1.47\",<br/>"
        "&nbsp;&nbsp;\"port\": 3306,<br/>"
        "&nbsp;&nbsp;\"user\": \"NOMEDOSISTEMA_master\",<br/>"
        "&nbsp;&nbsp;\"password\": \"SENHAFORTE\",<br/>"
        "&nbsp;&nbsp;\"database\": \"NOMEDOSISTEMA\",<br/>"
        "&nbsp;&nbsp;\"charset\": \"utf8mb4\"<br/>"
        "}",
        CODIGO,
    ))
    f.append(Paragraph(
        "<b>Atalho prático:</b> copia o arquivo de uma máquina já "
        "configurada via rede — em vez de digitar de novo em cada uma.",
        NOTA,
    ))

    # ============ Passo 9
    f.append(Paragraph("Passo 9 — Fazer backup automático", PASSO))
    f.append(Paragraph(
        "<b>Sem backup você corre risco de perder tudo</b> se o disco "
        "do servidor pifar.", TEXTO,
    ))
    f.append(Paragraph("<b>9.1 Script simples de backup (Windows)</b>", H2))
    f.append(Paragraph("Criar backup.bat na máquina servidor:", TEXTO))
    f.append(Paragraph(
        "@echo off<br/>"
        "set DATA=%date:~6,4%-%date:~3,2%-%date:~0,2%<br/>"
        "\"C:\\Program Files\\MariaDB 11.x\\bin\\mysqldump\" ^<br/>"
        "&nbsp;&nbsp;-u root -pSENHAROOT ^<br/>"
        "&nbsp;&nbsp;NOMEDOSISTEMA > \"C:\\Backups\\NOMEDOSISTEMA_%DATA%.sql\"",
        CODIGO,
    ))
    f.append(Paragraph("<b>9.2 Agendar</b>", H2))
    f.append(bullets([
        "Abrir <b>Agendador de Tarefas</b> (Windows) → "
        "<b>Criar Tarefa Básica</b>.",
        "Nome: 'Backup diário do sistema'.",
        "Disparador: <b>Diariamente</b> → 23:00 (fora do expediente).",
        "Ação: <b>Iniciar um programa</b> → escolha o backup.bat.",
        "Concluir.",
    ]))
    f.append(Paragraph("<b>9.3 Manter só os últimos 30 dias</b>", H2))
    f.append(Paragraph("Adiciona no fim do backup.bat:", TEXTO))
    f.append(Paragraph(
        "forfiles /P \"C:\\Backups\" /M *.sql /D -30 /C \"cmd /c del @path\"",
        CODIGO,
    ))
    f.append(Paragraph("<b>9.4 IMPORTANTE — copiar backup pra outro lugar</b>", H2))
    f.append(Paragraph(
        "O backup <b>na mesma máquina que tem o banco não vale nada</b> "
        "se a máquina inteira quebrar. Configure cópia pra:", TEXTO,
    ))
    f.append(bullets([
        "Outra máquina da rede (via robocopy num .bat).",
        "Um NAS.",
        "Nuvem (OneDrive, Dropbox, Google Drive).",
    ]))

    # ============ Checklist final
    f.append(Paragraph("Checklist final", PASSO))
    f.append(Paragraph(
        "Antes de considerar o serviço no ar:", TEXTO,
    ))
    f.append(bullets([
        "MariaDB instalado no servidor",
        "Usuário do sistema criado, com senha forte anotada",
        "Banco vazio criado",
        "Acesso pela rede liberado (bind-address + firewall)",
        "Tabelas criadas rodando o setup_db.py",
        "Dados antigos migrados (se havia)",
        "Teste de outra máquina funcionou",
        "Pelo menos 1 máquina cliente configurada e testada",
        "Backup automático rodando",
        "Backup sendo copiado pra outro lugar",
        "Senha do root e do usuário anotadas em lugar seguro",
        "Documento de como fazer isso guardado (esse arquivo!)",
    ]))

    # ============ Ajuda
    f.append(Paragraph("Se precisar de ajuda", PASSO))
    f.append(bullets([
        "Documentação MariaDB: mariadb.com/kb/",
        "PyMySQL: pymysql.readthedocs.io",
        "Referência real desse fluxo: pasta do Conciliador "
        "(setup_db.py, db.py, INSTALACAO_CLIENTE.md).",
    ]))

    doc.build(f)
    print(f"PDF gerado: {saida}")


if __name__ == "__main__":
    gerar()
