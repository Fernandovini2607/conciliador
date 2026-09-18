# Como colocar um sistema no banco de dados de rede

Passo a passo simples, sem jargão, pra quem nunca fez. Serve pra
qualquer sistema Python que hoje guarda dados em arquivos locais
(JSON, CSV, SQLite) e precisa passar pra um banco compartilhado
entre várias máquinas.

**Estimativa:** 2–4 horas na primeira vez.

Estrutura do que você vai fazer:

1. Escolher o servidor
2. Instalar o banco (MariaDB)
3. Criar o usuário e o banco vazio
4. Liberar acesso pela rede
5. Criar as tabelas do sistema
6. Migrar os dados que já existiam nos arquivos locais
7. Testar de outra máquina
8. Configurar cada máquina cliente
9. Fazer backup automático

---

## Passo 1 — Escolher o servidor

**O que é:** uma máquina que fica sempre ligada, na rede do escritório,
e vai guardar os dados que todo mundo compartilha.

**Pode ser:**
- Um computador antigo que fica só pra isso (mais barato).
- Um NAS (Synology, QNAP) — quase todos têm MariaDB pronto.
- Uma máquina virtual num servidor Windows Server.
- Um serviço na nuvem (AWS RDS, DigitalOcean, etc) — mais caro, mas
  não precisa cuidar de hardware.

**Anotar antes de começar:**
- **IP interno do servidor** (ex: `10.0.1.47`). Descubra com
  `ipconfig` no cmd da máquina servidor.
- **Sistema operacional** (Windows, Linux). Vou explicar pros dois.

---

## Passo 2 — Instalar o banco (MariaDB)

**Por que MariaDB:**
- É grátis e sem limite.
- Todos os operadores conseguem acessar ao mesmo tempo.
- Rápido pra dados de escritório (até milhões de linhas sem suar).

### 2.1 Baixar

Site oficial: <https://mariadb.org/download>

Escolha:
- **Windows**: o `.msi` de 64 bits.
- **Linux**: geralmente já vem no gerenciador de pacotes
  (`apt install mariadb-server` no Ubuntu).

### 2.2 Instalar (Windows)

1. Duplo-clique no instalador.
2. Aceite os termos → Next.
3. Deixe o caminho padrão.
4. Na tela **"Default instance properties"**:
   - **Senha do root**: crie uma senha forte e **anote num lugar
     seguro**. Você vai precisar dela depois.
   - **NÃO marque** "Enable access from remote machines" ainda —
     vamos configurar isso do jeito seguro no Passo 4.
   - **Porta**: pode deixar 3306, ou trocar pra 3307 se já tiver
     algum outro banco na porta 3306.
5. Próxima → Install → Finish.

### 2.3 Testar

Abra o cmd na máquina servidor:

```
mysql -u root -p
```

Digite a senha do root. Se aparecer `MariaDB [(none)]>`, deu certo.
Digite `exit` pra sair.

---

## Passo 3 — Criar o usuário e o banco vazio

**Por que:** o usuário `root` é só pra administrar. O sistema vai
usar um usuário próprio, com acesso só ao banco dele — mais seguro.

### 3.1 Entrar no MariaDB

No cmd do servidor:
```
mysql -u root -p
```

### 3.2 Rodar os comandos (uma linha por vez, apertando Enter)

Copia e cola. Troca `NOMEDOSISTEMA` pelo nome curto do seu sistema
(ex: `conciliador`, `financeiro`, `estoque`). Escolha uma senha forte
pra `SENHAFORTE` — anote.

```sql
CREATE DATABASE NOMEDOSISTEMA CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

CREATE USER 'NOMEDOSISTEMA_master'@'%' IDENTIFIED BY 'SENHAFORTE';

GRANT ALL PRIVILEGES ON NOMEDOSISTEMA.* TO 'NOMEDOSISTEMA_master'@'%';

FLUSH PRIVILEGES;

EXIT;
```

**O que cada linha faz (em português):**
- Linha 1: cria o banco vazio.
- Linha 2: cria o usuário que o sistema vai usar. O `'%'` significa
  "pode conectar de qualquer máquina da rede".
- Linha 3: dá permissão total pra esse usuário, só no banco dele.
- Linha 4: aplica as mudanças imediatamente.
- Linha 5: sai.

---

## Passo 4 — Liberar acesso pela rede

Sem isso, só a máquina servidor consegue conectar. Vamos abrir pra
rede local.

### 4.1 Editar a configuração do MariaDB

**Windows:** abre com o Bloco de Notas o arquivo:
```
C:\Program Files\MariaDB 11.x\data\my.ini
```

**Linux:**
```
sudo nano /etc/mysql/mariadb.conf.d/50-server.cnf
```

Procure a linha que começa com `bind-address`. Deve estar tipo:
```
bind-address = 127.0.0.1
```

Troque pra:
```
bind-address = 0.0.0.0
```

Isso diz "aceita conexão de qualquer IP", não só da própria máquina.

Salva e fecha.

### 4.2 Reiniciar o MariaDB

**Windows:**
1. Abre "Serviços" (aperta Win+R, digita `services.msc`).
2. Procura "MariaDB".
3. Botão direito → **Reiniciar**.

**Linux:**
```
sudo systemctl restart mariadb
```

### 4.3 Liberar a porta no Firewall

**Windows:**
1. Abre "Windows Defender Firewall com Segurança Avançada".
2. **Regras de Entrada** → **Nova Regra**.
3. Tipo: **Porta** → TCP → Porta específica: **3306** (ou 3307 se
   trocou).
4. Ação: **Permitir a conexão**.
5. Perfil: marque **Domínio + Privada** (rede local). Desmarque
   Pública (segurança).
6. Nome: `MariaDB`. Concluir.

**Linux (com ufw):**
```
sudo ufw allow 3306/tcp
```

---

## Passo 5 — Criar as tabelas do sistema

Cada sistema tem uma estrutura de tabelas diferente. **Você precisa
escrever um script Python que cria essas tabelas.**

### 5.1 Modelo pronto (adaptar)

No nosso Conciliador, esse script é o [`setup_db.py`](setup_db.py).
Use-o como referência — só troque o SQL das tabelas pelas suas.

Estrutura básica do script:

```python
import pymysql

# Configuração do banco (vai apontar pro servidor)
CONFIG = {
    "host": "10.0.1.47",         # IP do servidor
    "port": 3306,                # ou 3307
    "user": "NOMEDOSISTEMA_master",
    "password": "SENHAFORTE",
    "database": "NOMEDOSISTEMA",
    "charset": "utf8mb4",
}

# SQL de cada tabela do seu sistema
TABELAS = [
    """
    CREATE TABLE IF NOT EXISTS usuarios (
        id INT AUTO_INCREMENT PRIMARY KEY,
        username VARCHAR(50) UNIQUE NOT NULL,
        senha_hash VARCHAR(200) NOT NULL,
        criado_em DATETIME DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
    """,
    # ... acrescente uma entrada por tabela
]

# Conecta e cria
conn = pymysql.connect(**CONFIG)
with conn.cursor() as cur:
    for sql in TABELAS:
        cur.execute(sql)
conn.commit()
conn.close()
print("Tabelas criadas com sucesso!")
```

### 5.2 Rodar o script

Numa máquina que já tem Python + acesso ao servidor:

```
pip install pymysql
python setup_db.py
```

Se aparecer "Tabelas criadas com sucesso", pronto.

**Dica:** o `CREATE TABLE IF NOT EXISTS` só cria se não existir. Pode
rodar o script quantas vezes quiser sem quebrar dados existentes.

---

## Passo 6 — Migrar os dados que já existiam

Se o sistema já rodava com arquivos locais (JSON, CSV), a gente
precisa mover esses dados pro banco.

### 6.1 Estratégia

1. Abrir o arquivo local (JSON, por exemplo).
2. Pra cada registro, executar um `INSERT` no banco.
3. Deixar o arquivo antigo como backup (não apagar).

Exemplo de script:

```python
import json
import pymysql

CONFIG = { ... }  # mesmo do passo 5

# Carrega dados antigos
with open("config.json", "r", encoding="utf-8") as f:
    dados_antigos = json.load(f)

# Insere no banco
conn = pymysql.connect(**CONFIG)
with conn.cursor() as cur:
    for chave, valor in dados_antigos.items():
        cur.execute(
            "INSERT INTO app_config (chave, valor) VALUES (%s, %s) "
            "ON DUPLICATE KEY UPDATE valor = VALUES(valor)",
            (chave, json.dumps(valor)),
        )
conn.commit()
conn.close()
print("Migração concluída!")
```

O `ON DUPLICATE KEY UPDATE` faz o script ser seguro — pode rodar 2x
sem duplicar.

### 6.2 Backup antes

Antes de rodar, faz uma cópia da pasta com os dados antigos:

```
copy C:\SistemaAntigo\data\config.json C:\SistemaAntigo\data\config.json.bak
```

Se der problema, você volta o backup.

---

## Passo 7 — Testar de outra máquina

Vá pra uma máquina **que não é o servidor** e tem Python instalado.

### 7.1 Instalar o driver

```
pip install pymysql
```

### 7.2 Rodar um teste

Criar um arquivo `teste.py`:

```python
import pymysql

conn = pymysql.connect(
    host="10.0.1.47",           # IP do servidor
    port=3306,                  # ou 3307
    user="NOMEDOSISTEMA_master",
    password="SENHAFORTE",
    database="NOMEDOSISTEMA",
)
with conn.cursor() as cur:
    cur.execute("SELECT VERSION()")
    print("Conectado:", cur.fetchone()[0])
conn.close()
```

Rodar:
```
python teste.py
```

Se aparecer "Conectado: 10.x.x-MariaDB", **funcionou**. Você
consegue conectar do outro computador pelo servidor.

### 7.3 Se der erro

- **"Can't connect to server"**: firewall bloqueando. Volta no passo
  4.3.
- **"Access denied for user"**: senha errada ou usuário criado com
  `'localhost'` em vez de `'%'`. Volta no passo 3.2 e recria o
  usuário com `'%'`.
- **"Unknown database"**: o nome do banco no CONFIG está diferente
  do que criou. Confere.

---

## Passo 8 — Configurar cada máquina cliente

Em cada computador que vai usar o sistema:

### 8.1 Instalar Python e o sistema

Do jeito normal do sistema (git clone, pip install, etc). Veja o
[`INSTALACAO_CLIENTE.md`](INSTALACAO_CLIENTE.md) como exemplo.

### 8.2 Colocar as credenciais

Cria o arquivo de configuração (no Conciliador é
`data/db_config.json`):

```json
{
  "host": "10.0.1.47",
  "port": 3306,
  "user": "NOMEDOSISTEMA_master",
  "password": "SENHAFORTE",
  "database": "NOMEDOSISTEMA",
  "charset": "utf8mb4"
}
```

**Atalho prático:** copia o arquivo de uma máquina já configurada
via rede — em vez de digitar de novo em cada uma.

### 8.3 Testar

Abre o sistema. Se logar e ver os dados, tá tudo certo.

---

## Passo 9 — Fazer backup automático

**Sem backup você corre risco de perder tudo** se o disco do
servidor pifar.

### 9.1 Script simples de backup (Windows)

Criar `backup.bat` na máquina servidor:

```bat
@echo off
set DATA=%date:~6,4%-%date:~3,2%-%date:~0,2%
"C:\Program Files\MariaDB 11.x\bin\mysqldump" ^
  -u root -pSENHAROOT ^
  NOMEDOSISTEMA > "C:\Backups\NOMEDOSISTEMA_%DATA%.sql"
```

Troque `SENHAROOT` pela senha real do root e `NOMEDOSISTEMA` pelo
nome do seu banco.

### 9.2 Agendar

1. Abrir **Agendador de Tarefas** (Windows) → **Criar Tarefa
   Básica**.
2. Nome: "Backup diário do sistema".
3. Disparador: **Diariamente** → 23:00 (ou horário fora do
   expediente).
4. Ação: **Iniciar um programa** → Escolha o `backup.bat`.
5. Concluir.

### 9.3 Manter só os últimos 30 dias

Adiciona no fim do `backup.bat`:

```bat
forfiles /P "C:\Backups" /M *.sql /D -30 /C "cmd /c del @path"
```

Isso apaga arquivos com mais de 30 dias, evitando encher o disco.

### 9.4 IMPORTANTE — copiar backup pra outro lugar

O backup **na mesma máquina que tem o banco não vale nada** se a
máquina inteira quebrar. Configure:
- Cópia pra outra máquina da rede (via `robocopy` num .bat).
- Cópia pra um NAS.
- Cópia pra nuvem (OneDrive, Dropbox, Google Drive).

---

## Checklist final

Antes de considerar o serviço no ar:

- [ ] MariaDB instalado no servidor
- [ ] Usuário do sistema criado, com senha forte anotada
- [ ] Banco vazio criado
- [ ] Acesso pela rede liberado (bind-address + firewall)
- [ ] Tabelas criadas rodando o setup_db.py
- [ ] Dados antigos migrados (se havia)
- [ ] Teste de outra máquina funcionou
- [ ] Pelo menos 1 máquina cliente configurada e testada
- [ ] Backup automático rodando
- [ ] Backup sendo copiado pra outro lugar
- [ ] Senha do root e do usuário anotadas em lugar seguro
- [ ] Documento de como fazer isso guardado (esse arquivo!)

---

## Se precisar de ajuda

- Documentação MariaDB: <https://mariadb.com/kb/>
- PyMySQL: <https://pymysql.readthedocs.io/>
- Referência real desse fluxo: pasta do Conciliador
  (`setup_db.py`, `db.py`, `INSTALACAO_CLIENTE.md`).
