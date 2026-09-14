# Instalação do Conciliador em máquina cliente

Passo a passo para instalar o app numa máquina nova de operador
(ex: Aline, ou qualquer outro usuário) que vai conectar no servidor
MariaDB da rede.

**Estimativa:** 15–30 minutos por máquina.

---

## PARTE 1 — Pré-requisitos (só na 1ª instalação)

### 1.1 Instalar Python

1. Baixar: <https://www.python.org/downloads/>
2. Instalador oficial (versão 3.11 ou mais nova).
3. **IMPORTANTE:** na tela de instalação marque **"Add Python to PATH"**.
4. Confirme abrindo `cmd` e digitando:
   ```
   python --version
   ```
   Deve responder `Python 3.11.x` (ou similar).

### 1.2 Instalar Git

1. Baixar: <https://git-scm.com/download/win>
2. Instalador padrão (Next → Next → Next → Install).
3. Confirme:
   ```
   git --version
   ```

### 1.3 Acesso ao repositório privado do GitHub

O repositório `Fernandovini2607/conciliador` é **privado**. Escolha UMA das
duas opções:

**Opção A — Convidar o usuário como colaborador (mais simples):**
- No GitHub, vá em `Settings → Collaborators → Add people`.
- Adicione o usuário GitHub da pessoa (ex: `aline_janco`).
- Ela recebe convite por email e aceita.

**Opção B — Personal Access Token (não precisa conta GitHub por pessoa):**
- Você cria um PAT no seu GitHub (Settings → Developer settings → Tokens).
- Marque só o escopo `repo`.
- Copia o token e coloca na URL do clone (passo 2.2 abaixo).

### 1.4 Driver ODBC do Domínio

Necessário só se o operador vai carregar pagamentos/plano de contas do Domínio.

1. Instalar o **SQL Anywhere ODBC Driver** (vem junto com a instalação
   do Domínio Contábil, ou baixe do site da Thomson Reuters).
2. Criar a DSN de sistema:
   - Painel de Controle → `Fontes de Dados ODBC (64 bits)`.
   - Aba `DSN de Sistema` → **Adicionar** → escolher `SQL Anywhere 17`.
   - Nome da DSN: **`Contabil`** (exatamente esse nome — é o que o app
     procura por padrão).
   - Server name: `janco`
   - Database name: `contabil`
   - Salvar e testar.

---

## PARTE 2 — Instalar o app

### 2.1 Escolher pasta

Sugestão: `C:\Conciliador` (evita espaços/acentos no caminho).

Abra o `cmd` e vá pra pasta que vai conter o projeto:
```
cd C:\
```

### 2.2 Clonar o repositório

**Se usou Opção A (colaborador):**
```
git clone https://github.com/Fernandovini2607/conciliador.git Conciliador
```

**Se usou Opção B (PAT):** troque `SEU_TOKEN_AQUI` pelo token gerado:
```
git clone https://SEU_TOKEN_AQUI@github.com/Fernandovini2607/conciliador.git Conciliador
```

### 2.3 Criar ambiente virtual (venv)

```
cd C:\Conciliador
python -m venv .venv
```

### 2.4 Instalar dependências

```
.\.venv\Scripts\activate.bat
pip install -r requirements.txt
```

Se aparecer `Successfully installed ...`, tá pronto. Pode fechar o cmd.

---

## PARTE 3 — Configurar credenciais

Todos os arquivos de credencial ficam na pasta `data\` do projeto.
Ela vem vazia depois do clone — você precisa criá-los.

### 3.1 Conexão com o MariaDB da rede

Crie o arquivo `C:\Conciliador\data\db_config.json` com este conteúdo
**exato**:

```json
{
  "host": "10.0.1.47",
  "port": 3307,
  "user": "conciliador_master",
  "password": "RjE6ymYkTASbEyYpuFb6",
  "database": "conciliador",
  "charset": "utf8mb4"
}
```

**Atalho:** copie o `data\db_config.json` da sua máquina (que já está
apontando pro servidor) diretamente pra máquina do operador.

### 3.2 Conexão com o Domínio (ODBC)

Crie `C:\Conciliador\data\dominio_config.json`:

```json
{
  "dsn": "Contabil",
  "usuario": "SEU_USUARIO_DOMINIO",
  "senha": "SUA_SENHA_DOMINIO"
}
```

**Substitua** pelos dados de login que a pessoa usa no Domínio Contábil.
Cada operador pode ter um user próprio do Domínio, ou usar um user
compartilhado — depende da política do escritório.

### 3.3 Testar a conexão com o servidor

Abra `cmd` na pasta do projeto:
```
cd C:\Conciliador
.\.venv\Scripts\python.exe -c "import db; ok, msg = db.testar_conexao(); print(msg)"
```

Deve responder algo tipo `Conectado - MariaDB/MySQL 12.2.2-MariaDB`.

Se der erro:
- Verifique se o IP `10.0.1.47` está acessível da rede da máquina
  (`ping 10.0.1.47` no cmd).
- Confirme que o servidor está ligado e a porta `3307` liberada.
- Confira se o `db_config.json` foi criado no caminho certo.

---

## PARTE 4 — Uso normal

### 4.1 Iniciar o app

Duplo-clique em **`iniciar.bat`** dentro da pasta `C:\Conciliador`.

O que ele faz automaticamente:
1. Faz `git pull` (atualiza o código).
2. Ativa o venv.
3. Instala/atualiza dependências.
4. Abre o app.

### 4.2 Login

A tela de login aparece.
- **Username e senha:** os que o admin (Vinicius) cadastrou pra essa
  pessoa via **"Gerenciar usuários"** no app.

Se ela ainda não tem cadastro, o Vinicius precisa criar antes:
- Ele loga no app na máquina dele → **Gerenciar usuários** → **Novo
  usuário** → cadastra username, nome, email, senha inicial, marca
  "Perfil administrador" só se ela também for admin.
- Depois avisa a pessoa qual é o username e a senha inicial.
- Na primeira entrada, ela troca a senha em **"Minha senha"**.

---

## Atalho na área de trabalho

Pra facilitar o dia a dia do operador:

1. Botão direito em `C:\Conciliador\iniciar.bat` → **Enviar para** →
   **Área de trabalho (criar atalho)**.
2. Renomear o atalho pra **"Conciliador"**.
3. (Opcional) Alterar o ícone: botão direito no atalho → Propriedades
   → Alterar Ícone → escolher um ícone do sistema.

---

## Solução de problemas comuns

**"git não é reconhecido como comando"**
→ Git não foi instalado ou não está no PATH. Reinstale marcando "add to PATH".

**"python não é reconhecido"**
→ Idem, mas com Python.

**"Sem usuários cadastrados"**
→ O `db_config.json` está apontando pro banco errado (talvez local,
não o servidor). Corrija o IP.

**"Falha na auto-conexão do Domínio"**
→ O DSN `Contabil` não foi criado, ou o usuário/senha em
`dominio_config.json` está errado.

**"[ERRO] Setup do banco falhou"**
→ O `iniciar.bat` tentou rodar `setup_db.py` porque o `db_config.json`
não existe. Crie ele manualmente (passo 3.1) e reabra.

**Após atualização, o app não abre**
→ Rode manualmente:
```
cd C:\Conciliador
.\.venv\Scripts\activate.bat
python main.py
```
E copie o erro pra reportar ao admin.

---

## Checklist rápido pra cada máquina

- [ ] Python instalado
- [ ] Git instalado
- [ ] Acesso ao repositório privado
- [ ] Driver ODBC do Domínio + DSN "Contabil" configurada
- [ ] Repositório clonado em `C:\Conciliador`
- [ ] Venv criado + dependências instaladas
- [ ] `data\db_config.json` com IP do servidor
- [ ] `data\dominio_config.json` com credenciais Domínio
- [ ] Teste de conexão funcionou
- [ ] Usuário cadastrado no app (via Vinicius)
- [ ] Atalho na área de trabalho
- [ ] Login testado com sucesso
