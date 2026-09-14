# Instalação do Conciliador (versão executável)

Passo a passo para instalar o app numa máquina de operador usando o
**ZIP pronto** (`Conciliador_YYYY-MM-DD.zip`) — sem precisar de Python,
sem Git, sem pip.

**Estimativa:** 5–10 minutos por máquina.

**Arquivo de distribuição:**
`Conciliador_2026-09-14.zip` (25 MB) — gerado em
`C:\Users\PC\Projetos\conciliador\dist\`.

---

## PARTE 1 — Pré-requisitos (só na 1ª instalação da máquina)

### 1.1 Driver ODBC do Domínio + DSN

Necessário para que o operador consiga carregar dados do Domínio Contábil.
Na maioria das máquinas de contadores isso **já está instalado** (é
requisito do Domínio). Se não estiver:

1. Instalar o **SQL Anywhere ODBC Driver 17** (vem com a instalação do
   Domínio ou baixe do site da Thomson Reuters).
2. Criar a DSN de sistema:
   - Painel de Controle → **Fontes de Dados ODBC (64 bits)**.
   - Aba **DSN de Sistema** → **Adicionar** → escolher `SQL Anywhere 17`.
   - Nome da DSN: **`Contabil`** (nome exato, é o que o app procura).
   - Server name: `janco`
   - Database name: `contabil`
   - Salvar e testar.

> Se a máquina já usa o Domínio Contábil no dia a dia, esse passo
> normalmente **já está pronto** — pode pular pra Parte 2.

---

## PARTE 2 — Instalar o Conciliador

### 2.1 Copiar o ZIP pra máquina do operador

Escolha uma das opções:
- Pen drive
- Pasta compartilhada da rede
- OneDrive/Google Drive
- Email (25 MB, cabe)

### 2.2 Extrair a pasta

1. Botão direito no `Conciliador_2026-09-14.zip` → **Extrair Tudo…**
2. Extrair em **`C:\Conciliador`** (ou outro local fixo — mas evite
   `Downloads`, `Área de Trabalho` ou pastas temporárias).

Estrutura extraída:
```
C:\Conciliador\
├── Conciliador.exe          ← duplo-clique pra abrir
├── LEIA-ME.txt              ← instruções resumidas
├── data\
│   ├── db_config.json       ← já vem preenchido (servidor 10.0.1.47)
│   └── dominio_config.json.EXEMPLO
└── _internal\               ← libs (não mexer)
```

### 2.3 Configurar credenciais do Domínio da pessoa

Cada operador tem seu próprio login no Domínio Contábil.

1. Abrir `C:\Conciliador\data\`
2. **Copiar** o arquivo `dominio_config.json.EXEMPLO`
3. **Renomear** a cópia para `dominio_config.json` (tirar o `.EXEMPLO`)
4. Abrir no Bloco de Notas e substituir:

```json
{
  "dsn": "Contabil",
  "usuario": "USUARIO_DOMINIO_DA_PESSOA",
  "senha": "SENHA_DOMINIO_DA_PESSOA"
}
```

5. Salvar.

### 2.4 Criar atalho na área de trabalho (recomendado)

1. Botão direito em `C:\Conciliador\Conciliador.exe`
2. **Enviar para** → **Área de Trabalho (criar atalho)**
3. Renomear o atalho para **"Conciliador"**
4. (Opcional) Botão direito no atalho → **Propriedades** → **Alterar Ícone**

---

## PARTE 3 — Cadastrar o usuário no app (feito pelo admin)

Antes de a pessoa fazer login pela primeira vez, o **admin (Vinicius)**
precisa cadastrá-la no sistema:

1. Vinicius abre o app na máquina dele.
2. Clica em **Gerenciar usuários** (barra do topo).
3. **+ Novo usuário**.
4. Preenche:
   - **Username**: sem espaços, minúsculo (ex: `aline`, `carlos`)
   - **Nome completo**: nome real
   - **Email**: opcional
   - **Senha inicial**: qualquer coisa fácil (ex: `mudar123`)
   - **Perfil administrador**: deixar **desmarcado** (é operador)
5. Salvar.
6. Avisar a pessoa por Teams/WhatsApp/email:
   *"Seu login: `aline`, senha inicial `mudar123`. Troque assim que
   entrar em 'Minha senha'."*

---

## PARTE 4 — Uso normal

### 4.1 Primeira vez

1. Duplo-clique no atalho **Conciliador** na área de trabalho.
2. Tela de login aparece.
3. Digitar username e senha fornecidos pelo admin.
4. Entrar.
5. Clicar em **Minha senha** (barra do topo) → trocar pela senha
   pessoal.

### 4.2 Dia a dia

1. Duplo-clique no atalho.
2. Login com usuário e senha próprios.
3. Trabalhar normalmente. Empresa selecionada e trabalho ficam salvos
   no servidor — se abrir na terça de novo, volta onde estava.

---

## Como atualizar quando sair uma versão nova

Quando o Vinicius gerar uma versão nova:

1. Ele avisa que tem `Conciliador_YYYY-MM-DD.zip` novo (pasta
   compartilhada, email, etc.).
2. Na máquina do operador:
   - **Fechar o app** se estiver aberto.
   - Renomear a pasta atual (backup): `C:\Conciliador` → `C:\Conciliador.bak`
   - Extrair o novo ZIP em `C:\Conciliador`
   - **Copiar** o arquivo `data\dominio_config.json` do backup pra pasta
     nova (pra não perder as credenciais do Domínio da pessoa).
   - Rodar de novo.
   - Se tudo funcionar, deletar o `.bak`.

Alternativa mais simples (se o admin distribuir só a nova `Conciliador.exe`
sem tocar em `data/`):

- **Substituir apenas o `Conciliador.exe` e a pasta `_internal/`** na
  pasta existente. Os arquivos em `data/` ficam intocados.

---

## Solução de problemas

**"O aplicativo não pôde ser iniciado corretamente" ao abrir o .exe**
→ Faltou algum arquivo do build. Extraia o ZIP de novo, tudo junto.

**"Erro no banco: (2003, ...) Can't connect to MySQL server"**
→ A máquina não está enxergando o servidor `10.0.1.47:3307`.
- Verificar cabo/rede.
- Testar `ping 10.0.1.47` no cmd.
- Confirmar que o servidor está ligado.

**"Sem usuários cadastrados no banco"**
→ O `db_config.json` está errado (talvez apontando pra outro banco).
Substitua pelo original do ZIP.

**"Auto-conexão do Domínio falhou..."** (aparece no label do Domínio,
em vermelho)
→ Ou o `dominio_config.json` não existe / está errado, ou a DSN
"Contabil" do ODBC não foi criada.
- Verificar que o arquivo `data\dominio_config.json` existe (não
  `.EXEMPLO`).
- Testar a DSN no Painel de Controle → ODBC → seleciona "Contabil" →
  **Configurar** → **Testar Conexão**.

**Login diz "Usuário ou senha inválidos"**
→ Confirmar com o admin (Vinicius) que o cadastro foi feito e qual é a
senha inicial. Case-sensitive (Aline ≠ aline).

**Depois de digitar a senha, a tela de login fecha e nada acontece**
→ Provavelmente o EXE crashou. Rode ele pelo cmd pra ver o erro:
```
cd C:\Conciliador
Conciliador.exe
```

---

## Checklist rápido pra cada máquina

- [ ] Driver ODBC + DSN "Contabil" (normalmente já feito)
- [ ] ZIP extraído em `C:\Conciliador`
- [ ] `data\dominio_config.json` criado (a partir do `.EXEMPLO`)
- [ ] Atalho na área de trabalho
- [ ] Usuário cadastrado no app (via Vinicius no "Gerenciar usuários")
- [ ] Login testado com sucesso
- [ ] Senha trocada em "Minha senha"

---

## Diferenças em relação ao método antigo (Git)

| Item | Método antigo (Git) | Método novo (EXE) |
|---|---|---|
| Instalar Python | ✅ Precisava | ❌ Não precisa |
| Instalar Git | ✅ Precisava | ❌ Não precisa |
| Acesso GitHub | ✅ Precisava | ❌ Não precisa |
| Criar venv + pip | ✅ Precisava | ❌ Não precisa |
| Configurar db_config | Manual | ✅ Já vem no ZIP |
| Configurar dominio_config | Manual | Manual (idem) |
| DSN "Contabil" | Manual | Manual (idem) |
| Updates | `git pull` automático | Substituir a pasta manualmente |
| Tempo instalação | 15–30 min | 5–10 min |
