# Instalação do Conciliador (versão executável)

Passo a passo para instalar o app numa máquina de operador usando o
**ZIP pronto** (`Conciliador_YYYY-MM-DD.zip`) — sem precisar de Python,
sem Git, sem pip.

**Estimativa:** 5–10 minutos por máquina.

**Arquivo de distribuição:**
`Conciliador_2026-09-15.zip` (39 MB) — gerado pelo `publicar.bat` em
`C:\Users\PC\Projetos\conciliador\dist\`.

> O pacote passou de 25 MB para 39 MB quando entrou a leitura de
> comprovantes em PDF. **Não cabe mais em anexo de e-mail** — use a pasta
> de rede, pen drive ou nuvem.

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
├── Conciliador.exe              ← duplo-clique pra abrir
├── Atualizar Conciliador.bat    ← duplo-clique pra atualizar depois
├── LEIA-ME.txt                  ← instruções resumidas
├── VERSAO.txt                   ← versão instalada nesta máquina
├── data\
│   ├── db_config.json           ← já vem preenchido (servidor 10.0.1.47)
│   └── dominio_config.json.EXEMPLO
└── _internal\                   ← libs (não mexer)
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

### Na máquina do operador (rotina normal)

1. **Fechar o Conciliador.**
2. Duplo-clique em **`Atualizar Conciliador.bat`**, dentro de
   `C:\Conciliador`.
3. Esperar o `[OK] Atualizado.` — a mensagem já mostra qual versão ficou
   instalada.
4. Abrir o programa normalmente.

O atualizador copia a versão publicada na pasta de rede
(`\10.0.1.47\conciliador\atual`) por cima da instalação local:

- **preserva** o `data\dominio_config.json` (credenciais do Domínio da
  pessoa);
- **apaga** arquivos que saíram da versão nova — evita DLL velha
  sobrando, causa clássica de "funcionava ontem";
- **recusa rodar** com o app aberto, porque o Windows trava o `.exe` em
  uso e a cópia sairia pela metade.

Se aparecer `[ERRO]`, a atualização não foi concluída: chame o
administrador em vez de apagar pastas por conta própria.

### Do lado do administrador (publicar a versão)

No projeto, com o código já commitado:

```
publicar.bat
```

O script carimba a versão do dia em `versao.py` (é o número que aparece
na tela do app), roda o PyInstaller, monta o pacote — `data\` no nível do
`.exe`, `LEIA-ME.txt`, `VERSAO.txt` e o atualizador —, gera
`dist\Conciliador_AAAA-MM-DD.zip`, espelha tudo em
`\10.0.1.47\conciliador\atual` e guarda uma cópia do ZIP em
`\10.0.1.47\conciliador\historico\` para rollback.

Para publicar em outro lugar: `publicar.bat "\servidor\pasta\atual"`.
Se a pasta de rede não estiver acessível, o ZIP é gerado do mesmo jeito e
o script avisa.

> O endereço da pasta de rede está fixo em duas linhas: `DESTINO` no
> `publicar.bat` e `ORIGEM` no `pacote\Atualizar Conciliador.bat`. Se o
> servidor mudar, edite as duas e publique de novo — as máquinas recebem
> o atualizador corrigido junto com a versão.

### Sem pasta de rede (fallback manual)

1. **Fechar o app.**
2. Renomear a pasta atual: `C:\Conciliador` → `C:\Conciliador.bak`
3. Extrair o ZIP novo em `C:\Conciliador`
4. **Copiar** o `data\dominio_config.json` do backup pra pasta nova (pra
   não perder as credenciais do Domínio da pessoa).
5. Conferir que abre e deletar o `.bak`.

### Qual versão está rodando numa máquina

- No app: número cinza no topo, ao lado do nome do usuário — também no
  título da janela.
- Fora do app: `VERSAO.txt`, dentro de `C:\Conciliador`.

Se o número for menor que o publicado na rede, aquela máquina não rodou o
atualizador.

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

**Atualizei, mas a versão na tela continua a mesma**
→ O `Atualizar Conciliador.bat` foi rodado noutra pasta, ou a pasta de
rede ainda tem a versão antiga. Confira o `VERSAO.txt` da máquina contra
o de `\10.0.1.47\conciliador\atual`; se a rede estiver atrasada, o
administrador precisa rodar o `publicar.bat`.

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
- [ ] Máquina enxerga `\10.0.1.47\conciliador\atual` (é de onde vem a
      atualização; testar rodando o `Atualizar Conciliador.bat` uma vez)

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
| Updates | `git pull` automático | `Atualizar Conciliador.bat` (1 clique) |
| Tempo instalação | 15–30 min | 5–10 min |
