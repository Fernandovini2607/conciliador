# Conciliador OFX × Planilha × Domínio

Aplicativo desktop em Python que faz conciliação em três níveis:

1. **Planilha (Excel) × Extrato (OFX)** — confere se os pagamentos lançados na sua planilha bateram no extrato bancário.
2. **Conciliação manual e sugestões aproximadas** — pares com diferença de até 2 dias / R$ 10,00 viram sugestões para análise.
3. **Conciliados × Sistema Domínio (ODBC)** — confere se cada pagamento conciliado também está lançado na sua contabilidade.

## Instalação

1. Instale o Python 3.11+ ([python.org/downloads](https://www.python.org/downloads/)).
   Marque **"Add Python to PATH"** no instalador.
2. PowerShell na pasta do projeto:

   ```powershell
   cd C:\Users\PC\Projetos\conciliador
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

   Se o PowerShell bloquear o `Activate.ps1`:
   `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.

## Como usar

```powershell
python main.py
```

### Fluxo básico (planilha + OFX)

1. **Abrir planilha (.xlsx)** — selecione o arquivo. Vai aparecer o diálogo de mapeamento com preview ao vivo. Confirme/ajuste as colunas Data, Valor e Descrição.
2. **Importar OFX** — selecione o extrato. Recebimentos (valores positivos) são automaticamente ignorados; o app só compara pagamentos.
3. **Conciliar** — gera as abas:
   - **Conciliados**: pares verdes (automáticos por data+valor exatos) e azuis (manuais).
   - **Pendentes**: dois lados (só na planilha / só no OFX). Selecione um de cada lado e clique em "Conciliar selecionadas →" para casar manualmente.
   - **Sugestões**: pares com diferença de até 2 dias e até R$ 10,00. Selecione uma sugestão e clique em "Aceitar como conciliação" para promovê-la.

### Comparação com o Domínio (ODBC)

#### Pré-requisito: configurar o DSN ODBC

O Domínio usa **Sybase SQL Anywhere** como banco. Antes de usar o app, configure o DSN no Windows:

1. **Painel de Controle → Ferramentas Administrativas → ODBC Data Sources (64-bit)**.
2. Aba **Drivers**: confirme que existe `SQL Anywhere [versão]`. Se não existir, instale o **SQL Anywhere Client** correspondente à versão do seu Domínio.
3. Aba **System DSN → Add**: escolha o driver `SQL Anywhere`, preencha:
   - **Data Source Name**: `DOMINIO` (ou outro nome — você informa no app).
   - **Server Name**: nome do serviço SQL Anywhere do Domínio (procure no `services.msc` algo como `SQL Anywhere - Dominio`).
   - **Database Name**: a base contábil da empresa.
   - **Login**: usuário e senha (padrão antigo do Domínio: `dba` / `sql`).
4. Clique em **Test Connection** — precisa retornar "Connection successful".

#### No app

1. **Conectar Domínio** → informe DSN, usuário, senha. Se OK, fica salvo em `config.json` local.
2. **Configurar fonte** → escolha entre:
   - **Tabela + filtro**: dropdown lista todas as tabelas do banco. Selecione uma → "Carregar amostra" → veja as primeiras 15 linhas → mapeie as colunas Data/Valor/Descrição. Pode adicionar um WHERE opcional (ex.: `data_pagto BETWEEN '2026-01-01' AND '2026-01-31'`).
   - **Query SQL manual**: cole uma instrução SELECT pronta. Carregue a amostra e mapeie as colunas pelos nomes retornados.
3. **Carregar pagamentos** → o app executa a query e lista os pagamentos do Domínio.
4. **Comparar com Domínio** (habilita após conciliar e carregar) → preenche a aba **Domínio**:
   - **Verde** — Conciliado e no Domínio: tudo em ordem.
   - **Amarelo** — Conciliado mas falta no Domínio: pagamento bateu banco × planilha, mas não há lançamento contábil.
   - **Vermelho** — Só no Domínio: lançamento contábil sem reflexo no extrato/planilha.

A configuração (DSN, credenciais, tabela/SQL, mapeamento) é salva em `config.json` na pasta do projeto. **Senha fica em texto claro** — só rode na sua máquina.

## Estrutura do projeto

- `main.py` — janela Tkinter, ponto de entrada.
- `parser_xlsx.py` — leitura da planilha, detecção de cabeçalho e colunas.
- `parser_ofx.py` — leitura do OFX (filtra pagamentos).
- `parser_dominio.py` — conexão ODBC e extração de pagamentos do Domínio.
- `dialogos_dominio.py` — diálogos da UI para conectar e selecionar fonte do Domínio.
- `matcher.py` — algoritmo de conciliação e geração de sugestões.
- `config.py` — persistência simples em `config.json`.
- `requirements.txt` — dependências.
