# Manual do Conciliador — Como usar no dia a dia

Guia passo a passo para operar o sistema. Não é sobre instalar — é
sobre **usar** depois de instalado.

---

## 🔑 Passo 1 — Entrar no sistema

1. Duplo-clique no atalho **Conciliador** na área de trabalho.
2. Digite seu **usuário** e **senha** (o admin te forneceu).
3. Clique em **Entrar** (ou aperte Enter).

**Primeira vez usando?** Clique em **Minha senha** (barra do topo) e
troque a senha inicial pela sua senha pessoal.

---

## 🏢 Passo 2 — Selecionar a empresa que vai trabalhar

1. Na barra do topo, clique em **Selecionar empresa**.
2. Escolha a empresa da lista e clique em **OK**.
3. Se aparecer uma pergunta "Trocar de empresa? Os dados carregados
   serão descartados", clique em **Sim** (é normal — cada empresa
   trabalha com dados próprios).

> A empresa que você escolher fica salva na sua conta. Da próxima vez
> que você abrir o app, vai voltar direto pra ela.

---

## 📥 Passo 3 — Carregar os dados da empresa

Você vai trabalhar com **até 3 fontes de dados** ao mesmo tempo:

### 3.1 Planilha de contas a pagar (opcional)
- Clique em **Abrir planilha (.xlsx)** e escolha o arquivo.
- Se for a primeira vez com essa empresa, o sistema pergunta qual
  coluna é o quê. Marque:
  - **Vencimento** (obrigatório)
  - **Valor** (obrigatório)
  - Outros campos (Pagamento, Fornecedor, CNPJ, Nº NF, Histórico,
    Tipo) — se a planilha não tem alguma, deixe **(deixar vazia)**.
- Clique em **Confirmar**.

> Da próxima vez que abrir uma planilha da mesma empresa, o sistema já
> lembra o mapeamento — vai direto sem perguntar.

### 3.1.b Comprovantes PDF de pagamento (alternativa/complemento)

Se a empresa **não tem planilha de controle** mas você tem
comprovantes de boleto em PDF baixados do internet banking:

- Clique em **Importar comprovantes PDF**.
- Selecione um ou mais arquivos (Ctrl+clique).
- **Bancos suportados**: Sicoob e Bradesco. Outros bancos podem ser
  adicionados sob demanda (peça pro admin).
- Aguarde o processamento (rápido em PDFs pequenos, alguns segundos
  em PDFs de 100+ páginas).
- Os comprovantes viram uma "planilha virtual" — mesmas colunas,
  mesmo fluxo de conciliação.

**Combinar planilha + comprovantes:** você pode importar planilha
.xlsx primeiro e depois os PDFs. O sistema **detecta duplicatas**
(quando o mesmo lançamento está nos dois) e importa cada um só uma
vez. O popup avisa quantos entraram e quantos foram ignorados.

> Cada comprovante PDF traz **CNPJ + nome do fornecedor** — dados
> mais completos que a maioria das planilhas de controle. Isso ajuda
> muito a conciliação com o Domínio.

### 3.2 Extrato bancário OFX
- Clique em **Importar OFX**.
- Escolha **um ou mais arquivos .OFX** (Ctrl+clique pra selecionar
  vários bancos de uma vez).
- Confirme.

### 3.3 Domínio Contábil (parcelas do sistema)
- Clique em **Carregar pagamentos**.
- Aguarde ele buscar as parcelas do Domínio da empresa.

> 💡 **Empresa matriz com filiais?** O sistema detecta automaticamente
> pelo **CNPJ raiz** (primeiros 8 dígitos) e traz as parcelas de
> **todas as empresas do grupo** — útil quando a matriz paga boletos
> emitidos contra as filiais.
> - Aparece uma mensagem listando quais empresas foram incluídas.
> - A aba **Domínio dados** ganha a coluna **Empresa (código)**
>   mostrando de onde veio cada parcela.
> - O título da aba fica **"Domínio dados (N | X empresas)"**.
> - O **plano de contas continua sendo o da matriz** — as parcelas
>   das filiais são lançadas nas contas da matriz.

### 3.4 Plano de contas (só na primeira vez do dia)
- Clique em **Carregar plano contas**.
- Ele traz todas as contas analíticas da empresa (pra usar nos
  lançamentos contábeis).

> 💡 **Selecionou uma filial?** O sistema busca automaticamente a
> **matriz do grupo** (a que tem CNPJ terminando em `/0001-XX`) e
> carrega o **plano de contas da matriz** — porque o plano contábil
> é o mesmo pro grupo todo.
> - Uma mensagem aparece dizendo qual matriz foi usada.
> - As **regras de taxa continuam separadas por empresa** (contas
>   bancárias mudam por filial, então cada uma pode ter suas próprias
>   regras de conta bancária).

---

## 🔗 Passo 4 — Conciliar planilha × OFX

Depois de carregar planilha + OFX:

1. Clique no botão **Conciliar** (barra do meio).
2. O sistema tenta casar cada linha da planilha com uma linha do OFX
   por **data + valor**.

### O que acontece:

| Onde vai | O que é |
|---|---|
| Aba **Conciliados** | Casou tudo certinho (verde = automático, azul = manual) |
| Aba **Pendentes** | O que sobrou dos dois lados (não casou) |
| Aba **Sugestões** | Pares com pequena diferença (até 2 dias e até R$ 10) |

---

## 💡 Passo 5 — Revisar as sugestões

Se tiver algo na aba **Sugestões**:

1. Vá pra aba **Sugestões**.
2. Marque as linhas que fazem sentido casar (você reconhece os
   pagamentos).
3. Clique em **Aceitar selecionadas** (pra várias) ou **Selecionar
   tudo** + Aceitar (se todas estão OK).

Elas vão pra aba **Conciliados**.

---

## 🧹 Passo 6 — Tratar os pendentes

Aba **Pendentes** tem 2 blocos:

### Bloco de cima — **Só na planilha**
Linhas que tinham na planilha mas não bateram com OFX. Podem ser:
- Pagamento em dinheiro / Caixa geral (não passou pelo banco)
- Duplicidade / erro
- Ainda não foi pago

### Bloco de baixo — **Só no OFX**
Linhas do extrato que não bateram com a planilha. Normalmente são:
- Tarifas do banco (manutenção, IOF, TED)
- Juros / rendimentos
- Pagamentos não previstos na planilha

### Como resolver cada linha do OFX (tarifas, etc.):

**Se é uma tarifa que se repete todo mês** (ex: "TARIFA PACOTE
SERVICOS"):
1. Selecione a linha.
2. Clique em **Criar regra (memo)**.
3. Preencha:
   - **Padrão**: o texto que identifica (ex: `TARIFA PACOTE`)
   - **Banco** (opcional): o nome do banco (evita confusão se banco
     diferente tem tarifa com nome parecido)
   - **Histórico contábil**: como quer que apareça no lançamento.
     **Pode deixar em branco** — nesse caso, cada lançamento vai
     usar o **Memo do OFX** da linha como histórico.
   - **Conta contábil**: escolha da lista
4. Clique OK. Todas as tarifas iguais viram lançamento automático
   agora e nas próximas conciliações.

**Se é um caso único, não recorrente**:
1. Selecione a linha.
2. Clique em **Lançamento manual**.
3. Preencha histórico + conta contábil.
4. Clique OK.

### Como resolver cada linha da planilha:

**Se é um pagamento que se repete (ex: aluguel, condomínio)**:
1. Selecione a linha.
2. Clique em **Criar regra (fornecedor)**.
3. Preencha (o padrão vem preenchido com o Tipo/CNPJ/nome):
   - Confira o padrão
   - Ajuste o histórico. **Pode deixar em branco** — nesse caso,
     cada lançamento vai usar o **Histórico da planilha** da linha
     como histórico contábil.
   - Escolha a conta contábil
4. OK. Casos iguais viram lançamento automático.

**Se é um caso único**:
1. Selecione a linha.
2. Clique em **Lançamento manual**.
3. Preencha e OK.

> Pagamentos da planilha sem OFX correspondente entram como
> **Caixa geral** (sem banco associado).

---

## 🔍 Passo 7 — Comparar com o Domínio

Depois de conciliar planilha × OFX:

1. Clique em **Comparar com Domínio** (barra do meio).
2. Vai pra aba **Comparação** automaticamente.

### Como o sistema decide que um pagamento bate com o Domínio

Ele tenta **4 formas de match, uma depois da outra** (se casou numa,
não passa pra próxima):

1. **Match exato**: data de vencimento + valor + Nº NF iguais.
2. **2 de 3**: pelo menos 2 dentre (CNPJ, data, valor) iguais. O
   terceiro pode divergir — a diferença aparece na coluna Δ Domínio.
3. **Fornecedor + valor**: valor exato E (CNPJ ou nome do fornecedor
   bate). Data pode divergir — útil pra parcelas renegociadas /
   vencimento prorrogado.
4. **NF + fornecedor** (valor livre): Nº NF igual (obrigatório e
   não-vazio) E (CNPJ ou nome do fornecedor bate). **Valor pode
   divergir** — útil pra pagamentos com juros/multa/desconto.

> Cada parcela do Domínio só pode ser vinculada a **um pagamento** —
> não gera duplicidade.
>
> **Exceção — parcelas com status `Parcial`**: quando o Domínio marca
> a parcela como `Parcial` (indicando que ela recebe vários pagamentos
> parciais), a **Fase 4** permite vincular a mesma parcela a **vários
> pagamentos** ao mesmo tempo. Isso reflete o caso real de um boleto
> pago em partes — todos os pagamentos ficam ligados à mesma parcela
> do Domínio.

Na aba Comparação você vê **6 cores**:

| Cor | Significado | O que fazer |
|---|---|---|
| 🟢 Verde (OK) | Já bateu no Domínio, tudo certo | Nada — está fechado |
| 🟡 Amarelo (Falta dom) | Conciliado no banco, falta lançar no Domínio | Criar lançamento ou regra |
| 🔵 Azul (Caixa OK) | Caixa geral já no Domínio | Nada — está fechado |
| ⚪ Cinza (Caixa falta) | Caixa geral, falta lançar | Criar lançamento ou regra |
| 🩵 Ciano (OFX OK) | Extrato bateu no Domínio | Nada — está fechado |
| 🟠 Laranja (OFX falta) | Extrato sem lançamento no Domínio | Criar lançamento ou regra |

### Filtrar por cor

No topo da aba Comparação tem o campo **Filtrar por cor**. Escolha uma
das 6 cores (ou "Todos") pra ver **só os lançamentos daquela cor** —
útil quando você quer focar só nos amarelos, cinzas ou laranjas
(pendências). O botão **Limpar** volta pra "Todos".

> O filtro é só visual — não altera os totais no título da aba nem
> afeta a exportação de pendências.

### Ações na aba Comparação:

- **Editar dados**: se a planilha tem dados errados (NF errada, valor
  digitado errado) que impediram o match com o Domínio, você corrige
  aqui e o sistema tenta o match de novo.
- **Lançar manualmente**: cria lançamento contábil pra essa linha.
- **Criar regra de fornecedor**: cria uma regra que também vai pegar
  outras linhas similares.
- **Exportar pendências**: gera um Excel com todas as linhas **falta**
  (amarelas, cinzas, laranjas) — útil pra revisar em planilha ou
  mandar pra alguém.

---

## 📒 Passo 8 — Revisar os lançamentos contábeis

Aba **Lançamentos contábeis** mostra tudo que vai virar lançamento
no Domínio:

- Cada linha tem: data, banco, valor, conta, histórico, regra que
  gerou.
- **Botões:**
  - **Editar lançamento** — muda data/valor/conta/histórico de uma
    linha específica sem afetar a regra.
  - **Excluir lançamento** — remove essa linha. A transação volta
    pra aba Pendentes.
  - **Exportar para Excel** — gera .xlsx com todos os lançamentos +
    linha TOTAL no fim.

> Essa aba é o **produto final** do seu trabalho. É o que você vai
> lançar no Domínio (manualmente ou via importação, dependendo do
> processo do escritório).

---

## 📤 Passo 9 — Exportar relatórios (Excel)

Você pode exportar em várias abas pra revisar ou entregar pra alguém:

- **Aba Pendentes** → botão *Exportar para Excel*: gera .xlsx com 2
  abas (planilha e OFX pendentes).
- **Aba Conciliados × Domínio** → botão *Exportar para Excel*: tudo
  que já está fechado com o Domínio.
- **Aba Lançamentos contábeis** → botão *Exportar para Excel*: os
  lançamentos que vão pro Domínio, com totalização.
- **Aba Comparação** → botão *Exportar pendências*: só o que falta
  (amarelo + cinza + laranja).

Os arquivos são salvos com nome sugerido:
`<tipo>_<código_empresa>_<data>.xlsx` (ex: `lancamentos_55_2026-09-14.xlsx`).

---

## 🔁 Fluxo típico do dia a dia

Resumo do que você faz a cada empresa/período:

```
1. Abre o app → login
2. Seleciona empresa
3. Carrega planilha + OFX + Domínio
4. Clica em Conciliar
5. Aceita sugestões que fazem sentido
6. Trata os Pendentes (cria regras / lançamentos)
7. Clica em Comparar com Domínio
8. Trata linhas amarelas/cinzas/laranjas
9. Vai em Lançamentos contábeis → confere
10. Exporta pra Excel se precisar
11. Fecha o app
```

Da próxima vez que abrir para a **mesma empresa**, o sistema já
lembra:
- O mapeamento das colunas da planilha
- Todas as regras que você criou
- A empresa que estava selecionada

**As regras cadastradas por qualquer usuário ficam disponíveis pra
todos** que trabalham nessa empresa — não precisa cadastrar de novo
quando é outro operador.

---

## 🧹 Trocar OFX ou planilha no meio do trabalho

Se você percebeu que importou o OFX errado (ou a planilha errada) e
já fez conciliação em cima:

**Limpar OFX**:
- Fecha só o OFX. As **conciliações que você já fez ficam preservadas**.
- A planilha (ou os comprovantes PDF) que estão carregados continuam.
- Ao importar o novo OFX e clicar Conciliar de novo, o sistema
  **não vai refazer** as linhas que já estavam conciliadas — só as
  que sobraram como pendentes tentam casar com o novo OFX.

**Limpar planilha**:
- Simetrico: fecha só a planilha, OFX continua, conciliações
  preservadas.

> Utilíssimo pra corrigir um OFX incompleto sem perder 20 minutos de
> classificação já feita.

**Popup de confirmação** te avisa quantas conciliações e lançamentos
serão preservados antes de aplicar.

---

## 🧑‍💻 Uso compartilhado (multi-usuário)

- **Cada operador tem seu próprio usuário e senha.**
- **Cada operador pode estar em uma empresa diferente ao mesmo tempo**
  (Aline na empresa 55, Carlos na 82) — não interfere um no outro.
- **Regras e mapeamentos são da EMPRESA** (compartilhados). Se você
  criar uma regra de tarifa da empresa 55, seu colega vê a mesma regra
  quando trabalhar na 55.
- **Dados temporários** (planilha carregada, OFX importado) são só na
  sua tela — quando você fecha o app, some. Se precisar continuar
  depois, exporte pra Excel.

---

## ❓ Perguntas comuns

**Errei uma regra, como corrijo?**
- Botão **Configurar taxas** (barra do meio) → seleciona a regra →
  Editar ou Excluir.

**Meu operador colega criou uma regra errada, o que faço?**
- Igual acima. Regras são compartilhadas — qualquer um pode
  ajustar/apagar.

**Quero trocar de empresa no meio do trabalho, perco tudo?**
- Sim: planilha/OFX carregados são descartados (o sistema avisa antes).
  Exporte pra Excel se quiser guardar.

**Tem como "salvar" o trabalho pra continuar depois?**
- Só via export pra Excel. As **regras** e **mapeamentos** ficam
  salvos automaticamente no servidor — só os dados carregados
  (planilha/OFX) são temporários.

**Como sei que uma regra automática realmente pegou minha linha?**
- Vá na aba **Lançamentos contábeis** — se a linha está lá, virou
  lançamento (a coluna "Regra" mostra qual padrão casou).

**A conta contábil errada foi escolhida numa regra que já pegou
várias linhas. Como corrijo tudo?**
- Vai em **Configurar taxas** → edita a regra corrigindo a conta.
- O sistema **NÃO** reaplica automaticamente nas linhas que já
  viraram lançamento. Pra corrigir cada uma: vai em
  **Lançamentos contábeis** → seleciona a linha errada → **Editar
  lançamento** → muda a conta manualmente.

**Um lançamento automático saiu errado, mas a regra está certa. E
agora?**
- Vai em **Lançamentos contábeis** → seleciona → **Editar** (muda o
  que precisar) OU **Excluir** (a transação volta pra Pendentes).

**Como troco minha senha?**
- Botão **Minha senha** (barra do topo) → coloca senha atual + nova.

**Esqueci minha senha, como reseto?**
- Fale com um admin (Vinicius). Só admin pode resetar senhas de
  outros usuários no **Gerenciar usuários**.

---

## 🚨 Problemas comuns

**"Auto-conexão do Domínio falhou..." em vermelho no topo**
→ O sistema não conseguiu conectar no Domínio Contábil. Verifique se
o Domínio está aberto/rodando na sua máquina. Se persistir, fale com
o admin.

**Tela de login não aparece / app não abre**
→ Verifique se o servidor de banco está acessível. Falar com o admin.

**Após aceitar sugestões, algumas voltam pra Pendentes**
→ Verifique se as datas e valores realmente batem. Se batem mas o
sistema não aceita, exporte a Comparação pra planilha e mande pro
admin analisar.

**Clico em "Comparar com Domínio" e não abre a aba**
→ Confirma que você clicou em **Carregar pagamentos** antes (o
Domínio precisa estar carregado).

---

## 📞 Suporte

Problema no sistema? Fale com o admin (Vinicius) informando:
1. O que você estava fazendo (passo do manual acima).
2. Qual mensagem apareceu (screenshot ajuda muito).
3. Qual empresa estava trabalhando.
