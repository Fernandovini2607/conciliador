# Colunas da planilha de controle — Conciliador

Documento para enviar ao cliente/operador que vai fornecer a planilha
de contas a pagar que o Conciliador vai conciliar com o OFX bancário
e com o Domínio.

> **O Conciliador aceita qualquer planilha .xlsx**, desde que ela tenha
> **pelo menos duas colunas obrigatórias** (vencimento e valor). As
> demais são opcionais, mas enriquecem muito a conciliação — quanto
> mais informação, mais pagamentos casam automaticamente com o
> Domínio e com o OFX.

---

## 📋 Resumo rápido

| Campo | Obrigatório? | Para que serve |
|---|---|---|
| **Data de vencimento** | ✅ Sim | Chave de match com OFX e Domínio |
| **Valor** | ✅ Sim | Chave de match (valor pago/devido) |
| Data de pagamento | ⭕ Opcional | Match preciso com OFX (data em que o dinheiro saiu do banco) |
| Nº da nota fiscal (NF) | ⭕ Opcional | Match forte com o Domínio (triple-match) |
| CNPJ do fornecedor | ⭕ Opcional | Match por fornecedor quando a NF/data divergem |
| Nome do fornecedor | ⭕ Opcional | Match por nome quando CNPJ falta |
| Histórico contábil | ⭕ Opcional | Usado como histórico padrão nos lançamentos contábeis |
| Tipo / Categoria | ⭕ Opcional | Usado como auxílio de classificação |
| Data de emissão | ⭕ Opcional | Mostra data da nota fiscal nos relatórios |

---

## 🔑 Obrigatórias

### 1. **Data de vencimento**

A data em que o pagamento deveria ocorrer (ou efetivamente ocorreu, se
a planilha só traz uma data). O sistema aceita qualquer um destes nomes
de coluna (não precisa ser exatamente este):

- Data
- Data lançamento / Data do lançamento
- Data movimento / Data mov
- Vencimento / Data vencimento / Data venc / Venc
- Data pagamento / Data pgto / Data operação
- Data crédito / Data débito
- Data compensação

**Formato aceito:** `dd/mm/aaaa`, `dd-mm-aaaa`, `aaaa-mm-dd`, `dd.mm.aaaa`
ou célula já formatada como data no Excel.

### 2. **Valor**

O valor em reais do pagamento. Nomes aceitos:

- Valor / Vlr / Vl
- Valor (R$) / Valor R$ / R$
- Valor lançamento / Valor parcela / Valor da parcela
- Total / Montante / Quantia
- Crédito / Débito / Entrada / Saída
- Amount / Value

**Formato aceito:**
- Número puro: `1234.56` ou `1234,56`
- Com separador de milhar: `1.234,56` ou `1,234.56`
- Entre parênteses (valor negativo): `(123,45)` → vira `-123,45`
- Com símbolo: `R$ 123,45` (o símbolo é ignorado)

---

## ⭕ Opcionais — enriquecem a conciliação

### 3. **Data de pagamento**

Data em que o dinheiro efetivamente saiu do banco. **Muito
recomendada** quando a planilha controla vencimento e pagamento em
colunas separadas — permite o match exato com o OFX bancário.

Nomes aceitos:

- Data pagamento / Data pgto / Data de pagamento
- Pagamento / Pago em / Dt pgto / Dt pagamento
- Data baixa / Data quitação / Data liquidação / Data compensação

### 4. **Nº da nota fiscal (NF)**

Número da nota fiscal associada ao pagamento. Essencial para o
**triple-match** com o Domínio (data + valor + NF).

Nomes aceitos:

- NF / N NF / Nº NF / Número NF
- Número da NF / Num NF / N Nota / Nº Nota
- Número / Numero documento / Doc
- Nota / Nota fiscal / Número nota fiscal

### 5. **CNPJ do fornecedor**

Nomes aceitos:

- CNPJ / CGC / CNPJ/CPF
- CNPJ fornecedor / CNPJ do fornecedor
- Documento fornecedor / Doc fornecedor

**Formato aceito:** com ou sem máscara (`12.345.678/0001-90` ou
`12345678000190`).

### 6. **Nome do fornecedor / razão social**

Nomes aceitos:

- Fornecedor / Nome fornecedor / Nome do fornecedor
- Razão social / Razao social / Razão / Razao
- Beneficiário / Favorecido / Credor
- Pagamento a / Para / Destinatário

### 7. **Histórico contábil**

Texto descritivo que vira o histórico padrão do lançamento contábil
gerado. Nomes aceitos:

- Histórico / Hist / Histórico contábil
- Descrição / Complemento / Obs / Observação
- Detalhes / Narrativa

### 8. **Tipo / Categoria**

Classificação da despesa. Usado no match por fornecedor quando mais
de um pagamento tem o mesmo valor na mesma data. Nomes aceitos:

- Tipo / Tipo de despesa / Tipo lançamento
- Classificação / Categoria / Grupo
- Natureza / Natureza da despesa

### 9. **Data de emissão da nota**

Nomes aceitos:

- Data emissão / Emissão
- Data NF / Data da nota / Data documento
- Dt emissão

---

## 📝 Exemplo de planilha mínima (só obrigatórias)

| Vencimento | Valor |
|---|---|
| 05/10/2026 | 1.234,56 |
| 10/10/2026 | 500,00 |
| 15/10/2026 | 2.789,30 |

Esta planilha já funciona. Mas a conciliação vai depender só de **data
+ valor** — se dois pagamentos têm o mesmo valor na mesma data, o
sistema pode pedir confirmação manual.

---

## 📝 Exemplo de planilha completa (recomendada)

| Vencimento | Pagamento | Valor | NF | CNPJ | Fornecedor | Histórico | Tipo |
|---|---|---|---|---|---|---|---|
| 05/10/2026 | 05/10/2026 | 1.234,56 | 12345 | 12.345.678/0001-90 | ACME LTDA | Comissão outubro | Serviço |
| 10/10/2026 | 11/10/2026 | 500,00 | 67890 | 98.765.432/0001-10 | FORNECEDOR X | Material escritório | Material |
| 15/10/2026 |            | 2.789,30 | 11111 | 11.222.333/0001-44 | ENERGIA SA | Conta de luz | Utilidade |

- A **terceira linha** não tem data de pagamento preenchida — o
  sistema vai usar o vencimento como fallback, mas a conciliação com
  o OFX fica menos precisa.
- Com todas as colunas preenchidas, o Conciliador consegue:
  - **Fase 1** (triple-match exato): data + valor + NF iguais ao Domínio
  - **Fase 2** (match por 2 de 3): quando uma das três informações diverge
  - **Fase 3** (fornecedor + valor): quando a NF do Domínio foi renegociada
  - **Fase 4** (NF + fornecedor): quando o valor pago difere da parcela (juros/desconto)
  - **Fase 5** (aprovação manual): quando a diferença de valor é > 10%
  - **Fase 6** (fornecedor forte + ~valor): quando falta NF mas o fornecedor é único

---

## ⚠️ Boas práticas

1. **Primeira linha deve ser o cabeçalho** com os nomes das colunas.
   O sistema identifica automaticamente qual coluna é qual pelo nome.
2. **Não misture pagamentos de empresas diferentes** na mesma planilha.
   Em grupo empresarial, cada filial tem a sua.
3. **Evite células mescladas** (merge) nas linhas de dados — bagunçam
   a leitura. Pode haver mescla no título/cabeçalho acima da linha de
   cabeçalho propriamente dita, se a planilha tiver.
4. **Formate como "data"** as células de data (não como texto). Se
   estiver como texto, use `dd/mm/aaaa`.
5. **Valor em uma única coluna.** Se a planilha tem Débito/Crédito em
   colunas separadas, consolide numa só antes de importar (ou use só
   uma delas se os pagamentos forem sempre do mesmo lado).
6. **Linhas em branco** no meio da planilha são ignoradas, não há
   problema.
7. **PDFs de comprovantes** (boletos Sicoob/Bradesco e PIX Sicoob)
   podem substituir ou complementar a planilha — o sistema extrai
   data, valor, NF, CNPJ e fornecedor direto do PDF.

---

## 🛠️ Se o sistema não achar as colunas automaticamente

O Conciliador tenta identificar as colunas pelo nome do cabeçalho. Se
alguma coluna estiver com nome não-padrão (ex.: "Val Pgto" no lugar
de "Valor"), ao importar a planilha o sistema abre um diálogo
**"Editar colunas"** onde o operador associa manualmente cada campo
à coluna correta da planilha. O mapeamento é salvo **por empresa**,
então só precisa ser feito uma vez.
