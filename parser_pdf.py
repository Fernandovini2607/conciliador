"""Parser de comprovantes de pagamento em PDF (nativos digitais).

Extrai valor, data, beneficiário e CNPJ de comprovantes de boleto pagos
gerados pelo internet banking. Suporta PDFs com múltiplos comprovantes
por arquivo — cada comprovante vira uma Transacao no fluxo do app.

Bancos suportados hoje:
- Sicoob (SISBR / SISTEMA DE INFORMÁTICA DO SICOOB)
- Bradesco (Bradesco NET EMPRESA)

Para adicionar um banco novo:
1. Estude o texto extraído do PDF (extrair_texto_bruto ajuda).
2. Adicione função _detectar_<banco>(texto) → bool.
3. Adicione função _parsear_<banco>(texto) → list[Transacao].
4. Registre em BANCOS_SUPORTADOS.

Tolerante a acentos quebrados (\\ufffd) e strings ascii-quebradas
(comum em PDFs com fontes não-padrão).
"""

from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

import pdfplumber

from parser_xlsx import Transacao


# -------------------------------------------------------- utilidades

def _para_data(txt: str) -> date | None:
    txt = (txt or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(txt, fmt).date()
        except ValueError:
            continue
    return None


def _para_valor(txt: str) -> Decimal | None:
    """Converte 'R$ 2.233,99' → Decimal('2233.99')."""
    if not txt:
        return None
    limpo = re.sub(r"[^\d,.\-]", "", txt)
    if "," in limpo and "." in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    elif "," in limpo:
        limpo = limpo.replace(",", ".")
    try:
        return Decimal(limpo)
    except InvalidOperation:
        return None


def _limpar_cnpj(txt: str) -> str:
    """Mantém só dígitos + ponto/barra/hífen do CNPJ."""
    return re.sub(r"[^\d./\-]", "", txt or "").strip()


def _limpar_nome(txt: str) -> str:
    """Remove espaços duplicados e quebras estranhas."""
    return re.sub(r"\s+", " ", (txt or "")).strip()


# -------------------------------------------------------- Sicoob

def _detectar_sicoob(texto: str) -> bool:
    # Usa marcadores FORTES — só aparecem no cabeçalho do próprio comprovante
    # Sicoob. Não usa a palavra "SICOOB" solta porque ela pode aparecer no
    # nome do banco DESTINATÁRIO de um comprovante Bradesco/Itaú/etc.
    marcadores_fortes = (
        "SISBR",                      # abreviação exclusiva do sistema Sicoob
        "OUVIDORIA SICOOB",           # rodapé fixo
        "SISTEMA DE INFORM",          # "SISTEMA DE INFORMÁTICA DO SICOOB"
    )
    return any(m in texto.upper() for m in marcadores_fortes)


def _parsear_sicoob(texto: str, arquivo: str) -> list[Transacao]:
    """Sicoob: cada comprovante começa em 'COMPROVANTE DE ... PAGAMENTO DE BOLETO'
    e termina em 'Autenticação:'. Split por marcador de início."""
    marcador_inicio = "PAGAMENTO DE BOLETO"
    blocos: list[str] = []

    partes = texto.split(marcador_inicio)
    # partes[0] é lixo antes do 1º comprovante; partes[1:] são conteúdos
    for parte in partes[1:]:
        # Cada bloco vai até "Autenticação:" (inclusive) ou até o próximo "COMPROVANTE DE"
        m = re.search(r"Autentica[c\u00e7\ufffd]{1,3}o:\s*[\w-]+", parte)
        if m:
            blocos.append(parte[:m.end()])
        else:
            blocos.append(parte)

    transacoes: list[Transacao] = []
    for i, bloco in enumerate(blocos, 1):
        t = _extrair_transacao_sicoob(bloco, arquivo, ordem=i)
        if t is not None:
            transacoes.append(t)
    return transacoes


def _extrair_transacao_sicoob(bloco: str, arquivo: str, ordem: int) -> Transacao | None:
    # Valor pago
    m_valor = re.search(r"Pago:\s*R\$\s*([\d.,]+)", bloco)
    valor = _para_valor(m_valor.group(1)) if m_valor else None
    if valor is None:
        return None

    # Data de pagamento
    m_data = re.search(r"Pagamento:\s*(\d{2}/\d{2}/\d{4})", bloco)
    if not m_data:
        m_data = re.search(r"Realizado:\s*(\d{2}/\d{2}/\d{4})", bloco)
    data_pagto = _para_data(m_data.group(1)) if m_data else None
    if data_pagto is None:
        return None

    # Data de vencimento (opcional)
    m_venc = re.search(r"Vencimento:\s*(\d{2}/\d{2}/\d{4})", bloco)
    data_venc = _para_data(m_venc.group(1)) if m_venc else data_pagto

    # Beneficiário — bloco entre "Beneficiário:" e "Pagador:"
    # Prioridade: quando o comprovante traz "Benefici\u00e1rio final:" (boleto
    # que passou por intermedi\u00e1rio / agente de cobran\u00e7a), esse \u00e9 o
    # fornecedor real. Quando n\u00e3o traz, cai no bloco "Benefici\u00e1rio:".
    fornecedor = ""
    cnpj = ""
    m_bloco_final = re.search(
        r"Benefici[a\u00e1\ufffd]rio final:(.*?)Datas:", bloco, re.DOTALL,
    )
    bloco_benef_usado = m_bloco_final.group(1) if m_bloco_final else None
    if bloco_benef_usado is None:
        m_bloco_benef = re.search(
            r"Benefici[a\u00e1\ufffd]rio:(.*?)Pagador:", bloco, re.DOTALL,
        )
        if m_bloco_benef:
            bloco_benef_usado = m_bloco_benef.group(1)
    if bloco_benef_usado:
        m_nome = re.search(
            r"Nome/Raz[a\u00e3\ufffd]{1,2}o [Ss]ocial:\s*(.+)",
            bloco_benef_usado,
        )
        if m_nome:
            fornecedor = _limpar_nome(m_nome.group(1).split("\n")[0])
        m_cnpj = re.search(
            r"CPF/CNPJ:\s*([\d./\-]+)", bloco_benef_usado,
        )
        if m_cnpj:
            cnpj = _limpar_cnpj(m_cnpj.group(1))

    # Número do documento
    m_doc = re.search(r"N[u\u00fa\ufffd]mero do documento:\s*(\S+)", bloco)
    doc = m_doc.group(1).strip() if m_doc else ""

    return Transacao(
        data=data_venc,
        valor=valor,
        descricao="",
        origem="pdf",
        linha=ordem,
        extras={
            "fornecedor": fornecedor,
            "cnpj": cnpj,
            "numero_nf": doc,
            "historico": f"Boleto Sicoob {doc}".strip(),
            "arquivo": arquivo,
            "banco_pdf": "Sicoob",
        },
        data_pagamento=data_pagto,
    )


# -------------------------------------------------------- Bradesco

def _detectar_bradesco(texto: str) -> bool:
    marcadores = ("Bradesco NET EMPRESA", "Banco Bradesco S/A", "BRADESCO")
    tem_marca = any(m.upper() in texto.upper() for m in marcadores)
    tem_layout = "Comprovante de Transa" in texto or "Boleto de Cobran" in texto
    return tem_marca and tem_layout


def _parsear_bradesco(texto: str, arquivo: str) -> list[Transacao]:
    """Bradesco: cada comprovante começa em 'Comprovante de Transação Bancária'
    e vai até o próximo mesmo cabeçalho ou até o fim."""
    marcador = re.compile(r"Comprovante de Transa[c\u00e7\ufffd]{1,3}[a\u00e3\ufffd]o Banc[a\u00e1\ufffd]ria")
    posicoes = [m.start() for m in marcador.finditer(texto)]
    if not posicoes:
        return []

    blocos: list[str] = []
    for i, ini in enumerate(posicoes):
        fim = posicoes[i + 1] if i + 1 < len(posicoes) else len(texto)
        blocos.append(texto[ini:fim])

    transacoes: list[Transacao] = []
    for i, bloco in enumerate(blocos, 1):
        t = _extrair_transacao_bradesco(bloco, arquivo, ordem=i)
        if t is not None:
            transacoes.append(t)
    return transacoes


def _extrair_transacao_bradesco(bloco: str, arquivo: str, ordem: int) -> Transacao | None:
    # Valor total
    m_valor = re.search(r"Valor total:\s*R\$\s*([\d.,]+)", bloco)
    if not m_valor:
        m_valor = re.search(r"Valor\s+R\$\s*([\d.,]+)", bloco)
    valor = _para_valor(m_valor.group(1)) if m_valor else None
    if valor is None:
        return None

    # Data de débito (a real de pagamento no banco)
    m_data = re.search(r"Data de d[e\u00e9\ufffd]bito:\s*(\d{2}/\d{2}/\d{4})", bloco)
    if not m_data:
        m_data = re.search(r"Data da opera[c\u00e7\ufffd]{1,3}[a\u00e3\ufffd]o:\s*(\d{2}/\d{2}/\d{4})", bloco)
    data_pagto = _para_data(m_data.group(1)) if m_data else None
    if data_pagto is None:
        return None

    # Data de vencimento
    m_venc = re.search(r"Data de vencimento:\s*(\d{2}/\d{2}/\d{4})", bloco)
    data_venc = _para_data(m_venc.group(1)) if m_venc else data_pagto

    # Beneficiário — Bradesco quebra o label em 2 linhas: "Razão Social" \n texto
    m_nome = re.search(
        r"Raz[a\u00e3\ufffd]o Social\s+([^\n]+?)\s+Benefici[a\u00e1\ufffd]rio:",
        bloco, re.DOTALL,
    )
    fornecedor = _limpar_nome(m_nome.group(1)) if m_nome else ""
    if not fornecedor:
        # fallback: "Razão Social <NOME>" seguido de qualquer coisa
        m_nome = re.search(r"Raz[a\u00e3\ufffd]o Social\s+([A-Z0-9][^\n]{2,})", bloco)
        if m_nome:
            fornecedor = _limpar_nome(m_nome.group(1))

    # CNPJ beneficiário
    m_cnpj = re.search(r"CPF/CNPJ Benefici[a\u00e1\ufffd]rio:\s*([\d./\-]+)", bloco)
    cnpj = _limpar_cnpj(m_cnpj.group(1)) if m_cnpj else ""

    # Número do documento
    m_doc = re.search(r"Documento:\s*(\d+)", bloco)
    doc = m_doc.group(1).strip() if m_doc else ""

    return Transacao(
        data=data_venc,
        valor=valor,
        descricao="",
        origem="pdf",
        linha=ordem,
        extras={
            "fornecedor": fornecedor,
            "cnpj": cnpj,
            "numero_nf": doc,
            "historico": f"Boleto Bradesco {doc}".strip(),
            "arquivo": arquivo,
            "banco_pdf": "Bradesco",
        },
        data_pagamento=data_pagto,
    )


# -------------------------------------------------------- pipeline

BANCOS_SUPORTADOS: list[tuple[str, Callable[[str], bool], Callable[[str, str], list[Transacao]]]] = [
    ("Sicoob",   _detectar_sicoob,   _parsear_sicoob),
    ("Bradesco", _detectar_bradesco, _parsear_bradesco),
]


def extrair_texto_bruto(caminho: str | Path) -> str:
    """Extrai texto de todas as páginas do PDF (streaming — funciona bem
    até uns milhares de páginas em PDFs nativos)."""
    textos: list[str] = []
    with pdfplumber.open(str(caminho)) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                textos.append(texto)
    return "\n".join(textos)


def ler_comprovantes_pdf(
    caminho: str | Path,
    progresso: Callable[[int, int], None] | None = None,
) -> tuple[list[Transacao], str]:
    """Lê um PDF de comprovantes e devolve (transacoes, banco_detectado).

    Se ``progresso`` for passado, é chamado a cada página processada com
    (pagina_atual, total_paginas) — útil pra UI mostrar progresso em PDFs
    grandes (500+ páginas).

    Levanta ValueError se o PDF é vazio, ou se o banco não é reconhecido.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise ValueError(f"Arquivo não encontrado: {caminho}")

    textos: list[str] = []
    with pdfplumber.open(str(caminho)) as pdf:
        total = len(pdf.pages)
        if total == 0:
            raise ValueError(f"PDF vazio: {caminho.name}")
        for i, pagina in enumerate(pdf.pages, 1):
            texto = pagina.extract_text()
            if texto:
                textos.append(texto)
            if progresso:
                progresso(i, total)

    texto_completo = "\n".join(textos)
    if not texto_completo.strip():
        raise ValueError(
            f"Não consegui extrair texto de {caminho.name}. "
            "Pode ser PDF escaneado (imagem) — precisaria de OCR."
        )

    # Detecta o banco
    for nome_banco, detectar, parsear in BANCOS_SUPORTADOS:
        if detectar(texto_completo):
            transacoes = parsear(texto_completo, caminho.name)
            return transacoes, nome_banco

    # Nenhum banco reconhecido
    bancos_suportados = ", ".join(b for b, _, _ in BANCOS_SUPORTADOS)
    raise ValueError(
        f"Banco não reconhecido em {caminho.name}. "
        f"Suportados: {bancos_suportados}. "
        "Envie um exemplo pro admin pra adicionar suporte."
    )


def ler_comprovantes_pdfs(
    caminhos: list[str | Path],
    progresso: Callable[[int, int, str], None] | None = None,
) -> tuple[list[Transacao], dict[str, tuple[str, int]]]:
    """Lê múltiplos PDFs, agrega as transações. Devolve:
    - lista completa de Transacoes
    - dict {arquivo: (banco_detectado, n_transacoes)} — pra relatório

    ``progresso`` (opcional) é chamado como (arquivo_atual, total_arquivos, nome).
    """
    todas: list[Transacao] = []
    relatorio: dict[str, tuple[str, int]] = {}
    total = len(caminhos)
    for i, caminho in enumerate(caminhos, 1):
        caminho = Path(caminho)
        if progresso:
            progresso(i, total, caminho.name)
        try:
            transacoes, banco = ler_comprovantes_pdf(caminho)
            todas.extend(transacoes)
            relatorio[caminho.name] = (banco, len(transacoes))
        except ValueError as e:
            relatorio[caminho.name] = (f"ERRO: {e}", 0)
    return todas, relatorio


if __name__ == "__main__":
    # Modo teste — roda com: python parser_pdf.py caminho.pdf
    import sys
    if len(sys.argv) < 2:
        print("Uso: python parser_pdf.py caminho.pdf [caminho2.pdf ...]")
        raise SystemExit(1)
    for arq in sys.argv[1:]:
        print(f"\n=== {arq} ===")
        try:
            txs, banco = ler_comprovantes_pdf(arq)
            print(f"Banco detectado: {banco}")
            print(f"Comprovantes extraidos: {len(txs)}")
            for t in txs[:5]:
                print(
                    f"  {t.data_pagamento}  R$ {t.valor:>10.2f}  "
                    f"{(t.extras.get('fornecedor') or '')[:40]:40s}  "
                    f"{t.extras.get('cnpj', '')}"
                )
            if len(txs) > 5:
                print(f"  ... (+{len(txs) - 5} restantes)")
        except ValueError as e:
            print(f"ERRO: {e}")
