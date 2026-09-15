"""Exportação para Excel (.xlsx) das abas Conciliados × Domínio e
Lançamentos contábeis.

Usa openpyxl (já é dependência do projeto — parser_xlsx.py).
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

if TYPE_CHECKING:
    from lancamentos import LancamentoContabil
    from matcher import Par
    from parser_xlsx import Transacao


# Paleta consistente com a UI Tkinter (mesmas cores)
COR_HEADER = "1F3A68"           # azul escuro do título
COR_ABERTO = "D4EDDA"           # verde
COR_PARCIAL = "CFE2FF"          # azul claro
COR_PAGA = "E9ECEF"             # cinza

FONTE_HEADER = Font(name="Arial", size=10, bold=True, color="FFFFFF")
FONTE_CELULA = Font(name="Arial", size=10)


def _fmt_data(d) -> str:
    if d is None:
        return ""
    if isinstance(d, (date, datetime)):
        return d.strftime("%d/%m/%Y")
    return str(d)


def _tag_status(status: str) -> str:
    sl = (status or "").lower()
    if sl.startswith("pag"):
        return COR_PAGA
    if sl.startswith("parc"):
        return COR_PARCIAL
    if sl.startswith("ab"):
        return COR_ABERTO
    return ""


def _aplica_header(ws, colunas: list[tuple[str, int]]) -> None:
    """Escreve cabeçalho colorido na linha 1 e ajusta larguras."""
    fill = PatternFill(start_color=COR_HEADER, end_color=COR_HEADER, fill_type="solid")
    align = Alignment(horizontal="left", vertical="center", wrap_text=False)
    for i, (titulo, largura) in enumerate(colunas, start=1):
        cell = ws.cell(row=1, column=i, value=titulo)
        cell.font = FONTE_HEADER
        cell.fill = fill
        cell.alignment = align
        ws.column_dimensions[get_column_letter(i)].width = largura
    ws.freeze_panes = "A2"      # header sempre visível ao rolar


def _pinta_linha(ws, row_idx: int, n_cols: int, cor: str) -> None:
    if not cor:
        return
    fill = PatternFill(start_color=cor, end_color=cor, fill_type="solid")
    for col in range(1, n_cols + 1):
        ws.cell(row=row_idx, column=col).fill = fill


def _pega(*fontes, chave: str) -> str:
    """Devolve o primeiro valor não-vazio entre as fontes (dicts de extras)."""
    for f in fontes:
        if f is None:
            continue
        v = f.get(chave, "")
        if v:
            return v
    return ""


def exportar_conciliados_dominio(
    caminho: str | Path,
    pares_triple: list["Par"],
    pendentes_caixa_dominio: list[tuple["Transacao", dict]],
    pendentes_ofx_dominio: list[tuple["Transacao", dict]] | None = None,
) -> int:
    """Exporta a aba "Conciliados × Domínio" para .xlsx.

    PRIORIDADE de dados: Domínio > planilha/PDF > OFX. O CNPJ, fornecedor
    e NF vêm sempre do Domínio quando disponíveis — o Domínio é a fonte
    contábil confiável (planilhas e comprovantes podem ter dados errados
    ou vazios).

    Args:
        caminho: destino do .xlsx
        pares_triple: pares P×OFX que casaram com Domínio
        pendentes_caixa_dominio: [(transacao_planilha, match_dict), ...]
        pendentes_ofx_dominio: [(transacao_ofx, match_dict), ...] — opcional

    Devolve o total de linhas escritas.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliados x Dominio"

    colunas = [
        ("Tipo", 8),
        ("Origem", 22),
        ("Vencimento", 12),
        ("Pagamento", 12),
        ("Valor", 14),
        ("Emissão", 12),
        ("Nº NF", 12),
        ("CNPJ", 20),
        ("Fornecedor", 32),
        # Empresa: só populada quando o grupo matriz+filiais foi
        # carregado — mostra em qual empresa a parcela foi lançada.
        ("Empresa (código)", 30),
        ("Memo OFX / Histórico", 40),
        ("Δ Domínio", 16),
        ("Status (Domínio)", 15),
    ]
    _aplica_header(ws, colunas)

    def _empresa(t_dom) -> str:
        if t_dom is None:
            return ""
        codi = t_dom.extras.get("codi_emp_origem")
        if codi is None:
            return ""
        razao = t_dom.extras.get("razao_empresa", "") or ""
        return f"{codi} - {razao[:40]}" if razao else str(codi)

    linha = 2

    # 1) Pares P×OFX — prioridade Domínio > planilha > OFX
    for par in pares_triple:
        tipo_txt = "Auto" if par.tipo == "auto" else "Manual"
        origem = par.ofx.extras.get("banco", "") or "OFX"
        pagto = par.planilha.data_pagamento or par.ofx.data
        dom_extras = par.dominio.extras if par.dominio else {}
        p_extras = par.planilha.extras
        o_extras = par.ofx.extras
        status = dom_extras.get("status", "") or ""

        cnpj = _pega(dom_extras, p_extras, o_extras, chave="cnpj")
        fornecedor = _pega(dom_extras, p_extras, o_extras, chave="fornecedor")
        numero_nf = _pega(dom_extras, p_extras, chave="numero_nf")
        emissao = _pega(dom_extras, p_extras, chave="data_emissao")

        diff_dom = ""
        if par.diff_dias_dominio or par.diff_valor_dominio:
            diff_dom = f"{par.diff_dias_dominio}d, R$ {par.diff_valor_dominio:.2f}"

        valores = [
            tipo_txt,
            origem,
            _fmt_data(par.planilha.data),
            _fmt_data(pagto),
            float(par.planilha.valor),
            _fmt_data(emissao) if hasattr(emissao, "strftime") else str(emissao or ""),
            str(numero_nf),
            str(cnpj),
            str(fornecedor),
            _empresa(par.dominio),
            (par.ofx.descricao or ""),
            diff_dom,
            status,
        ]
        for i, v in enumerate(valores, start=1):
            cell = ws.cell(row=linha, column=i, value=v)
            cell.font = FONTE_CELULA
            if i == 5:  # coluna Valor
                cell.number_format = '#,##0.00'
        _pinta_linha(ws, linha, len(colunas), _tag_status(status))
        linha += 1

    # 2) Pendentes da planilha (Caixa geral) — prioridade Domínio > planilha
    for t_p, match in pendentes_caixa_dominio:
        t_dom = match.get("dominio")
        d_d = match.get("diff_dias", 0)
        d_v = match.get("diff_valor", Decimal("0"))
        pagto = t_p.data_pagamento or t_p.data
        dom_extras = t_dom.extras if t_dom else {}
        p_extras = t_p.extras
        status = dom_extras.get("status", "") or ""

        cnpj = _pega(dom_extras, p_extras, chave="cnpj")
        fornecedor = _pega(dom_extras, p_extras, chave="fornecedor")
        numero_nf = _pega(dom_extras, p_extras, chave="numero_nf")
        emissao = _pega(dom_extras, p_extras, chave="data_emissao")

        diff_dom = ""
        if d_d or d_v:
            diff_dom = f"{d_d}d, R$ {d_v:.2f}"

        memo_txt = (
            f"Histórico: {t_p.extras.get('historico', '')}"
            if t_p.extras.get("historico") else "(sem OFX)"
        )

        valores = [
            "Caixa",
            "Caixa geral",
            _fmt_data(t_p.data),
            _fmt_data(pagto),
            float(t_p.valor),
            _fmt_data(emissao) if hasattr(emissao, "strftime") else str(emissao or ""),
            str(numero_nf),
            str(cnpj),
            str(fornecedor),
            _empresa(t_dom),
            memo_txt,
            diff_dom,
            status,
        ]
        for i, v in enumerate(valores, start=1):
            cell = ws.cell(row=linha, column=i, value=v)
            cell.font = FONTE_CELULA
            if i == 5:
                cell.number_format = '#,##0.00'
        _pinta_linha(ws, linha, len(colunas), _tag_status(status))
        linha += 1

    # 3) Pendentes OFX que casaram com Domínio (sem planilha)
    #    prioridade Domínio > OFX (enriquecido ou não)
    if pendentes_ofx_dominio:
        for t_o, match in pendentes_ofx_dominio:
            t_dom = match.get("dominio")
            d_d = match.get("diff_dias", 0)
            d_v = match.get("diff_valor", Decimal("0"))
            dom_extras = t_dom.extras if t_dom else {}
            o_extras = t_o.extras
            status = dom_extras.get("status", "") or ""

            cnpj = _pega(dom_extras, o_extras, chave="cnpj")
            fornecedor = _pega(dom_extras, o_extras, chave="fornecedor")
            numero_nf = _pega(dom_extras, o_extras, chave="numero_nf")
            emissao = dom_extras.get("data_emissao")
            origem = t_o.extras.get("banco", "") or "OFX"

            diff_dom = ""
            if d_d or d_v:
                diff_dom = f"{d_d}d, R$ {d_v:.2f}"

            valores = [
                "OFX",
                origem,
                _fmt_data(t_o.data),
                _fmt_data(t_o.data),
                float(t_o.valor),
                _fmt_data(emissao) if hasattr(emissao, "strftime") else "",
                str(numero_nf),
                str(cnpj),
                str(fornecedor),
                _empresa(t_dom),
                (t_o.descricao or ""),
                diff_dom,
                status,
            ]
            for i, v in enumerate(valores, start=1):
                cell = ws.cell(row=linha, column=i, value=v)
                cell.font = FONTE_CELULA
                if i == 5:
                    cell.number_format = '#,##0.00'
            _pinta_linha(ws, linha, len(colunas), _tag_status(status))
            linha += 1

    wb.save(str(caminho))
    return linha - 2  # total de linhas de dados


def exportar_pendencias_comparacao(
    caminho: str | Path,
    pares_amarelos: list["Par"],
    caixa_cinzas: list["Transacao"],
    ofx_laranjas: list["Transacao"],
) -> tuple[int, int, int]:
    """Exporta as pendências da aba Comparação para .xlsx com abas separadas
    por tipo de pendência (evita misturar estruturas diferentes):
    - "Falta Domínio (P×OFX)" — pares amarelos
    - "Caixa geral (falta Dom)" — pendentes planilha cinzas
    - "OFX (falta Dom)" — pendentes OFX laranjas (só existe sem planilha)

    Só cria as abas que têm dados.

    Devolve (n_amarelos, n_cinzas, n_laranjas).
    """
    wb = Workbook()
    ws_default = wb.active
    primeira_criada = False

    # Aba 1: Amarelos (P×OFX sem Domínio)
    if pares_amarelos:
        ws = ws_default
        ws.title = "Falta Dominio (PxOFX)"
        primeira_criada = True
        colunas = [
            ("Vencimento", 12),
            ("Pagamento", 12),
            ("Valor", 14),
            ("Emissão", 12),
            ("Nº NF", 12),
            ("CNPJ", 20),
            ("Fornecedor", 32),
            ("Histórico", 32),
            ("Tipo", 20),
            ("Banco (OFX)", 20),
            ("Memo OFX", 40),
        ]
        _aplica_header(ws, colunas)
        linha = 2
        for par in pares_amarelos:
            emissao = par.planilha.extras.get("data_emissao")
            pagto = par.planilha.data_pagamento or par.ofx.data
            valores = [
                _fmt_data(par.planilha.data),
                _fmt_data(pagto),
                float(par.planilha.valor),
                _fmt_data(emissao),
                str(par.planilha.extras.get("numero_nf", "") or ""),
                str(par.planilha.extras.get("cnpj", "") or ""),
                str(par.planilha.extras.get("fornecedor", "") or ""),
                str(par.planilha.extras.get("historico", "") or ""),
                str(par.planilha.extras.get("tipo", "") or ""),
                str(par.ofx.extras.get("banco", "") or ""),
                (par.ofx.descricao or ""),
            ]
            for i, v in enumerate(valores, start=1):
                cell = ws.cell(row=linha, column=i, value=v)
                cell.font = FONTE_CELULA
                if i == 3:
                    cell.number_format = '#,##0.00'
            _pinta_linha(ws, linha, len(colunas), "FFF3CD")
            linha += 1
    n_amarelos = len(pares_amarelos)

    # Aba 2: Cinzas (Caixa geral sem Domínio)
    if caixa_cinzas:
        ws = ws_default if not primeira_criada else wb.create_sheet()
        ws.title = "Caixa (falta Dom)"
        primeira_criada = True
        colunas = [
            ("Vencimento", 12),
            ("Pagamento", 12),
            ("Valor", 14),
            ("Emissão", 12),
            ("Nº NF", 12),
            ("CNPJ", 20),
            ("Fornecedor", 32),
            ("Histórico", 40),
            ("Tipo", 20),
        ]
        _aplica_header(ws, colunas)
        linha = 2
        for t in caixa_cinzas:
            valores = [
                _fmt_data(t.data),
                _fmt_data(t.data_pagamento),
                float(t.valor),
                _fmt_data(t.extras.get("data_emissao")),
                str(t.extras.get("numero_nf", "") or ""),
                str(t.extras.get("cnpj", "") or ""),
                str(t.extras.get("fornecedor", "") or ""),
                str(t.extras.get("historico", "") or ""),
                str(t.extras.get("tipo", "") or ""),
            ]
            for i, v in enumerate(valores, start=1):
                cell = ws.cell(row=linha, column=i, value=v)
                cell.font = FONTE_CELULA
                if i == 3:
                    cell.number_format = '#,##0.00'
            _pinta_linha(ws, linha, len(colunas), "E2E3E5")
            linha += 1
    n_cinzas = len(caixa_cinzas)

    # Aba 3: Laranjas (OFX sem planilha, sem Domínio)
    if ofx_laranjas:
        ws = ws_default if not primeira_criada else wb.create_sheet()
        ws.title = "OFX (falta Dom)"
        primeira_criada = True
        colunas = [
            ("Data pagamento", 14),
            ("Banco", 22),
            ("Documento", 15),
            ("Valor", 14),
            ("Memo OFX", 50),
        ]
        _aplica_header(ws, colunas)
        linha = 2
        for t in ofx_laranjas:
            valores = [
                _fmt_data(t.data),
                str(t.extras.get("banco", "") or ""),
                str(t.extras.get("documento", "") or ""),
                float(t.valor),
                (t.descricao or ""),
            ]
            for i, v in enumerate(valores, start=1):
                cell = ws.cell(row=linha, column=i, value=v)
                cell.font = FONTE_CELULA
                if i == 4:
                    cell.number_format = '#,##0.00'
            _pinta_linha(ws, linha, len(colunas), "FFE5CC")
            linha += 1
    n_laranjas = len(ofx_laranjas)

    # Se nenhuma aba foi criada, remove a default e cria uma vazia
    if not primeira_criada:
        ws_default.title = "Sem pendencias"
        ws_default.cell(row=1, column=1, value="Nenhuma pendência para exportar.")

    wb.save(str(caminho))
    return n_amarelos, n_cinzas, n_laranjas


def exportar_pendentes(
    caminho: str | Path,
    pendentes_planilha: list["Transacao"],
    pendentes_ofx: list["Transacao"],
) -> tuple[int, int]:
    """Exporta a aba Pendentes para .xlsx com DUAS abas — uma pra planilha,
    outra pra OFX. Devolve (n_planilha, n_ofx)."""
    wb = Workbook()
    # Aba 1: planilha (usa a aba padrão criada pelo Workbook)
    ws_p = wb.active
    ws_p.title = "Pendentes Planilha"

    colunas_p = [
        ("Vencimento", 12),
        ("Pagamento", 12),
        ("Valor", 14),
        ("Nº NF", 12),
        ("Fornecedor", 32),
        ("Histórico", 40),
        ("Tipo", 18),
        ("CNPJ", 20),
        ("Data emissão", 12),
    ]
    _aplica_header(ws_p, colunas_p)
    linha = 2
    for t in pendentes_planilha:
        valores = [
            _fmt_data(t.data),
            _fmt_data(t.data_pagamento),
            float(t.valor),
            str(t.extras.get("numero_nf", "") or ""),
            str(t.extras.get("fornecedor", "") or ""),
            str(t.extras.get("historico", "") or ""),
            str(t.extras.get("tipo", "") or ""),
            str(t.extras.get("cnpj", "") or ""),
            _fmt_data(t.extras.get("data_emissao")),
        ]
        for i, v in enumerate(valores, start=1):
            cell = ws_p.cell(row=linha, column=i, value=v)
            cell.font = FONTE_CELULA
            if i == 3:
                cell.number_format = '#,##0.00'
        linha += 1
    n_planilha = linha - 2

    # Aba 2: OFX
    ws_o = wb.create_sheet(title="Pendentes OFX")
    colunas_o = [
        ("Data pagamento", 14),
        ("Banco", 22),
        ("Documento", 15),
        ("Valor", 14),
        ("Memo OFX", 50),
    ]
    _aplica_header(ws_o, colunas_o)
    linha = 2
    for t in pendentes_ofx:
        valores = [
            _fmt_data(t.data),
            str(t.extras.get("banco", "") or ""),
            str(t.extras.get("documento", "") or ""),
            float(t.valor),
            (t.descricao or ""),
        ]
        for i, v in enumerate(valores, start=1):
            cell = ws_o.cell(row=linha, column=i, value=v)
            cell.font = FONTE_CELULA
            if i == 4:
                cell.number_format = '#,##0.00'
        linha += 1
    n_ofx = linha - 2

    wb.save(str(caminho))
    return n_planilha, n_ofx


def exportar_lancamentos_contabeis(
    caminho: str | Path,
    lancamentos: list["LancamentoContabil"],
) -> int:
    """Exporta a aba "Lançamentos contábeis" para .xlsx."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Lancamentos contabeis"

    colunas = [
        ("Data pagto", 12),
        ("Banco", 22),
        ("Valor", 14),
        ("Conta", 14),
        ("Histórico contábil", 40),
        ("Memo (OFX)", 40),
        ("Regra", 22),
        ("Tipo", 22),
        ("Fornecedor", 30),
        ("CNPJ", 20),
    ]
    _aplica_header(ws, colunas)

    # Legenda amigável para tipo_regra
    tipo_legivel = {
        "memo": "Regra por memo (OFX)",
        "fornecedor": "Regra por fornecedor (par P×OFX)",
        "fornecedor_planilha": "Regra por fornecedor (planilha, Caixa geral)",
        "manual": "Manual (par P×OFX)",
        "manual_ofx": "Manual (pendente OFX)",
        "manual_planilha": "Manual (pendente planilha)",
    }

    linha = 2
    for l in lancamentos:
        valores = [
            _fmt_data(l.data),
            l.banco or "",
            float(l.valor),
            l.conta or "",
            l.historico or "",
            l.memo_original or "",
            l.padrao_match or "",
            tipo_legivel.get(l.tipo_regra, l.tipo_regra or ""),
            l.fornecedor or "",
            l.cnpj or "",
        ]
        for i, v in enumerate(valores, start=1):
            cell = ws.cell(row=linha, column=i, value=v)
            cell.font = FONTE_CELULA
            if i == 3:
                cell.number_format = '#,##0.00'
        linha += 1

    # Total no final (somando a coluna Valor)
    if lancamentos:
        cell_tot_rot = ws.cell(row=linha, column=2, value="TOTAL")
        cell_tot_rot.font = Font(name="Arial", size=10, bold=True)
        cell_tot_rot.alignment = Alignment(horizontal="right")
        cell_tot = ws.cell(
            row=linha, column=3,
            value=f"=SUM(C2:C{linha - 1})",
        )
        cell_tot.font = Font(name="Arial", size=10, bold=True)
        cell_tot.number_format = '#,##0.00'

    wb.save(str(caminho))
    return linha - 2
