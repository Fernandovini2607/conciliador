import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import config
import parser_dominio
from dialogos_dominio import DialogoConexao, DialogoFonte, DialogoSelecionarEmpresa
from matcher import (
    Par,
    Resultado,
    conciliar_automatico,
    diferenca,
    gerar_sugestoes,
)
from parser_ofx import ler_ofx
from parser_xlsx import (
    EstruturaPlanilha,
    Transacao,
    descobrir_estrutura,
    extrair_transacoes,
    para_data,
    para_decimal,
)


CAMPOS = [
    ("data", "Data"),
    ("valor", "Valor"),
    ("descricao", "Descrição"),
]

CAMPOS_EXTRAS_UI = [
    ("data_emissao", "Data emissão"),
    ("numero_nf", "Nº NF"),
    ("cnpj", "CNPJ fornecedor"),
    ("fornecedor", "Fornecedor"),
]

OPCAO_VAZIA = "(não usar)"


class DialogoMapeamento(tk.Toplevel):
    """Modal para o usuário escolher quais colunas da planilha são Data/Valor/Descrição."""

    PREVIEW_LINHAS = 10

    def __init__(self, master: tk.Misc, estrutura: EstruturaPlanilha) -> None:
        super().__init__(master)
        self.title("Mapear colunas da planilha")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.estrutura = estrutura
        self.mapeamento: dict[str, int] | None = None

        info = ttk.Label(
            self,
            text=(
                f"Cabeçalho detectado na linha {estrutura.linha_cabecalho}. "
                "Escolha qual coluna corresponde a cada campo — o preview embaixo "
                "mostra como cada linha será lida. Vermelho = não pôde ser convertido."
            ),
            wraplength=720,
        )
        info.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w")

        self.opcoes = [
            f"{idx + 1}. {nome if nome else '(sem nome)'}"
            for idx, nome in enumerate(estrutura.cabecalho)
        ]
        self.combos: dict[str, ttk.Combobox] = {}

        for i, (campo, rotulo) in enumerate(CAMPOS, start=1):
            ttk.Label(self, text=f"{rotulo}:").grid(row=i, column=0, padx=12, pady=4, sticky="e")
            cb = ttk.Combobox(self, values=self.opcoes, state="readonly", width=48)
            cb.grid(row=i, column=1, padx=(0, 12), pady=4, sticky="w")
            idx = estrutura.sugestao.get(campo)
            if idx is not None and idx < len(self.opcoes):
                cb.current(idx)
            cb.bind("<<ComboboxSelected>>", lambda _e: self._atualiza_preview())
            self.combos[campo] = cb

        # ----- Campos extras opcionais (data emissão, NF, CNPJ, fornecedor)
        opcoes_extras = [OPCAO_VAZIA] + self.opcoes
        ttk.Separator(self, orient="horizontal").grid(
            row=len(CAMPOS) + 1, column=0, columnspan=2, sticky="we", padx=12, pady=(6, 2),
        )
        ttk.Label(
            self, text="Campos extras (opcionais — deixe vazio se a planilha não tem):",
            font=("TkDefaultFont", 9, "italic"),
        ).grid(row=len(CAMPOS) + 2, column=0, columnspan=2, padx=12, pady=(2, 4), sticky="w")
        for j, (campo, rotulo) in enumerate(CAMPOS_EXTRAS_UI):
            row = len(CAMPOS) + 3 + j
            ttk.Label(self, text=f"{rotulo}:").grid(row=row, column=0, padx=12, pady=2, sticky="e")
            cb = ttk.Combobox(self, values=opcoes_extras, state="readonly", width=48)
            cb.grid(row=row, column=1, padx=(0, 12), pady=2, sticky="w")
            idx_sug = estrutura.sugestao.get(campo)
            if idx_sug is not None and 0 <= idx_sug < len(self.opcoes):
                cb.current(idx_sug + 1)  # +1 porque tem OPCAO_VAZIA no início
            else:
                cb.current(0)
            self.combos[campo] = cb

        row_preview = len(CAMPOS) + 3 + len(CAMPOS_EXTRAS_UI)
        ttk.Label(self, text="Preview das primeiras linhas:").grid(
            row=row_preview, column=0, columnspan=2, padx=12, pady=(8, 2), sticky="w",
        )

        preview_frame = ttk.Frame(self)
        preview_frame.grid(row=row_preview + 1, column=0, columnspan=2, padx=12, sticky="we")

        cols = ("linha", "data", "valor", "descricao")
        self.preview = ttk.Treeview(
            preview_frame, columns=cols, show="headings", height=self.PREVIEW_LINHAS,
        )
        self.preview.heading("linha", text="Linha")
        self.preview.heading("data", text="Data")
        self.preview.heading("valor", text="Valor")
        self.preview.heading("descricao", text="Descrição")
        self.preview.column("linha", width=55, anchor="center")
        self.preview.column("data", width=140, anchor="w")
        self.preview.column("valor", width=130, anchor="e")
        self.preview.column("descricao", width=360, anchor="w")
        self.preview.tag_configure("erro", background="#f8d7da")
        self.preview.pack(side="left", fill="both", expand=True)

        self.lbl_status = ttk.Label(self, text="")
        self.lbl_status.grid(
            row=row_preview + 2, column=0, columnspan=2, padx=12, pady=(4, 0), sticky="w",
        )

        botoes = ttk.Frame(self)
        botoes.grid(row=row_preview + 3, column=0, columnspan=2, pady=(8, 12), padx=12, sticky="e")
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(side="right", padx=6)
        ttk.Button(botoes, text="Confirmar", command=self._confirmar).pack(side="right", padx=6)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Return>", lambda _e: self._confirmar())
        self.bind("<Escape>", lambda _e: self._cancelar())

        self._atualiza_preview()

    def _atualiza_preview(self) -> None:
        for item in self.preview.get_children():
            self.preview.delete(item)

        idxs = {campo: self.combos[campo].current() for campo, _ in CAMPOS}
        base = self.estrutura.linha_cabecalho + 1

        total_validas = 0
        total_consideradas = 0
        amostra = self.estrutura.linhas[: self.PREVIEW_LINHAS]

        for offset, linha in enumerate(amostra):
            if not linha or all(c is None or c == "" for c in linha):
                continue
            total_consideradas += 1

            def _cel(i: int):
                return linha[i] if 0 <= i < len(linha) else None

            cel_data = _cel(idxs["data"])
            cel_valor = _cel(idxs["valor"])
            cel_desc = _cel(idxs["descricao"])

            data_parsed = para_data(cel_data) if idxs["data"] >= 0 else None
            valor_parsed = para_decimal(cel_valor) if idxs["valor"] >= 0 else None

            if data_parsed is not None and valor_parsed is not None:
                total_validas += 1
                tag = ""
            else:
                tag = "erro"

            data_txt = (
                data_parsed.strftime("%d/%m/%Y") if data_parsed
                else (f"✗ {cel_data!r}" if cel_data not in (None, "") else "—")
            )
            valor_txt = (
                f"{valor_parsed:.2f}" if valor_parsed is not None
                else (f"✗ {cel_valor!r}" if cel_valor not in (None, "") else "—")
            )
            desc_txt = "" if cel_desc is None else str(cel_desc)

            self.preview.insert(
                "", "end",
                values=(base + offset, data_txt, valor_txt, desc_txt),
                tags=(tag,) if tag else (),
            )

        total_linhas = len([l for l in self.estrutura.linhas if l and any(c is not None and c != "" for c in l)])
        if total_consideradas == 0:
            self.lbl_status.config(text="Planilha sem dados após o cabeçalho.")
        else:
            self.lbl_status.config(
                text=(
                    f"Preview: {total_validas}/{total_consideradas} linhas válidas. "
                    f"Total de linhas com dados na planilha: {total_linhas}."
                )
            )

    def _confirmar(self) -> None:
        mapa: dict[str, int] = {}
        for campo, _ in CAMPOS:
            sel = self.combos[campo].current()
            if sel < 0:
                messagebox.showwarning(
                    "Mapeamento incompleto",
                    f"Selecione a coluna para '{campo}'.",
                    parent=self,
                )
                return
            mapa[campo] = sel
        if len(set(mapa.values())) < 3:
            messagebox.showwarning(
                "Colunas duplicadas",
                "Os campos obrigatórios (Data/Valor/Descrição) precisam apontar "
                "para colunas diferentes.",
                parent=self,
            )
            return
        # Campos extras: índice 0 = "(não usar)", >0 = mapeado (subtrai 1)
        for campo, _ in CAMPOS_EXTRAS_UI:
            sel = self.combos[campo].current()
            if sel > 0:
                mapa[campo] = sel - 1
        self.mapeamento = mapa
        self.destroy()

    def _cancelar(self) -> None:
        self.mapeamento = None
        self.destroy()


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Conciliador OFX × Planilha")
        self.geometry("1200x720")

        # Dados originais
        self.transacoes_planilha: list[Transacao] = []
        self.transacoes_ofx: list[Transacao] = []
        self.caminho_planilha: Path | None = None
        self.caminho_ofx: Path | None = None
        self.estrutura_planilha: EstruturaPlanilha | None = None
        self.mapeamento_planilha: dict[str, int] | None = None

        # Estado de conciliação
        self.pares_conciliados: list[Par] = []          # auto + manual + sugestões aceitas
        self.pendentes_planilha: list[Transacao] = []
        self.pendentes_ofx: list[Transacao] = []
        self.sugestoes: list[Par] = []
        self.itens_pares: dict[str, Par] = {}            # iid → Par (aba conciliados)
        self.itens_pendentes_p: dict[str, Transacao] = {}
        self.itens_pendentes_o: dict[str, Transacao] = {}
        self.itens_sugestoes: dict[str, Par] = {}

        # Domínio
        self.conn_dominio = None
        self.transacoes_dominio: list[Transacao] = []
        self.comparacao_dominio: list[tuple[str, Par | Transacao]] = []
        # ↑ status, registro (Par para casados; Transacao para faltantes)
        self.cfg = config.carregar()
        self._migrar_config_legado()

        self._monta_ui()

    def _migrar_config_legado(self) -> None:
        """Se a config antiga tinha credenciais em cfg["dominio"], move pra
        data/dominio_config.json (padrão Janco) e renomeia o restante para
        cfg["dominio_fonte"]."""
        legado = self.cfg.get("dominio")
        if not isinstance(legado, dict):
            return
        cred_keys = {"dsn", "usuario", "senha"}
        if cred_keys & legado.keys():
            auth_atual = parser_dominio.load_odbc_config()
            for k in cred_keys:
                if k in legado:
                    auth_atual[k] = legado.pop(k)
            if auth_atual.get("dsn"):
                parser_dominio.save_odbc_config(auth_atual)
        if legado:
            self.cfg.setdefault("dominio_fonte", {}).update(legado)
        del self.cfg["dominio"]
        config.salvar(self.cfg)

    # ------------------------------------------------------------------ UI

    def _monta_ui(self) -> None:
        topo = ttk.Frame(self, padding=10)
        topo.pack(fill="x")
        ttk.Button(topo, text="Abrir planilha (.xlsx)", command=self._abrir_planilha).pack(side="left", padx=4)
        self.btn_editar_colunas = ttk.Button(
            topo, text="Editar colunas", command=self._editar_colunas, state="disabled",
        )
        self.btn_editar_colunas.pack(side="left", padx=4)
        self.lbl_planilha = ttk.Label(topo, text="(nenhuma planilha carregada)")
        self.lbl_planilha.pack(side="left", padx=8)

        topo2 = ttk.Frame(self, padding=(10, 0, 10, 10))
        topo2.pack(fill="x")
        ttk.Button(topo2, text="Importar OFX", command=self._abrir_ofx).pack(side="left", padx=4)
        self.lbl_ofx = ttk.Label(topo2, text="(nenhum OFX carregado)")
        self.lbl_ofx.pack(side="left", padx=8)

        topo3 = ttk.Frame(self, padding=(10, 0, 10, 6))
        topo3.pack(fill="x")
        ttk.Button(topo3, text="Conectar Domínio", command=self._conectar_dominio).pack(side="left", padx=4)
        self.btn_empresa = ttk.Button(
            topo3, text="Selecionar empresa", command=self._selecionar_empresa, state="disabled",
        )
        self.btn_empresa.pack(side="left", padx=4)
        self.btn_fonte = ttk.Button(
            topo3, text="Configurar fonte", command=self._configurar_fonte_dominio, state="disabled",
        )
        self.btn_fonte.pack(side="left", padx=4)
        self.btn_carregar_dominio = ttk.Button(
            topo3, text="Carregar pagamentos", command=self._carregar_dominio, state="disabled",
        )
        self.btn_carregar_dominio.pack(side="left", padx=4)
        self.lbl_dominio = ttk.Label(topo3, text="(Domínio não conectado)")
        self.lbl_dominio.pack(side="left", padx=8)

        acoes = ttk.Frame(self, padding=(10, 0, 10, 8))
        acoes.pack(fill="x")
        self.btn_conciliar = ttk.Button(
            acoes, text="Conciliar", command=self._executar_conciliacao, state="disabled",
        )
        self.btn_conciliar.pack(side="left", padx=4)
        self.btn_comparar_dominio = ttk.Button(
            acoes, text="Comparar com Domínio",
            command=self._comparar_com_dominio, state="disabled",
        )
        self.btn_comparar_dominio.pack(side="left", padx=4)
        self.lbl_resumo = ttk.Label(acoes, text="")
        self.lbl_resumo.pack(side="left", padx=12)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self._monta_aba_conciliados()
        self._monta_aba_pendentes()
        self._monta_aba_sugestoes()
        self._monta_aba_dominio()

    def _monta_aba_conciliados(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Conciliados (0)")
        self._aba_conciliados = aba

        cols = ("tipo", "data", "valor", "nf", "fornecedor", "desc_p", "desc_o", "diff")
        tree = ttk.Treeview(aba, columns=cols, show="headings")
        tree.heading("tipo", text="Tipo")
        tree.heading("data", text="Data")
        tree.heading("valor", text="Valor")
        tree.heading("nf", text="Nº NF")
        tree.heading("fornecedor", text="Fornecedor")
        tree.heading("desc_p", text="Descrição (planilha)")
        tree.heading("desc_o", text="Descrição (OFX)")
        tree.heading("diff", text="Diferenças")
        tree.column("tipo", width=70, anchor="w")
        tree.column("data", width=85, anchor="center")
        tree.column("valor", width=100, anchor="e")
        tree.column("nf", width=80, anchor="center")
        tree.column("fornecedor", width=200, anchor="w")
        tree.column("desc_p", width=240, anchor="w")
        tree.column("desc_o", width=240, anchor="w")
        tree.column("diff", width=130, anchor="w")
        tree.tag_configure("auto", background="#d4edda")
        tree.tag_configure("manual", background="#cfe2ff")

        sb = ttk.Scrollbar(aba, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        botoes = ttk.Frame(aba)
        botoes.pack(side="bottom", fill="x")
        ttk.Button(
            botoes, text="Desfazer conciliação manual",
            command=self._desfazer_conciliacao,
        ).pack(side="left", padx=6, pady=6)

        self.tree_conciliados = tree

    def _monta_aba_pendentes(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Pendentes (0)")
        self._aba_pendentes = aba

        instr = ttk.Label(
            aba,
            text=(
                "Selecione uma linha em cada tabela e clique em 'Conciliar selecionadas'. "
                "Use Ctrl+clique para desmarcar."
            ),
        )
        instr.pack(anchor="w", padx=6, pady=(6, 4))

        corpo = ttk.Frame(aba)
        corpo.pack(fill="both", expand=True)

        lado_p = ttk.LabelFrame(corpo, text="Só na planilha")
        lado_p.pack(side="left", fill="both", expand=True, padx=(6, 3), pady=(0, 4))

        cols = ("data", "valor", "descricao")
        self.tree_pend_p = ttk.Treeview(
            lado_p, columns=cols, show="headings", selectmode="browse",
        )
        for c, w, a in [("data", 90, "center"), ("valor", 110, "e"), ("descricao", 260, "w")]:
            self.tree_pend_p.heading(c, text=c.capitalize())
            self.tree_pend_p.column(c, width=w, anchor=a)
        sb_p = ttk.Scrollbar(lado_p, orient="vertical", command=self.tree_pend_p.yview)
        self.tree_pend_p.configure(yscrollcommand=sb_p.set)
        self.tree_pend_p.pack(side="left", fill="both", expand=True)
        sb_p.pack(side="right", fill="y")

        lado_o = ttk.LabelFrame(corpo, text="Só no OFX")
        lado_o.pack(side="left", fill="both", expand=True, padx=(3, 6), pady=(0, 4))

        self.tree_pend_o = ttk.Treeview(
            lado_o, columns=cols, show="headings", selectmode="browse",
        )
        for c, w, a in [("data", 90, "center"), ("valor", 110, "e"), ("descricao", 260, "w")]:
            self.tree_pend_o.heading(c, text=c.capitalize())
            self.tree_pend_o.column(c, width=w, anchor=a)
        sb_o = ttk.Scrollbar(lado_o, orient="vertical", command=self.tree_pend_o.yview)
        self.tree_pend_o.configure(yscrollcommand=sb_o.set)
        self.tree_pend_o.pack(side="left", fill="both", expand=True)
        sb_o.pack(side="right", fill="y")

        botoes = ttk.Frame(aba)
        botoes.pack(side="bottom", fill="x")
        ttk.Button(
            botoes, text="Conciliar selecionadas →",
            command=self._conciliar_selecionadas,
        ).pack(side="left", padx=6, pady=6)

    def _monta_aba_sugestoes(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Sugestões (0)")
        self._aba_sugestoes = aba

        instr = ttk.Label(
            aba,
            text=(
                "Pares com diferença de até 2 dias e até R$ 10,00. "
                "Selecione uma sugestão e clique em 'Aceitar como conciliação'."
            ),
        )
        instr.pack(anchor="w", padx=6, pady=(6, 4))

        cols = ("data_p", "valor_p", "desc_p", "data_o", "valor_o", "desc_o", "diff_dias", "diff_valor")
        tree = ttk.Treeview(aba, columns=cols, show="headings", selectmode="browse")
        for c, t, w, a in [
            ("data_p", "Data (pla)", 90, "center"),
            ("valor_p", "Valor (pla)", 90, "e"),
            ("desc_p", "Descrição (pla)", 220, "w"),
            ("data_o", "Data (OFX)", 90, "center"),
            ("valor_o", "Valor (OFX)", 90, "e"),
            ("desc_o", "Descrição (OFX)", 220, "w"),
            ("diff_dias", "Δ dias", 60, "center"),
            ("diff_valor", "Δ R$", 80, "e"),
        ]:
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=a)
        tree.tag_configure("destaque", background="#fff3cd")

        sb = ttk.Scrollbar(aba, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        botoes = ttk.Frame(aba)
        botoes.pack(side="bottom", fill="x")
        ttk.Button(
            botoes, text="Aceitar como conciliação",
            command=self._aceitar_sugestao,
        ).pack(side="left", padx=6, pady=6)

        self.tree_sugestoes = tree

    def _monta_aba_dominio(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Domínio (0)")
        self._aba_dominio = aba

        instr = ttk.Label(
            aba,
            text=(
                "Comparação dos pagamentos conciliados (planilha×OFX) com o Domínio. "
                "Match por data + valor exatos."
            ),
        )
        instr.pack(anchor="w", padx=6, pady=(6, 4))

        cols = (
            "status", "vencimento", "valor", "emissao", "nf",
            "cnpj", "fornecedor", "desc_concil", "desc_dominio",
        )
        tree = ttk.Treeview(aba, columns=cols, show="headings")
        tree.heading("status", text="Status")
        tree.heading("vencimento", text="Vencimento")
        tree.heading("valor", text="Valor")
        tree.heading("emissao", text="Emissão")
        tree.heading("nf", text="Nº NF")
        tree.heading("cnpj", text="CNPJ")
        tree.heading("fornecedor", text="Fornecedor")
        tree.heading("desc_concil", text="Descrição (planilha/OFX)")
        tree.heading("desc_dominio", text="Descrição (Domínio)")
        tree.column("status", width=160, anchor="w")
        tree.column("vencimento", width=85, anchor="center")
        tree.column("valor", width=95, anchor="e")
        tree.column("emissao", width=85, anchor="center")
        tree.column("nf", width=75, anchor="center")
        tree.column("cnpj", width=130, anchor="w")
        tree.column("fornecedor", width=200, anchor="w")
        tree.column("desc_concil", width=200, anchor="w")
        tree.column("desc_dominio", width=200, anchor="w")
        tree.tag_configure("ok", background="#d4edda")
        tree.tag_configure("falta_dominio", background="#fff3cd")
        tree.tag_configure("falta_concil", background="#f8d7da")

        sb = ttk.Scrollbar(aba, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree_dominio = tree

    # --------------------------------------------------------- Domínio

    def _conectar_dominio(self) -> None:
        dlg = DialogoConexao(self)
        self.wait_window(dlg)
        if dlg.conn is None:
            return
        self.conn_dominio = dlg.conn
        self.btn_empresa.config(state="normal")
        self.btn_fonte.config(state="normal")
        if self.cfg.get("dominio_fonte", {}).get("mapeamento"):
            self.btn_carregar_dominio.config(state="normal")
        self._atualiza_label_dominio()

    def _selecionar_empresa(self) -> None:
        if self.conn_dominio is None:
            return
        dlg = DialogoSelecionarEmpresa(
            self, self.conn_dominio, self.cfg.get("dominio_empresa"),
        )
        self.wait_window(dlg)
        if dlg.empresa is None:
            return
        self.cfg["dominio_empresa"] = dlg.empresa
        config.salvar(self.cfg)
        self._atualiza_label_dominio()

    def _atualiza_label_dominio(self) -> None:
        cred = parser_dominio.load_odbc_config()
        emp = self.cfg.get("dominio_empresa")
        partes = [f"Conectado (read-only): DSN={cred.get('dsn', '?')}"]
        if emp:
            partes.append(f"Empresa: {emp['codi_emp']} — {emp['razao'][:40]}")
        if self.transacoes_dominio:
            partes.append(f"{len(self.transacoes_dominio)} pagamentos carregados")
        self.lbl_dominio.config(text="  |  ".join(partes))

    def _configurar_fonte_dominio(self) -> None:
        if self.conn_dominio is None:
            return
        fonte_atual = self.cfg.get("dominio_fonte", {})
        dlg = DialogoFonte(self, self.conn_dominio, fonte_atual)
        self.wait_window(dlg)
        if dlg.fonte is None:
            return
        self.cfg["dominio_fonte"] = dlg.fonte
        config.salvar(self.cfg)
        self.btn_carregar_dominio.config(state="normal")

    def _carregar_dominio(self) -> None:
        if self.conn_dominio is None:
            return
        fonte = self.cfg.get("dominio_fonte", {})
        if not fonte.get("mapeamento"):
            messagebox.showinfo(
                "Sem fonte",
                "Configure a fonte de pagamentos primeiro.",
            )
            return
        emp = self.cfg.get("dominio_empresa") or {}
        codi_emp = emp.get("codi_emp")
        if codi_emp is None and fonte.get("modo") == "tabela":
            if not messagebox.askyesno(
                "Sem empresa selecionada",
                "Você não selecionou uma empresa — a query vai retornar "
                "lançamentos de TODAS as empresas misturadas.\n\n"
                "Deseja continuar mesmo assim?",
            ):
                return
        try:
            self.transacoes_dominio = parser_dominio.extrair_pagamentos(
                self.conn_dominio, fonte, codi_emp=codi_emp,
            )
        except Exception as e:
            messagebox.showerror("Erro ao ler Domínio", str(e))
            return
        self._atualiza_label_dominio()
        self._atualiza_botao_comparar()

    def _atualiza_botao_comparar(self) -> None:
        pode = bool(self.pares_conciliados and self.transacoes_dominio)
        self.btn_comparar_dominio.config(state="normal" if pode else "disabled")

    def _comparar_com_dominio(self) -> None:
        from collections import defaultdict
        from decimal import Decimal

        idx_dom: dict[tuple, list[Transacao]] = defaultdict(list)
        for t in self.transacoes_dominio:
            idx_dom[(t.data, t.valor.quantize(Decimal("0.01")))].append(t)

        resultados: list[tuple[str, Par, Transacao | None]] = []
        usados: set[int] = set()
        for par in self.pares_conciliados:
            chave = (par.planilha.data, par.planilha.valor.quantize(Decimal("0.01")))
            candidatos = [t for t in idx_dom.get(chave, []) if id(t) not in usados]
            if candidatos:
                t_dom = candidatos[0]
                usados.add(id(t_dom))
                resultados.append(("ok", par, t_dom))
            else:
                resultados.append(("falta_dominio", par, None))

        # Pagamentos no Domínio que ninguém casou
        sobras_dominio = [t for t in self.transacoes_dominio if id(t) not in usados]

        self._render_aba_dominio(resultados, sobras_dominio)
        self.notebook.select(self._aba_dominio)

    def _render_aba_dominio(
        self,
        resultados: list[tuple[str, Par, Transacao | None]],
        sobras_dominio: list[Transacao],
    ) -> None:
        for item in self.tree_dominio.get_children():
            self.tree_dominio.delete(item)

        def _fmt_data(d) -> str:
            return d.strftime("%d/%m/%Y") if d else ""

        n_ok = n_falta_dom = 0
        for status, par, t_dom in resultados:
            ok = status == "ok"
            if ok:
                n_ok += 1
                rotulo = "Conciliado e no Domínio"
            else:
                n_falta_dom += 1
                rotulo = "Conciliado, falta no Domínio"
            # Extras: prioriza Domínio se houver, depois planilha
            origem_extras = t_dom.extras if t_dom else par.planilha.extras
            extras_fallback = par.planilha.extras
            emissao = origem_extras.get("data_emissao") or extras_fallback.get("data_emissao")
            nf = origem_extras.get("numero_nf") or extras_fallback.get("numero_nf", "")
            cnpj = origem_extras.get("cnpj") or extras_fallback.get("cnpj", "")
            fornecedor = origem_extras.get("fornecedor") or extras_fallback.get("fornecedor", "")
            desc_dom = t_dom.descricao if t_dom else ""

            self.tree_dominio.insert(
                "", "end",
                values=(
                    rotulo,
                    par.planilha.data.strftime("%d/%m/%Y"),
                    f"{par.planilha.valor:.2f}",
                    _fmt_data(emissao),
                    nf,
                    cnpj,
                    fornecedor,
                    par.planilha.descricao or par.ofx.descricao,
                    desc_dom,
                ),
                tags=("ok" if ok else "falta_dominio",),
            )

        for t in sobras_dominio:
            self.tree_dominio.insert(
                "", "end",
                values=(
                    "No Domínio, sem conciliação",
                    t.data.strftime("%d/%m/%Y"),
                    f"{t.valor:.2f}",
                    _fmt_data(t.extras.get("data_emissao")),
                    t.extras.get("numero_nf", ""),
                    t.extras.get("cnpj", ""),
                    t.extras.get("fornecedor", ""),
                    "",
                    t.descricao,
                ),
                tags=("falta_concil",),
            )

        self.notebook.tab(
            3,
            text=(
                f"Domínio (ok {n_ok} | falta {n_falta_dom} | só dom {len(sobras_dominio)})"
            ),
        )

    # ------------------------------------------------------ Carregar dados

    def _abrir_planilha(self) -> None:
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if not caminho:
            return
        try:
            estrutura = descobrir_estrutura(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))
            return
        if not estrutura.cabecalho:
            messagebox.showerror("Planilha vazia", "A planilha não contém dados.")
            return

        dlg = DialogoMapeamento(self, estrutura)
        self.wait_window(dlg)
        if dlg.mapeamento is None:
            return

        try:
            transacoes = extrair_transacoes(estrutura, dlg.mapeamento)
        except Exception as e:
            messagebox.showerror("Erro ao extrair dados", str(e))
            return

        if not transacoes:
            messagebox.showwarning(
                "Nenhum lançamento identificado",
                "Nenhuma linha pôde ser convertida em lançamento.\n"
                "Use 'Editar colunas' para revisar o mapeamento.",
            )

        self.caminho_planilha = Path(caminho)
        self.estrutura_planilha = estrutura
        self.mapeamento_planilha = dlg.mapeamento
        self.transacoes_planilha = transacoes
        self._atualiza_label_planilha()
        self.btn_editar_colunas.config(state="normal")
        self._atualiza_botao()
        self._limpa_resultados()

    def _editar_colunas(self) -> None:
        if not self.estrutura_planilha:
            return
        estrutura = self.estrutura_planilha
        if self.mapeamento_planilha:
            estrutura.sugestao = dict(self.mapeamento_planilha)
        dlg = DialogoMapeamento(self, estrutura)
        self.wait_window(dlg)
        if dlg.mapeamento is None:
            return
        try:
            transacoes = extrair_transacoes(estrutura, dlg.mapeamento)
        except Exception as e:
            messagebox.showerror("Erro ao extrair dados", str(e))
            return
        if not transacoes:
            messagebox.showwarning(
                "Nenhum lançamento identificado",
                "Nenhuma linha pôde ser convertida em lançamento com esse mapeamento.\n"
                "O preview no diálogo destaca as linhas inválidas em vermelho.",
            )
        self.transacoes_planilha = transacoes
        self.mapeamento_planilha = dlg.mapeamento
        self._atualiza_label_planilha()
        self._atualiza_botao()
        self._limpa_resultados()

    def _atualiza_label_planilha(self) -> None:
        if not self.caminho_planilha or self.mapeamento_planilha is None:
            return
        cab = self.estrutura_planilha.cabecalho if self.estrutura_planilha else []
        partes = []
        for campo, rotulo in CAMPOS:
            idx = self.mapeamento_planilha.get(campo)
            if idx is None:
                nome = "?"
            elif idx < len(cab) and cab[idx]:
                nome = cab[idx]
            else:
                nome = f"col {idx + 1}"
            partes.append(f"{rotulo}={nome}")
        self.lbl_planilha.config(
            text=(
                f"{self.caminho_planilha.name} — {len(self.transacoes_planilha)} lançamentos  "
                f"[{' | '.join(partes)}]"
            )
        )

    def _abrir_ofx(self) -> None:
        caminho = filedialog.askopenfilename(
            title="Selecione o OFX",
            filetypes=[("OFX", "*.ofx"), ("Todos", "*.*")],
        )
        if not caminho:
            return
        try:
            self.transacoes_ofx, ignorados = ler_ofx(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao ler OFX", str(e))
            return
        self.caminho_ofx = Path(caminho)
        extra = f" ({ignorados} recebimentos ignorados)" if ignorados else ""
        self.lbl_ofx.config(
            text=f"{self.caminho_ofx.name} — {len(self.transacoes_ofx)} pagamentos{extra}"
        )
        self._atualiza_botao()
        self._limpa_resultados()

    def _atualiza_botao(self) -> None:
        pode = bool(self.transacoes_planilha and self.transacoes_ofx)
        self.btn_conciliar.config(state="normal" if pode else "disabled")

    # ---------------------------------------------------- Lógica de matching

    def _limpa_resultados(self) -> None:
        self.pares_conciliados = []
        self.pendentes_planilha = []
        self.pendentes_ofx = []
        self.sugestoes = []
        self._redesenha_abas()
        self.lbl_resumo.config(text="")
        for item in self.tree_dominio.get_children():
            self.tree_dominio.delete(item)
        self.notebook.tab(3, text="Domínio (0)")
        self._atualiza_botao_comparar()

    def _executar_conciliacao(self) -> None:
        pares, pend_p, pend_o = conciliar_automatico(
            self.transacoes_planilha, self.transacoes_ofx,
        )
        self.pares_conciliados = pares
        self.pendentes_planilha = pend_p
        self.pendentes_ofx = pend_o
        self._recalcula_sugestoes()
        self._redesenha_abas()
        self._atualiza_resumo()
        self._atualiza_botao_comparar()

    def _recalcula_sugestoes(self) -> None:
        self.sugestoes = gerar_sugestoes(self.pendentes_planilha, self.pendentes_ofx)

    def _atualiza_resumo(self) -> None:
        self.lbl_resumo.config(
            text=(
                f"  Conciliados: {len(self.pares_conciliados)}   "
                f"Pendentes planilha: {len(self.pendentes_planilha)}   "
                f"Pendentes OFX: {len(self.pendentes_ofx)}   "
                f"Sugestões: {len(self.sugestoes)}"
            )
        )

    # ------------------------------------------------- Ações nas abas

    def _conciliar_selecionadas(self) -> None:
        sel_p = self.tree_pend_p.selection()
        sel_o = self.tree_pend_o.selection()
        if not sel_p or not sel_o:
            messagebox.showinfo(
                "Selecione lançamentos",
                "Escolha uma linha na lista da planilha e outra na do OFX.",
            )
            return
        t_p = self.itens_pendentes_p[sel_p[0]]
        t_o = self.itens_pendentes_o[sel_o[0]]
        d_dias, d_valor = diferenca(t_p, t_o)
        self._aceitar_par(t_p, t_o, d_dias, d_valor)

    def _aceitar_sugestao(self) -> None:
        sel = self.tree_sugestoes.selection()
        if not sel:
            messagebox.showinfo("Selecione uma sugestão", "Escolha uma linha de sugestão.")
            return
        par = self.itens_sugestoes[sel[0]]
        self._aceitar_par(par.planilha, par.ofx, par.diff_dias, par.diff_valor)

    def _aceitar_par(
        self, t_p: Transacao, t_o: Transacao, d_dias: int, d_valor,
    ) -> None:
        novo = Par(
            planilha=t_p, ofx=t_o, tipo="manual",
            diff_dias=d_dias, diff_valor=d_valor,
        )
        self.pares_conciliados.append(novo)
        if t_p in self.pendentes_planilha:
            self.pendentes_planilha.remove(t_p)
        if t_o in self.pendentes_ofx:
            self.pendentes_ofx.remove(t_o)
        self._recalcula_sugestoes()
        self._redesenha_abas()
        self._atualiza_resumo()

    def _desfazer_conciliacao(self) -> None:
        sel = self.tree_conciliados.selection()
        if not sel:
            messagebox.showinfo("Selecione um lançamento", "Escolha uma linha para desfazer.")
            return
        par = self.itens_pares[sel[0]]
        if par.tipo == "auto":
            if not messagebox.askyesno(
                "Confirmar",
                "Esse par foi conciliado automaticamente. Desfazer mesmo assim?",
            ):
                return
        self.pares_conciliados.remove(par)
        self.pendentes_planilha.append(par.planilha)
        self.pendentes_ofx.append(par.ofx)
        self.pendentes_planilha.sort(key=lambda t: (t.data, t.valor))
        self.pendentes_ofx.sort(key=lambda t: (t.data, t.valor))
        self._recalcula_sugestoes()
        self._redesenha_abas()
        self._atualiza_resumo()

    # ----------------------------------------------- Render das tabelas

    def _redesenha_abas(self) -> None:
        self._render_conciliados()
        self._render_pendentes()
        self._render_sugestoes()
        self.notebook.tab(0, text=f"Conciliados ({len(self.pares_conciliados)})")
        self.notebook.tab(
            1, text=f"Pendentes ({len(self.pendentes_planilha)}/{len(self.pendentes_ofx)})",
        )
        self.notebook.tab(2, text=f"Sugestões ({len(self.sugestoes)})")

    def _render_conciliados(self) -> None:
        for item in self.tree_conciliados.get_children():
            self.tree_conciliados.delete(item)
        self.itens_pares.clear()
        for par in self.pares_conciliados:
            diff_txt = ""
            if par.diff_dias or par.diff_valor:
                diff_txt = f"Δ {par.diff_dias}d, R$ {par.diff_valor:.2f}"
            tipo_txt = "Auto" if par.tipo == "auto" else "Manual"
            nf = par.planilha.extras.get("numero_nf", "") or par.ofx.extras.get("numero_nf", "")
            fornecedor = par.planilha.extras.get("fornecedor", "") or par.ofx.extras.get("fornecedor", "")
            iid = self.tree_conciliados.insert(
                "", "end",
                values=(
                    tipo_txt,
                    par.planilha.data.strftime("%d/%m/%Y"),
                    f"{par.planilha.valor:.2f}",
                    nf,
                    fornecedor,
                    par.planilha.descricao,
                    par.ofx.descricao,
                    diff_txt,
                ),
                tags=(par.tipo,),
            )
            self.itens_pares[iid] = par

    def _render_pendentes(self) -> None:
        for item in self.tree_pend_p.get_children():
            self.tree_pend_p.delete(item)
        self.itens_pendentes_p.clear()
        for t in self.pendentes_planilha:
            iid = self.tree_pend_p.insert(
                "", "end",
                values=(t.data.strftime("%d/%m/%Y"), f"{t.valor:.2f}", t.descricao),
            )
            self.itens_pendentes_p[iid] = t

        for item in self.tree_pend_o.get_children():
            self.tree_pend_o.delete(item)
        self.itens_pendentes_o.clear()
        for t in self.pendentes_ofx:
            iid = self.tree_pend_o.insert(
                "", "end",
                values=(t.data.strftime("%d/%m/%Y"), f"{t.valor:.2f}", t.descricao),
            )
            self.itens_pendentes_o[iid] = t

    def _render_sugestoes(self) -> None:
        for item in self.tree_sugestoes.get_children():
            self.tree_sugestoes.delete(item)
        self.itens_sugestoes.clear()
        for par in self.sugestoes:
            iid = self.tree_sugestoes.insert(
                "", "end",
                values=(
                    par.planilha.data.strftime("%d/%m/%Y"),
                    f"{par.planilha.valor:.2f}",
                    par.planilha.descricao,
                    par.ofx.data.strftime("%d/%m/%Y"),
                    f"{par.ofx.valor:.2f}",
                    par.ofx.descricao,
                    str(par.diff_dias),
                    f"{par.diff_valor:.2f}",
                ),
                tags=("destaque",),
            )
            self.itens_sugestoes[iid] = par


if __name__ == "__main__":
    App().mainloop()
