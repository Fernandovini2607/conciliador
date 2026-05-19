"""Diálogos para conectar ao Domínio e selecionar a fonte de pagamentos."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any

import pyodbc

import parser_dominio


class DialogoConexao(tk.Toplevel):
    """Pede DSN, usuário e senha. Tenta abrir conexão read-only e devolve o handle.

    As credenciais são persistidas em ``data/dominio_config.json`` via
    ``parser_dominio.save_odbc_config``.
    """

    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self.title("Conectar ao Domínio (ODBC, read-only)")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.conn: pyodbc.Connection | None = None

        cfg_atual = parser_dominio.load_odbc_config()

        info = ttk.Label(
            self,
            text=(
                "Informe o DSN ODBC do Domínio (configurado no Windows) e as "
                "credenciais do banco SQL Anywhere. A conexão é aberta como "
                "read-only — toda escrita exige RPA, fora do escopo deste app."
            ),
            wraplength=460,
        )
        info.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w")

        ttk.Label(self, text="DSN:").grid(row=1, column=0, padx=12, pady=4, sticky="e")
        self.e_dsn = ttk.Entry(self, width=36)
        self.e_dsn.grid(row=1, column=1, padx=(0, 12), pady=4, sticky="w")
        self.e_dsn.insert(0, cfg_atual.get("dsn", ""))

        ttk.Label(self, text="Usuário:").grid(row=2, column=0, padx=12, pady=4, sticky="e")
        self.e_user = ttk.Entry(self, width=36)
        self.e_user.grid(row=2, column=1, padx=(0, 12), pady=4, sticky="w")
        self.e_user.insert(0, cfg_atual.get("usuario", ""))

        ttk.Label(self, text="Senha:").grid(row=3, column=0, padx=12, pady=4, sticky="e")
        self.e_pass = ttk.Entry(self, width=36, show="•")
        self.e_pass.grid(row=3, column=1, padx=(0, 12), pady=4, sticky="w")
        self.e_pass.insert(0, cfg_atual.get("senha", ""))

        self.lbl_status = ttk.Label(self, text="")
        self.lbl_status.grid(row=4, column=0, columnspan=2, padx=12, pady=(4, 0), sticky="w")

        botoes = ttk.Frame(self)
        botoes.grid(row=5, column=0, columnspan=2, padx=12, pady=(8, 12), sticky="e")
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(side="right", padx=4)
        ttk.Button(botoes, text="Conectar", command=self._conectar).pack(side="right", padx=4)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Return>", lambda _e: self._conectar())
        self.bind("<Escape>", lambda _e: self._cancelar())

    def _conectar(self) -> None:
        cfg = {
            "dsn": self.e_dsn.get().strip(),
            "usuario": self.e_user.get().strip(),
            "senha": self.e_pass.get(),
        }
        if not cfg["dsn"]:
            messagebox.showwarning("DSN vazio", "Informe o DSN.", parent=self)
            return
        self.lbl_status.config(text="Conectando…")
        self.update_idletasks()
        try:
            self.conn = parser_dominio.open_connection(readonly=True, cfg=cfg)
        except pyodbc.Error as e:
            self.lbl_status.config(text="")
            messagebox.showerror("Falha na conexão", str(e), parent=self)
            return
        parser_dominio.save_odbc_config(cfg)
        self.destroy()

    def _cancelar(self) -> None:
        self.conn = None
        self.destroy()


class DialogoSelecionarEmpresa(tk.Toplevel):
    """Lista as empresas de ``bethadba.geempre`` com campo de busca.

    Recebe a conexão e devolve em ``self.empresa`` o dicionário escolhido
    ``{"codi_emp": int, "razao": str, "cnpj": str}`` (ou ``None`` se cancelar).
    """

    def __init__(
        self,
        master: tk.Misc,
        conn: pyodbc.Connection,
        empresa_atual: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Selecionar empresa")
        self.transient(master)
        self.grab_set()
        self.geometry("680x520")

        self.conn = conn
        self.empresa: dict[str, Any] | None = None
        self.empresas: list[dict[str, Any]] = []
        self.codi_atual = (empresa_atual or {}).get("codi_emp")

        topo = ttk.Frame(self, padding=(10, 10, 10, 4))
        topo.pack(fill="x")
        ttk.Label(topo, text="Buscar:").pack(side="left")
        self.var_busca = tk.StringVar()
        self.var_busca.trace_add("write", lambda *_a: self._aplica_filtro())
        ent = ttk.Entry(topo, textvariable=self.var_busca, width=40)
        ent.pack(side="left", padx=6)
        ent.focus_set()

        ttk.Button(topo, text="Recarregar", command=self._carregar).pack(side="left", padx=4)

        meio = ttk.Frame(self, padding=(10, 0, 10, 4))
        meio.pack(fill="both", expand=True)
        cols = ("codi", "razao", "cnpj")
        self.tree = ttk.Treeview(meio, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("codi", text="Código")
        self.tree.heading("razao", text="Razão social")
        self.tree.heading("cnpj", text="CNPJ")
        self.tree.column("codi", width=70, anchor="center")
        self.tree.column("razao", width=380, anchor="w")
        self.tree.column("cnpj", width=140, anchor="w")
        sb = ttk.Scrollbar(meio, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-Button-1>", lambda _e: self._confirmar())

        rodape = ttk.Frame(self, padding=(10, 4, 10, 10))
        rodape.pack(fill="x")
        self.lbl_status = ttk.Label(rodape, text="")
        self.lbl_status.pack(side="left")
        ttk.Button(rodape, text="Cancelar", command=self._cancelar).pack(side="right", padx=4)
        ttk.Button(rodape, text="Selecionar", command=self._confirmar).pack(side="right", padx=4)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Escape>", lambda _e: self._cancelar())
        self.bind("<Return>", lambda _e: self._confirmar())

        self._carregar()

    def _carregar(self) -> None:
        self.lbl_status.config(text="Carregando empresas…")
        self.update_idletasks()
        try:
            self.empresas = parser_dominio.listar_empresas(self.conn)
        except Exception as e:
            messagebox.showerror("Erro ao listar empresas", str(e), parent=self)
            self.empresas = []
        self.lbl_status.config(text=f"{len(self.empresas)} empresas")
        self._aplica_filtro()

    def _aplica_filtro(self) -> None:
        termo = self.var_busca.get().strip().lower()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for emp in self.empresas:
            if termo:
                buscavel = f"{emp['codi_emp']} {emp['razao']} {emp['cnpj']}".lower()
                if termo not in buscavel:
                    continue
            iid = self.tree.insert(
                "", "end", values=(emp["codi_emp"], emp["razao"], emp["cnpj"]),
            )
            if emp["codi_emp"] == self.codi_atual:
                self.tree.selection_set(iid)
                self.tree.see(iid)

    def _confirmar(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sem seleção", "Escolha uma empresa.", parent=self)
            return
        vals = self.tree.item(sel[0], "values")
        codi = int(vals[0]) if str(vals[0]).isdigit() else vals[0]
        self.empresa = {"codi_emp": codi, "razao": vals[1], "cnpj": vals[2]}
        self.destroy()

    def _cancelar(self) -> None:
        self.empresa = None
        self.destroy()


class DialogoFonte(tk.Toplevel):
    """Depois de conectar: escolhe tabela (ou query SQL), mostra amostra
    e mapeia colunas Data/Valor/Descrição.

    Recebe e devolve um dict de **fonte** (``dominio_fonte`` no config.json).
    """

    PREVIEW_LINHAS = 15
    CAMPOS = [("data", "Data"), ("valor", "Valor"), ("descricao", "Descrição")]

    def __init__(
        self,
        master: tk.Misc,
        conn: pyodbc.Connection,
        fonte_atual: dict[str, Any],
    ) -> None:
        super().__init__(master)
        self.title("Selecionar fonte de pagamentos no Domínio")
        self.transient(master)
        self.grab_set()
        self.geometry("960x680")

        self.conn = conn
        self.fonte: dict[str, Any] | None = None
        self.colunas_atuais: list[str] = []
        self.tabelas: list[str] = []
        self.var_todas_tabelas = tk.BooleanVar(value=False)

        self._monta_ui(fonte_atual)
        self._carrega_tabelas()

    def _monta_ui(self, cfg: dict[str, Any]) -> None:
        topo = ttk.Frame(self)
        topo.pack(fill="x", padx=10, pady=10)

        self.modo_var = tk.StringVar(value=cfg.get("modo", "tabela"))
        ttk.Radiobutton(
            topo, text="Tabela + filtro", variable=self.modo_var,
            value="tabela", command=self._troca_modo,
        ).pack(side="left", padx=4)
        ttk.Radiobutton(
            topo, text="Query SQL manual", variable=self.modo_var,
            value="sql", command=self._troca_modo,
        ).pack(side="left", padx=4)

        # ----- modo "tabela"
        self.frame_tabela = ttk.Frame(self)

        linha1 = ttk.Frame(self.frame_tabela)
        linha1.grid(row=0, column=0, columnspan=2, sticky="we", pady=2)
        ttk.Label(linha1, text="Tabela:").pack(side="left", padx=(0, 4))
        self.cb_tabela = ttk.Combobox(linha1, state="readonly", width=50)
        self.cb_tabela.pack(side="left", padx=4)
        self.cb_tabela.bind("<<ComboboxSelected>>", lambda _e: self._carrega_amostra())
        if cfg.get("tabela"):
            self.cb_tabela.set(cfg["tabela"])
        ttk.Checkbutton(
            linha1, text=f"Mostrar todas (não só {parser_dominio.SCHEMA_PADRAO})",
            variable=self.var_todas_tabelas,
            command=self._carrega_tabelas,
        ).pack(side="left", padx=8)

        ttk.Label(self.frame_tabela, text="WHERE (opcional):").grid(
            row=1, column=0, padx=4, pady=4, sticky="ne",
        )
        self.txt_where = tk.Text(self.frame_tabela, width=60, height=2)
        self.txt_where.grid(row=1, column=1, padx=4, pady=4, sticky="w")
        self.txt_where.insert("1.0", cfg.get("where", ""))

        # ----- modo "sql"
        self.frame_sql = ttk.Frame(self)
        ttk.Label(self.frame_sql, text="SELECT (use bethadba.<tabela>):").pack(anchor="w")
        self.txt_sql = tk.Text(self.frame_sql, width=100, height=5)
        self.txt_sql.pack(fill="x")
        self.txt_sql.insert("1.0", cfg.get("sql", ""))

        # ----- ação carregar amostra
        self.frame_acao = ttk.Frame(self)
        ttk.Button(
            self.frame_acao, text="Carregar amostra",
            command=self._carrega_amostra,
        ).pack(side="left", padx=4)

        # ----- mapeamento
        self.frame_map = ttk.Frame(self)
        self.combos: dict[str, ttk.Combobox] = {}
        for i, (campo, rotulo) in enumerate(self.CAMPOS):
            ttk.Label(self.frame_map, text=f"{rotulo}:").grid(
                row=0, column=i * 2, padx=(0, 4), sticky="e",
            )
            cb = ttk.Combobox(self.frame_map, state="readonly", width=24)
            cb.grid(row=0, column=i * 2 + 1, padx=(0, 12), sticky="w")
            mapeamento = cfg.get("mapeamento", {})
            if mapeamento.get(campo):
                cb.set(mapeamento[campo])
            self.combos[campo] = cb

        # ----- preview
        self.frame_preview = ttk.Frame(self)
        self.tree = ttk.Treeview(self.frame_preview, show="headings")
        sb = ttk.Scrollbar(self.frame_preview, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ----- botões finais
        self.frame_fim = ttk.Frame(self)
        ttk.Button(self.frame_fim, text="Cancelar", command=self._cancelar).pack(side="right", padx=4)
        ttk.Button(self.frame_fim, text="Confirmar", command=self._confirmar).pack(side="right", padx=4)

        # layout em ordem
        self.frame_tabela.pack(fill="x", padx=10, pady=4)
        self.frame_sql.pack(fill="x", padx=10, pady=4)
        self.frame_acao.pack(fill="x", padx=10, pady=4)
        self.frame_map.pack(fill="x", padx=10, pady=6)
        ttk.Label(self, text="Amostra (15 linhas):").pack(anchor="w", padx=10)
        self.frame_preview.pack(fill="both", expand=True, padx=10, pady=4)
        self.frame_fim.pack(fill="x", padx=10, pady=8)

        self._troca_modo()
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

    def _troca_modo(self) -> None:
        if self.modo_var.get() == "tabela":
            self.frame_sql.pack_forget()
            if not self.frame_tabela.winfo_ismapped():
                self.frame_tabela.pack(fill="x", padx=10, pady=4, before=self.frame_sql)
        else:
            self.frame_tabela.pack_forget()
            if not self.frame_sql.winfo_ismapped():
                self.frame_sql.pack(fill="x", padx=10, pady=4)

    def _carrega_tabelas(self) -> None:
        schema = None if self.var_todas_tabelas.get() else parser_dominio.SCHEMA_PADRAO
        try:
            self.tabelas = parser_dominio.listar_tabelas(self.conn, schema=schema)
        except Exception as e:
            messagebox.showerror("Erro ao listar tabelas", str(e), parent=self)
            return
        self.cb_tabela["values"] = self.tabelas

    def _carrega_amostra(self) -> None:
        try:
            if self.modo_var.get() == "tabela":
                tabela = self.cb_tabela.get().strip()
                if not tabela:
                    messagebox.showinfo("Tabela", "Escolha uma tabela.", parent=self)
                    return
                colunas, linhas = parser_dominio.amostra(self.conn, tabela, self.PREVIEW_LINHAS)
            else:
                sql = self.txt_sql.get("1.0", "end").strip()
                if not sql:
                    messagebox.showinfo("SQL", "Cole uma query SQL.", parent=self)
                    return
                colunas, linhas = parser_dominio.executar_query(self.conn, sql)
                linhas = linhas[: self.PREVIEW_LINHAS]
        except Exception as e:
            messagebox.showerror("Erro ao carregar", str(e), parent=self)
            return

        self.colunas_atuais = colunas
        for c in self.tree["columns"]:
            self.tree.heading(c, text="")
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = colunas
        for c in colunas:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="w")
        for linha in linhas:
            self.tree.insert("", "end", values=tuple(str(c) if c is not None else "" for c in linha))

        for campo, cb in self.combos.items():
            atual = cb.get()
            cb["values"] = colunas
            if atual in colunas:
                cb.set(atual)
            else:
                cb.set("")

    def _confirmar(self) -> None:
        if not self.colunas_atuais:
            messagebox.showinfo(
                "Carregue uma amostra",
                "Clique em 'Carregar amostra' antes de confirmar.",
                parent=self,
            )
            return
        mapa: dict[str, str] = {}
        for campo, _ in self.CAMPOS:
            valor = self.combos[campo].get().strip()
            if not valor:
                messagebox.showwarning(
                    "Mapeamento incompleto",
                    f"Escolha a coluna correspondente a {campo}.",
                    parent=self,
                )
                return
            mapa[campo] = valor

        modo = self.modo_var.get()
        if modo == "tabela":
            tabela = self.cb_tabela.get().strip()
            if not tabela:
                messagebox.showwarning("Tabela vazia", "Escolha uma tabela.", parent=self)
                return
            self.fonte = {
                "modo": "tabela",
                "tabela": tabela,
                "where": self.txt_where.get("1.0", "end").strip(),
                "mapeamento": mapa,
            }
        else:
            sql = self.txt_sql.get("1.0", "end").strip()
            if not sql:
                messagebox.showwarning("SQL vazio", "Cole uma query SQL.", parent=self)
                return
            self.fonte = {"modo": "sql", "sql": sql, "mapeamento": mapa}
        self.destroy()

    def _cancelar(self) -> None:
        self.fonte = None
        self.destroy()
