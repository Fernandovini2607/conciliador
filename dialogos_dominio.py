"""Diálogos para conectar ao Domínio e selecionar a fonte de pagamentos."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any

import pyodbc

import parser_dominio


class DialogoConexao(tk.Toplevel):
    """Pede DSN, usuário e senha. Tenta conectar e devolve o objeto Connection."""

    def __init__(self, master: tk.Misc, config_atual: dict[str, Any]) -> None:
        super().__init__(master)
        self.title("Conectar ao Domínio")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.conn: pyodbc.Connection | None = None
        self.dados: dict[str, str] | None = None

        info = ttk.Label(
            self,
            text=(
                "Informe o DSN ODBC do Domínio (configurado no Windows) "
                "e as credenciais do banco."
            ),
            wraplength=420,
        )
        info.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w")

        ttk.Label(self, text="DSN:").grid(row=1, column=0, padx=12, pady=4, sticky="e")
        self.e_dsn = ttk.Entry(self, width=36)
        self.e_dsn.grid(row=1, column=1, padx=(0, 12), pady=4, sticky="w")
        self.e_dsn.insert(0, config_atual.get("dsn", ""))

        ttk.Label(self, text="Usuário:").grid(row=2, column=0, padx=12, pady=4, sticky="e")
        self.e_user = ttk.Entry(self, width=36)
        self.e_user.grid(row=2, column=1, padx=(0, 12), pady=4, sticky="w")
        self.e_user.insert(0, config_atual.get("usuario", ""))

        ttk.Label(self, text="Senha:").grid(row=3, column=0, padx=12, pady=4, sticky="e")
        self.e_pass = ttk.Entry(self, width=36, show="•")
        self.e_pass.grid(row=3, column=1, padx=(0, 12), pady=4, sticky="w")
        self.e_pass.insert(0, config_atual.get("senha", ""))

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
        dsn = self.e_dsn.get().strip()
        usuario = self.e_user.get().strip()
        senha = self.e_pass.get()
        if not dsn:
            messagebox.showwarning("DSN vazio", "Informe o DSN.", parent=self)
            return
        self.lbl_status.config(text="Conectando…")
        self.update_idletasks()
        try:
            self.conn = parser_dominio.conectar(dsn, usuario, senha)
        except pyodbc.Error as e:
            self.lbl_status.config(text="")
            messagebox.showerror("Falha na conexão", str(e), parent=self)
            return
        self.dados = {"dsn": dsn, "usuario": usuario, "senha": senha}
        self.destroy()

    def _cancelar(self) -> None:
        self.conn = None
        self.dados = None
        self.destroy()


class DialogoFonte(tk.Toplevel):
    """Depois de conectar: escolhe tabela (ou query SQL), mostra amostra
    e mapeia colunas Data/Valor/Descrição."""

    PREVIEW_LINHAS = 15
    CAMPOS = [("data", "Data"), ("valor", "Valor"), ("descricao", "Descrição")]

    def __init__(
        self,
        master: tk.Misc,
        conn: pyodbc.Connection,
        config_atual: dict[str, Any],
    ) -> None:
        super().__init__(master)
        self.title("Selecionar fonte de pagamentos no Domínio")
        self.transient(master)
        self.grab_set()
        self.geometry("960x640")

        self.conn = conn
        self.fonte: dict[str, Any] | None = None
        self.colunas_atuais: list[str] = []
        self.tabelas: list[str] = []

        self._monta_ui(config_atual)
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

        # área "tabela"
        self.frame_tabela = ttk.Frame(self)
        ttk.Label(self.frame_tabela, text="Tabela:").grid(row=0, column=0, padx=4, pady=4, sticky="e")
        self.cb_tabela = ttk.Combobox(self.frame_tabela, state="readonly", width=50)
        self.cb_tabela.grid(row=0, column=1, padx=4, pady=4, sticky="w")
        self.cb_tabela.bind("<<ComboboxSelected>>", lambda _e: self._carrega_amostra())
        if cfg.get("tabela"):
            self.cb_tabela.set(cfg["tabela"])
        ttk.Label(self.frame_tabela, text="WHERE (opcional):").grid(row=1, column=0, padx=4, pady=4, sticky="ne")
        self.txt_where = tk.Text(self.frame_tabela, width=60, height=2)
        self.txt_where.grid(row=1, column=1, padx=4, pady=4, sticky="w")
        self.txt_where.insert("1.0", cfg.get("where", ""))

        # área "sql"
        self.frame_sql = ttk.Frame(self)
        ttk.Label(self.frame_sql, text="SELECT ...:").pack(anchor="w")
        self.txt_sql = tk.Text(self.frame_sql, width=100, height=5)
        self.txt_sql.pack(fill="x")
        self.txt_sql.insert("1.0", cfg.get("sql", ""))

        # botão "Carregar amostra"
        self.frame_acao = ttk.Frame(self)
        ttk.Button(
            self.frame_acao, text="Carregar amostra",
            command=self._carrega_amostra,
        ).pack(side="left", padx=4)

        # mapeamento
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

        # preview
        self.frame_preview = ttk.Frame(self)
        self.tree = ttk.Treeview(self.frame_preview, show="headings")
        sb = ttk.Scrollbar(self.frame_preview, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # botões finais
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
        try:
            self.tabelas = parser_dominio.listar_tabelas(self.conn)
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
        # reconstrói treeview com as colunas
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

        # atualiza comboboxes mantendo a seleção atual se possível
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
