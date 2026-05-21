"""Diálogos para configurar regras de classificação de taxas bancárias."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable


class DialogoNovaRegra(tk.Toplevel):
    """Sub-diálogo simples: pede padrão (substring do memo) + histórico contábil."""

    def __init__(
        self,
        master: tk.Misc,
        regra_atual: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Nova regra" if not regra_atual else "Editar regra")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.regra: dict[str, str] | None = None

        ttk.Label(
            self,
            text="Padrão do memo (texto que aparece no OFX, sem distinguir maiúsculas):",
        ).grid(row=0, column=0, padx=10, pady=(10, 2), sticky="w")
        self.entry_padrao = ttk.Entry(self, width=50)
        self.entry_padrao.grid(row=1, column=0, padx=10, pady=2, sticky="we")

        ttk.Label(
            self,
            text="Histórico contábil (texto que vai no lançamento):",
        ).grid(row=2, column=0, padx=10, pady=(10, 2), sticky="w")
        self.entry_historico = ttk.Entry(self, width=50)
        self.entry_historico.grid(row=3, column=0, padx=10, pady=2, sticky="we")

        ttk.Label(
            self,
            text="Conta contábil (código que vai no lançamento — opcional):",
        ).grid(row=4, column=0, padx=10, pady=(10, 2), sticky="w")
        self.entry_conta = ttk.Entry(self, width=50)
        self.entry_conta.grid(row=5, column=0, padx=10, pady=2, sticky="we")

        if regra_atual:
            self.entry_padrao.insert(0, regra_atual.get("padrao", ""))
            self.entry_historico.insert(0, regra_atual.get("historico", ""))
            self.entry_conta.insert(0, regra_atual.get("conta", ""))

        botoes = ttk.Frame(self)
        botoes.grid(row=6, column=0, padx=10, pady=(12, 10), sticky="e")
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(side="right", padx=4)
        ttk.Button(botoes, text="OK", command=self._confirmar).pack(side="right", padx=4)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Return>", lambda _e: self._confirmar())
        self.bind("<Escape>", lambda _e: self._cancelar())
        self.entry_padrao.focus_set()

    def _confirmar(self) -> None:
        padrao = self.entry_padrao.get().strip()
        historico = self.entry_historico.get().strip()
        conta = self.entry_conta.get().strip()
        if not padrao:
            messagebox.showwarning("Campo vazio", "Informe o padrão do memo.", parent=self)
            return
        if not historico:
            messagebox.showwarning(
                "Campo vazio", "Informe o histórico contábil.", parent=self,
            )
            return
        self.regra = {"padrao": padrao, "historico": historico, "conta": conta}
        self.destroy()

    def _cancelar(self) -> None:
        self.regra = None
        self.destroy()


class DialogoConfigurarTaxas(tk.Toplevel):
    """Listagem editável de regras de taxas/movimentações bancárias.

    Recebe a lista atual e um callback ``on_change(regras_novas)`` que é
    chamado a cada alteração (adicionar/remover/editar) — o caller persiste.
    """

    def __init__(
        self,
        master: tk.Misc,
        regras: list[dict[str, Any]],
        on_change: Callable[[list[dict[str, Any]]], None],
    ) -> None:
        super().__init__(master)
        self.title("Configurar regras de taxas / movimentações")
        self.transient(master)
        self.grab_set()
        self.geometry("680x420")

        # Cópia de trabalho — só persiste via on_change
        self.regras: list[dict[str, Any]] = [dict(r) for r in regras]
        self.on_change = on_change

        info = ttk.Label(
            self,
            text=(
                "Cada lançamento OFX cujo memo contenha um destes padrões "
                "(substring, sem distinguir maiúsculas) será classificado "
                "automaticamente como lançamento contábil e aparecerá na "
                "aba 'Lançamentos contábeis'."
            ),
            wraplength=660,
        )
        info.pack(fill="x", padx=10, pady=(10, 4))

        botoes_top = ttk.Frame(self)
        botoes_top.pack(fill="x", padx=10, pady=4)
        ttk.Button(botoes_top, text="+ Adicionar", command=self._adicionar).pack(side="left", padx=2)
        ttk.Button(botoes_top, text="✎ Editar", command=self._editar).pack(side="left", padx=2)
        ttk.Button(botoes_top, text="− Remover", command=self._remover).pack(side="left", padx=2)
        self.lbl_count = ttk.Label(botoes_top, text="", foreground="#666")
        self.lbl_count.pack(side="left", padx=8)

        corpo = ttk.Frame(self)
        corpo.pack(fill="both", expand=True, padx=10, pady=4)
        cols = ("padrao", "historico", "conta")
        self.tree = ttk.Treeview(corpo, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("padrao", text="Padrão (memo)")
        self.tree.heading("historico", text="Histórico contábil")
        self.tree.heading("conta", text="Conta contábil")
        self.tree.column("padrao", width=200, anchor="w")
        self.tree.column("historico", width=300, anchor="w")
        self.tree.column("conta", width=120, anchor="w")
        sb = ttk.Scrollbar(corpo, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-Button-1>", lambda _e: self._editar())

        fim = ttk.Frame(self)
        fim.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Button(fim, text="Fechar", command=self.destroy).pack(side="right")

        self._renderiza()

    def _renderiza(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in self.regras:
            self.tree.insert(
                "", "end",
                values=(
                    r.get("padrao", ""),
                    r.get("historico", ""),
                    r.get("conta", ""),
                ),
            )
        self.lbl_count.config(text=f"{len(self.regras)} regra(s)")

    def _adicionar(self) -> None:
        dlg = DialogoNovaRegra(self)
        self.wait_window(dlg)
        if dlg.regra:
            self.regras.append(dlg.regra)
            self._renderiza()
            self.on_change(self.regras)

    def _remover(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sem seleção", "Selecione uma regra para remover.", parent=self)
            return
        idx = self.tree.index(sel[0])
        regra = self.regras[idx]
        if messagebox.askyesno(
            "Confirmar remoção",
            f"Remover a regra '{regra.get('padrao')}'?",
            parent=self,
        ):
            del self.regras[idx]
            self._renderiza()
            self.on_change(self.regras)

    def _editar(self) -> None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Sem seleção", "Selecione uma regra para editar.", parent=self)
            return
        idx = self.tree.index(sel[0])
        dlg = DialogoNovaRegra(self, self.regras[idx])
        self.wait_window(dlg)
        if dlg.regra:
            self.regras[idx] = dlg.regra
            self._renderiza()
            self.on_change(self.regras)
