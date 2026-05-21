"""Diálogos para configurar regras de classificação de taxas bancárias."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable


class DialogoNovaRegra(tk.Toplevel):
    """Sub-diálogo simples: pede padrão (substring do memo) + histórico contábil
    + conta contábil. Quando ``plano_contas`` é fornecido, a conta vira um
    Combobox com filtro substring (digite parte do código ou descrição)."""

    def __init__(
        self,
        master: tk.Misc,
        regra_atual: dict[str, Any] | None = None,
        plano_contas: list | None = None,
        tipo: str = "memo",
    ) -> None:
        super().__init__(master)
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.regra: dict[str, str] | None = None
        self.plano_contas = plano_contas or []
        self.tipo = (regra_atual or {}).get("tipo") or tipo
        # opções formatadas: "1.01.001 - CAIXA"
        self._opcoes_conta: list[str] = [
            f"{c.codigo} - {c.descricao}" for c in self.plano_contas
        ]
        # Map código → descrição (pra validar/buscar)
        self._codigos_validos: set[str] = {c.codigo for c in self.plano_contas}

        if self.tipo == "fornecedor":
            self.title("Regra por fornecedor" if not regra_atual else "Editar regra por fornecedor")
            label_padrao = "Padrão (CNPJ ou parte do nome do fornecedor):"
        else:
            self.title("Regra por memo" if not regra_atual else "Editar regra por memo")
            label_padrao = "Padrão do memo (texto que aparece no OFX):"

        ttk.Label(self, text=label_padrao).grid(
            row=0, column=0, padx=10, pady=(10, 2), sticky="w",
        )
        self.entry_padrao = ttk.Entry(self, width=60)
        self.entry_padrao.grid(row=1, column=0, padx=10, pady=2, sticky="we")

        ttk.Label(
            self,
            text="Histórico contábil (texto que vai no lançamento):",
        ).grid(row=2, column=0, padx=10, pady=(10, 2), sticky="w")
        self.entry_historico = ttk.Entry(self, width=60)
        self.entry_historico.grid(row=3, column=0, padx=10, pady=2, sticky="we")

        rotulo_conta = (
            "Conta contábil (digite parte do código ou descrição — selecione no dropdown):"
            if self.plano_contas
            else "Conta contábil (código que vai no lançamento — opcional):"
        )
        ttk.Label(self, text=rotulo_conta).grid(
            row=4, column=0, padx=10, pady=(10, 2), sticky="w",
        )
        if self.plano_contas:
            self.entry_conta: ttk.Entry = ttk.Combobox(
                self, width=60, values=self._opcoes_conta,
            )
            self.entry_conta.bind("<KeyRelease>", self._filtra_contas)
        else:
            self.entry_conta = ttk.Entry(self, width=60)
        self.entry_conta.grid(row=5, column=0, padx=10, pady=2, sticky="we")

        if regra_atual:
            self.entry_padrao.insert(0, regra_atual.get("padrao", ""))
            self.entry_historico.insert(0, regra_atual.get("historico", ""))
            conta_atual = regra_atual.get("conta", "")
            if conta_atual:
                # Tenta achar a descrição no plano pra mostrar formato completo
                opt_completa = next(
                    (o for o in self._opcoes_conta if o.startswith(f"{conta_atual} -")),
                    conta_atual,
                )
                self.entry_conta.insert(0, opt_completa)

        botoes = ttk.Frame(self)
        botoes.grid(row=6, column=0, padx=10, pady=(12, 10), sticky="e")
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(side="right", padx=4)
        ttk.Button(botoes, text="OK", command=self._confirmar).pack(side="right", padx=4)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Return>", lambda _e: self._confirmar())
        self.bind("<Escape>", lambda _e: self._cancelar())
        self.entry_padrao.focus_set()

    def _filtra_contas(self, event) -> None:
        """Filtra as opções do Combobox conforme o usuário digita."""
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        termo = self.entry_conta.get().strip().lower()
        if not termo:
            self.entry_conta["values"] = self._opcoes_conta
        else:
            filtradas = [o for o in self._opcoes_conta if termo in o.lower()]
            self.entry_conta["values"] = filtradas

    def _confirmar(self) -> None:
        padrao = self.entry_padrao.get().strip()
        historico = self.entry_historico.get().strip()
        conta_raw = self.entry_conta.get().strip()
        # Quando há plano carregado e o usuário selecionou "código - descrição",
        # extrai só o código pra salvar limpo.
        if self.plano_contas and " - " in conta_raw:
            codigo_candidato = conta_raw.split(" - ", 1)[0].strip()
            if codigo_candidato in self._codigos_validos:
                conta = codigo_candidato
            else:
                conta = conta_raw
        else:
            conta = conta_raw
        if not padrao:
            messagebox.showwarning("Campo vazio", "Informe o padrão do memo.", parent=self)
            return
        if not historico:
            messagebox.showwarning(
                "Campo vazio", "Informe o histórico contábil.", parent=self,
            )
            return
        self.regra = {
            "tipo": self.tipo,
            "padrao": padrao,
            "historico": historico,
            "conta": conta,
        }
        self.destroy()

    def _cancelar(self) -> None:
        self.regra = None
        self.destroy()


class DialogoLancamentoManual(tk.Toplevel):
    """Diálogo simples para criar UM lançamento contábil avulso a partir de
    um par P×OFX que falta no Domínio. Pede só conta + histórico — NÃO cria
    regra. Devolve dict ``{"historico": ..., "conta": ...}`` em ``self.resultado``
    ou ``None`` se cancelar.
    """

    def __init__(
        self,
        master: tk.Misc,
        par,                              # matcher.Par
        plano_contas: list | None = None,
        sugestao_historico: str = "",
    ) -> None:
        super().__init__(master)
        self.title("Lançamento contábil manual")
        self.transient(master)
        self.grab_set()
        self.resizable(False, False)

        self.par = par
        self.plano_contas = plano_contas or []
        self.resultado: dict[str, str] | None = None
        self._opcoes_conta = [
            f"{c.codigo} - {c.descricao}" for c in self.plano_contas
        ]
        self._codigos_validos = {c.codigo for c in self.plano_contas}

        # Resumo do par selecionado (read-only)
        cab = ttk.LabelFrame(self, text="Par selecionado")
        cab.grid(row=0, column=0, padx=10, pady=(10, 4), sticky="we")
        vcto = par.planilha.data.strftime("%d/%m/%Y")
        valor = f"R$ {par.planilha.valor:.2f}"
        forn = par.planilha.extras.get("fornecedor", "") or ""
        cnpj = par.planilha.extras.get("cnpj", "") or ""
        nf = par.planilha.extras.get("numero_nf", "") or ""
        memo = par.ofx.descricao or ""
        for i, (rotulo, txt) in enumerate([
            ("Vencimento", vcto),
            ("Valor",     valor),
            ("Nº NF",     nf),
            ("Fornecedor", f"{forn} ({cnpj})" if cnpj else forn),
            ("Memo OFX",  memo[:80]),
        ]):
            ttk.Label(cab, text=f"{rotulo}:", foreground="#555").grid(
                row=i, column=0, padx=6, pady=1, sticky="e",
            )
            ttk.Label(cab, text=txt).grid(row=i, column=1, padx=6, pady=1, sticky="w")

        ttk.Label(self, text="Histórico contábil:").grid(
            row=1, column=0, padx=10, pady=(8, 2), sticky="w",
        )
        self.entry_historico = ttk.Entry(self, width=60)
        self.entry_historico.grid(row=2, column=0, padx=10, pady=2, sticky="we")
        if sugestao_historico:
            self.entry_historico.insert(0, sugestao_historico)

        rotulo_conta = (
            "Conta contábil (digite parte do código ou descrição):"
            if self.plano_contas
            else "Conta contábil:"
        )
        ttk.Label(self, text=rotulo_conta).grid(
            row=3, column=0, padx=10, pady=(8, 2), sticky="w",
        )
        if self.plano_contas:
            self.entry_conta: ttk.Entry = ttk.Combobox(
                self, width=60, values=self._opcoes_conta,
            )
            self.entry_conta.bind("<KeyRelease>", self._filtra_contas)
        else:
            self.entry_conta = ttk.Entry(self, width=60)
        self.entry_conta.grid(row=4, column=0, padx=10, pady=2, sticky="we")

        botoes = ttk.Frame(self)
        botoes.grid(row=5, column=0, padx=10, pady=(12, 10), sticky="e")
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(side="right", padx=4)
        ttk.Button(botoes, text="Lançar", command=self._confirmar).pack(side="right", padx=4)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Return>", lambda _e: self._confirmar())
        self.bind("<Escape>", lambda _e: self._cancelar())
        self.entry_historico.focus_set()

    def _filtra_contas(self, event) -> None:
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        termo = self.entry_conta.get().strip().lower()
        if not termo:
            self.entry_conta["values"] = self._opcoes_conta
        else:
            self.entry_conta["values"] = [
                o for o in self._opcoes_conta if termo in o.lower()
            ]

    def _confirmar(self) -> None:
        historico = self.entry_historico.get().strip()
        conta_raw = self.entry_conta.get().strip()
        if self.plano_contas and " - " in conta_raw:
            codigo_candidato = conta_raw.split(" - ", 1)[0].strip()
            conta = (
                codigo_candidato if codigo_candidato in self._codigos_validos
                else conta_raw
            )
        else:
            conta = conta_raw
        if not historico:
            messagebox.showwarning(
                "Campo vazio", "Informe o histórico contábil.", parent=self,
            )
            return
        self.resultado = {"historico": historico, "conta": conta}
        self.destroy()

    def _cancelar(self) -> None:
        self.resultado = None
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
        plano_contas: list | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Configurar regras de taxas / movimentações")
        self.transient(master)
        self.grab_set()
        self.geometry("680x420")

        # Cópia de trabalho — só persiste via on_change
        self.regras: list[dict[str, Any]] = [dict(r) for r in regras]
        self.on_change = on_change
        self.plano_contas = plano_contas or []

        info = ttk.Label(
            self,
            text=(
                "Regras de classificação de lançamentos. Dois tipos:\n"
                "• Memo: casa contra o memo do OFX nos pendentes (tarifas, IOF, juros).\n"
                "• Fornecedor: casa contra CNPJ/nome do fornecedor nos pares "
                "conciliados que faltam no Domínio."
            ),
            wraplength=660,
        )
        info.pack(fill="x", padx=10, pady=(10, 4))

        botoes_top = ttk.Frame(self)
        botoes_top.pack(fill="x", padx=10, pady=4)
        ttk.Button(
            botoes_top, text="+ Regra por memo",
            command=lambda: self._adicionar(tipo="memo"),
        ).pack(side="left", padx=2)
        ttk.Button(
            botoes_top, text="+ Regra por fornecedor",
            command=lambda: self._adicionar(tipo="fornecedor"),
        ).pack(side="left", padx=2)
        ttk.Button(botoes_top, text="✎ Editar", command=self._editar).pack(side="left", padx=2)
        ttk.Button(botoes_top, text="− Remover", command=self._remover).pack(side="left", padx=2)
        self.lbl_count = ttk.Label(botoes_top, text="", foreground="#666")
        self.lbl_count.pack(side="left", padx=8)

        corpo = ttk.Frame(self)
        corpo.pack(fill="both", expand=True, padx=10, pady=4)
        cols = ("tipo", "padrao", "historico", "conta")
        self.tree = ttk.Treeview(corpo, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("padrao", text="Padrão")
        self.tree.heading("historico", text="Histórico contábil")
        self.tree.heading("conta", text="Conta contábil")
        self.tree.column("tipo", width=90, anchor="center")
        self.tree.column("padrao", width=180, anchor="w")
        self.tree.column("historico", width=260, anchor="w")
        self.tree.column("conta", width=110, anchor="w")
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
            tipo = (r.get("tipo") or "memo").capitalize()
            self.tree.insert(
                "", "end",
                values=(
                    tipo,
                    r.get("padrao", ""),
                    r.get("historico", ""),
                    r.get("conta", ""),
                ),
            )
        self.lbl_count.config(text=f"{len(self.regras)} regra(s)")

    def _adicionar(self, tipo: str = "memo") -> None:
        dlg = DialogoNovaRegra(self, plano_contas=self.plano_contas, tipo=tipo)
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
        dlg = DialogoNovaRegra(
            self, self.regras[idx], plano_contas=self.plano_contas,
        )
        self.wait_window(dlg)
        if dlg.regra:
            self.regras[idx] = dlg.regra
            self._renderiza()
            self.on_change(self.regras)
