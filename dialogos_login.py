"""Diálogos de autenticação e gerenciamento de usuários.

- DialogoLogin           — tela de login inicial (bloqueia até logar/cancelar)
- DialogoMudarSenha      — troca senha própria (ou de outro usuário se admin)
- DialogoCadastroUsuario — cria/edita usuário (só admin)
- DialogoGerenciarUsuarios — lista + botões (só admin)
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Callable

import auth


class DialogoLogin(tk.Toplevel):
    """Tela de login. Usado no startup do app.

    Após confirmar sucesso, ``self.usuario`` fica populado com o dict do
    usuário logado. Se o usuário fechar sem logar (Esc, X), ``self.usuario``
    permanece None e o app deve encerrar."""

    def __init__(self, master: tk.Misc | None = None) -> None:
        super().__init__(master)
        self.title("Conciliador — Login")
        self.resizable(False, False)
        # NÃO usa transient(master) porque no Windows em algumas versões
        # isso deixa o diálogo escondido atrás da janela principal.
        # grab_set garante que só esse diálogo aceita input.
        self.grab_set()

        self.usuario: dict[str, Any] | None = None

        frm = ttk.Frame(self, padding=20)
        frm.grid(row=0, column=0)

        ttk.Label(
            frm, text="Conciliador OFX × Planilha × Domínio",
            font=("TkDefaultFont", 11, "bold"),
            foreground="#1f3a68",
        ).grid(row=0, column=0, columnspan=2, pady=(0, 12))

        ttk.Label(frm, text="Usuário:").grid(
            row=1, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_user = ttk.Entry(frm, width=28)
        self.entry_user.grid(row=1, column=1, pady=4, sticky="w")

        ttk.Label(frm, text="Senha:").grid(
            row=2, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_senha = ttk.Entry(frm, width=28, show="•")
        self.entry_senha.grid(row=2, column=1, pady=4, sticky="w")

        self.lbl_erro = ttk.Label(frm, text="", foreground="#c0392b")
        self.lbl_erro.grid(row=3, column=0, columnspan=2, pady=(6, 2))

        botoes = ttk.Frame(frm)
        botoes.grid(row=4, column=0, columnspan=2, pady=(8, 0))
        ttk.Button(botoes, text="Entrar", command=self._entrar).pack(
            side="left", padx=4,
        )
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(
            side="left", padx=4,
        )

        self.bind("<Return>", lambda _e: self._entrar())
        self.bind("<Escape>", lambda _e: self._cancelar())
        self.protocol("WM_DELETE_WINDOW", self._cancelar)

        # Foco inicial no username + forçar janela pro topo (senão fica
        # atrás de outras janelas do Windows quando abre)
        self.after(50, self._destaque_inicial)
        self._centralizar()

    def _destaque_inicial(self) -> None:
        try:
            self.lift()
            self.attributes("-topmost", True)
            self.focus_force()
            self.entry_user.focus_set()
            # Tira o topmost depois de meio segundo (senão fica sempre em cima)
            self.after(500, lambda: self.attributes("-topmost", False))
        except tk.TclError:
            pass

    def _centralizar(self) -> None:
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _entrar(self) -> None:
        username = self.entry_user.get().strip()
        senha = self.entry_senha.get()
        try:
            self.usuario = auth.login(username, senha)
        except auth.LoginErro as e:
            self.lbl_erro.config(text=str(e))
            self.entry_senha.delete(0, "end")
            self.entry_senha.focus_set()
            return
        except Exception as e:
            self.lbl_erro.config(text=f"Erro ao conectar ao banco: {e}")
            return
        self.destroy()

    def _cancelar(self) -> None:
        self.usuario = None
        self.destroy()


class DialogoMudarSenha(tk.Toplevel):
    """Diálogo pra trocar senha. Se ``exigir_senha_atual=True``, o usuário
    precisa confirmar a senha vigente (uso normal). Se False, admin trocando
    senha de outro user — só pede a nova."""

    def __init__(
        self,
        master: tk.Misc,
        usuario_id: int,
        exigir_senha_atual: bool = True,
        nome_para_titulo: str = "",
    ) -> None:
        super().__init__(master)
        titulo = "Mudar minha senha" if exigir_senha_atual else \
                 f"Mudar senha — {nome_para_titulo}"
        self.title(titulo)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self.usuario_id = usuario_id
        self.exigir_senha_atual = exigir_senha_atual
        self.trocou = False

        frm = ttk.Frame(self, padding=16)
        frm.grid(row=0, column=0)

        row = 0
        if exigir_senha_atual:
            ttk.Label(frm, text="Senha atual:").grid(
                row=row, column=0, sticky="e", padx=(0, 6), pady=4,
            )
            self.entry_atual = ttk.Entry(frm, width=28, show="•")
            self.entry_atual.grid(row=row, column=1, pady=4, sticky="w")
            row += 1

        ttk.Label(frm, text="Nova senha:").grid(
            row=row, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_nova = ttk.Entry(frm, width=28, show="•")
        self.entry_nova.grid(row=row, column=1, pady=4, sticky="w")
        row += 1

        ttk.Label(frm, text="Confirmar:").grid(
            row=row, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_conf = ttk.Entry(frm, width=28, show="•")
        self.entry_conf.grid(row=row, column=1, pady=4, sticky="w")
        row += 1

        self.lbl_erro = ttk.Label(frm, text="", foreground="#c0392b")
        self.lbl_erro.grid(row=row, column=0, columnspan=2, pady=(6, 2))
        row += 1

        botoes = ttk.Frame(frm)
        botoes.grid(row=row, column=0, columnspan=2, pady=(8, 0))
        ttk.Button(botoes, text="Salvar", command=self._salvar).pack(
            side="left", padx=4,
        )
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(
            side="left", padx=4,
        )

        self.bind("<Return>", lambda _e: self._salvar())
        self.bind("<Escape>", lambda _e: self.destroy())

    def _salvar(self) -> None:
        nova = self.entry_nova.get()
        conf = self.entry_conf.get()
        if nova != conf:
            self.lbl_erro.config(text="As senhas não conferem.")
            return
        if self.exigir_senha_atual:
            # Valida senha atual: refaz o login com username do próprio usuário
            usuario = auth.buscar_usuario(self.usuario_id)
            if not usuario:
                self.lbl_erro.config(text="Usuário não encontrado.")
                return
            try:
                auth.login(usuario["username"], self.entry_atual.get())
            except auth.LoginErro:
                self.lbl_erro.config(text="Senha atual incorreta.")
                return
        try:
            auth.mudar_senha(self.usuario_id, nova)
        except auth.UsuarioErro as e:
            self.lbl_erro.config(text=str(e))
            return
        self.trocou = True
        messagebox.showinfo("Sucesso", "Senha alterada.", parent=self)
        self.destroy()


class DialogoCadastroUsuario(tk.Toplevel):
    """Cria novo usuário OU edita existente (menos senha). Só admin usa."""

    def __init__(
        self,
        master: tk.Misc,
        usuario: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(master)
        self.title("Novo usuário" if usuario is None else f"Editar — {usuario['username']}")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        self.editar_id = usuario["id"] if usuario else None
        self.salvou = False

        frm = ttk.Frame(self, padding=16)
        frm.grid(row=0, column=0)

        ttk.Label(frm, text="Username:").grid(
            row=0, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_username = ttk.Entry(frm, width=30)
        self.entry_username.grid(row=0, column=1, pady=4, sticky="w")
        if usuario:
            self.entry_username.insert(0, usuario["username"])
            self.entry_username.config(state="readonly")

        ttk.Label(frm, text="Nome completo:").grid(
            row=1, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_nome = ttk.Entry(frm, width=30)
        self.entry_nome.grid(row=1, column=1, pady=4, sticky="w")
        if usuario:
            self.entry_nome.insert(0, usuario.get("nome", "") or "")

        ttk.Label(frm, text="Email:").grid(
            row=2, column=0, sticky="e", padx=(0, 6), pady=4,
        )
        self.entry_email = ttk.Entry(frm, width=30)
        self.entry_email.grid(row=2, column=1, pady=4, sticky="w")
        if usuario:
            self.entry_email.insert(0, usuario.get("email", "") or "")

        # Senha só ao criar novo. Admin pode redefinir depois pelo botão
        # "Trocar senha" na tela de gerenciar usuários.
        if usuario is None:
            ttk.Label(frm, text="Senha inicial:").grid(
                row=3, column=0, sticky="e", padx=(0, 6), pady=4,
            )
            self.entry_senha = ttk.Entry(frm, width=30, show="•")
            self.entry_senha.grid(row=3, column=1, pady=4, sticky="w")

            ttk.Label(frm, text="Confirmar:").grid(
                row=4, column=0, sticky="e", padx=(0, 6), pady=4,
            )
            self.entry_senha2 = ttk.Entry(frm, width=30, show="•")
            self.entry_senha2.grid(row=4, column=1, pady=4, sticky="w")
        else:
            self.entry_senha = None
            self.entry_senha2 = None

        self.var_admin = tk.BooleanVar(value=bool(usuario and usuario.get("admin")))
        ttk.Checkbutton(
            frm, text="Perfil administrador",
            variable=self.var_admin,
        ).grid(row=5, column=1, sticky="w", pady=4)

        if usuario is not None:
            self.var_ativo = tk.BooleanVar(value=bool(usuario.get("ativo", True)))
            ttk.Checkbutton(
                frm, text="Ativo (pode fazer login)",
                variable=self.var_ativo,
            ).grid(row=6, column=1, sticky="w", pady=4)
        else:
            self.var_ativo = None

        self.lbl_erro = ttk.Label(frm, text="", foreground="#c0392b")
        self.lbl_erro.grid(row=7, column=0, columnspan=2, pady=(6, 2))

        botoes = ttk.Frame(frm)
        botoes.grid(row=8, column=0, columnspan=2, pady=(8, 0))
        ttk.Button(botoes, text="Salvar", command=self._salvar).pack(
            side="left", padx=4,
        )
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(
            side="left", padx=4,
        )

        self.bind("<Return>", lambda _e: self._salvar())
        self.bind("<Escape>", lambda _e: self.destroy())

    def _salvar(self) -> None:
        nome = self.entry_nome.get().strip()
        email = self.entry_email.get().strip()
        admin = bool(self.var_admin.get())

        try:
            if self.editar_id is None:
                username = self.entry_username.get().strip()
                senha = self.entry_senha.get()
                senha2 = self.entry_senha2.get()
                if senha != senha2:
                    self.lbl_erro.config(text="As senhas não conferem.")
                    return
                auth.criar_usuario(
                    username, senha, nome=nome, email=email, admin=admin,
                )
            else:
                auth.atualizar_usuario(
                    self.editar_id,
                    nome=nome, email=email, admin=admin,
                    ativo=bool(self.var_ativo.get()),
                )
        except auth.UsuarioErro as e:
            self.lbl_erro.config(text=str(e))
            return
        except Exception as e:
            self.lbl_erro.config(text=f"Erro: {e}")
            return

        self.salvou = True
        self.destroy()


class DialogoGerenciarUsuarios(tk.Toplevel):
    """Lista + botões pra criar/editar/desativar/trocar senha. Só admin."""

    def __init__(self, master: tk.Misc, usuario_logado: dict[str, Any]) -> None:
        super().__init__(master)
        self.title("Gerenciar usuários")
        self.geometry("720x420")
        self.transient(master)
        self.grab_set()

        self.usuario_logado = usuario_logado

        info = ttk.Label(
            self,
            text=(
                "Cadastre os usuários que podem acessar o app. Cada um tem "
                "seu próprio login e sua própria empresa ativa; regras e "
                "mapeamentos continuam por empresa (compartilhados)."
            ),
            wraplength=680, foreground="#1f3a68",
            font=("TkDefaultFont", 9, "italic"),
        )
        info.pack(side="top", fill="x", padx=6, pady=(6, 4))

        # Rodapé de botões — packado antes pra ficar ancorado embaixo
        rodape = ttk.Frame(self)
        rodape.pack(side="bottom", fill="x", padx=6, pady=(4, 6))
        ttk.Button(rodape, text="+ Novo usuário", command=self._novo).pack(
            side="left", padx=2,
        )
        ttk.Button(rodape, text="Editar selecionado", command=self._editar).pack(
            side="left", padx=2,
        )
        ttk.Button(
            rodape, text="Trocar senha do selecionado",
            command=self._trocar_senha,
        ).pack(side="left", padx=2)
        ttk.Button(
            rodape, text="Ativar/desativar",
            command=self._toggle_ativo,
        ).pack(side="left", padx=2)
        ttk.Button(rodape, text="Fechar", command=self.destroy).pack(
            side="right", padx=2,
        )

        # Tabela
        cols = ("username", "nome", "email", "admin", "ativo", "ultimo_login")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for c, t, w, a in [
            ("username", "Username", 130, "w"),
            ("nome", "Nome", 200, "w"),
            ("email", "Email", 180, "w"),
            ("admin", "Perfil", 70, "center"),
            ("ativo", "Ativo", 60, "center"),
            ("ultimo_login", "Último login", 130, "center"),
        ]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 4))
        sb.pack(side="right", fill="y", padx=(0, 6), pady=(0, 4))

        self.tree.tag_configure("inativo", background="#f8d7da")

        self._itens: dict[str, dict[str, Any]] = {}
        self._recarregar()

        self.tree.bind("<Double-1>", lambda _e: self._editar())
        self.bind("<Escape>", lambda _e: self.destroy())

    def _recarregar(self) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._itens.clear()
        for u in auth.listar_usuarios(incluir_inativos=True):
            ult = u["ultimo_login"]
            ult_txt = ult.strftime("%d/%m/%Y %H:%M") if ult else "—"
            iid = self.tree.insert(
                "", "end",
                values=(
                    u["username"],
                    u["nome"],
                    u["email"],
                    "admin" if u["admin"] else "operador",
                    "sim" if u["ativo"] else "não",
                    ult_txt,
                ),
                tags=() if u["ativo"] else ("inativo",),
            )
            self._itens[iid] = u

    def _selecionado(self) -> dict[str, Any] | None:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um usuário na lista.", parent=self,
            )
            return None
        return self._itens.get(sel[0])

    def _novo(self) -> None:
        dlg = DialogoCadastroUsuario(self)
        self.wait_window(dlg)
        if dlg.salvou:
            self._recarregar()

    def _editar(self) -> None:
        u = self._selecionado()
        if not u:
            return
        dlg = DialogoCadastroUsuario(self, usuario=u)
        self.wait_window(dlg)
        if dlg.salvou:
            self._recarregar()

    def _trocar_senha(self) -> None:
        u = self._selecionado()
        if not u:
            return
        eh_ele_mesmo = u["id"] == self.usuario_logado["id"]
        dlg = DialogoMudarSenha(
            self, usuario_id=u["id"],
            exigir_senha_atual=eh_ele_mesmo,
            nome_para_titulo=u["username"],
        )
        self.wait_window(dlg)

    def _toggle_ativo(self) -> None:
        u = self._selecionado()
        if not u:
            return
        if u["id"] == self.usuario_logado["id"] and u["ativo"]:
            messagebox.showwarning(
                "Ação inválida",
                "Você não pode desativar o próprio usuário logado.",
                parent=self,
            )
            return
        novo_estado = not u["ativo"]
        acao = "ativar" if novo_estado else "desativar"
        if not messagebox.askyesno(
            "Confirmar",
            f"Deseja {acao} o usuário '{u['username']}'?",
            parent=self,
        ):
            return
        auth.atualizar_usuario(u["id"], ativo=novo_estado)
        self._recarregar()
