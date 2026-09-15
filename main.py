import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import config
import parser_dominio
from dialogos_dominio import DialogoConexao, DialogoFonte, DialogoSelecionarEmpresa
from dialogos_taxas import (
    DialogoConfigurarTaxas,
    DialogoEditarLancamento,
    DialogoEditarPar,
    DialogoEditarTransacao,
    DialogoLancamentoManual,
    DialogoLancamentoManualAvulso,
    DialogoNovaRegra,
)
from exportar_xlsx import (
    exportar_conciliados_dominio,
    exportar_lancamentos_contabeis,
    exportar_pendencias_comparacao,
    exportar_pendentes,
)
from lancamentos import LancamentoContabil, gerar_lancamentos_contabeis
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
    ("data", "Data vencimento"),
    ("data_pagamento", "Data pagamento"),
    ("data_emissao", "Data emissão"),
    ("valor", "Valor"),
    ("numero_nf", "Nº NF"),
    ("cnpj", "CNPJ fornecedor"),
    ("fornecedor", "Fornecedor"),
    ("historico", "Histórico"),
    ("tipo", "Tipo"),
]
# Apenas data (vencimento) e valor são chave de match — os demais podem
# ser deixados em branco se a planilha não tiver a coluna.
CAMPOS_OPCIONAIS = {
    "data_pagamento", "data_emissao",
    "numero_nf", "cnpj", "fornecedor", "historico", "tipo",
}


class DialogoFiltroColuna(tk.Toplevel):
    """Dropdown estilo Excel ao clicar no cabeçalho de uma coluna.

    Mostra checkbox por valor único da coluna + busca interna + marcar/
    desmarcar tudo. Devolve em ``self.resultado`` o conjunto de valores
    selecionados, ou ``None`` se TUDO está marcado (= sem filtro).
    ``self.cancelado`` indica que o usuário desistiu.
    """

    def __init__(
        self,
        master: tk.Misc,
        titulo: str,
        valores_unicos: list[str],
        selecionados: set[str] | None,
        x: int | None = None,
        y: int | None = None,
    ) -> None:
        super().__init__(master)
        self.title(titulo)
        self.transient(master)
        self.grab_set()
        self.geometry(f"320x420{'+' + str(x) if x else ''}{'+' + str(y) if y else ''}")

        self.valores_unicos = sorted(valores_unicos, key=lambda v: (v == "", v))
        self.resultado: set[str] | None = None
        self.cancelado = False
        self.checkbuttons: dict[str, tuple[tk.BooleanVar, ttk.Checkbutton]] = {}

        # Topo: marcar/desmarcar tudo
        topo = ttk.Frame(self)
        topo.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Button(topo, text="Marcar tudo", command=self._marcar_tudo).pack(side="left", padx=2)
        ttk.Button(topo, text="Desmarcar tudo", command=self._desmarcar_tudo).pack(side="left", padx=2)

        # Busca dentro do filtro
        busca_frame = ttk.Frame(self)
        busca_frame.pack(fill="x", padx=8, pady=2)
        ttk.Label(busca_frame, text="Buscar:").pack(side="left")
        self.busca_var = tk.StringVar()
        self.busca_var.trace_add("write", lambda *_a: self._aplica_busca())
        ttk.Entry(busca_frame, textvariable=self.busca_var).pack(
            side="left", fill="x", expand=True, padx=4,
        )

        # Lista scrollable com checkboxes
        lista_frame = ttk.Frame(self)
        lista_frame.pack(fill="both", expand=True, padx=8, pady=4)
        canvas = tk.Canvas(lista_frame, highlightthickness=0)
        sb = ttk.Scrollbar(lista_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        # Permite scroll do mouse
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self._canvas = canvas

        marcados_iniciais = (
            set(selecionados) if selecionados is not None
            else set(self.valores_unicos)
        )
        for valor in self.valores_unicos:
            var = tk.BooleanVar(value=(valor in marcados_iniciais))
            label = str(valor) if valor != "" else "(em branco)"
            cb = ttk.Checkbutton(self.inner, text=label, variable=var)
            cb.pack(anchor="w", padx=4, pady=1)
            self.checkbuttons[valor] = (var, cb)

        # Botões finais
        botoes = ttk.Frame(self)
        botoes.pack(fill="x", padx=8, pady=8)
        ttk.Button(botoes, text="Cancelar", command=self._cancelar).pack(side="right", padx=2)
        ttk.Button(botoes, text="OK", command=self._confirmar).pack(side="right", padx=2)

        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self.bind("<Return>", lambda _e: self._confirmar())
        self.bind("<Escape>", lambda _e: self._cancelar())

    def _aplica_busca(self) -> None:
        termo = self.busca_var.get().strip().lower()
        for valor, (_var, cb) in self.checkbuttons.items():
            mostra = (not termo) or (termo in str(valor).lower())
            if mostra:
                cb.pack(anchor="w", padx=4, pady=1)
            else:
                cb.pack_forget()

    def _marcar_tudo(self) -> None:
        # Marca apenas o que está visível (respeita busca)
        termo = self.busca_var.get().strip().lower()
        for valor, (var, _cb) in self.checkbuttons.items():
            if (not termo) or (termo in str(valor).lower()):
                var.set(True)

    def _desmarcar_tudo(self) -> None:
        termo = self.busca_var.get().strip().lower()
        for valor, (var, _cb) in self.checkbuttons.items():
            if (not termo) or (termo in str(valor).lower()):
                var.set(False)

    def _confirmar(self) -> None:
        marcados = {v for v, (var, _) in self.checkbuttons.items() if var.get()}
        # Se TUDO marcado → sem filtro (None)
        if marcados == set(self.valores_unicos):
            self.resultado = None
        else:
            self.resultado = marcados
        # Desfaz binding global do MouseWheel
        try:
            self._canvas.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        self.destroy()

    def _cancelar(self) -> None:
        self.cancelado = True
        try:
            self._canvas.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        self.destroy()


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
                "Escolha qual coluna corresponde a cada campo. Campos opcionais "
                "podem ficar como '(deixar vazia)' se a planilha não tiver "
                "essa coluna. Apenas Vencimento e Valor são obrigatórios."
            ),
            wraplength=720,
        )
        info.grid(row=0, column=0, columnspan=2, padx=12, pady=(12, 8), sticky="w")

        # Primeira opção do dropdown = "deixar vazia". As demais são as
        # colunas da planilha. Como o índice 0 é a opção vazia, na hora
        # de resolver o idx real fazemos current() - 1.
        self.SENTINELA_VAZIO = "(deixar vazia)"
        self.opcoes = [self.SENTINELA_VAZIO] + [
            f"{idx + 1}. {nome if nome else '(sem nome)'}"
            for idx, nome in enumerate(estrutura.cabecalho)
        ]
        self.combos: dict[str, ttk.Combobox] = {}

        for i, (campo, rotulo) in enumerate(CAMPOS, start=1):
            sufixo = " (opcional)" if campo in CAMPOS_OPCIONAIS else ""
            cor = "#555" if campo in CAMPOS_OPCIONAIS else "black"
            ttk.Label(self, text=f"{rotulo}{sufixo}:", foreground=cor).grid(
                row=i, column=0, padx=12, pady=4, sticky="e",
            )
            cb = ttk.Combobox(self, values=self.opcoes, state="readonly", width=48)
            cb.grid(row=i, column=1, padx=(0, 12), pady=4, sticky="w")
            idx = estrutura.sugestao.get(campo)
            if idx is not None and 0 <= idx < len(estrutura.cabecalho):
                cb.current(idx + 1)  # +1 por causa do sentinela "(deixar vazia)"
            else:
                cb.current(0)  # padrão = vazia (será obrigado a escolher se obrigatório)
            cb.bind("<<ComboboxSelected>>", lambda _e: self._atualiza_preview())
            self.combos[campo] = cb

        row_preview = len(CAMPOS) + 1
        ttk.Label(self, text="Preview das primeiras linhas:").grid(
            row=row_preview, column=0, columnspan=2, padx=12, pady=(8, 2), sticky="w",
        )

        preview_frame = ttk.Frame(self)
        preview_frame.grid(row=row_preview + 1, column=0, columnspan=2, padx=12, sticky="we")

        cols = (
            "linha", "data", "data_pagamento", "data_emissao",
            "valor", "numero_nf", "cnpj", "fornecedor", "historico", "tipo",
        )
        self.preview = ttk.Treeview(
            preview_frame, columns=cols, show="headings", height=self.PREVIEW_LINHAS,
        )
        self.preview.heading("linha", text="Linha")
        self.preview.heading("data", text="Vencimento")
        self.preview.heading("data_pagamento", text="Pagamento")
        self.preview.heading("data_emissao", text="Emissão")
        self.preview.heading("valor", text="Valor")
        self.preview.heading("numero_nf", text="Nº NF")
        self.preview.heading("cnpj", text="CNPJ")
        self.preview.heading("fornecedor", text="Fornecedor")
        self.preview.heading("historico", text="Histórico")
        self.preview.heading("tipo", text="Tipo")
        self.preview.column("linha", width=45, anchor="center")
        self.preview.column("data", width=80, anchor="w")
        self.preview.column("data_pagamento", width=80, anchor="w")
        self.preview.column("data_emissao", width=80, anchor="w")
        self.preview.column("valor", width=85, anchor="e")
        self.preview.column("numero_nf", width=70, anchor="center")
        self.preview.column("cnpj", width=120, anchor="w")
        self.preview.column("fornecedor", width=180, anchor="w")
        self.preview.column("historico", width=180, anchor="w")
        self.preview.column("tipo", width=110, anchor="w")
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

        # Resolve idx real: current() - 1 (porque opção 0 é "(deixar vazia)").
        # Vazia → -1 (mesmo significado de "não selecionado" do código antigo).
        idxs = {campo: self.combos[campo].current() - 1 for campo, _ in CAMPOS}
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
            cel_pgto = _cel(idxs["data_pagamento"])
            cel_emis = _cel(idxs["data_emissao"])
            cel_valor = _cel(idxs["valor"])
            cel_nf = _cel(idxs["numero_nf"])
            cel_cnpj = _cel(idxs["cnpj"])
            cel_forn = _cel(idxs["fornecedor"])
            cel_hist = _cel(idxs["historico"])
            cel_tipo = _cel(idxs["tipo"])

            data_parsed = para_data(cel_data) if idxs["data"] >= 0 else None
            pgto_parsed = para_data(cel_pgto) if idxs["data_pagamento"] >= 0 else None
            emis_parsed = para_data(cel_emis) if idxs["data_emissao"] >= 0 else None
            valor_parsed = para_decimal(cel_valor) if idxs["valor"] >= 0 else None

            # Tag de erro: alguma das datas obrigatórias ou valor não converteu
            if (
                data_parsed is None or valor_parsed is None
                or pgto_parsed is None or emis_parsed is None
            ):
                tag = "erro"
            else:
                total_validas += 1
                tag = ""

            def _fmt_data_txt(parsed, raw):
                if parsed:
                    return parsed.strftime("%d/%m/%Y")
                if raw in (None, ""):
                    return "—"
                return f"✗ {raw!r}"

            def _fmt_str(raw):
                if raw in (None, ""):
                    return "—"
                return str(raw)

            self.preview.insert(
                "", "end",
                values=(
                    base + offset,
                    _fmt_data_txt(data_parsed, cel_data),
                    _fmt_data_txt(pgto_parsed, cel_pgto),
                    _fmt_data_txt(emis_parsed, cel_emis),
                    (
                        f"{valor_parsed:.2f}" if valor_parsed is not None
                        else (f"✗ {cel_valor!r}" if cel_valor not in (None, "") else "—")
                    ),
                    _fmt_str(cel_nf),
                    _fmt_str(cel_cnpj),
                    _fmt_str(cel_forn),
                    _fmt_str(cel_hist),
                    _fmt_str(cel_tipo),
                ),
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
        for campo, rotulo in CAMPOS:
            # current() = 0 → "(deixar vazia)"; 1..N → coluna (idx = current-1)
            sel_raw = self.combos[campo].current()
            sel = sel_raw - 1
            if sel < 0:
                if campo in CAMPOS_OPCIONAIS:
                    continue  # campo opcional sem mapeamento — segue ok
                messagebox.showwarning(
                    "Mapeamento incompleto",
                    f"Selecione a coluna para '{rotulo}' "
                    "(esse campo é obrigatório).",
                    parent=self,
                )
                return
            mapa[campo] = sel
        if len(set(mapa.values())) < len(mapa):
            messagebox.showwarning(
                "Colunas duplicadas",
                "Cada campo precisa apontar para uma coluna diferente da planilha.",
                parent=self,
            )
            return
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

        # Mostra a janela principal já visível (com placeholder) — assim
        # o Toplevel do login aparece corretamente na frente. Esconder a
        # janela principal (withdraw) fazia o dialog filho ficar invisível
        # em algumas versões do Windows.
        placeholder = ttk.Label(
            self,
            text="Aguardando login...",
            font=("TkDefaultFont", 14),
            foreground="#1f3a68",
        )
        placeholder.pack(expand=True, padx=40, pady=40)
        self.update()  # força renderização antes de abrir o diálogo

        # Usuário logado — populado pelo login. Se None ao final do
        # _pedir_login, encerra o app.
        self.usuario_atual: dict | None = None
        if not self._pedir_login():
            self.destroy()
            return
        # Registra no config pra que ele saiba de quem é a empresa_ativa
        config.set_usuario_atual(self.usuario_atual["id"])

        # Remove o placeholder — a UI real será montada em _monta_ui()
        placeholder.destroy()

        # Dados originais
        self.transacoes_planilha: list[Transacao] = []
        self.transacoes_ofx: list[Transacao] = []
        self.caminho_planilha: Path | None = None
        self.caminhos_ofx: list[Path] = []
        self.estrutura_planilha: EstruturaPlanilha | None = None
        self.mapeamento_planilha: dict[str, int] | None = None

        # Estado de conciliação
        self.pares_conciliados: list[Par] = []          # auto + manual + sugestões aceitas
        self.pendentes_planilha: list[Transacao] = []
        # Mesma lógica de "brutos" da OFX: pendentes da planilha sem desconto
        # dos que viraram lançamento contábil (regra fornecedor ou manual).
        self.pendentes_planilha_brutos: list[Transacao] = []
        # Match no Domínio dos pendentes da planilha (sem OFX = Caixa geral):
        # id(t_planilha) → {dominio: Transacao | None, diff_dias, diff_valor}
        self.pendentes_planilha_dominio: dict[int, dict] = {}
        # IDs de transacao_origem cuja regra automática deve ser IGNORADA
        # (usuário excluiu/editou o lançamento contábil).
        self.lancamentos_ignorados: set[int] = set()
        # Match no Domínio dos pendentes do OFX (comparação OFX×Domínio sem
        # planilha): id(t_ofx) → {dominio, diff_dias, diff_valor}
        self.pendentes_ofx_dominio: dict[int, dict] = {}
        self.pendentes_ofx: list[Transacao] = []
        # "brutos": pendentes OFX sem desconto dos que viraram lançamentos
        # contábeis. self.pendentes_ofx (visível) = brutos - classificados.
        self.pendentes_ofx_brutos: list[Transacao] = []
        self.sugestoes: list[Par] = []
        self.itens_pares: dict[str, Par] = {}            # iid → Par (aba conciliados)
        self.itens_pendentes_p: dict[str, Transacao] = {}
        self.itens_pendentes_o: dict[str, Transacao] = {}
        self.itens_sugestoes: dict[str, Par] = {}

        # Domínio
        self.conn_dominio = None
        self.transacoes_dominio: list[Transacao] = []
        self.plano_contas: list[parser_dominio.ContaContabil] = []
        self.comparacao_dominio: list[tuple[str, Par | Transacao]] = []
        # ↑ status, registro (Par para casados; Transacao para faltantes)
        self.lancamentos_contabeis: list[LancamentoContabil] = []
        # Lançamentos manuais (avulsos): persistem entre re-cálculos automáticos
        self.lancamentos_manuais: list[LancamentoContabil] = []
        self.ids_pares_classificados: set[int] = set()

        self.cfg = config.carregar()
        self._migrar_config_legado()

        self._monta_ui()

        # Conecta o Domínio automaticamente usando as credenciais salvas
        # (após a UI estar montada, senão os botões que ele habilita ainda
        # não existem). Silencioso se falhar — o usuário pode conectar manualmente.
        self.after(100, self._auto_conectar_dominio)

    # ------------------------------------------------------ Login

    def _pedir_login(self) -> bool:
        """Mostra o diálogo de login. Retorna True se autenticou, False
        se cancelou/fechou. Se o banco não tem nenhum usuário cadastrado,
        avisa e retorna False (usuário precisa rodar setup_db.py)."""
        from dialogos_login import DialogoLogin
        import auth

        try:
            tem_usuario = auth.existe_algum_usuario()
        except Exception as e:
            messagebox.showerror(
                "Erro no banco",
                f"Não consegui conectar ao banco de dados:\n\n{e}\n\n"
                "Verifique se o MariaDB está rodando e se o data/db_config.json "
                "está correto. Se ainda não configurou, rode:\n"
                "    python setup_db.py",
            )
            return False

        if not tem_usuario:
            messagebox.showerror(
                "Sem usuários cadastrados",
                "O banco ainda não tem nenhum usuário do app. "
                "Rode primeiro:\n\n    python setup_db.py\n\n"
                "e cadastre o admin.",
            )
            return False

        dlg = DialogoLogin(self)
        self.wait_window(dlg)
        if dlg.usuario is None:
            return False
        self.usuario_atual = dlg.usuario
        return True

    def _trocar_usuario(self) -> None:
        """Fecha o app atual — usuário reabre e loga com outra conta."""
        if not messagebox.askyesno(
            "Trocar usuário",
            "Isso vai fechar o aplicativo. Você precisa reabri-lo pra "
            "logar com outra conta.\n\nQuaisquer conciliações não "
            "exportadas serão perdidas. Continuar?",
        ):
            return
        self.destroy()

    def _gerenciar_usuarios(self) -> None:
        if not self.usuario_atual.get("admin"):
            messagebox.showwarning(
                "Permissão negada",
                "Apenas usuários com perfil admin podem gerenciar usuários.",
            )
            return
        from dialogos_login import DialogoGerenciarUsuarios
        dlg = DialogoGerenciarUsuarios(self, self.usuario_atual)
        self.wait_window(dlg)

    def _mudar_minha_senha(self) -> None:
        from dialogos_login import DialogoMudarSenha
        dlg = DialogoMudarSenha(
            self, usuario_id=self.usuario_atual["id"],
            exigir_senha_atual=True,
        )
        self.wait_window(dlg)

    def _migrar_config_legado(self) -> None:
        """Migrações de formato de config:
        1) cfg["dominio"] (credenciais misturadas) → data/dominio_config.json
           + cfg["dominio_fonte"].
        2) cfg["regras_taxas"] (lista global) → cfg["regras_taxas_por_empresa"]
           vinculada à empresa atualmente selecionada (se houver)."""
        precisa_salvar = False

        legado = self.cfg.get("dominio")
        if isinstance(legado, dict):
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
            precisa_salvar = True

        # Migração 2: regras_taxas globais → por empresa atual
        if "regras_taxas" in self.cfg:
            regras_legacy = self.cfg.pop("regras_taxas", [])
            emp = self.cfg.get("dominio_empresa") or {}
            codi = emp.get("codi_emp")
            if regras_legacy and codi is not None:
                por_emp = self.cfg.setdefault("regras_taxas_por_empresa", {})
                por_emp[str(codi)] = regras_legacy
            precisa_salvar = True

        # Migração 3: dominio_fonte (singular) → dominio_fonte_pagamentos
        if "dominio_fonte" in self.cfg and "dominio_fonte_pagamentos" not in self.cfg:
            self.cfg["dominio_fonte_pagamentos"] = self.cfg.pop("dominio_fonte")
            precisa_salvar = True
        elif "dominio_fonte" in self.cfg:
            del self.cfg["dominio_fonte"]  # já tem o novo
            precisa_salvar = True

        if precisa_salvar:
            config.salvar(self.cfg)

    # ------------------------------------------------------------------ UI

    def _monta_ui(self) -> None:
        # --- Linha 0: Usuário logado (barra fina no topo) ---
        topo_user = ttk.Frame(self, padding=(10, 6, 10, 2))
        topo_user.pack(fill="x")
        eh_admin = bool(self.usuario_atual.get("admin"))
        nome = self.usuario_atual.get("nome") or self.usuario_atual["username"]
        perfil = "admin" if eh_admin else "operador"
        ttk.Label(
            topo_user,
            text=f"👤 {nome} ({self.usuario_atual['username']}) — {perfil}",
            foreground="#1f3a68",
            font=("TkDefaultFont", 9, "bold"),
        ).pack(side="left")
        ttk.Button(
            topo_user, text="Trocar usuário",
            command=self._trocar_usuario,
        ).pack(side="right", padx=2)
        ttk.Button(
            topo_user, text="Minha senha",
            command=self._mudar_minha_senha,
        ).pack(side="right", padx=2)
        # "Gerenciar usuarios" so aparece pra admin (operadores nao precisam)
        self.btn_gerenciar_usuarios = ttk.Button(
            topo_user, text="Gerenciar usuários",
            command=self._gerenciar_usuarios,
        )
        if eh_admin:
            self.btn_gerenciar_usuarios.pack(side="right", padx=2)

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # --- Linha 1: Domínio (sistema contábil) ---
        topo_dom = ttk.Frame(self, padding=(10, 6, 10, 4))
        topo_dom.pack(fill="x")
        ttk.Button(topo_dom, text="Conectar Domínio", command=self._conectar_dominio).pack(side="left", padx=4)
        self.btn_empresa = ttk.Button(
            topo_dom, text="Selecionar empresa", command=self._selecionar_empresa, state="disabled",
        )
        self.btn_empresa.pack(side="left", padx=4)
        # Botões de FONTE (SQL Domínio) são configuração do sistema —
        # só admin vê. Operadores usam a fonte que o admin já configurou.
        self.btn_fonte = ttk.Button(
            topo_dom, text="Fonte: pagamentos", command=self._configurar_fonte_dominio, state="disabled",
        )
        self.btn_fonte_plano = ttk.Button(
            topo_dom, text="Fonte: plano contas",
            command=self._configurar_fonte_plano_contas, state="disabled",
        )
        if eh_admin:
            self.btn_fonte.pack(side="left", padx=4)
            self.btn_fonte_plano.pack(side="left", padx=4)
        self.btn_carregar_dominio = ttk.Button(
            topo_dom, text="Carregar pagamentos", command=self._carregar_dominio, state="disabled",
        )
        self.btn_carregar_dominio.pack(side="left", padx=4)
        self.btn_carregar_plano = ttk.Button(
            topo_dom, text="Carregar plano contas",
            command=self._carregar_plano_contas, state="disabled",
        )
        self.btn_carregar_plano.pack(side="left", padx=4)
        self.lbl_dominio = ttk.Label(topo_dom, text="(Domínio não conectado)")
        self.lbl_dominio.pack(side="left", padx=8)

        # --- Linha 2: Planilha ---
        topo = ttk.Frame(self, padding=(10, 0, 10, 4))
        topo.pack(fill="x")
        ttk.Button(topo, text="Abrir planilha (.xlsx)", command=self._abrir_planilha).pack(side="left", padx=4)
        ttk.Button(
            topo, text="Importar comprovantes PDF",
            command=self._importar_comprovantes_pdf,
        ).pack(side="left", padx=4)
        self.btn_editar_colunas = ttk.Button(
            topo, text="Editar colunas", command=self._editar_colunas, state="disabled",
        )
        self.btn_editar_colunas.pack(side="left", padx=4)
        self.btn_limpar_planilha = ttk.Button(
            topo, text="Limpar planilha", command=self._limpar_planilha, state="disabled",
        )
        self.btn_limpar_planilha.pack(side="left", padx=4)
        self.lbl_planilha = ttk.Label(topo, text="(nenhuma planilha carregada)")
        self.lbl_planilha.pack(side="left", padx=8)

        # --- Linha 3: OFX ---
        topo2 = ttk.Frame(self, padding=(10, 0, 10, 6))
        topo2.pack(fill="x")
        ttk.Button(topo2, text="Importar OFX", command=self._abrir_ofx).pack(side="left", padx=4)
        self.btn_limpar_ofx = ttk.Button(
            topo2, text="Limpar OFX", command=self._limpar_ofx, state="disabled",
        )
        self.btn_limpar_ofx.pack(side="left", padx=4)
        self.lbl_ofx = ttk.Label(topo2, text="(nenhum OFX carregado)")
        self.lbl_ofx.pack(side="left", padx=8)

        acoes = ttk.Frame(self, padding=(10, 0, 10, 8))
        acoes.pack(fill="x")
        self.btn_conciliar = ttk.Button(
            acoes, text="Conciliar", command=self._executar_conciliacao, state="disabled",
        )
        self.btn_conciliar.pack(side="left", padx=4)
        self.btn_comparar_dominio = ttk.Button(
            acoes, text="Comparar com Domínio",
            command=self._comparar_com_dominio,
        )
        self.btn_comparar_dominio.pack(side="left", padx=4)
        ttk.Button(
            acoes, text="Configurar taxas", command=self._abrir_config_taxas,
        ).pack(side="left", padx=4)
        self.lbl_resumo = ttk.Label(acoes, text="")
        self.lbl_resumo.pack(side="left", padx=12)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Abas de dados crus (origem) — vêm primeiro no fluxo de leitura
        self._monta_aba_planilha_dados()
        self._monta_aba_ofx_dados()
        self._monta_aba_dominio_dados()
        # Abas de conciliação (resultado)
        self._monta_aba_conciliados()
        self._monta_aba_pendentes()
        self._monta_aba_sugestoes()
        self._monta_aba_conciliados_dominio()
        self._monta_aba_dominio()
        self._monta_aba_lancamentos()
        self._monta_aba_plano_contas()

    # --------------- Filtros estilo Excel (popup ao clicar no cabeçalho) ---

    def _label_coluna_filtro(self, base: str, ativo: bool) -> str:
        """Adiciona indicador visual no cabeçalho conforme estado do filtro."""
        return f"{base}  ▼ ★" if ativo else f"{base}  ▾"

    def _abrir_filtro_excel(
        self,
        tree: ttk.Treeview,
        col_ids: tuple[str, ...],
        col_labels: dict[str, str],
        filtros: dict[str, set[str] | None],
        rows_iter,
        event: tk.Event,
        on_apply,
    ) -> None:
        """Detecta clique no cabeçalho e abre o DialogoFiltroColuna correspondente.

        - ``tree``: Treeview da aba
        - ``col_ids``: tupla com os IDs das colunas (ordem)
        - ``col_labels``: dict id → rótulo amigável
        - ``filtros``: estado atual {col_id: set ou None}
        - ``rows_iter``: callable que devolve as tuplas (uma por transação) na
          ordem original (sem filtro aplicado)
        - ``event``: evento de clique
        - ``on_apply``: callable chamado após aplicar filtro
        """
        region = tree.identify_region(event.x, event.y)
        if region != "heading":
            return
        col_str = tree.identify_column(event.x)
        if not col_str or not col_str.startswith("#"):
            return
        idx = int(col_str.lstrip("#")) - 1
        if idx < 0 or idx >= len(col_ids):
            return
        col_id = col_ids[idx]

        valores_unicos = sorted({str(row[idx]) for row in rows_iter()})
        if not valores_unicos:
            return

        x = event.x_root
        y = event.y_root + 15
        dlg = DialogoFiltroColuna(
            self,
            titulo=f"Filtrar: {col_labels[col_id]}",
            valores_unicos=valores_unicos,
            selecionados=filtros.get(col_id),
            x=x, y=y,
        )
        self.wait_window(dlg)
        if dlg.cancelado:
            return
        filtros[col_id] = dlg.resultado  # None = sem filtro, set = filtro ativo
        on_apply()

    def _monta_aba_planilha_dados(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Planilha (0)")
        self._aba_planilha = aba

        # Barra de filtro
        topo = ttk.Frame(aba)
        topo.pack(side="top", fill="x", padx=4, pady=4)
        ttk.Label(topo, text="Buscar:").pack(side="left", padx=(0, 4))
        self.filtro_planilha = tk.StringVar()
        self.filtro_planilha.trace_add("write", lambda *_a: self._render_aba_planilha())
        ttk.Entry(topo, textvariable=self.filtro_planilha, width=40).pack(side="left")
        ttk.Button(topo, text="Limpar", command=self._limpa_filtros_planilha).pack(
            side="left", padx=4,
        )
        ttk.Button(
            topo, text="Editar lançamento selecionado",
            command=self._editar_lancamento_planilha,
        ).pack(side="left", padx=4)
        self.lbl_filtro_planilha = ttk.Label(topo, text="", foreground="#666")
        self.lbl_filtro_planilha.pack(side="left", padx=8)

        # Estado dos filtros por coluna (estilo Excel)
        self.filtros_col_planilha: dict[str, set[str] | None] = {}
        # iid → Transacao (para resolver seleção do botão Editar)
        self.itens_tree_planilha: dict[str, Transacao] = {}

        # Treeview + scrollbar
        corpo = ttk.Frame(aba)
        corpo.pack(side="top", fill="both", expand=True)
        cols = ("linha", "venc", "pagto", "emis", "valor", "nf", "cnpj", "fornecedor", "historico", "tipo")
        tree = ttk.Treeview(corpo, columns=cols, show="headings")
        for c, t, w, a in [
            ("linha", "Linha", 55, "center"),
            ("venc", "Vencimento", 100, "center"),
            ("pagto", "Pagamento", 100, "center"),
            ("emis", "Emissão", 100, "center"),
            ("valor", "Valor", 105, "e"),
            ("nf", "Nº NF", 85, "center"),
            ("cnpj", "CNPJ", 130, "w"),
            ("fornecedor", "Fornecedor", 220, "w"),
            ("historico", "Histórico", 220, "w"),
            ("tipo", "Tipo", 130, "w"),
        ]:
            tree.heading(c, text=self._label_coluna_filtro(t, False))
            tree.column(c, width=w, anchor=a)
        sb = ttk.Scrollbar(corpo, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree_planilha = tree
        tree.bind("<Button-1>", self._on_click_header_planilha)

    def _row_planilha(self, t) -> tuple:
        return (
            str(t.linha) if t.linha is not None else "",
            t.data.strftime("%d/%m/%Y"),
            self._fmt_data(t.data_pagamento),
            self._fmt_data(t.extras.get("data_emissao")),
            f"{t.valor:.2f}",
            t.extras.get("numero_nf", "") or "",
            t.extras.get("cnpj", "") or "",
            t.extras.get("fornecedor", "") or "",
            t.extras.get("historico", "") or "",
            t.extras.get("tipo", "") or "",
        )

    COLS_PLANILHA = (
        "linha", "venc", "pagto", "emis", "valor",
        "nf", "cnpj", "fornecedor", "historico", "tipo",
    )
    LABELS_PLANILHA = {
        "linha": "Linha", "venc": "Vencimento", "pagto": "Pagamento",
        "emis": "Emissão", "valor": "Valor", "nf": "Nº NF",
        "cnpj": "CNPJ", "fornecedor": "Fornecedor",
        "historico": "Histórico", "tipo": "Tipo",
    }

    def _on_click_header_planilha(self, event: tk.Event) -> None:
        self._abrir_filtro_excel(
            tree=self.tree_planilha,
            col_ids=self.COLS_PLANILHA,
            col_labels=self.LABELS_PLANILHA,
            filtros=self.filtros_col_planilha,
            rows_iter=lambda: (self._row_planilha(t) for t in self.transacoes_planilha),
            event=event,
            on_apply=lambda: (self._atualiza_headers_planilha(), self._render_aba_planilha()),
        )

    def _atualiza_headers_planilha(self) -> None:
        for col in self.COLS_PLANILHA:
            ativo = self.filtros_col_planilha.get(col) is not None
            self.tree_planilha.heading(
                col, text=self._label_coluna_filtro(self.LABELS_PLANILHA[col], ativo),
            )

    def _limpa_filtros_planilha(self) -> None:
        self.filtro_planilha.set("")
        self.filtros_col_planilha.clear()
        self._atualiza_headers_planilha()

    def _editar_lancamento_planilha(self) -> None:
        """Edita o lançamento da planilha selecionado. Como mudar dados de
        entrada invalida pares e pendentes já calculados, limpamos os
        resultados — o usuário re-executa a conciliação depois."""
        sel = self.tree_planilha.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione uma linha na aba Planilha para editar.",
            )
            return
        t = self.itens_tree_planilha.get(sel[0])
        if t is None:
            return

        dlg = DialogoEditarTransacao(self, t)
        self.wait_window(dlg)
        if not dlg.resultado:
            return
        r = dlg.resultado

        # Aplica no objeto Transacao (in-place — referências em
        # transacoes_planilha continuam apontando pro mesmo objeto)
        t.data = r["data"]
        t.data_pagamento = r["data_pagamento"]
        t.valor = r["valor"]
        # extras: setar/remover conforme valor preenchido
        if r["data_emissao"]:
            t.extras["data_emissao"] = r["data_emissao"]
        elif "data_emissao" in t.extras:
            del t.extras["data_emissao"]
        # data_pagamento também espelhado em extras (compat com outros lugares)
        if r["data_pagamento"]:
            t.extras["data_pagamento"] = r["data_pagamento"]
        elif "data_pagamento" in t.extras:
            del t.extras["data_pagamento"]
        for k in ("numero_nf", "cnpj", "fornecedor", "historico"):
            if r[k]:
                t.extras[k] = r[k]
            elif k in t.extras:
                del t.extras[k]

        self._render_aba_planilha()
        # Invalida conciliação/resultados — usuário precisa rodar de novo
        self._limpa_resultados()
        messagebox.showinfo(
            "Lançamento atualizado",
            "Dados salvos. Os resultados anteriores de conciliação foram "
            "limpos — clique em 'Conciliar' para refazer com os novos dados.",
        )

    def _monta_aba_ofx_dados(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="OFX (0)")
        self._aba_ofx = aba

        # Barra de filtro
        topo = ttk.Frame(aba)
        topo.pack(side="top", fill="x", padx=4, pady=4)
        ttk.Label(topo, text="Buscar:").pack(side="left", padx=(0, 4))
        self.filtro_ofx = tk.StringVar()
        self.filtro_ofx.trace_add("write", lambda *_a: self._render_aba_ofx())
        ttk.Entry(topo, textvariable=self.filtro_ofx, width=40).pack(side="left")
        ttk.Button(topo, text="Limpar", command=self._limpa_filtros_ofx).pack(
            side="left", padx=4,
        )
        self.lbl_filtro_ofx = ttk.Label(topo, text="", foreground="#666")
        self.lbl_filtro_ofx.pack(side="left", padx=8)

        # Estado dos filtros por coluna (estilo Excel)
        self.filtros_col_ofx: dict[str, set[str] | None] = {}

        # Treeview + scrollbar
        corpo = ttk.Frame(aba)
        corpo.pack(side="top", fill="both", expand=True)
        # Fornecedor e CNPJ só ficam preenchidos quando o OFX foi
        # enriquecido por comprovante PDF que casou por data+valor.
        cols = ("data", "banco", "documento", "valor", "memo",
                "fornecedor", "cnpj")
        tree = ttk.Treeview(corpo, columns=cols, show="headings")
        for c, t, w, a in [
            ("data", "Data pagamento", 110, "center"),
            ("banco", "Banco", 120, "w"),
            ("documento", "Documento", 100, "w"),
            ("valor", "Valor", 100, "e"),
            ("memo", "Memo", 260, "w"),
            ("fornecedor", "Fornecedor (via PDF)", 200, "w"),
            ("cnpj", "CNPJ (via PDF)", 130, "w"),
        ]:
            tree.heading(c, text=self._label_coluna_filtro(t, False))
            tree.column(c, width=w, anchor=a)
        sb = ttk.Scrollbar(corpo, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree_ofx = tree
        # Tag pra destacar linhas enriquecidas por PDF (fundo azul claro)
        tree.tag_configure("enriquecido_pdf", background="#e7f3ff")
        tree.bind("<Button-1>", self._on_click_header_ofx)

    def _row_ofx(self, t) -> tuple:
        return (
            t.data.strftime("%d/%m/%Y"),
            t.extras.get("banco", "") or "",
            t.extras.get("documento", "") or "",
            f"{t.valor:.2f}",
            t.descricao or "",
            t.extras.get("fornecedor", "") or "",
            t.extras.get("cnpj", "") or "",
        )

    COLS_OFX = ("data", "banco", "documento", "valor", "memo",
                "fornecedor", "cnpj")
    LABELS_OFX = {
        "data": "Data pagamento", "banco": "Banco",
        "documento": "Documento",
        "valor": "Valor", "memo": "Memo",
        "fornecedor": "Fornecedor (via PDF)",
        "cnpj": "CNPJ (via PDF)",
    }

    def _on_click_header_ofx(self, event: tk.Event) -> None:
        self._abrir_filtro_excel(
            tree=self.tree_ofx,
            col_ids=self.COLS_OFX,
            col_labels=self.LABELS_OFX,
            filtros=self.filtros_col_ofx,
            rows_iter=lambda: (self._row_ofx(t) for t in self.transacoes_ofx),
            event=event,
            on_apply=lambda: (self._atualiza_headers_ofx(), self._render_aba_ofx()),
        )

    def _atualiza_headers_ofx(self) -> None:
        for col in self.COLS_OFX:
            ativo = self.filtros_col_ofx.get(col) is not None
            self.tree_ofx.heading(
                col, text=self._label_coluna_filtro(self.LABELS_OFX[col], ativo),
            )

    def _limpa_filtros_ofx(self) -> None:
        self.filtro_ofx.set("")
        self.filtros_col_ofx.clear()
        self._atualiza_headers_ofx()

    def _monta_aba_dominio_dados(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Domínio dados (0)")
        self._aba_dominio_dados = aba

        # Barra de filtro
        topo = ttk.Frame(aba)
        topo.pack(side="top", fill="x", padx=4, pady=4)
        ttk.Label(topo, text="Buscar:").pack(side="left", padx=(0, 4))
        self.filtro_dominio = tk.StringVar()
        self.filtro_dominio.trace_add(
            "write", lambda *_a: self._render_aba_dominio_dados(),
        )
        ttk.Entry(topo, textvariable=self.filtro_dominio, width=40).pack(side="left")
        ttk.Label(topo, text="Status:").pack(side="left", padx=(12, 4))
        self.filtro_dominio_status = tk.StringVar(value="Todos")
        cb_status = ttk.Combobox(
            topo, textvariable=self.filtro_dominio_status, state="readonly",
            values=["Todos", "Aberto", "Parcial", "Paga"], width=10,
        )
        cb_status.pack(side="left")
        cb_status.bind(
            "<<ComboboxSelected>>", lambda _e: self._render_aba_dominio_dados(),
        )
        ttk.Button(topo, text="Limpar", command=self._limpa_filtros_dominio).pack(
            side="left", padx=4,
        )
        self.lbl_filtro_dominio = ttk.Label(topo, text="", foreground="#666")
        self.lbl_filtro_dominio.pack(side="left", padx=8)

        # Estado dos filtros por coluna (estilo Excel)
        self.filtros_col_dominio: dict[str, set[str] | None] = {}

        # Treeview + scrollbar
        corpo = ttk.Frame(aba)
        corpo.pack(side="top", fill="both", expand=True)
        cols = (
            "venc", "emis", "valor", "pago", "status",
            "nf", "cnpj", "fornecedor", "empresa",
        )
        tree = ttk.Treeview(corpo, columns=cols, show="headings")
        for c, t, w, a in [
            ("venc", "Vencimento", 95, "center"),
            ("emis", "Emissão", 95, "center"),
            ("valor", "Valor parcela", 100, "e"),
            ("pago", "Valor pago", 95, "e"),
            ("status", "Status", 80, "center"),
            ("nf", "Nº NF", 75, "center"),
            ("cnpj", "CNPJ", 130, "w"),
            ("fornecedor", "Fornecedor", 200, "w"),
            # Empresa: só populada quando grupo empresarial (matriz + filiais)
            ("empresa", "Empresa (código)", 130, "w"),
        ]:
            tree.heading(c, text=self._label_coluna_filtro(t, False))
            tree.column(c, width=w, anchor=a)
        tree.tag_configure("aberto", background="#d4edda")    # verde
        tree.tag_configure("parcial", background="#cfe2ff")   # azul claro
        tree.tag_configure("paga", background="#e9ecef")      # cinza (já liquidada)
        sb = ttk.Scrollbar(corpo, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree_dominio_dados = tree
        tree.bind("<Button-1>", self._on_click_header_dominio)

    def _row_dominio(self, t) -> tuple:
        v_pago = t.extras.get("valor_pago")
        pago_txt = f"{v_pago:.2f}" if v_pago is not None else ""
        # Empresa: "código - razao" quando marcada pelo grupo empresarial
        codi = t.extras.get("codi_emp_origem")
        razao = t.extras.get("razao_empresa", "")
        if codi is not None:
            empresa_txt = f"{codi} - {razao[:30]}" if razao else str(codi)
        else:
            empresa_txt = ""
        return (
            t.data.strftime("%d/%m/%Y"),
            self._fmt_data(t.extras.get("data_emissao")),
            f"{t.valor:.2f}",
            pago_txt,
            t.extras.get("status", "") or "",
            t.extras.get("numero_nf", "") or "",
            t.extras.get("cnpj", "") or "",
            t.extras.get("fornecedor", "") or "",
            empresa_txt,
        )

    COLS_DOMINIO = (
        "venc", "emis", "valor", "pago", "status",
        "nf", "cnpj", "fornecedor", "empresa",
    )
    LABELS_DOMINIO = {
        "venc": "Vencimento", "emis": "Emissão", "valor": "Valor parcela",
        "pago": "Valor pago", "status": "Status", "nf": "Nº NF",
        "cnpj": "CNPJ", "fornecedor": "Fornecedor",
        "empresa": "Empresa (código)",
    }

    def _on_click_header_dominio(self, event: tk.Event) -> None:
        self._abrir_filtro_excel(
            tree=self.tree_dominio_dados,
            col_ids=self.COLS_DOMINIO,
            col_labels=self.LABELS_DOMINIO,
            filtros=self.filtros_col_dominio,
            rows_iter=lambda: (self._row_dominio(t) for t in self.transacoes_dominio),
            event=event,
            on_apply=lambda: (self._atualiza_headers_dominio(), self._render_aba_dominio_dados()),
        )

    def _atualiza_headers_dominio(self) -> None:
        for col in self.COLS_DOMINIO:
            ativo = self.filtros_col_dominio.get(col) is not None
            self.tree_dominio_dados.heading(
                col, text=self._label_coluna_filtro(self.LABELS_DOMINIO[col], ativo),
            )

    def _limpa_filtros_dominio(self) -> None:
        self.filtro_dominio.set("")
        self.filtro_dominio_status.set("Todos")
        if hasattr(self, "filtros_col_dominio"):
            self.filtros_col_dominio.clear()
        if hasattr(self, "tree_dominio_dados"):
            self._atualiza_headers_dominio()

    def _monta_aba_conciliados(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Conciliados (0)")
        self._aba_conciliados = aba

        cols = (
            "tipo", "origem", "data", "pagto", "valor", "emissao",
            "nf", "cnpj", "fornecedor", "memo_ofx", "diff",
        )
        tree = ttk.Treeview(aba, columns=cols, show="headings")
        tree.heading("tipo", text="Tipo")
        tree.heading("origem", text="Origem")
        tree.heading("data", text="Vencimento")
        tree.heading("pagto", text="Pagamento")
        tree.heading("valor", text="Valor")
        tree.heading("emissao", text="Emissão")
        tree.heading("nf", text="Nº NF")
        tree.heading("cnpj", text="CNPJ")
        tree.heading("fornecedor", text="Fornecedor")
        tree.heading("memo_ofx", text="Memo OFX")
        tree.heading("diff", text="Diferenças")
        tree.column("tipo", width=60, anchor="w")
        tree.column("origem", width=130, anchor="w")
        tree.column("data", width=85, anchor="center")
        tree.column("pagto", width=85, anchor="center")
        tree.column("valor", width=95, anchor="e")
        tree.column("emissao", width=85, anchor="center")
        tree.column("nf", width=70, anchor="center")
        tree.column("cnpj", width=130, anchor="w")
        tree.column("fornecedor", width=170, anchor="w")
        tree.column("memo_ofx", width=180, anchor="w")
        tree.column("diff", width=110, anchor="w")
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

        # IMPORTANTE: o rodapé é packado ANTES dos blocos com expand=True,
        # pra ficar ancorado embaixo mesmo se a janela encolher (caso contrário
        # ele sai da tela porque os blocos expandidos roubam o espaço).
        rodape = ttk.Frame(aba)
        rodape.pack(side="bottom", fill="x", padx=6, pady=(2, 6))
        ttk.Button(
            rodape, text="Conciliar selecionadas (escolha 1 linha em cada bloco) →",
            command=self._conciliar_selecionadas,
        ).pack(side="left", padx=4, pady=2)
        ttk.Button(
            rodape, text="Exportar para Excel (.xlsx)",
            command=self._exportar_pendentes,
        ).pack(side="left", padx=4, pady=2)

        # ----- Bloco PLANILHA (em cima): tabela + ações da planilha
        lado_p = ttk.LabelFrame(aba, text="Só na planilha")
        lado_p.pack(side="top", fill="both", expand=True, padx=6, pady=(0, 3))

        # Botões da planilha — packados PRIMEIRO no fundo do LabelFrame pra
        # ficarem grudados embaixo da tabela mesmo com a tabela expandindo.
        acoes_p = ttk.Frame(lado_p)
        acoes_p.pack(side="bottom", fill="x", padx=4, pady=(2, 4))
        ttk.Button(
            acoes_p, text="Lançamento manual",
            command=self._lancamento_manual_pend_planilha,
        ).pack(side="left", padx=2)
        ttk.Button(
            acoes_p, text="Criar regra (fornecedor)",
            command=self._criar_lancamento_padrao_planilha,
        ).pack(side="left", padx=2)

        # Tabela planilha
        tabela_p = ttk.Frame(lado_p)
        tabela_p.pack(side="top", fill="both", expand=True)
        cols_p = ("data", "pagto", "valor", "nf", "fornecedor", "historico", "tipo")
        self.tree_pend_p = ttk.Treeview(
            tabela_p, columns=cols_p, show="headings", selectmode="browse",
        )
        for c, t, w, a in [
            ("data", "Vencimento", 90, "center"),
            ("pagto", "Pagamento", 90, "center"),
            ("valor", "Valor", 100, "e"),
            ("nf", "Nº NF", 70, "center"),
            ("fornecedor", "Fornecedor", 220, "w"),
            ("historico", "Histórico", 220, "w"),
            ("tipo", "Tipo", 130, "w"),
        ]:
            self.tree_pend_p.heading(c, text=t)
            self.tree_pend_p.column(c, width=w, anchor=a)
        sb_p = ttk.Scrollbar(tabela_p, orient="vertical", command=self.tree_pend_p.yview)
        self.tree_pend_p.configure(yscrollcommand=sb_p.set)
        self.tree_pend_p.pack(side="left", fill="both", expand=True)
        sb_p.pack(side="right", fill="y")

        # ----- Bloco OFX (embaixo): tabela + ações do OFX
        lado_o = ttk.LabelFrame(aba, text="Só no OFX")
        lado_o.pack(side="top", fill="both", expand=True, padx=6, pady=(3, 4))

        acoes_o = ttk.Frame(lado_o)
        acoes_o.pack(side="bottom", fill="x", padx=4, pady=(2, 4))
        ttk.Button(
            acoes_o, text="Lançamento manual",
            command=self._lancamento_manual_pend_ofx,
        ).pack(side="left", padx=2)
        ttk.Button(
            acoes_o, text="Criar regra (memo)",
            command=self._criar_lancamento_padrao,
        ).pack(side="left", padx=2)

        tabela_o = ttk.Frame(lado_o)
        tabela_o.pack(side="top", fill="both", expand=True)
        # Fornecedor e CNPJ ficam preenchidos quando o OFX foi enriquecido
        # por comprovante PDF que casou por data+valor durante a conciliação
        # anterior — mesmo que depois tenha voltado pra Pendentes.
        cols_o = ("data", "banco", "documento", "valor", "descricao",
                  "fornecedor", "cnpj")
        self.tree_pend_o = ttk.Treeview(
            tabela_o, columns=cols_o, show="headings", selectmode="browse",
        )
        for c, t, w, a in [
            ("data", "Data pagamento", 100, "center"),
            ("banco", "Banco", 130, "w"),
            ("documento", "Documento", 100, "w"),
            ("valor", "Valor", 100, "e"),
            ("descricao", "Memo OFX", 280, "w"),
            ("fornecedor", "Fornecedor (via PDF)", 180, "w"),
            ("cnpj", "CNPJ (via PDF)", 130, "w"),
        ]:
            self.tree_pend_o.heading(c, text=t)
            self.tree_pend_o.column(c, width=w, anchor=a)
        # Tag pra destacar linhas com dados enriquecidos por PDF
        self.tree_pend_o.tag_configure("enriquecido_pdf", background="#e7f3ff")
        sb_o = ttk.Scrollbar(tabela_o, orient="vertical", command=self.tree_pend_o.yview)
        self.tree_pend_o.configure(yscrollcommand=sb_o.set)
        self.tree_pend_o.pack(side="left", fill="both", expand=True)
        sb_o.pack(side="right", fill="y")

    def _monta_aba_sugestoes(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Sugestões (0)")
        self._aba_sugestoes = aba

        instr = ttk.Label(
            aba,
            text=(
                "Pares com diferença de até 2 dias e até R$ 10,00. "
                "Marque uma ou mais sugestões (Ctrl+clique ou Shift+clique, "
                "ou Ctrl+A pra todas) e clique em 'Aceitar selecionadas'."
            ),
        )
        instr.pack(anchor="w", padx=6, pady=(6, 4))

        cols = ("data_p", "valor_p", "nf_p", "forn_p", "data_o", "valor_o", "memo_o", "diff_dias", "diff_valor")
        tree = ttk.Treeview(aba, columns=cols, show="headings", selectmode="extended")
        for c, t, w, a in [
            ("data_p", "Venc. (pla)", 90, "center"),
            ("valor_p", "Valor (pla)", 90, "e"),
            ("nf_p", "NF (pla)", 70, "center"),
            ("forn_p", "Fornecedor (pla)", 200, "w"),
            ("data_o", "Pagto (OFX)", 90, "center"),
            ("valor_o", "Valor (OFX)", 90, "e"),
            ("memo_o", "Memo (OFX)", 200, "w"),
            ("diff_dias", "Δ dias", 55, "center"),
            ("diff_valor", "Δ R$", 75, "e"),
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
            botoes, text="Selecionar tudo",
            command=self._selecionar_todas_sugestoes,
        ).pack(side="left", padx=6, pady=6)
        ttk.Button(
            botoes, text="Aceitar selecionadas",
            command=self._aceitar_sugestao,
        ).pack(side="left", padx=6, pady=6)

        self.tree_sugestoes = tree
        # Atalho Ctrl+A pra selecionar tudo
        tree.bind("<Control-a>", lambda _e: self._selecionar_todas_sugestoes())
        tree.bind("<Control-A>", lambda _e: self._selecionar_todas_sugestoes())

    def _monta_aba_conciliados_dominio(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Conciliados × Domínio (0)")
        self._aba_conciliados_dominio = aba

        instr = ttk.Label(
            aba,
            text=(
                "Lançamentos que batem no Domínio (data + valor + NF). "
                "Inclui pares Planilha × OFX e também pendentes da planilha "
                "do Caixa geral (sem OFX) que casaram com o Domínio."
            ),
            foreground="#1f3a68",
            font=("TkDefaultFont", 9, "italic"),
        )
        instr.pack(side="top", fill="x", padx=6, pady=(6, 0))

        cols = (
            "tipo", "origem", "data", "pagto", "valor", "emissao",
            "nf", "cnpj", "fornecedor", "memo_ofx", "diff_dom", "status_dom",
        )
        tree = ttk.Treeview(aba, columns=cols, show="headings")
        tree.heading("tipo", text="Tipo")
        tree.heading("origem", text="Origem")
        tree.heading("data", text="Vencimento")
        tree.heading("pagto", text="Pagamento")
        tree.heading("valor", text="Valor")
        tree.heading("emissao", text="Emissão")
        tree.heading("nf", text="Nº NF")
        tree.heading("cnpj", text="CNPJ")
        tree.heading("fornecedor", text="Fornecedor")
        tree.heading("memo_ofx", text="Memo OFX")
        tree.heading("diff_dom", text="Δ Domínio")
        tree.heading("status_dom", text="Status (Domínio)")
        tree.column("tipo", width=55, anchor="w")
        tree.column("origem", width=120, anchor="w")
        tree.column("data", width=85, anchor="center")
        tree.column("pagto", width=85, anchor="center")
        tree.column("valor", width=90, anchor="e")
        tree.column("emissao", width=85, anchor="center")
        tree.column("nf", width=70, anchor="center")
        tree.column("cnpj", width=130, anchor="w")
        tree.column("fornecedor", width=140, anchor="w")
        tree.column("memo_ofx", width=130, anchor="w")
        tree.column("diff_dom", width=110, anchor="center")
        tree.column("status_dom", width=110, anchor="center")
        tree.tag_configure("aberto", background="#d4edda")    # verde
        tree.tag_configure("parcial", background="#cfe2ff")   # azul claro
        tree.tag_configure("paga", background="#e9ecef")      # cinza (já liquidada)

        # Rodapé com botão de exportação — packado ANTES do tree pra ficar
        # ancorado embaixo mesmo com tree expandindo
        rodape_cd = ttk.Frame(aba)
        rodape_cd.pack(side="bottom", fill="x", padx=6, pady=(2, 6))
        ttk.Button(
            rodape_cd, text="Exportar para Excel (.xlsx)",
            command=self._exportar_conciliados_dominio,
        ).pack(side="left", padx=2)

        sb = ttk.Scrollbar(aba, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        self.tree_conciliados_dominio = tree

    def _monta_aba_dominio(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Domínio (0)")
        self._aba_dominio = aba

        # ---- Cabeçalho: instrução + legenda de cores lado a lado
        topo_comp = ttk.Frame(aba)
        topo_comp.pack(side="top", fill="x", padx=6, pady=(6, 4))

        instr = ttk.Label(
            topo_comp,
            text=(
                "Comparação de 3 fontes com o Domínio:\n"
                "• Pares Planilha×OFX conciliados\n"
                "• Pendentes só na planilha (Caixa geral)\n"
                "• Pendentes só no OFX (quando não há planilha importada)"
            ),
            wraplength=500,
            justify="left",
        )
        instr.pack(side="left", anchor="nw")

        # Legenda das cores — 6 chips coloridos com significado
        legenda = ttk.LabelFrame(topo_comp, text="Legenda de cores")
        legenda.pack(side="right", padx=(10, 0))
        cores = [
            ("#d4edda", "OK",          "Conciliado P×OFX e no Domínio"),
            ("#fff3cd", "Falta dom",   "Conciliado P×OFX, falta no Domínio"),
            ("#cce5ff", "Caixa OK",    "Pendente da planilha (Caixa geral) no Domínio"),
            ("#e2e3e5", "Caixa falta", "Pendente da planilha (Caixa geral) falta no Domínio"),
            ("#d1ecf1", "OFX OK",      "Pendente do OFX (sem planilha) no Domínio"),
            ("#ffe5cc", "OFX falta",   "Pendente do OFX (sem planilha) falta no Domínio"),
        ]
        for i, (cor, rotulo, desc) in enumerate(cores):
            # tk.Label aceita background — ttk.Label ignora em alguns temas
            chip = tk.Label(
                legenda, text=f"  {rotulo}  ",
                background=cor, foreground="#111",
                font=("TkDefaultFont", 8, "bold"),
                relief="solid", borderwidth=1,
            )
            chip.grid(row=i, column=0, padx=(6, 4), pady=1, sticky="w")
            ttk.Label(
                legenda, text=desc, font=("TkDefaultFont", 8),
                foreground="#333",
            ).grid(row=i, column=1, padx=(0, 8), pady=1, sticky="w")

        cols = (
            "status", "vencimento", "valor", "emissao", "nf",
            "cnpj", "fornecedor", "memo_ofx",
        )
        tree = ttk.Treeview(aba, columns=cols, show="headings")
        tree.heading("status", text="Status")
        tree.heading("vencimento", text="Vencimento")
        tree.heading("valor", text="Valor")
        tree.heading("emissao", text="Emissão")
        tree.heading("nf", text="Nº NF")
        tree.heading("cnpj", text="CNPJ")
        tree.heading("fornecedor", text="Fornecedor")
        tree.heading("memo_ofx", text="Memo OFX")
        tree.column("status", width=180, anchor="w")
        tree.column("vencimento", width=85, anchor="center")
        tree.column("valor", width=100, anchor="e")
        tree.column("emissao", width=85, anchor="center")
        tree.column("nf", width=75, anchor="center")
        tree.column("cnpj", width=130, anchor="w")
        tree.column("fornecedor", width=220, anchor="w")
        tree.column("memo_ofx", width=200, anchor="w")
        tree.tag_configure("ok", background="#d4edda")
        tree.tag_configure("falta_dominio", background="#fff3cd")
        tree.tag_configure("falta_concil", background="#f8d7da")
        # Pendentes da planilha (Caixa geral, sem OFX)
        tree.tag_configure("caixa_ok", background="#cce5ff")          # azul claro
        tree.tag_configure("caixa_falta", background="#e2e3e5")       # cinza claro
        # Pendentes do OFX (sem planilha) — comparação OFX × Domínio direta
        tree.tag_configure("ofx_ok", background="#d1ecf1")            # ciano claro
        tree.tag_configure("ofx_falta", background="#ffe5cc")         # laranja claro

        sb = ttk.Scrollbar(aba, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.tree_dominio = tree
        # iid → Par (para linhas P×OFX) OU Transacao (para Caixa geral)
        self._itens_comparacao: dict[str, Par | Transacao] = {}

        botoes = ttk.Frame(aba)
        botoes.pack(side="bottom", fill="x")
        ttk.Button(
            botoes, text="Editar dados",
            command=self._editar_par_amarelo,
        ).pack(side="top", anchor="w", padx=6, pady=(6, 2))
        ttk.Button(
            botoes, text="Lançar manualmente",
            command=self._lancar_manual,
        ).pack(side="top", anchor="w", padx=6, pady=2)
        ttk.Button(
            botoes, text="Criar regra de fornecedor (do amarelo selecionado)",
            command=self._criar_regra_fornecedor,
        ).pack(side="top", anchor="w", padx=6, pady=(2, 2))
        ttk.Button(
            botoes, text="Exportar pendências (amarelos + cinzas + laranjas) para Excel",
            command=self._exportar_pendencias_comparacao,
        ).pack(side="top", anchor="w", padx=6, pady=(2, 6))

    def _editar_par_amarelo(self) -> None:
        """Edita os dados do lado da planilha de um par amarelo (Conciliado,
        falta no Domínio). Útil pra corrigir NF/valor/CNPJ que estavam
        errados e impediam o match com o Domínio."""
        sel = self.tree_dominio.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione uma linha amarela (Conciliado, falta no Domínio) "
                "para editar.",
            )
            return
        item = self._itens_comparacao.get(sel[0])
        if not isinstance(item, Par):
            messagebox.showwarning(
                "Linha inválida",
                "Edição via diálogo só está disponível para linhas amarelas "
                "(pares Planilha×OFX). Para editar lançamentos de Caixa geral, "
                "use a aba 'Planilha'.",
            )
            return
        par = item
        if par.dominio is not None:
            messagebox.showwarning(
                "Linha inválida",
                "Edição só faz sentido para linhas amarelas (Conciliado, "
                "falta no Domínio).",
            )
            return

        dlg = DialogoEditarPar(self, par)
        self.wait_window(dlg)
        if not dlg.resultado:
            return

        # Aplica as alterações na Transacao da planilha
        r = dlg.resultado
        par.planilha.data = r["data"]
        par.planilha.valor = r["valor"]
        if r["data_emissao"]:
            par.planilha.extras["data_emissao"] = r["data_emissao"]
        elif "data_emissao" in par.planilha.extras:
            del par.planilha.extras["data_emissao"]
        for k in ("numero_nf", "cnpj", "fornecedor"):
            par.planilha.extras[k] = r[k]

        # Re-tenta match com Domínio + atualiza tudo
        self._filtrar_conciliados_por_dominio()
        self._gerar_lancamentos_contabeis()
        self._render_aba_planilha()      # planilha mudou
        self._render_conciliados()        # par mudou de valores
        self._comparar_com_dominio()      # re-renderiza Comparação + Conciliados × Domínio

        # Feedback claro pro usuário
        if par.dominio is not None:
            messagebox.showinfo(
                "Match!",
                "Após edição o par agora bate com o Domínio. Foi movido "
                "para a aba 'Conciliados × Domínio'.",
            )
        else:
            messagebox.showinfo(
                "Dados atualizados",
                "Dados salvos, mas o par ainda não casa com o Domínio. "
                "Verifique se Vencimento + Valor + Nº NF estão exatamente "
                "iguais ao registro do Domínio.",
            )

    def _lancar_manual(self) -> None:
        """Cria UM lançamento contábil avulso a partir da linha selecionada.
        Funciona em:
        - Linha amarela (Par P×OFX sem Domínio): banco vem do OFX
        - Linha cinza (Caixa geral, pendente planilha sem Domínio):
          banco='Caixa geral'
        """
        sel = self.tree_dominio.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione uma linha (amarela ou cinza) sem match no Domínio.",
            )
            return
        item = self._itens_comparacao.get(sel[0])
        if item is None:
            return

        if isinstance(item, Par):
            par = item
            if par.dominio is not None:
                messagebox.showwarning(
                    "Linha inválida",
                    "O lançamento manual só faz sentido para linhas SEM "
                    "match no Domínio.",
                )
                return
            forn = par.planilha.extras.get("fornecedor", "") or ""
            sugestao = f"Pagto. {forn}" if forn else ""
            dlg = DialogoLancamentoManual(
                self, par, plano_contas=self.plano_contas,
                sugestao_historico=sugestao,
            )
            self.wait_window(dlg)
            if not dlg.resultado:
                return
            lanc = LancamentoContabil(
                data=par.ofx.data,
                historico=dlg.resultado["historico"],
                valor=par.planilha.valor,
                banco=par.ofx.extras.get("banco", "") or "",
                memo_original=par.ofx.descricao or "",
                padrao_match="(manual)",
                conta=dlg.resultado["conta"],
                tipo_regra="manual",
                fornecedor=forn,
                cnpj=par.planilha.extras.get("cnpj", "") or "",
                transacao_origem=par.ofx,
                par_origem=par,
            )
        else:
            # Pendente da planilha (Caixa geral, sem OFX)
            t = item
            # Bloqueia se já tem match no Domínio
            match = self.pendentes_planilha_dominio.get(id(t))
            if match and match.get("dominio") is not None:
                messagebox.showwarning(
                    "Linha inválida",
                    "Esse lançamento já tem match no Domínio (linha azul). "
                    "Não precisa lançar manualmente.",
                )
                return
            forn = t.extras.get("fornecedor", "") or ""
            hist_plan = t.extras.get("historico", "") or ""
            sugestao = f"Pagto. {forn}" if forn else hist_plan
            dlg = DialogoLancamentoManualAvulso(
                self, t, origem="planilha",
                plano_contas=self.plano_contas,
                sugestao_historico=sugestao,
            )
            self.wait_window(dlg)
            if not dlg.resultado:
                return
            lanc = LancamentoContabil(
                data=t.data_pagamento or t.data,
                historico=dlg.resultado["historico"],
                valor=t.valor,
                banco="Caixa geral",
                memo_original="",
                padrao_match="(manual planilha)",
                conta=dlg.resultado["conta"],
                tipo_regra="manual_planilha",
                fornecedor=forn,
                cnpj=t.extras.get("cnpj", "") or "",
                transacao_origem=t,
                par_origem=None,
            )

        self.lancamentos_manuais.append(lanc)
        self._gerar_lancamentos_contabeis()
        self._comparar_com_dominio()  # remove a linha da Comparação

    def _criar_regra_fornecedor(self) -> None:
        """Atalho: cria regra do tipo 'fornecedor' a partir da linha
        selecionada na aba Comparação. Funciona tanto em par amarelo
        (P×OFX sem Domínio) quanto em pendente Caixa geral (cinza)."""
        if not self._exigir_empresa("criar regras"):
            return
        sel = self.tree_dominio.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione uma linha (amarela ou cinza) sem match no "
                "Domínio para criar uma regra de fornecedor.",
            )
            return
        item = self._itens_comparacao.get(sel[0])
        if item is None:
            return

        # Extrai dados da planilha (tanto Par quanto Transacao têm extras)
        if isinstance(item, Par):
            par = item
            if par.dominio is not None:
                messagebox.showwarning(
                    "Linha inválida",
                    "Regra só faz sentido para linhas SEM match no Domínio.",
                )
                return
            t_planilha = par.planilha
        else:
            t_planilha = item
            match = self.pendentes_planilha_dominio.get(id(t_planilha))
            if match and match.get("dominio") is not None:
                messagebox.showwarning(
                    "Linha inválida",
                    "Esse lançamento já tem match no Domínio. Não precisa "
                    "criar regra.",
                )
                return
        cnpj = t_planilha.extras.get("cnpj", "") or ""
        fornecedor = t_planilha.extras.get("fornecedor", "") or ""
        historico = t_planilha.extras.get("historico", "") or ""
        tipo_col = t_planilha.extras.get("tipo", "") or ""
        # Prioridade: CNPJ → fornecedor → tipo → histórico
        sugestao = (
            cnpj.strip() or fornecedor.strip()
            or tipo_col.strip() or historico.strip()
        )
        if not sugestao:
            messagebox.showwarning(
                "Sem dados",
                "A linha selecionada não tem CNPJ, fornecedor, tipo nem "
                "histórico preenchidos.",
            )
            return
        dlg = DialogoNovaRegra(
            self,
            regra_atual={"padrao": sugestao, "historico": "", "conta": ""},
            plano_contas=self.plano_contas,
            tipo="fornecedor",
        )
        emp = self._empresa_selecionada()
        dlg.title(f"Regra de fornecedor — {emp['razao'][:60]} (empresa {emp['codi_emp']})")
        self.wait_window(dlg)
        if not dlg.regra:
            return
        regras = self._get_regras_empresa()
        regras.append(dlg.regra)
        self._set_regras_empresa(regras)
        self._gerar_lancamentos_contabeis()
        self._comparar_com_dominio()  # re-renderiza tirando o par classificado

    def _monta_aba_lancamentos(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Lançamentos contábeis (0)")
        self._aba_lancamentos = aba

        info = ttk.Label(
            aba,
            text=(
                "Lançamentos contábeis gerados a partir dos pendentes do OFX "
                "que casam com as regras de taxas configuradas. Use "
                "'Configurar taxas' (linha de ações) para adicionar/remover. "
                "As regras são vinculadas à empresa selecionada no Domínio."
            ),
            wraplength=1100,
            foreground="#1f3a68",
            font=("TkDefaultFont", 9, "italic"),
        )
        info.pack(side="top", fill="x", padx=6, pady=(6, 2))

        self.lbl_lancamentos_empresa = ttk.Label(
            aba, text="", foreground="#555",
        )
        self.lbl_lancamentos_empresa.pack(side="top", fill="x", padx=6, pady=(0, 4))

        # Rodapé com botões — packado antes do corpo pra ficar ancorado embaixo
        rodape_lanc = ttk.Frame(aba)
        rodape_lanc.pack(side="bottom", fill="x", padx=6, pady=(2, 6))
        ttk.Button(
            rodape_lanc, text="Editar lançamento",
            command=self._editar_lancamento_contabil,
        ).pack(side="left", padx=2)
        ttk.Button(
            rodape_lanc, text="Excluir lançamento",
            command=self._excluir_lancamento_contabil,
        ).pack(side="left", padx=2)
        ttk.Button(
            rodape_lanc, text="Exportar para Excel (.xlsx)",
            command=self._exportar_lancamentos_contabeis,
        ).pack(side="left", padx=2)

        corpo_lanc = ttk.Frame(aba)
        corpo_lanc.pack(side="top", fill="both", expand=True)
        cols = ("data", "banco", "valor", "conta", "historico", "memo", "padrao")
        tree = ttk.Treeview(corpo_lanc, columns=cols, show="headings")
        for c, t, w, a in [
            ("data", "Data pagto", 100, "center"),
            ("banco", "Banco", 120, "w"),
            ("valor", "Valor", 100, "e"),
            ("conta", "Conta", 100, "w"),
            ("historico", "Histórico contábil", 220, "w"),
            ("memo", "Memo (OFX)", 240, "w"),
            ("padrao", "Regra", 120, "w"),
        ]:
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=a)
        sb = ttk.Scrollbar(corpo_lanc, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree_lancamentos = tree
        # iid → LancamentoContabil (para resolver seleção no botão)
        self.itens_lancamentos: dict[str, LancamentoContabil] = {}

    def _monta_aba_plano_contas(self) -> None:
        aba = ttk.Frame(self.notebook)
        self.notebook.add(aba, text="Plano de contas (0)")
        self._aba_plano_contas = aba

        info = ttk.Label(
            aba,
            text=(
                "Plano de contas carregado do Domínio para a empresa "
                "selecionada. Configure a fonte em 'Fonte: plano contas' "
                "(barra do Domínio) antes de carregar."
            ),
            foreground="#1f3a68",
            font=("TkDefaultFont", 9, "italic"),
        )
        info.pack(side="top", fill="x", padx=6, pady=(6, 2))

        # Filtro de busca
        topo_f = ttk.Frame(aba)
        topo_f.pack(side="top", fill="x", padx=6, pady=2)
        ttk.Label(topo_f, text="Buscar:").pack(side="left", padx=(0, 4))
        self.filtro_plano = tk.StringVar()
        self.filtro_plano.trace_add(
            "write", lambda *_a: self._render_aba_plano_contas(),
        )
        ttk.Entry(topo_f, textvariable=self.filtro_plano, width=40).pack(side="left")
        ttk.Button(
            topo_f, text="Limpar", command=lambda: self.filtro_plano.set(""),
        ).pack(side="left", padx=4)
        self.lbl_filtro_plano = ttk.Label(topo_f, text="", foreground="#666")
        self.lbl_filtro_plano.pack(side="left", padx=8)

        cols = ("codigo", "descricao", "tipo")
        tree = ttk.Treeview(aba, columns=cols, show="headings")
        tree.heading("codigo", text="Código")
        tree.heading("descricao", text="Descrição")
        tree.heading("tipo", text="Tipo")
        tree.column("codigo", width=130, anchor="w")
        tree.column("descricao", width=500, anchor="w")
        tree.column("tipo", width=60, anchor="center")
        sb = ttk.Scrollbar(aba, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree_plano_contas = tree

    def _render_aba_plano_contas(self) -> None:
        for item in self.tree_plano_contas.get_children():
            self.tree_plano_contas.delete(item)
        termo = (
            self.filtro_plano.get().strip().lower()
            if hasattr(self, "filtro_plano") else ""
        )
        mostradas = 0
        for c in self.plano_contas:
            row = (c.codigo, c.descricao, c.tipo)
            if termo and termo not in " ".join(str(x) for x in row).lower():
                continue
            self.tree_plano_contas.insert("", "end", values=row)
            mostradas += 1
        total = len(self.plano_contas)
        idx = self.notebook.index(self._aba_plano_contas)
        self.notebook.tab(idx, text=f"Plano de contas ({total})")
        if hasattr(self, "lbl_filtro_plano"):
            self.lbl_filtro_plano.config(
                text=f"Mostrando {mostradas} de {total}" if termo else f"{total} contas",
            )

    # --------------------------------------------------------- Domínio

    def _conectar_dominio(self) -> None:
        dlg = DialogoConexao(self)
        self.wait_window(dlg)
        if dlg.conn is None:
            return
        self.conn_dominio = dlg.conn
        self._pos_conexao_dominio()

    def _pos_conexao_dominio(self) -> None:
        """Ações depois que self.conn_dominio foi estabelecido: habilita
        botões dependentes e atualiza o label. Usado tanto pela conexão
        manual (via diálogo) quanto pela auto-conexão do startup."""
        self.btn_empresa.config(state="normal")
        self.btn_fonte.config(state="normal")
        self.btn_fonte_plano.config(state="normal")
        if self.cfg.get("dominio_fonte_pagamentos", {}).get("mapeamento"):
            self.btn_carregar_dominio.config(state="normal")
        if self.cfg.get("dominio_fonte_plano_contas", {}).get("mapeamento"):
            self.btn_carregar_plano.config(state="normal")
        self._atualiza_label_dominio()

    def _auto_conectar_dominio(self) -> None:
        """Tenta abrir a conexão com o Domínio automaticamente usando as
        credenciais salvas em data/dominio_config.json. Silencioso — se
        falhar, o usuário pode conectar manualmente no botão."""
        if self.conn_dominio is not None:
            return
        try:
            cred = parser_dominio.load_odbc_config()
        except Exception:
            return
        if not cred.get("dsn"):
            return  # ainda não configurou credenciais
        try:
            self.conn_dominio = parser_dominio.open_connection(cfg=cred)
        except Exception as e:
            # Não travar o app — só avisa no label do Domínio pra o usuário
            # saber que a auto-conexão falhou e pode tentar manual.
            self.lbl_dominio.config(
                text=f"(Auto-conexão falhou: {e}. Clique em 'Conectar Domínio')",
                foreground="#c0392b",
            )
            return
        self._pos_conexao_dominio()

    def _selecionar_empresa(self) -> None:
        if self.conn_dominio is None:
            return
        emp_anterior = self.cfg.get("dominio_empresa") or {}
        dlg = DialogoSelecionarEmpresa(
            self, self.conn_dominio, emp_anterior or None,
        )
        self.wait_window(dlg)
        if dlg.empresa is None:
            return

        # Detecta TROCA de empresa (não primeira seleção, ID diferente)
        codi_ant = emp_anterior.get("codi_emp") if emp_anterior else None
        codi_nov = dlg.empresa.get("codi_emp")
        trocou = codi_ant is not None and codi_ant != codi_nov

        if trocou:
            # Tem algo importado da empresa anterior? Confirma antes de descartar.
            tem_dados = bool(
                self.transacoes_planilha or self.transacoes_ofx
                or self.transacoes_dominio or self.plano_contas
            )
            if tem_dados:
                if not messagebox.askyesno(
                    "Trocar empresa",
                    f"Trocar de '{emp_anterior.get('razao', codi_ant)}' "
                    f"para '{dlg.empresa.get('razao', codi_nov)}'?\n\n"
                    "Os dados de Planilha, OFX, Domínio e Plano de contas "
                    "carregados serão DESCARTADOS (são específicos da "
                    "empresa atual). Você terá que importar novamente os "
                    "dados da nova empresa.",
                ):
                    return

            # Limpa tudo que é específico da empresa antiga
            self._limpar_dados_empresa()

        self.cfg["dominio_empresa"] = dlg.empresa
        config.salvar(self.cfg)
        self._atualiza_label_dominio()
        # Regras de taxas mudam com a empresa — recalcula lançamentos
        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()

    def _limpar_dados_empresa(self) -> None:
        """Limpa planilha, OFX, Domínio e plano de contas — invocado ao
        trocar de empresa. As regras de taxas continuam salvas por empresa
        no config.json (não são apagadas)."""
        # Planilha
        self.transacoes_planilha = []
        self.caminho_planilha = None
        self.estrutura_planilha = None
        self.mapeamento_planilha = None
        self.lbl_planilha.config(text="(nenhuma planilha carregada)")
        self.btn_editar_colunas.config(state="disabled")
        self.btn_limpar_planilha.config(state="disabled")
        # OFX
        self.transacoes_ofx = []
        self.caminhos_ofx = []
        self.lbl_ofx.config(text="(nenhum OFX carregado)")
        self.btn_limpar_ofx.config(state="disabled")
        # Domínio (pagamentos da empresa antiga)
        self.transacoes_dominio = []
        # Plano de contas (da empresa antiga)
        self.plano_contas = []
        # Re-renderiza abas de origem agora vazias + zera resultados
        self._render_aba_planilha()
        self._render_aba_ofx()
        if hasattr(self, "tree_dominio_dados"):
            for item in self.tree_dominio_dados.get_children():
                self.tree_dominio_dados.delete(item)
        if hasattr(self, "tree_plano_contas"):
            for item in self.tree_plano_contas.get_children():
                self.tree_plano_contas.delete(item)
        self._atualiza_botao()
        self._limpa_resultados()

    def _atualiza_label_dominio(self) -> None:
        cred = parser_dominio.load_odbc_config()
        emp = self.cfg.get("dominio_empresa")
        partes = [f"Conectado (read-only): DSN={cred.get('dsn', '?')}"]
        if emp:
            partes.append(f"Empresa: {emp['codi_emp']} — {emp['razao'][:40]}")
        if self.transacoes_dominio:
            partes.append(f"{len(self.transacoes_dominio)} pagamentos")
        if self.plano_contas:
            partes.append(f"{len(self.plano_contas)} contas no plano")
        self.lbl_dominio.config(text="  |  ".join(partes))

    def _configurar_fonte_dominio(self) -> None:
        """Configura a fonte de pagamentos do Domínio (parcelas a pagar)."""
        if self.conn_dominio is None:
            return
        fonte_atual = self.cfg.get("dominio_fonte_pagamentos", {})
        codi_emp = (self.cfg.get("dominio_empresa") or {}).get("codi_emp")
        dlg = DialogoFonte(
            self, self.conn_dominio, fonte_atual, codi_emp=codi_emp,
            titulo="Selecionar fonte de PAGAMENTOS no Domínio",
        )
        self.wait_window(dlg)
        if dlg.fonte is None:
            return
        self.cfg["dominio_fonte_pagamentos"] = dlg.fonte
        config.salvar(self.cfg)
        self.btn_carregar_dominio.config(state="normal")

    def _configurar_fonte_plano_contas(self) -> None:
        """Configura a fonte do plano de contas do Domínio."""
        if self.conn_dominio is None:
            return
        fonte_atual = self.cfg.get("dominio_fonte_plano_contas", {})
        codi_emp = (self.cfg.get("dominio_empresa") or {}).get("codi_emp")
        dlg = DialogoFonte(
            self, self.conn_dominio, fonte_atual, codi_emp=codi_emp,
            campos=DialogoFonte.CAMPOS_PLANO_CONTAS,
            titulo="Selecionar fonte do PLANO DE CONTAS no Domínio",
            opcionais={"tipo"},
        )
        self.wait_window(dlg)
        if dlg.fonte is None:
            return
        self.cfg["dominio_fonte_plano_contas"] = dlg.fonte
        config.salvar(self.cfg)
        self.btn_carregar_plano.config(state="normal")

    def _carregar_dominio(self) -> None:
        if self.conn_dominio is None:
            return
        fonte = self.cfg.get("dominio_fonte_pagamentos", {})
        if not fonte.get("mapeamento"):
            messagebox.showinfo(
                "Sem fonte",
                "Configure a fonte de pagamentos primeiro.",
            )
            return
        emp = self.cfg.get("dominio_empresa") or {}
        codi_emp = emp.get("codi_emp")
        cnpj_matriz = emp.get("cnpj", "")
        if codi_emp is None and fonte.get("modo") == "tabela":
            if not messagebox.askyesno(
                "Sem empresa selecionada",
                "Você não selecionou uma empresa — a query vai retornar "
                "lançamentos de TODAS as empresas misturadas.\n\n"
                "Deseja continuar mesmo assim?",
            ):
                return

        # ---- Detecta filiais do mesmo grupo (mesmo CNPJ raiz)
        # A matriz paga boletos das filiais frequentemente; se não
        # carregarmos as parcelas de todas as empresas do grupo, muitas
        # notas ficam como "falta no Domínio" (amarelo) sem motivo.
        empresas_pra_carregar: list[dict] = []
        if codi_emp is not None and cnpj_matriz:
            try:
                filiais = parser_dominio.listar_filiais(
                    self.conn_dominio, cnpj_matriz,
                )
                if filiais:
                    empresas_pra_carregar = filiais
            except Exception:
                pass

        # Fallback: se não achou filiais (ou sem CNPJ da matriz),
        # carrega só a empresa selecionada
        if not empresas_pra_carregar:
            empresas_pra_carregar = [{
                "codi_emp": codi_emp,
                "razao": emp.get("razao", ""),
                "cnpj": cnpj_matriz,
            }]

        # ---- Carrega parcelas de cada empresa e marca a origem
        try:
            todas_transacoes: list[Transacao] = []
            for e in empresas_pra_carregar:
                codi = e.get("codi_emp")
                txs = parser_dominio.extrair_pagamentos(
                    self.conn_dominio, fonte, codi_emp=codi,
                )
                # Marca cada Transacao com a empresa de origem
                for t in txs:
                    t.extras["codi_emp_origem"] = codi
                    t.extras["razao_empresa"] = e.get("razao", "")
                todas_transacoes.extend(txs)
            self.transacoes_dominio = todas_transacoes
        except Exception as e:
            messagebox.showerror("Erro ao ler Domínio", str(e))
            return

        # Se carregou de várias empresas, avisa
        if len(empresas_pra_carregar) > 1:
            resumo = "\n".join(
                f"• [{e['codi_emp']}] {e.get('razao','')[:60]} — "
                f"CNPJ {e.get('cnpj','')}"
                for e in empresas_pra_carregar
            )
            messagebox.showinfo(
                "Grupo empresarial detectado",
                f"Identifiquei {len(empresas_pra_carregar)} empresa(s) com "
                f"o mesmo CNPJ raiz da empresa selecionada.\n\n"
                f"Parcelas carregadas de TODAS:\n{resumo}\n\n"
                f"Total: {len(self.transacoes_dominio)} parcela(s).\n\n"
                "Isso permite conciliar boletos que a matriz pagou pelas "
                "filiais (ou vice-versa). O plano de contas do grupo é "
                "sempre o da matriz — se você selecionou uma filial, o "
                "sistema puxa o plano da matriz ao clicar em "
                "'Carregar plano contas'."
            )

        self._atualiza_label_dominio()
        self._render_aba_dominio_dados()
        # Se já houver conciliação P×O feita, refiltra e atualiza a aba Conciliados
        if self.pares_conciliados:
            self._filtrar_conciliados_por_dominio()
            self._render_conciliados()
            self._redesenha_abas()
        self._atualiza_botao_comparar()

    def _carregar_plano_contas(self) -> None:
        if self.conn_dominio is None:
            return
        fonte = self.cfg.get("dominio_fonte_plano_contas", {})
        if not fonte.get("mapeamento"):
            messagebox.showinfo(
                "Sem fonte",
                "Configure a fonte do plano de contas primeiro.",
            )
            return
        emp = self._empresa_selecionada() or {}
        codi_emp = emp.get("codi_emp")
        cnpj_sel = emp.get("cnpj", "")

        # Se a empresa selecionada é filial, o plano de contas contábil
        # tem que ser o da MATRIZ (mesmo plano pra todo o grupo, apenas
        # as contas bancárias mudam por filial nas regras de taxa).
        matriz_usada = None
        if codi_emp is not None and cnpj_sel:
            try:
                matriz = parser_dominio.encontrar_matriz(
                    self.conn_dominio, cnpj_sel,
                )
                if matriz and matriz.get("codi_emp") != codi_emp:
                    matriz_usada = matriz
                    codi_emp = matriz.get("codi_emp")
            except Exception:
                pass

        try:
            self.plano_contas = parser_dominio.extrair_plano_contas(
                self.conn_dominio, fonte, codi_emp=codi_emp,
            )
        except Exception as e:
            messagebox.showerror("Erro ao ler plano de contas", str(e))
            return
        self._render_aba_plano_contas()
        self._atualiza_label_dominio()

        if matriz_usada is not None:
            messagebox.showinfo(
                "Plano de contas — matriz do grupo",
                f"A empresa selecionada é filial ({emp.get('razao', '')}).\n\n"
                f"Carreguei o plano de contas da matriz:\n"
                f"[{matriz_usada['codi_emp']}] "
                f"{matriz_usada.get('razao', '')}\n"
                f"CNPJ {matriz_usada.get('cnpj', '')}\n\n"
                f"Total: {len(self.plano_contas)} conta(s) analítica(s).\n\n"
                "Isso garante que os lançamentos usem o mesmo plano contábil "
                "do grupo. Ajuste as regras de taxa por empresa se as "
                "contas bancárias forem diferentes."
            )

    def _atualiza_botao_comparar(self) -> None:
        """Mantido por compatibilidade — botão Comparar agora fica sempre
        habilitado e a validação acontece em _comparar_com_dominio."""
        pass

    def _comparar_com_dominio(self) -> None:
        # Valida pré-condições com mensagens claras
        if (
            not self.pares_conciliados
            and not self.pendentes_planilha_brutos
            and not self.pendentes_ofx_brutos
        ):
            messagebox.showwarning(
                "Sem dados",
                "Antes de comparar com o Domínio, é preciso ter pelo menos "
                "uma dessas fontes:\n"
                "• Planilha (.xlsx) importada\n"
                "• OFX importado\n\n"
                "Depois clique em 'Conciliar' para gerar pares/pendentes.",
            )
            return
        if not self.transacoes_dominio:
            messagebox.showwarning(
                "Domínio não carregado",
                "Carregue os pagamentos do Domínio antes de comparar:\n"
                "1. Conectar Domínio\n"
                "2. Selecionar empresa\n"
                "3. Fonte: pagamentos (configurar SQL)\n"
                "4. Clicar em 'Carregar pagamentos'",
            )
            return
        # Refiltra conciliados E pendentes pela regra triple
        self._filtrar_conciliados_por_dominio()
        # Regenera lançamentos: pares sem Domínio podem virar lançamento
        self._gerar_lancamentos_contabeis()
        self._render_conciliados()
        self._redesenha_abas()
        self._recalcular_comparacao()
        self.notebook.select(self._aba_dominio)

    def _recalcular_comparacao(self) -> None:
        """Recalcula e re-renderiza a aba Comparação. Diferente de
        _comparar_com_dominio: NÃO mostra mensagens de erro nem força
        foco na aba. Pode ser chamado de qualquer ponto que mude o
        estado (lançamento manual, criação de regra, etc)."""
        if not self.transacoes_dominio:
            # Nada a fazer — aba Comparação só existe com Domínio carregado
            return
        # Refiltra (caso pendentes_planilha_brutos tenha mudado)
        # — não chama _filtrar_conciliados_por_dominio aqui pra evitar loop
        # com _gerar_lancamentos_contabeis. O filtro já foi feito antes.
        self._renderizar_comparacao()

    def _renderizar_comparacao(self) -> None:
        """Monta a lista de resultados e chama _render_aba_dominio.
        Reutilizada por _comparar_com_dominio e _recalcular_comparacao."""
        # Monta a aba Comparação detalhada.
        # Linhas com regra/manual já aplicada são ocultadas.
        # Cada item: (status, planilha, ofx_ou_None, dominio_ou_None,
        #             diff_dias_dominio, diff_valor_dominio)
        from decimal import Decimal
        resultados: list[tuple] = []
        usados: set[int] = set()

        # 1) Pares P×OFX
        for par in self.pares_conciliados:
            if par.dominio is not None:
                resultados.append((
                    "ok", par.planilha, par.ofx, par.dominio,
                    par.diff_dias_dominio, par.diff_valor_dominio, par,
                ))
                usados.add(id(par.dominio))
            else:
                if id(par) in self.ids_pares_classificados:
                    continue  # já virou lançamento (regra/manual)
                resultados.append((
                    "falta_dominio", par.planilha, par.ofx, None,
                    0, Decimal("0"), par,
                ))

        # 2) Pendentes da planilha (Caixa geral, sem OFX)
        # Esconde os que já viraram lançamento contábil (regra fornecedor
        # ou lançamento manual avulso da planilha)
        ids_p_classificadas = {
            id(l.transacao_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("manual_planilha", "fornecedor_planilha")
            and l.transacao_origem is not None
        }
        for t_p in self.pendentes_planilha_brutos:
            if id(t_p) in ids_p_classificadas:
                continue
            match = self.pendentes_planilha_dominio.get(id(t_p))
            if match and match.get("dominio") is not None:
                resultados.append((
                    "caixa_ok", t_p, None, match["dominio"],
                    match["diff_dias"], match["diff_valor"], None,
                ))
                usados.add(id(match["dominio"]))
            else:
                resultados.append((
                    "caixa_falta", t_p, None, None,
                    0, Decimal("0"), None,
                ))

        # 3) Pendentes do OFX (sem planilha): comparação direta OFX × Domínio
        # Esconde os já classificados (regra memo ou manual OFX).
        ids_o_classificadas = {
            id(l.transacao_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("memo", "manual_ofx")
            and l.transacao_origem is not None
        }
        for t_o in self.pendentes_ofx_brutos:
            if id(t_o) in ids_o_classificadas:
                continue
            match = self.pendentes_ofx_dominio.get(id(t_o))
            if match and match.get("dominio") is not None:
                # Passamos t_o como "planilha" (é a origem para render) mas
                # também como t_ofx pra Memo aparecer
                resultados.append((
                    "ofx_ok", t_o, t_o, match["dominio"],
                    match["diff_dias"], match["diff_valor"], None,
                ))
                usados.add(id(match["dominio"]))
            else:
                resultados.append((
                    "ofx_falta", t_o, t_o, None,
                    0, Decimal("0"), None,
                ))

        # Pagamentos no Domínio que ninguém casou
        sobras_dominio = [t for t in self.transacoes_dominio if id(t) not in usados]

        self._render_aba_dominio(resultados, sobras_dominio)

    def _render_aba_dominio(
        self,
        resultados: list[tuple],
        sobras_dominio: list[Transacao],
    ) -> None:
        for item in self.tree_dominio.get_children():
            self.tree_dominio.delete(item)
        self._itens_comparacao.clear()

        def _fmt_data(d) -> str:
            return d.strftime("%d/%m/%Y") if d else ""

        rotulos = {
            "ok": "Conciliado P×OFX e no Domínio",
            "falta_dominio": "Conciliado P×OFX, falta no Domínio",
            "caixa_ok": "Caixa geral (no Domínio)",
            "caixa_falta": "Caixa geral (falta no Domínio)",
            "ofx_ok": "OFX (sem planilha) no Domínio",
            "ofx_falta": "OFX (sem planilha) falta no Domínio",
        }
        n_ok = n_falta_dom = n_caixa_ok = n_caixa_falta = 0
        n_ofx_ok = n_ofx_falta = 0
        for status, t_planilha, t_ofx, t_dom, _diff_d, _diff_v, par in resultados:
            if status == "ok":
                n_ok += 1
            elif status == "falta_dominio":
                n_falta_dom += 1
            elif status == "caixa_ok":
                n_caixa_ok += 1
            elif status == "caixa_falta":
                n_caixa_falta += 1
            elif status == "ofx_ok":
                n_ofx_ok += 1
            else:
                n_ofx_falta += 1
            rotulo = rotulos.get(status, status)
            # Extras: prioriza Domínio se houver, depois planilha
            origem_extras = t_dom.extras if t_dom else t_planilha.extras
            extras_fallback = t_planilha.extras
            emissao = origem_extras.get("data_emissao") or extras_fallback.get("data_emissao")
            nf = origem_extras.get("numero_nf") or extras_fallback.get("numero_nf", "")
            cnpj = origem_extras.get("cnpj") or extras_fallback.get("cnpj", "")
            fornecedor = origem_extras.get("fornecedor") or extras_fallback.get("fornecedor", "")
            memo = t_ofx.descricao if t_ofx else "(Caixa geral — sem OFX)"

            iid = self.tree_dominio.insert(
                "", "end",
                values=(
                    rotulo,
                    t_planilha.data.strftime("%d/%m/%Y"),
                    f"{t_planilha.valor:.2f}",
                    _fmt_data(emissao),
                    nf,
                    cnpj,
                    fornecedor,
                    memo,
                ),
                tags=(status,),
            )
            # Pares vão pro dict; pendentes da planilha (sem par) também
            # entram, mas com a Transacao da planilha — handlers fazem isinstance
            self._itens_comparacao[iid] = par if par is not None else t_planilha

        # Pares só no Domínio (que ninguém conciliou) NÃO são mostrados aqui.
        partes = [f"ok {n_ok}", f"falta dom {n_falta_dom}"]
        if n_caixa_ok or n_caixa_falta:
            partes.append(f"caixa {n_caixa_ok}/{n_caixa_falta}")
        if n_ofx_ok or n_ofx_falta:
            partes.append(f"OFX {n_ofx_ok}/{n_ofx_falta}")
        self.notebook.tab(7, text=f"Comparação ({' | '.join(partes)})")

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

        # Tenta aplicar o mapeamento salvo da empresa atual. Se todos os
        # campos casarem pelo nome da coluna, pula o diálogo e usa direto.
        # Se faltar algum, abre o diálogo com os campos resolvidos
        # pré-preenchidos para o usuário completar.
        mapa_resolvido, faltando = self._resolver_mapeamento_salvo(estrutura.cabecalho)
        usou_salvo_direto = False
        mapeamento_final: dict[str, int] | None = None

        if mapa_resolvido and not faltando:
            mapeamento_final = mapa_resolvido
            usou_salvo_direto = True
        else:
            # Pré-preenche a sugestão da estrutura com o que conseguimos resolver
            if mapa_resolvido:
                estrutura.sugestao = {**estrutura.sugestao, **mapa_resolvido}
            dlg = DialogoMapeamento(self, estrutura)
            self.wait_window(dlg)
            if dlg.mapeamento is None:
                return
            mapeamento_final = dlg.mapeamento

        try:
            transacoes = extrair_transacoes(estrutura, mapeamento_final)
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
        self.mapeamento_planilha = mapeamento_final
        self.transacoes_planilha = transacoes
        # Persiste o mapeamento da empresa (regrava sempre — re-sincroniza
        # nomes caso a planilha tenha mudado os rótulos)
        self._set_mapeamento_empresa(mapeamento_final, estrutura.cabecalho)
        self._atualiza_label_planilha()
        self.btn_editar_colunas.config(state="normal")
        self.btn_limpar_planilha.config(state="normal")
        self._atualiza_botao()
        self._render_aba_planilha()
        self._limpa_resultados()

        if usou_salvo_direto:
            emp = self._empresa_selecionada() or {}
            messagebox.showinfo(
                "Mapeamento aplicado",
                "Mapeamento salvo da empresa "
                f"'{emp.get('nome', emp.get('codi_emp', ''))}' aplicado "
                "automaticamente. Use 'Editar colunas' para revisar/alterar.",
            )

    def _importar_comprovantes_pdf(self) -> None:
        """Importa comprovantes de pagamento de boleto em PDF (nativos)
        de bancos suportados (Sicoob, Bradesco). Cada comprovante vira
        uma Transacao — vai pra 'planilha virtual' e entra no fluxo
        normal de conciliação com OFX/Domínio."""
        caminhos = filedialog.askopenfilenames(
            title="Selecione um ou mais PDFs de comprovantes",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
        )
        if not caminhos:
            return

        # Import tardio pra não travar o startup se pdfplumber não estiver ok
        import parser_pdf

        # Progresso via label popup — bom pra PDFs grandes (500 pgs)
        progresso_win = tk.Toplevel(self)
        progresso_win.title("Processando PDFs...")
        progresso_win.geometry("500x100")
        progresso_win.transient(self)
        progresso_win.resizable(False, False)
        lbl = ttk.Label(
            progresso_win, text="Iniciando...",
            font=("TkDefaultFont", 10),
        )
        lbl.pack(padx=20, pady=20, fill="x")
        progresso_win.update()

        def _on_prog(atual: int, total: int, nome: str) -> None:
            lbl.config(text=f"[{atual}/{total}] Lendo {nome}...")
            progresso_win.update()

        try:
            transacoes, relatorio = parser_pdf.ler_comprovantes_pdfs(
                list(caminhos), progresso=_on_prog,
            )
        except Exception as e:
            progresso_win.destroy()
            messagebox.showerror("Erro ao ler PDFs", str(e))
            return
        finally:
            try:
                progresso_win.destroy()
            except tk.TclError:
                pass

        if not transacoes:
            resumo = "\n".join(
                f"• {arq}: {banco} — {n} comprovante(s)"
                for arq, (banco, n) in relatorio.items()
            )
            messagebox.showwarning(
                "Nenhum comprovante extraído",
                f"Não consegui extrair nenhuma transação dos PDFs.\n\n"
                f"Relatório:\n{resumo}",
            )
            return

        # DEDUPLICAÇÃO: se já há transações na planilha (xlsx ou PDFs
        # anteriores), evita adicionar o mesmo lançamento duas vezes.
        # Cada PDF novo é comparado com as transações existentes por
        # (valor + data + CNPJ/nome). Se duplica, enriquece a existente
        # com dados que ela não tenha e ignora o novo.
        ja_existentes = list(self.transacoes_planilha)
        novos: list[Transacao] = []
        duplicatas: list[Transacao] = []  # os PDFs que foram ignorados
        for t_novo in transacoes:
            duplicata_de = None
            for t_exist in ja_existentes:
                if self._eh_mesma_transacao(t_novo, t_exist):
                    duplicata_de = t_exist
                    break
            if duplicata_de is not None:
                # Enriquece a existente com o que faltar (CNPJ/fornecedor)
                self._enriquecer_transacao_com(duplicata_de, t_novo)
                duplicatas.append(t_novo)
            else:
                novos.append(t_novo)
                # Já entra na lista pra comparar contra os próximos
                # (evita adicionar 2 vezes o mesmo PDF em batch)
                ja_existentes.append(t_novo)

        # Acumula ao invés de substituir
        modo_acumulado = bool(self.transacoes_planilha)
        self.transacoes_planilha = list(self.transacoes_planilha) + novos

        n_arqs = len(caminhos)
        n_extraidos = len(transacoes)
        n_novos = len(novos)
        n_duplicados = len(duplicatas)
        n_total = len(self.transacoes_planilha)

        # Só atualiza estrutura/mapeamento/caminho se está começando do zero.
        # Se está acumulando sobre uma planilha xlsx existente, mantém tudo.
        if not modo_acumulado:
            self.caminho_planilha = Path(caminhos[0])
            self.estrutura_planilha = None
            self.mapeamento_planilha = None
            self.btn_editar_colunas.config(state="disabled")

        # Label
        bancos = sorted({b for _, (b, n) in relatorio.items() if n > 0})
        if modo_acumulado:
            self.lbl_planilha.config(
                text=(
                    f"{n_total} lançamentos "
                    f"(+{n_novos} de {n_arqs} PDF novo{'s' if n_arqs > 1 else ''}"
                    f"{f', {n_duplicados} duplicado(s) ignorado(s)' if n_duplicados else ''})"
                )
            )
        else:
            prefixo = (
                f"{n_arqs} PDF{'s' if n_arqs > 1 else ''}"
                f" ({', '.join(bancos)})"
                if bancos else f"{n_arqs} PDF(s)"
            )
            self.lbl_planilha.config(
                text=f"{prefixo} — {n_novos} comprovante(s) importado(s)"
            )

        self.btn_limpar_planilha.config(state="normal")
        self._atualiza_botao()
        self._render_aba_planilha()
        self._limpa_resultados()

        # Relatório resumido
        resumo = "\n".join(
            f"• {arq}: {banco} — {n} comprovante(s)"
            for arq, (banco, n) in relatorio.items()
        )
        msg = f"Total extraído: {n_extraidos} de {n_arqs} PDF(s).\n\n"
        if n_duplicados > 0:
            msg += (
                f"• {n_novos} novo(s) comprovante(s) adicionado(s)\n"
                f"• {n_duplicados} DUPLICADO(S) ignorado(s) "
                "(já estavam na planilha; dados enriquecidos onde faltava)\n\n"
            )
        else:
            msg += f"• {n_novos} novo(s) comprovante(s) adicionado(s)\n\n"
        msg += f"Detalhes:\n{resumo}\n\n"
        if modo_acumulado:
            msg += (
                f"Planilha agora tem {n_total} lançamentos no total.\n\n"
            )
        msg += (
            "Os comprovantes aparecem na aba 'Planilha' e podem ser "
            "conciliados com o OFX e comparados com o Domínio."
        )
        messagebox.showinfo("Comprovantes importados", msg)

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
        # Re-salva o mapeamento da empresa com a versão editada
        self._set_mapeamento_empresa(dlg.mapeamento, estrutura.cabecalho)
        self._atualiza_label_planilha()
        self._atualiza_botao()
        self._render_aba_planilha()
        self._limpa_resultados()

    def _atualiza_label_planilha(self) -> None:
        if not self.caminho_planilha or self.mapeamento_planilha is None:
            return
        cab = self.estrutura_planilha.cabecalho if self.estrutura_planilha else []
        partes = []
        for campo, rotulo in CAMPOS:
            idx = self.mapeamento_planilha.get(campo)
            if idx is None:
                if campo in CAMPOS_OPCIONAIS:
                    continue  # opcional ausente — nem mostra
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
        caminhos = filedialog.askopenfilenames(
            title="Selecione um ou mais OFX (Ctrl+clique pra vários)",
            filetypes=[("OFX", "*.ofx"), ("Todos", "*.*")],
        )
        if not caminhos:
            return

        self.transacoes_ofx = []
        self.caminhos_ofx = []
        total_ignorados = 0
        erros: list[str] = []

        for caminho in caminhos:
            try:
                txs, ignorados = ler_ofx(caminho)
            except Exception as e:
                erros.append(f"{Path(caminho).name}: {e}")
                continue
            self.transacoes_ofx.extend(txs)
            self.caminhos_ofx.append(Path(caminho))
            total_ignorados += ignorados

        if erros:
            messagebox.showerror(
                "Erro ao ler um ou mais OFX",
                "\n".join(erros),
            )

        if not self.transacoes_ofx:
            self.lbl_ofx.config(text="(nenhum OFX carregado)")
            self._atualiza_botao()
            return

        n_arquivos = len(self.caminhos_ofx)
        bancos = sorted({t.extras.get("banco", "") for t in self.transacoes_ofx if t.extras.get("banco")})
        n_bancos = len(bancos)
        extra_ign = f" ({total_ignorados} recebimentos ignorados)" if total_ignorados else ""
        if n_arquivos == 1:
            prefixo = self.caminhos_ofx[0].name
        else:
            prefixo = f"{n_arquivos} arquivos OFX"
            if n_bancos > 0:
                prefixo += f" ({n_bancos} banco{'s' if n_bancos > 1 else ''})"
        self.lbl_ofx.config(
            text=f"{prefixo} — {len(self.transacoes_ofx)} pagamentos{extra_ign}"
        )
        self.btn_limpar_ofx.config(state="normal")
        self._atualiza_botao()
        self._render_aba_ofx()
        self._limpa_resultados()

    def _limpar_planilha(self) -> None:
        """Remove a planilha importada. Pendentes e sugestões são
        descartados, MAS os pares já conciliados, os matches com Domínio
        e os lançamentos contábeis já gerados permanecem intactos.

        Motivo: se você limpar a planilha e importar outra, os pares já
        conciliados com a planilha anterior NÃO vão ser refeitos — só o
        que ainda não foi conciliado da nova planilha vai ser processado.
        """
        if not self.transacoes_planilha and not self.caminho_planilha:
            return
        n_lanc = len(self.lancamentos_contabeis)
        n_pares = len(self.pares_conciliados)
        msg = (
            "Limpar a planilha importada?\n\n"
            "Pendentes e sugestões serão descartados. O OFX e o Domínio "
            "continuam carregados."
        )
        detalhes = []
        if n_pares:
            detalhes.append(
                f"{n_pares} conciliação(ões) já feita(s) (aba Conciliados "
                "e Conciliados × Domínio)"
            )
        if n_lanc:
            detalhes.append(f"{n_lanc} lançamento(s) contábil(is)")
        if detalhes:
            msg += (
                "\n\nSerão PRESERVADOS:\n• "
                + "\n• ".join(detalhes)
            )
        if not messagebox.askyesno("Confirmar", msg):
            return
        self.transacoes_planilha = []
        self.caminho_planilha = None
        self.estrutura_planilha = None
        self.mapeamento_planilha = None
        self.lbl_planilha.config(text="(nenhuma planilha carregada)")
        self.btn_editar_colunas.config(state="disabled")
        self.btn_limpar_planilha.config(state="disabled")
        self._atualiza_botao()
        self._render_aba_planilha()
        # Preserva pendentes do OFX que ainda estão carregados (não deve
        # apagar pendentes OFX só porque a planilha foi limpa).
        self._limpa_resultados(
            preservar_lancamentos=True,
            preservar_pendentes_ofx=True,
        )

    def _limpar_ofx(self) -> None:
        """Remove o(s) OFX importado(s). Pendentes e sugestões são
        descartados, MAS os pares já conciliados, os matches com Domínio
        e os lançamentos contábeis já gerados permanecem intactos.

        Motivo: se você limpar OFX-A e importar OFX-B, as linhas da
        planilha que já casaram com OFX-A NÃO vão ser reconciliadas com
        OFX-B — só as linhas ainda pendentes do OFX-B vão tentar casar
        com as planilhas que ainda estão sem par.
        """
        if not self.transacoes_ofx and not self.caminhos_ofx:
            return
        n_lanc = len(self.lancamentos_contabeis)
        n_pares = len(self.pares_conciliados)
        msg = (
            "Limpar o(s) OFX importado(s)?\n\n"
            "Pendentes e sugestões serão descartados. A planilha e o "
            "Domínio continuam carregados."
        )
        detalhes = []
        if n_pares:
            detalhes.append(
                f"{n_pares} conciliação(ões) já feita(s) (aba Conciliados "
                "e Conciliados × Domínio)"
            )
        if n_lanc:
            detalhes.append(f"{n_lanc} lançamento(s) contábil(is)")
        if detalhes:
            msg += (
                "\n\nSerão PRESERVADOS:\n• "
                + "\n• ".join(detalhes)
            )
        if not messagebox.askyesno("Confirmar", msg):
            return
        self.transacoes_ofx = []
        self.caminhos_ofx = []
        self.lbl_ofx.config(text="(nenhum OFX carregado)")
        self.btn_limpar_ofx.config(state="disabled")
        self._atualiza_botao()
        self._render_aba_ofx()
        # Preserva pendentes da planilha ainda carregada (não deve apagar
        # pendentes da planilha só porque o OFX foi limpo).
        self._limpa_resultados(
            preservar_lancamentos=True,
            preservar_pendentes_planilha=True,
        )

    def _atualiza_botao(self) -> None:
        # Basta ter planilha OU OFX carregado. Sem planilha, o fluxo
        # é comparar direto OFX × Domínio (todos os OFX ficam em pendentes_ofx).
        # Sem OFX, é planilha × Domínio (tudo vira Caixa geral).
        pode = bool(self.transacoes_planilha or self.transacoes_ofx)
        self.btn_conciliar.config(state="normal" if pode else "disabled")

    # ---------------------------------------------------- Lógica de matching

    def _limpa_resultados(
        self,
        preservar_lancamentos: bool = False,
        preservar_pendentes_planilha: bool = False,
        preservar_pendentes_ofx: bool = False,
    ) -> None:
        """Zera o estado de conciliação, com preservação seletiva.

        Parâmetros:
        - ``preservar_lancamentos``: mantém pares conciliados, matches com
          Domínio, lançamentos contábeis e manuais. Usado ao limpar
          planilha ou OFX (mas NÃO ao trocar de empresa).
        - ``preservar_pendentes_planilha``: mantém pendentes da planilha.
          Usado quando o OFX foi limpo — as planilhas ainda pendentes
          devem continuar visíveis na aba Pendentes.
        - ``preservar_pendentes_ofx``: mantém pendentes do OFX. Usado
          quando a planilha foi limpa — os OFX ainda pendentes devem
          continuar visíveis na aba Pendentes.

        Sugestões são sempre zeradas — são recalculadas no próximo
        Conciliar a partir dos pendentes atualizados.
        """
        if not preservar_lancamentos:
            # Zera TUDO (comportamento antigo — usado em troca de empresa
            # e outras situações que descartam o estado inteiro).
            self.pares_conciliados = []
            self.pendentes_planilha_dominio = {}
            self.pendentes_ofx_dominio = {}
            self.lancamentos_ignorados = set()
            self.lancamentos_contabeis = []
            self.lancamentos_manuais = []
            self.ids_pares_classificados = set()

        # Pendentes: preserva o lado oposto do que foi limpo.
        if not preservar_pendentes_planilha:
            self.pendentes_planilha_brutos = []
        if not preservar_pendentes_ofx:
            self.pendentes_ofx_brutos = []
        # Visíveis serão re-derivados (dos brutos, se houver)
        self.pendentes_planilha = list(self.pendentes_planilha_brutos)
        self.pendentes_ofx = list(self.pendentes_ofx_brutos)
        self.sugestoes = []

        # Se preservando algo, regera lançamentos automáticos + deriva
        # pendentes visíveis (removendo os que viraram lançamento).
        if preservar_lancamentos and (
            self.pendentes_planilha_brutos or self.pendentes_ofx_brutos
            or self.pares_conciliados or self.lancamentos_manuais
        ):
            self._gerar_lancamentos_contabeis()

        self._redesenha_abas()
        if hasattr(self, "tree_lancamentos"):
            self._render_aba_lancamentos()
        self.lbl_resumo.config(text="")

        if not preservar_lancamentos:
            for item in self.tree_dominio.get_children():
                self.tree_dominio.delete(item)
            self.notebook.tab(7, text="Comparação (0)")
            self._atualiza_botao_comparar()

    def _executar_conciliacao(self) -> None:
        # PRESERVA pares já conciliados (de rodadas anteriores, se o
        # usuário limpou planilha/OFX e importou outros). Isso evita
        # dupla conciliação: uma linha da planilha que já casou com
        # OFX-A NÃO vai casar de novo com OFX-B ao importar o novo OFX.
        ids_planilha_ja_pareada = {id(p.planilha) for p in self.pares_conciliados}
        ids_ofx_ja_pareado = {id(p.ofx) for p in self.pares_conciliados}

        planilha_pra_conciliar = [
            t for t in self.transacoes_planilha
            if id(t) not in ids_planilha_ja_pareada
        ]
        ofx_pra_conciliar = [
            t for t in self.transacoes_ofx
            if id(t) not in ids_ofx_ja_pareado
        ]

        novos_pares, pend_p, pend_o = conciliar_automatico(
            planilha_pra_conciliar, ofx_pra_conciliar,
        )
        # Adiciona os NOVOS pares aos existentes (preservados)
        self.pares_conciliados.extend(novos_pares)
        # Brutos são a fonte da verdade; visível é derivado depois.
        self.pendentes_planilha_brutos = list(pend_p)
        self.pendentes_planilha = list(pend_p)
        self.pendentes_ofx_brutos = list(pend_o)
        self.pendentes_ofx = list(pend_o)
        # Enriquece as Transacoes do OFX com CNPJ/nome/nº doc vindos dos
        # comprovantes PDF (quando a Transacao da planilha foi importada
        # de PDF). Faz aqui pra que o match com Domínio já use esses dados.
        self._enriquecer_ofx_com_pdf()
        # Primeiro classifica taxas (remove de pendentes_ofx visível)
        self._gerar_lancamentos_contabeis()
        # Sugestões usam pendentes_ofx visível (sem os classificados)
        self._recalcula_sugestoes()
        # Segunda fase: triple-match com Domínio
        self._filtrar_conciliados_por_dominio()
        self._redesenha_abas()
        self._atualiza_resumo()
        self._atualiza_botao_comparar()

    def _enriquecer_ofx_com_pdf(self) -> None:
        """Copia beneficiário/CNPJ/nº documento dos comprovantes PDF para
        as Transacoes do OFX que casaram com eles.

        Quando um Par tem ``planilha.origem == "pdf"``, o lado do OFX passa
        a saber pra quem foi pago — algo que o extrato bancário sozinho
        não fornece. Isso melhora:
        - Aba OFX (dados crus): usuário vê o beneficiário sem precisar
          abrir o comprovante.
        - Aba Comparação: se depois o par não casar com Domínio, o CNPJ
          fica disponível pra regras de fornecedor.
        - Exportação Excel: dados mais completos.

        Chaves gerenciadas em ``extras`` do OFX:
        - Zeradas em toda Transacao OFX antes de reenriquecer (evita
          resquício de conciliação anterior com outra planilha).
        - Repopuladas nos pares onde planilha veio de PDF.
        """
        CHAVES_ENRIQ = ("fornecedor", "cnpj", "numero_nf", "historico",
                        "enriquecido_por_pdf")

        # Zera enriquecimento anterior em todas as Transacoes do OFX
        # (pares atuais + pendentes brutos)
        for t in self.transacoes_ofx:
            for chave in CHAVES_ENRIQ:
                t.extras.pop(chave, None)

        # Aplica enriquecimento nos pares onde planilha veio de PDF
        n_enriq = 0
        for par in self.pares_conciliados:
            if getattr(par.planilha, "origem", "") != "pdf":
                continue
            p_extras = par.planilha.extras
            o_extras = par.ofx.extras
            for chave in ("fornecedor", "cnpj", "numero_nf", "historico"):
                valor = p_extras.get(chave, "")
                if valor:
                    o_extras[chave] = valor
            o_extras["enriquecido_por_pdf"] = True
            n_enriq += 1
        # Guarda quantos foram enriquecidos pra debug/relatório
        self._n_ofx_enriquecidos = n_enriq

    def _recalcula_sugestoes(self) -> None:
        self.sugestoes = gerar_sugestoes(self.pendentes_planilha, self.pendentes_ofx)

    @staticmethod
    def _normaliza_nf(v) -> str:
        return str(v).strip() if v is not None else ""

    @staticmethod
    def _normaliza_cnpj(v) -> str:
        """Mantém só dígitos pra comparação robusta (ignora pontuação)."""
        return "".join(c for c in str(v or "") if c.isdigit())

    @staticmethod
    def _normaliza_nome_fornecedor(v) -> str:
        """Normaliza pra comparação de fornecedor: casefold, sem acentos,
        sem pontuação, sem sufixos societários (LTDA/ME/EPP/S/A etc).

        Ex: 'Comercial Oeste Ltda ME' → 'comercial oeste'
            'COMERCIAL OESTE LTDA-ME' → 'comercial oeste'
        Assim os dois casam por substring.
        """
        import re
        import unicodedata
        if not v:
            return ""
        s = str(v).strip().casefold()
        # Remove acentos
        s = "".join(
            c for c in unicodedata.normalize("NFD", s)
            if unicodedata.category(c) != "Mn"
        )
        # Substitui pontuação por espaço
        s = re.sub(r"[^\w\s]", " ", s)
        # Colapsa espaços
        s = re.sub(r"\s+", " ", s).strip()
        # Remove sufixos societários no fim (iterativo — remove combinações
        # tipo "ltda me" ou "eireli epp")
        sufixos = (
            "ltda me", "eireli epp", "eireli", "epp", "me", "mei",
            "s a", "s/a", "sa", "ltda", "s c ltda",
        )
        while True:
            alterou = False
            for suf in sufixos:
                if s.endswith(" " + suf):
                    s = s[: -(len(suf) + 1)].strip()
                    alterou = True
                elif s == suf:
                    s = ""
                    alterou = True
            if not alterou:
                break
        return s

    @staticmethod
    def _nomes_batem(nome_a: str, nome_b: str) -> bool:
        """True se um nome contém o outro (substring, já normalizados).
        Requer pelo menos 4 caracteres pra evitar falso positivo em nomes
        curtos (ex: 'PB' bate com 'PB LTDA')."""
        if not nome_a or not nome_b:
            return False
        if len(nome_a) < 4 or len(nome_b) < 4:
            return nome_a == nome_b
        return nome_a in nome_b or nome_b in nome_a

    @classmethod
    def _eh_mesma_transacao(cls, t_a, t_b) -> bool:
        """Decide se duas Transacoes representam o MESMO lançamento
        (útil pra deduplicar planilha × comprovantes PDF).

        Critério (todos devem passar):
        1) VALOR igual (quantizado a 2 casas).
        2) DATA igual em algum dos pares (vencimento OU pagamento).
        3) IDENTIDADE: mesmo CNPJ (normalizado) OU nomes batem por
           substring normalizada.

        Se nenhum dos dois tiver CNPJ nem nome, NÃO considera duplicata
        (pra não colar coisas diferentes só porque valor+data coincidiram
        — ex: 2 tarifas iguais no mesmo dia).
        """
        from decimal import Decimal

        def _quant(v):
            return Decimal(str(v)).quantize(Decimal("0.01"))

        # 1) Valor
        if _quant(t_a.valor) != _quant(t_b.valor):
            return False

        # 2) Data (qualquer combinação vcto/pgto que bata)
        datas_a = {t_a.data, getattr(t_a, "data_pagamento", None)}
        datas_b = {t_b.data, getattr(t_b, "data_pagamento", None)}
        datas_a.discard(None)
        datas_b.discard(None)
        if not (datas_a & datas_b):
            return False

        # 3) Identidade — CNPJ ou nome
        cnpj_a = cls._normaliza_cnpj(t_a.extras.get("cnpj", ""))
        cnpj_b = cls._normaliza_cnpj(t_b.extras.get("cnpj", ""))
        # Se ambos tem CNPJ, só bate se forem iguais
        if cnpj_a and cnpj_b:
            if cnpj_a == cnpj_b:
                return True
            # CNPJs diferentes explícitos → não é duplicata
            return False

        nome_a = cls._normaliza_nome_fornecedor(
            t_a.extras.get("fornecedor", "")
        )
        nome_b = cls._normaliza_nome_fornecedor(
            t_b.extras.get("fornecedor", "")
        )
        if nome_a and nome_b:
            return cls._nomes_batem(nome_a, nome_b)

        # Um lado tem só CNPJ e o outro só nome (ou nenhum) — pra não
        # arriscar falso positivo, exige que pelo menos um valor bata
        # com o outro (o que não pode acontecer aqui) → não é duplicata
        return False

    @staticmethod
    def _enriquecer_transacao_com(alvo, doador) -> None:
        """Copia do ``doador`` pro ``alvo`` os campos que o alvo NÃO tem.
        Não sobrescreve valores existentes no alvo. Usado quando uma
        transação da planilha é dedupada contra um comprovante PDF —
        a linha da planilha ganha CNPJ/fornecedor/nº doc do PDF se
        estava faltando.
        """
        # Campos escalares que podem vir do doador
        for campo in ("fornecedor", "cnpj", "numero_nf", "historico"):
            if not alvo.extras.get(campo) and doador.extras.get(campo):
                alvo.extras[campo] = doador.extras[campo]
        # data_pagamento (se alvo não tem e doador tem)
        if (
            getattr(alvo, "data_pagamento", None) is None
            and getattr(doador, "data_pagamento", None) is not None
        ):
            alvo.data_pagamento = doador.data_pagamento
        # data_emissao (só em extras)
        if not alvo.extras.get("data_emissao") and doador.extras.get("data_emissao"):
            alvo.extras["data_emissao"] = doador.extras["data_emissao"]
        # Marca que foi enriquecida cruzando com PDF (útil pra debug/log)
        alvo.extras["dedup_enriquecida"] = True

    def _filtrar_conciliados_por_dominio(self) -> None:
        """Match com Domínio em TRÊS fases, aplicado a TRÊS fontes:
        - pares_conciliados (Planilha×OFX) — atualiza par.dominio
        - pendentes_planilha_brutos (sem OFX = Caixa geral) — atualiza
          self.pendentes_planilha_dominio[id(t)]
        - pendentes_ofx_brutos (sem planilha) — atualiza
          self.pendentes_ofx_dominio[id(t)] (comparação OFX×Domínio direta,
          quando o usuário não importou planilha)

        FASE 1 (exato): data_vencimento + valor + NF iguais.
        FASE 2 (aproximado): pelo menos 2 de 3 critérios (CNPJ, data_venc,
            valor) iguais. O critério restante pode ter diferença.
        FASE 3 (fornecedor + valor): valor bate exato E (CNPJ bate OU
            nome do fornecedor bate por substring normalizada). Data
            usada apenas como desempate quando há vários candidatos —
            NÃO precisa bater. Útil quando o Domínio tem a mesma parcela
            mas com vencimento renegociado/prorrogado.

        Ordem de prioridade (cada Transacao do Domínio só casa com 1 item):
        pares > pendentes planilha > pendentes OFX (nas 3 fases).
        """
        from collections import defaultdict
        from decimal import Decimal

        # Limpa associações anteriores
        for par in self.pares_conciliados:
            par.dominio = None
            par.diff_dias_dominio = 0
            par.diff_valor_dominio = Decimal("0")
        self.pendentes_planilha_dominio = {}
        self.pendentes_ofx_dominio = {}

        if not self.transacoes_dominio:
            return

        def _quant(v: Decimal) -> Decimal:
            return v.quantize(Decimal("0.01"))

        # ---------- FASE 1: match exato (data + valor + NF)
        indice: dict[tuple, list[Transacao]] = defaultdict(list)
        for t in self.transacoes_dominio:
            chave = (
                t.data,
                _quant(t.valor),
                self._normaliza_nf(t.extras.get("numero_nf", "")),
            )
            indice[chave].append(t)

        usados: set[int] = set()

        # Pares têm prioridade na FASE 1
        pares_sem_match: list[Par] = []
        for par in self.pares_conciliados:
            chave = (
                par.planilha.data,
                _quant(par.planilha.valor),
                self._normaliza_nf(par.planilha.extras.get("numero_nf", "")),
            )
            candidatos = [t for t in indice.get(chave, []) if id(t) not in usados]
            if candidatos:
                par.dominio = candidatos[0]
                usados.add(id(par.dominio))
            else:
                pares_sem_match.append(par)

        # Pendentes da planilha (Caixa geral) — FASE 1 nos restantes
        pendentes_sem_match: list[Transacao] = []
        for t_p in self.pendentes_planilha_brutos:
            chave = (
                t_p.data,
                _quant(t_p.valor),
                self._normaliza_nf(t_p.extras.get("numero_nf", "")),
            )
            candidatos = [t for t in indice.get(chave, []) if id(t) not in usados]
            if candidatos:
                self.pendentes_planilha_dominio[id(t_p)] = {
                    "dominio": candidatos[0],
                    "diff_dias": 0,
                    "diff_valor": Decimal("0"),
                }
                usados.add(id(candidatos[0]))
            else:
                pendentes_sem_match.append(t_p)

        # Pendentes do OFX (sem planilha correspondente) — FASE 1.
        # OFX geralmente não tem Nº NF, então o match exato aqui é
        # essencialmente (data, valor) com NF vazio dos dois lados.
        pendentes_ofx_sem_match: list[Transacao] = []
        for t_o in self.pendentes_ofx_brutos:
            chave = (
                t_o.data,
                _quant(t_o.valor),
                self._normaliza_nf(t_o.extras.get("numero_nf", "")),
            )
            candidatos = [t for t in indice.get(chave, []) if id(t) not in usados]
            if candidatos:
                self.pendentes_ofx_dominio[id(t_o)] = {
                    "dominio": candidatos[0],
                    "diff_dias": 0,
                    "diff_valor": Decimal("0"),
                }
                usados.add(id(candidatos[0]))
            else:
                pendentes_ofx_sem_match.append(t_o)

        # ---------- FASE 2: match aproximado (2 de 3 — CNPJ, data, valor)
        dominio_disponivel = [
            t for t in self.transacoes_dominio if id(t) not in usados
        ]

        def _melhor_match_dominio(
            cnpj_p_norm: str, data_p, valor_p,
        ) -> tuple[int | None, int, Decimal]:
            """Procura na lista dominio_disponivel o melhor match 2-de-3.
            Devolve (idx ou None, diff_dias, diff_valor)."""
            melhor_idx: int | None = None
            melhor_score: tuple | None = None
            melhor_d = 0
            melhor_v = Decimal("0")
            for i, t in enumerate(dominio_disponivel):
                cnpj_d = self._normaliza_cnpj(t.extras.get("cnpj", ""))
                valor_d = _quant(t.valor)
                matches = 0
                if cnpj_p_norm and cnpj_p_norm == cnpj_d:
                    matches += 1
                if data_p == t.data:
                    matches += 1
                if valor_p == valor_d:
                    matches += 1
                if matches < 2:
                    continue
                dd = abs((data_p - t.data).days)
                dv = abs(valor_p - valor_d)
                score = (-matches, dd, dv)
                if melhor_score is None or score < melhor_score:
                    melhor_score = score
                    melhor_idx = i
                    melhor_d = dd
                    melhor_v = dv
            return melhor_idx, melhor_d, melhor_v

        # Pares têm prioridade na FASE 2 também
        pares_sem_match_f3: list[Par] = []
        for par in pares_sem_match:
            cnpj_p = self._normaliza_cnpj(par.planilha.extras.get("cnpj", ""))
            idx, dd, dv = _melhor_match_dominio(
                cnpj_p, par.planilha.data, _quant(par.planilha.valor),
            )
            if idx is not None:
                t_dom = dominio_disponivel.pop(idx)
                par.dominio = t_dom
                par.diff_dias_dominio = dd
                par.diff_valor_dominio = dv
                usados.add(id(t_dom))
            else:
                pares_sem_match_f3.append(par)

        # Pendentes da planilha (Caixa geral) — FASE 2 no que sobrou
        pendentes_sem_match_f3: list[Transacao] = []
        for t_p in pendentes_sem_match:
            cnpj_p = self._normaliza_cnpj(t_p.extras.get("cnpj", ""))
            idx, dd, dv = _melhor_match_dominio(
                cnpj_p, t_p.data, _quant(t_p.valor),
            )
            if idx is not None:
                t_dom = dominio_disponivel.pop(idx)
                self.pendentes_planilha_dominio[id(t_p)] = {
                    "dominio": t_dom,
                    "diff_dias": dd,
                    "diff_valor": dv,
                }
                usados.add(id(t_dom))
            else:
                pendentes_sem_match_f3.append(t_p)

        # Pendentes do OFX — FASE 2 no que sobrou.
        # CNPJ do OFX raramente existe, então normalmente o match aqui é
        # 2-de-3 usando data + valor (o CNPJ empresa nem sempre bate).
        pendentes_ofx_sem_match_f3: list[Transacao] = []
        for t_o in pendentes_ofx_sem_match:
            cnpj_o = self._normaliza_cnpj(t_o.extras.get("cnpj", ""))
            idx, dd, dv = _melhor_match_dominio(
                cnpj_o, t_o.data, _quant(t_o.valor),
            )
            if idx is not None:
                t_dom = dominio_disponivel.pop(idx)
                self.pendentes_ofx_dominio[id(t_o)] = {
                    "dominio": t_dom,
                    "diff_dias": dd,
                    "diff_valor": dv,
                }
                usados.add(id(t_dom))
            else:
                pendentes_ofx_sem_match_f3.append(t_o)

        # ---------- FASE 3: valor exato + (CNPJ OU nome do fornecedor)
        # A data de vencimento vira apenas critério de desempate — quando
        # há vários candidatos válidos, prefere o de data mais próxima.
        # Útil pra casos onde a data no Domínio foi alterada (renegociação,
        # prorrogação de vencimento, etc).
        def _melhor_match_fase3(
            cnpj_p_norm: str, nome_p_norm: str, valor_p, data_p,
        ) -> tuple[int | None, int, Decimal]:
            """Fase 3: exige valor exato E (CNPJ bate OU nome bate).
            Data usada só pra desempate quando há múltiplos candidatos."""
            melhor_idx: int | None = None
            melhor_score: tuple | None = None
            melhor_d = 0
            for i, t in enumerate(dominio_disponivel):
                valor_d = _quant(t.valor)
                if valor_p != valor_d:
                    continue  # valor tem que bater exato
                cnpj_d = self._normaliza_cnpj(t.extras.get("cnpj", ""))
                nome_d = self._normaliza_nome_fornecedor(
                    t.extras.get("fornecedor", "")
                )
                bate_cnpj = bool(cnpj_p_norm) and cnpj_p_norm == cnpj_d
                bate_nome = self._nomes_batem(nome_p_norm, nome_d)
                if not (bate_cnpj or bate_nome):
                    continue
                dd = abs((data_p - t.data).days)
                # Prioridade: CNPJ > nome; menor diff de dias primeiro
                prioridade = 0 if bate_cnpj else 1
                score = (prioridade, dd)
                if melhor_score is None or score < melhor_score:
                    melhor_score = score
                    melhor_idx = i
                    melhor_d = dd
            return melhor_idx, melhor_d, Decimal("0")

        # Pares P×OFX — FASE 3
        for par in pares_sem_match_f3:
            cnpj_p = self._normaliza_cnpj(par.planilha.extras.get("cnpj", ""))
            nome_p = self._normaliza_nome_fornecedor(
                par.planilha.extras.get("fornecedor", "")
            )
            idx, dd, dv = _melhor_match_fase3(
                cnpj_p, nome_p, _quant(par.planilha.valor), par.planilha.data,
            )
            if idx is not None:
                t_dom = dominio_disponivel.pop(idx)
                par.dominio = t_dom
                par.diff_dias_dominio = dd
                par.diff_valor_dominio = dv
                usados.add(id(t_dom))

        # Pendentes planilha (Caixa geral) — FASE 3
        for t_p in pendentes_sem_match_f3:
            cnpj_p = self._normaliza_cnpj(t_p.extras.get("cnpj", ""))
            nome_p = self._normaliza_nome_fornecedor(
                t_p.extras.get("fornecedor", "")
            )
            idx, dd, dv = _melhor_match_fase3(
                cnpj_p, nome_p, _quant(t_p.valor), t_p.data,
            )
            if idx is not None:
                t_dom = dominio_disponivel.pop(idx)
                self.pendentes_planilha_dominio[id(t_p)] = {
                    "dominio": t_dom,
                    "diff_dias": dd,
                    "diff_valor": dv,
                }
                usados.add(id(t_dom))

        # Pendentes OFX — FASE 3 (pega quando OFX foi enriquecido por PDF
        # e ganhou fornecedor/CNPJ, mesmo que a data diverja do Domínio)
        for t_o in pendentes_ofx_sem_match_f3:
            cnpj_o = self._normaliza_cnpj(t_o.extras.get("cnpj", ""))
            nome_o = self._normaliza_nome_fornecedor(
                t_o.extras.get("fornecedor", "")
            )
            idx, dd, dv = _melhor_match_fase3(
                cnpj_o, nome_o, _quant(t_o.valor), t_o.data,
            )
            if idx is not None:
                t_dom = dominio_disponivel.pop(idx)
                self.pendentes_ofx_dominio[id(t_o)] = {
                    "dominio": t_dom,
                    "diff_dias": dd,
                    "diff_valor": dv,
                }
                usados.add(id(t_dom))

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

    def _selecionar_todas_sugestoes(self) -> None:
        """Marca todas as linhas atualmente exibidas na aba Sugestões."""
        todos = self.tree_sugestoes.get_children()
        if todos:
            self.tree_sugestoes.selection_set(todos)
            self.tree_sugestoes.focus(todos[0])

    def _aceitar_sugestao(self) -> None:
        """Aceita uma OU múltiplas sugestões. Pula automaticamente pares
        que envolvem transações já usadas em outra sugestão aceita
        (mesma transação não pode ser conciliada 2x)."""
        sels = self.tree_sugestoes.selection()
        if not sels:
            messagebox.showinfo(
                "Sem seleção",
                "Marque pelo menos uma sugestão (Ctrl+clique ou Ctrl+A pra todas).",
            )
            return

        # Captura os pares ANTES de aceitar (cada _aceitar_par recalcula sugestoes)
        pares = [self.itens_sugestoes[iid] for iid in sels if iid in self.itens_sugestoes]
        if not pares:
            return

        usados_p: set[int] = set()
        usados_o: set[int] = set()
        aceitos = 0
        puladas = 0
        for par in pares:
            if id(par.planilha) in usados_p or id(par.ofx) in usados_o:
                puladas += 1
                continue
            self._aceitar_par(par.planilha, par.ofx, par.diff_dias, par.diff_valor)
            usados_p.add(id(par.planilha))
            usados_o.add(id(par.ofx))
            aceitos += 1

        # Feedback no fim, só se houve algum conflito
        if puladas > 0:
            messagebox.showinfo(
                "Aceitação parcial",
                f"{aceitos} sugestão(ões) aceita(s).\n"
                f"{puladas} pulada(s) porque a planilha ou o OFX já tinha "
                f"sido usado em outra sugestão aceita.",
            )

    def _aceitar_par(
        self, t_p: Transacao, t_o: Transacao, d_dias: int, d_valor,
    ) -> None:
        novo = Par(
            planilha=t_p, ofx=t_o, tipo="manual",
            diff_dias=d_dias, diff_valor=d_valor,
        )
        self.pares_conciliados.append(novo)
        # Remove o pendente planilha da fonte (brutos); visível derivado depois
        if t_p in self.pendentes_planilha_brutos:
            self.pendentes_planilha_brutos.remove(t_p)
        if t_p in self.pendentes_planilha:
            self.pendentes_planilha.remove(t_p)
        # Remove o pendente OFX da fonte (brutos) e da visível
        if t_o in self.pendentes_ofx_brutos:
            self.pendentes_ofx_brutos.remove(t_o)
        # _gerar_lancamentos_contabeis recalcula pendentes_ofx a partir de brutos
        self._gerar_lancamentos_contabeis()
        self._recalcula_sugestoes()
        self._filtrar_conciliados_por_dominio()
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
        # Devolve os pendentes às fontes (brutos); visíveis serão recalculados
        self.pendentes_planilha_brutos.append(par.planilha)
        self.pendentes_ofx_brutos.append(par.ofx)
        self.pendentes_planilha_brutos.sort(key=lambda t: (t.data, t.valor))
        self.pendentes_ofx_brutos.sort(key=lambda t: (t.data, t.valor))
        self._gerar_lancamentos_contabeis()  # re-deriva pendentes_ofx visível
        self._recalcula_sugestoes()
        self._filtrar_conciliados_por_dominio()
        self._redesenha_abas()
        self._atualiza_resumo()

    # ----------------------------------------------- Render das tabelas

    def _redesenha_abas(self) -> None:
        self._render_conciliados()
        self._render_pendentes()
        self._render_sugestoes()
        self._render_aba_conciliados_dominio()
        # Também re-renderiza a aba OFX (dados crus) — assim o
        # enriquecimento por PDF (fornecedor/CNPJ nas colunas + tag azul)
        # aparece imediatamente após clicar em Conciliar.
        if hasattr(self, "tree_ofx"):
            self._render_aba_ofx()
        # Atualiza Comparação também (no-op se Domínio não carregado)
        self._recalcular_comparacao()
        self.notebook.tab(3, text=f"Conciliados ({len(self.pares_conciliados)})")
        self.notebook.tab(
            4, text=f"Pendentes ({len(self.pendentes_planilha)}/{len(self.pendentes_ofx)})",
        )
        self.notebook.tab(5, text=f"Sugestões ({len(self.sugestoes)})")

    def _render_conciliados(self) -> None:
        for item in self.tree_conciliados.get_children():
            self.tree_conciliados.delete(item)
        self.itens_pares.clear()
        for par in self.pares_conciliados:
            diff_txt = ""
            if par.diff_dias or par.diff_valor:
                diff_txt = f"Δ {par.diff_dias}d, R$ {par.diff_valor:.2f}"
            tipo_txt = "Auto" if par.tipo == "auto" else "Manual"
            emissao = par.planilha.extras.get("data_emissao")
            emissao_txt = emissao.strftime("%d/%m/%Y") if emissao else ""
            # Pagamento: prioriza data_pagamento da planilha; fallback = data do OFX
            pagto = par.planilha.data_pagamento or par.ofx.data
            pagto_txt = pagto.strftime("%d/%m/%Y") if pagto else ""
            origem = par.ofx.extras.get("banco", "") or "OFX"
            iid = self.tree_conciliados.insert(
                "", "end",
                values=(
                    tipo_txt,
                    origem,
                    par.planilha.data.strftime("%d/%m/%Y"),
                    pagto_txt,
                    f"{par.planilha.valor:.2f}",
                    emissao_txt,
                    par.planilha.extras.get("numero_nf", ""),
                    par.planilha.extras.get("cnpj", ""),
                    par.planilha.extras.get("fornecedor", ""),
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
                values=(
                    t.data.strftime("%d/%m/%Y"),
                    self._fmt_data(t.data_pagamento),
                    f"{t.valor:.2f}",
                    t.extras.get("numero_nf", ""),
                    t.extras.get("fornecedor", ""),
                    t.extras.get("historico", "") or "",
                    t.extras.get("tipo", "") or "",
                ),
            )
            self.itens_pendentes_p[iid] = t

        for item in self.tree_pend_o.get_children():
            self.tree_pend_o.delete(item)
        self.itens_pendentes_o.clear()
        for t in self.pendentes_ofx:
            enriquecido = bool(t.extras.get("enriquecido_por_pdf"))
            tags = ("enriquecido_pdf",) if enriquecido else ()
            iid = self.tree_pend_o.insert(
                "", "end",
                values=(
                    t.data.strftime("%d/%m/%Y"),
                    t.extras.get("banco", "") or "",
                    t.extras.get("documento", "") or "",
                    f"{t.valor:.2f}",
                    t.descricao,
                    t.extras.get("fornecedor", "") or "",
                    t.extras.get("cnpj", "") or "",
                ),
                tags=tags,
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
                    par.planilha.extras.get("numero_nf", ""),
                    par.planilha.extras.get("fornecedor", ""),
                    par.ofx.data.strftime("%d/%m/%Y"),
                    f"{par.ofx.valor:.2f}",
                    par.ofx.descricao,
                    str(par.diff_dias),
                    f"{par.diff_valor:.2f}",
                ),
                tags=("destaque",),
            )
            self.itens_sugestoes[iid] = par

    def _render_aba_conciliados_dominio(self) -> None:
        for item in self.tree_conciliados_dominio.get_children():
            self.tree_conciliados_dominio.delete(item)

        def _tag_status(status: str) -> str:
            sl = (status or "").lower()
            if sl.startswith("pag"):
                return "paga"
            if sl.startswith("parc"):
                return "parcial"
            if sl.startswith("ab"):
                return "aberto"
            return ""

        def _pega(*fontes, chave: str) -> str:
            """Devolve o primeiro valor não-vazio entre as fontes. Cada
            fonte é um dict de extras (ou None). Ordem = prioridade."""
            for f in fontes:
                if f is None:
                    continue
                v = f.get(chave, "")
                if v:
                    return v
            return ""

        # 1) Pares P×OFX triple-matched
        # PRIORIDADE de dados: Domínio > planilha/PDF > OFX
        # (Domínio é a fonte mais confiável — planilhas e comprovantes
        # podem ter CNPJ errado ou vazio)
        pares = [p for p in self.pares_conciliados if p.dominio is not None]
        for par in pares:
            tipo_txt = "Auto" if par.tipo == "auto" else "Manual"
            origem = par.ofx.extras.get("banco", "") or "OFX"
            pagto = par.planilha.data_pagamento or par.ofx.data
            pagto_txt = pagto.strftime("%d/%m/%Y") if pagto else ""

            dom_extras = par.dominio.extras if par.dominio else {}
            p_extras = par.planilha.extras
            o_extras = par.ofx.extras

            cnpj = _pega(dom_extras, p_extras, o_extras, chave="cnpj")
            fornecedor = _pega(dom_extras, p_extras, o_extras, chave="fornecedor")
            numero_nf = _pega(dom_extras, p_extras, chave="numero_nf")
            emissao_val = _pega(dom_extras, p_extras, chave="data_emissao")
            emissao_txt = (
                emissao_val.strftime("%d/%m/%Y")
                if hasattr(emissao_val, "strftime") else str(emissao_val or "")
            )

            status = dom_extras.get("status", "") or ""

            # Diferenças com o Domínio (fase 2/3 — match aproximado)
            diff_dom = ""
            if par.diff_dias_dominio or par.diff_valor_dominio:
                diff_dom = (
                    f"Δ {par.diff_dias_dominio}d, R$ {par.diff_valor_dominio:.2f}"
                )
            self.tree_conciliados_dominio.insert(
                "", "end",
                values=(
                    tipo_txt,
                    origem,
                    par.planilha.data.strftime("%d/%m/%Y"),
                    pagto_txt,
                    f"{par.planilha.valor:.2f}",
                    emissao_txt,
                    numero_nf,
                    cnpj,
                    fornecedor,
                    par.ofx.descricao,
                    diff_dom,
                    status,
                ),
                tags=(_tag_status(status),) if _tag_status(status) else (),
            )

        # 2) Pendentes da planilha (Caixa geral) que casaram com Domínio
        caixa_dominio = [
            t for t in self.pendentes_planilha_brutos
            if (m := self.pendentes_planilha_dominio.get(id(t)))
            and m.get("dominio") is not None
        ]
        for t_p in caixa_dominio:
            match = self.pendentes_planilha_dominio[id(t_p)]
            t_dom = match["dominio"]
            d_d = match.get("diff_dias", 0)
            d_v = match.get("diff_valor", 0)

            dom_extras = t_dom.extras if t_dom else {}
            p_extras = t_p.extras

            cnpj = _pega(dom_extras, p_extras, chave="cnpj")
            fornecedor = _pega(dom_extras, p_extras, chave="fornecedor")
            numero_nf = _pega(dom_extras, p_extras, chave="numero_nf")
            emissao_val = _pega(dom_extras, p_extras, chave="data_emissao")
            emissao_txt = (
                emissao_val.strftime("%d/%m/%Y")
                if hasattr(emissao_val, "strftime") else str(emissao_val or "")
            )

            pagto = t_p.data_pagamento or t_p.data
            pagto_txt = pagto.strftime("%d/%m/%Y") if pagto else ""

            status = dom_extras.get("status", "") or ""

            diff_dom = ""
            if d_d or d_v:
                diff_dom = f"Δ {d_d}d, R$ {d_v:.2f}"

            memo_txt = (
                f"Histórico: {t_p.extras.get('historico', '')}"
                if t_p.extras.get("historico") else "(sem OFX)"
            )

            self.tree_conciliados_dominio.insert(
                "", "end",
                values=(
                    "Caixa",                              # Tipo
                    "Caixa geral",                        # Origem
                    t_p.data.strftime("%d/%m/%Y"),
                    pagto_txt,
                    f"{t_p.valor:.2f}",
                    emissao_txt,
                    numero_nf,
                    cnpj,
                    fornecedor,
                    memo_txt,
                    diff_dom,
                    status,
                ),
                tags=(_tag_status(status),) if _tag_status(status) else (),
            )

        # 3) Pendentes do OFX (sem planilha) que casaram com Domínio
        ofx_dominio = [
            t for t in self.pendentes_ofx_brutos
            if (m := self.pendentes_ofx_dominio.get(id(t)))
            and m.get("dominio") is not None
        ]
        for t_o in ofx_dominio:
            match = self.pendentes_ofx_dominio[id(t_o)]
            t_dom = match["dominio"]
            d_d = match.get("diff_dias", 0)
            d_v = match.get("diff_valor", 0)

            dom_extras = t_dom.extras if t_dom else {}
            o_extras = t_o.extras

            # OFX pode estar enriquecido por PDF (fornecedor/CNPJ), mas o
            # Domínio ainda tem prioridade.
            cnpj = _pega(dom_extras, o_extras, chave="cnpj")
            fornecedor = _pega(dom_extras, o_extras, chave="fornecedor")
            numero_nf = _pega(dom_extras, o_extras, chave="numero_nf")
            emissao_val = dom_extras.get("data_emissao")
            emissao_txt = (
                emissao_val.strftime("%d/%m/%Y")
                if hasattr(emissao_val, "strftime") else str(emissao_val or "")
            )

            pagto_txt = t_o.data.strftime("%d/%m/%Y")
            origem = t_o.extras.get("banco", "") or "OFX"

            status = dom_extras.get("status", "") or ""

            diff_dom = ""
            if d_d or d_v:
                diff_dom = f"Δ {d_d}d, R$ {d_v:.2f}"

            self.tree_conciliados_dominio.insert(
                "", "end",
                values=(
                    "OFX",                                # Tipo
                    origem,                               # Origem = banco do OFX
                    t_o.data.strftime("%d/%m/%Y"),
                    pagto_txt,
                    f"{t_o.valor:.2f}",
                    emissao_txt,
                    numero_nf,
                    cnpj,
                    fornecedor,
                    t_o.descricao or "",
                    diff_dom,
                    status,
                ),
                tags=(_tag_status(status),) if _tag_status(status) else (),
            )

        total = len(pares) + len(caixa_dominio) + len(ofx_dominio)
        self.notebook.tab(6, text=f"Conciliados × Domínio ({total})")

    # ---------------- Abas de dados crus (origem) ----------------

    @staticmethod
    def _fmt_data(d) -> str:
        return d.strftime("%d/%m/%Y") if d else ""

    def _render_aba_planilha(self) -> None:
        for item in self.tree_planilha.get_children():
            self.tree_planilha.delete(item)
        if hasattr(self, "itens_tree_planilha"):
            self.itens_tree_planilha.clear()
        termo = self.filtro_planilha.get().strip().lower() if hasattr(self, "filtro_planilha") else ""
        cols = self.COLS_PLANILHA
        filtros = getattr(self, "filtros_col_planilha", {})
        tem_filtro_col = any(v is not None for v in filtros.values())
        mostradas = 0
        for t in self.transacoes_planilha:
            row = self._row_planilha(t)
            if termo and termo not in " ".join(row).lower():
                continue
            pula = False
            for i, col in enumerate(cols):
                permitidos = filtros.get(col)
                if permitidos is not None and str(row[i]) not in permitidos:
                    pula = True
                    break
            if pula:
                continue
            iid = self.tree_planilha.insert("", "end", values=row)
            if hasattr(self, "itens_tree_planilha"):
                self.itens_tree_planilha[iid] = t
            mostradas += 1
        total = len(self.transacoes_planilha)
        self.notebook.tab(0, text=f"Planilha ({total})")
        if hasattr(self, "lbl_filtro_planilha"):
            tem_filtro = termo or tem_filtro_col
            self.lbl_filtro_planilha.config(
                text=f"Mostrando {mostradas} de {total}" if tem_filtro else f"{total} lançamentos",
            )

    def _render_aba_ofx(self) -> None:
        for item in self.tree_ofx.get_children():
            self.tree_ofx.delete(item)
        termo = self.filtro_ofx.get().strip().lower() if hasattr(self, "filtro_ofx") else ""
        cols = self.COLS_OFX
        filtros = getattr(self, "filtros_col_ofx", {})
        tem_filtro_col = any(v is not None for v in filtros.values())
        mostradas = 0
        for t in self.transacoes_ofx:
            row = self._row_ofx(t)
            if termo and termo not in " ".join(row).lower():
                continue
            pula = False
            for i, col in enumerate(cols):
                permitidos = filtros.get(col)
                if permitidos is not None and str(row[i]) not in permitidos:
                    pula = True
                    break
            if pula:
                continue
            # Linhas enriquecidas por PDF ficam com fundo azul claro
            tags = ("enriquecido_pdf",) if t.extras.get("enriquecido_por_pdf") else ()
            self.tree_ofx.insert("", "end", values=row, tags=tags)
            mostradas += 1
        total = len(self.transacoes_ofx)
        n_enriq = sum(
            1 for t in self.transacoes_ofx
            if t.extras.get("enriquecido_por_pdf")
        )
        label_tab = f"OFX ({total})"
        if n_enriq:
            label_tab = f"OFX ({total} | {n_enriq} enriquecido por PDF)"
        self.notebook.tab(1, text=label_tab)
        if hasattr(self, "lbl_filtro_ofx"):
            tem_filtro = termo or tem_filtro_col
            self.lbl_filtro_ofx.config(
                text=f"Mostrando {mostradas} de {total}" if tem_filtro else f"{total} pagamentos",
            )

    def _render_aba_dominio_dados(self) -> None:
        for item in self.tree_dominio_dados.get_children():
            self.tree_dominio_dados.delete(item)
        termo = self.filtro_dominio.get().strip().lower() if hasattr(self, "filtro_dominio") else ""
        status_pedido = self.filtro_dominio_status.get() if hasattr(self, "filtro_dominio_status") else "Todos"
        cols = self.COLS_DOMINIO
        filtros = getattr(self, "filtros_col_dominio", {})
        tem_filtro_col = any(v is not None for v in filtros.values())
        mostradas = 0
        for t in self.transacoes_dominio:
            row = self._row_dominio(t)
            status = row[4]  # coluna "status"
            tag = ""
            sl = status.lower()
            if sl.startswith("pag"):
                tag = "paga"
            elif sl.startswith("parc"):
                tag = "parcial"
            elif sl.startswith("ab"):
                tag = "aberto"
            # Filtro de status (dropdown)
            if status_pedido != "Todos":
                if status_pedido == "Aberto" and tag != "aberto":
                    continue
                if status_pedido == "Parcial" and tag != "parcial":
                    continue
                if status_pedido == "Paga" and tag != "paga":
                    continue
            # Filtro global (texto)
            if termo and termo not in " ".join(row).lower():
                continue
            # Filtros por coluna (estilo Excel — set de valores)
            pula = False
            for i, col in enumerate(cols):
                permitidos = filtros.get(col)
                if permitidos is not None and str(row[i]) not in permitidos:
                    pula = True
                    break
            if pula:
                continue
            self.tree_dominio_dados.insert(
                "", "end", values=row, tags=(tag,) if tag else (),
            )
            mostradas += 1
        total = len(self.transacoes_dominio)
        # Detecta se veio de grupo empresarial (múltiplas empresas)
        empresas_unicas = {
            t.extras.get("codi_emp_origem")
            for t in self.transacoes_dominio
            if t.extras.get("codi_emp_origem") is not None
        }
        if len(empresas_unicas) > 1:
            self.notebook.tab(
                2, text=f"Domínio dados ({total} | {len(empresas_unicas)} empresas)",
            )
        else:
            self.notebook.tab(2, text=f"Domínio dados ({total})")
        if hasattr(self, "lbl_filtro_dominio"):
            tem_filtro = termo or status_pedido != "Todos" or tem_filtro_col
            self.lbl_filtro_dominio.config(
                text=f"Mostrando {mostradas} de {total}" if tem_filtro else f"{total} parcelas",
            )

    # ------------------------------------ Lançamentos contábeis ------------

    def _empresa_selecionada(self) -> dict | None:
        emp = self.cfg.get("dominio_empresa")
        if isinstance(emp, dict) and emp.get("codi_emp") is not None:
            return emp
        return None

    def _get_regras_empresa(self) -> list[dict]:
        """Devolve a lista de regras de taxas vinculadas à empresa atual.
        Lista vazia quando nenhuma empresa está selecionada."""
        emp = self._empresa_selecionada()
        if not emp:
            return []
        chave = str(emp["codi_emp"])
        return list(self.cfg.get("regras_taxas_por_empresa", {}).get(chave, []))

    def _set_regras_empresa(self, regras: list[dict]) -> None:
        emp = self._empresa_selecionada()
        if not emp:
            return
        chave = str(emp["codi_emp"])
        self.cfg.setdefault("regras_taxas_por_empresa", {})[chave] = regras
        config.salvar(self.cfg)

    def _exigir_empresa(self, acao: str) -> bool:
        """Mostra aviso e retorna False se não há empresa selecionada."""
        if self._empresa_selecionada():
            return True
        messagebox.showwarning(
            "Selecione uma empresa",
            f"Selecione a empresa do Domínio antes de {acao}.\n\n"
            "As regras de taxas são salvas por empresa (cada empresa tem seu "
            "próprio conjunto de regras).",
        )
        return False

    # ----- Mapeamento da planilha persistido por empresa -----
    # Salvo POR NOME DE COLUNA pra resistir a reordenação. Cada empresa tem
    # o seu próprio dict {campo: nome_da_coluna}. Sem empresa selecionada,
    # nada é salvo — usuário precisa mapear manualmente.

    @staticmethod
    def _normaliza_nome_coluna(nome: object) -> str:
        if nome is None:
            return ""
        return str(nome).strip().casefold()

    def _get_mapeamento_empresa(self) -> dict[str, str]:
        """Devolve o mapeamento salvo (campo → nome de coluna) da empresa
        atual. Dict vazio quando não há empresa ou não há mapeamento salvo."""
        emp = self._empresa_selecionada()
        if not emp:
            return {}
        chave = str(emp["codi_emp"])
        salvo = self.cfg.get("mapeamentos_planilha_por_empresa", {}).get(chave)
        return dict(salvo) if isinstance(salvo, dict) else {}

    def _set_mapeamento_empresa(
        self, mapa_idx: dict[str, int], cabecalho: list,
    ) -> None:
        """Salva o mapeamento da empresa atual convertendo idx → nome da
        coluna. Sem empresa selecionada, vira no-op (não persiste)."""
        emp = self._empresa_selecionada()
        if not emp:
            return
        chave = str(emp["codi_emp"])
        por_nome: dict[str, str] = {}
        for campo, idx in mapa_idx.items():
            if 0 <= idx < len(cabecalho):
                nome = cabecalho[idx]
                if nome:
                    por_nome[campo] = str(nome).strip()
        if not por_nome:
            return
        self.cfg.setdefault("mapeamentos_planilha_por_empresa", {})[chave] = por_nome
        config.salvar(self.cfg)

    def _resolver_mapeamento_salvo(
        self, cabecalho: list,
    ) -> tuple[dict[str, int], list[str]]:
        """Traduz o mapeamento salvo (nome → idx) usando o cabeçalho atual.
        Devolve (mapa_idx_resolvido, campos_faltando). Campos cujo nome não
        bate exatamente ficam fora do dict e entram na lista de faltando."""
        salvo = self._get_mapeamento_empresa()
        if not salvo:
            # Sem mapeamento salvo, todos os obrigatórios "faltam"
            return {}, [c for c, _ in CAMPOS if c not in CAMPOS_OPCIONAIS]
        # Índice: nome normalizado → idx (primeira ocorrência ganha)
        idx_por_nome: dict[str, int] = {}
        for i, nome in enumerate(cabecalho):
            chave = self._normaliza_nome_coluna(nome)
            if chave and chave not in idx_por_nome:
                idx_por_nome[chave] = i
        resolvido: dict[str, int] = {}
        faltando: list[str] = []
        usados: set[int] = set()
        for campo, _ in CAMPOS:
            nome_salvo = salvo.get(campo)
            if not nome_salvo:
                # Opcional sem mapeamento salvo é OK — não conta como faltando
                if campo not in CAMPOS_OPCIONAIS:
                    faltando.append(campo)
                continue
            idx = idx_por_nome.get(self._normaliza_nome_coluna(nome_salvo))
            if idx is None or idx in usados:
                # Opcional que estava salvo mas não bate com cabeçalho atual:
                # ignora silenciosamente (planilha pode não ter mais essa col)
                if campo not in CAMPOS_OPCIONAIS:
                    faltando.append(campo)
                continue
            resolvido[campo] = idx
            usados.add(idx)
        return resolvido, faltando

    def _abrir_config_taxas(self) -> None:
        if not self._exigir_empresa("configurar regras de taxas"):
            return
        emp = self._empresa_selecionada()
        regras = self._get_regras_empresa()

        def _on_change(novas: list[dict]) -> None:
            self._set_regras_empresa(novas)
            self._gerar_lancamentos_contabeis()
            self._redesenha_abas()

        dlg = DialogoConfigurarTaxas(
            self, regras, _on_change, plano_contas=self.plano_contas,
        )
        dlg.title(
            f"Configurar regras — {emp['razao'][:60]} "
            f"(empresa {emp['codi_emp']})"
        )
        self.wait_window(dlg)

    def _lancamento_manual_pend_planilha(self) -> None:
        """Lança um lançamento contábil avulso a partir de UM pendente da
        planilha selecionado. Não cria regra. A transação some dos Pendentes
        (vai pra aba Lançamentos contábeis)."""
        sel = self.tree_pend_p.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um lançamento na lista da PLANILHA (lado esquerdo) "
                "para fazer o lançamento manual.",
            )
            return
        t = self.itens_pendentes_p[sel[0]]
        forn = t.extras.get("fornecedor", "") or ""
        sugestao = f"Pagto. {forn}" if forn else ""
        dlg = DialogoLancamentoManualAvulso(
            self, t, origem="planilha",
            plano_contas=self.plano_contas,
            sugestao_historico=sugestao,
        )
        self.wait_window(dlg)
        if not dlg.resultado:
            return
        lanc = LancamentoContabil(
            data=t.data_pagamento or t.data,  # prioriza pagamento se mapeado
            historico=dlg.resultado["historico"],
            valor=t.valor,
            # Sem OFX correspondente — é dinheiro/caixa, não passou pelo banco
            banco="Caixa geral",
            memo_original="",
            padrao_match="(manual planilha)",
            conta=dlg.resultado["conta"],
            tipo_regra="manual_planilha",
            fornecedor=forn,
            cnpj=t.extras.get("cnpj", "") or "",
            transacao_origem=t,
            par_origem=None,
        )
        self.lancamentos_manuais.append(lanc)
        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()
        self._atualiza_resumo()

    def _lancamento_manual_pend_ofx(self) -> None:
        """Lança um lançamento contábil avulso a partir de UM pendente do OFX
        selecionado. Não cria regra. A transação some dos Pendentes (vai pra
        aba Lançamentos contábeis)."""
        sel = self.tree_pend_o.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um lançamento na lista do OFX (lado direito) "
                "para fazer o lançamento manual.",
            )
            return
        t = self.itens_pendentes_o[sel[0]]
        memo = (t.descricao or "").strip()
        # Dados enriquecidos por PDF (se houver)
        fornecedor_pdf = (t.extras.get("fornecedor", "") or "").strip()
        cnpj_pdf = (t.extras.get("cnpj", "") or "").strip()

        # Se OFX foi enriquecido por PDF, sugere histórico com fornecedor
        if fornecedor_pdf:
            sugestao = f"PAGAMENTO REF. A {fornecedor_pdf}"
        else:
            sugestao = memo[:60] if memo else ""
        dlg = DialogoLancamentoManualAvulso(
            self, t, origem="ofx",
            plano_contas=self.plano_contas,
            sugestao_historico=sugestao,
        )
        self.wait_window(dlg)
        if not dlg.resultado:
            return
        lanc = LancamentoContabil(
            data=t.data,
            historico=dlg.resultado["historico"],
            valor=t.valor,
            banco=t.extras.get("banco", "") or "",
            memo_original=t.descricao or "",
            padrao_match="(manual OFX)",
            conta=dlg.resultado["conta"],
            tipo_regra="manual_ofx",
            fornecedor=fornecedor_pdf,
            cnpj=cnpj_pdf,
            transacao_origem=t,
            par_origem=None,
        )
        self.lancamentos_manuais.append(lanc)
        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()
        self._atualiza_resumo()

    def _criar_lancamento_padrao_planilha(self) -> None:
        """Atalho: cria uma regra do tipo 'fornecedor' a partir do pendente
        da planilha selecionado. A regra é salva imediatamente nas regras
        da empresa atual e vai gerar lançamento sempre que o CNPJ/nome
        aparecer."""
        if not self._exigir_empresa("criar regras de taxas"):
            return
        sel = self.tree_pend_p.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um lançamento na lista da PLANILHA (lado esquerdo) "
                "para criar uma regra a partir dele.",
            )
            return
        t = self.itens_pendentes_p[sel[0]]
        cnpj = (t.extras.get("cnpj", "") or "").strip()
        fornecedor = (t.extras.get("fornecedor", "") or "").strip()
        historico = (t.extras.get("historico", "") or "").strip()
        tipo_col = (t.extras.get("tipo", "") or "").strip()
        # PRIORIZA Tipo — regra por tipo captura categorias inteiras
        # (ex.: "CONDOMÍNIO" pega todos os condomínios de qualquer fornecedor)
        sugestao = tipo_col or cnpj or fornecedor or historico
        if not sugestao:
            messagebox.showwarning(
                "Sem dados",
                "O lançamento selecionado não tem Tipo, CNPJ, fornecedor "
                "nem histórico — não dá pra gerar um padrão automático.",
            )
            return

        # Sugestão do histórico contábil da regra: prefixo fixo "PAGAMENTO
        # REF. A " + o que estiver na coluna Histórico da planilha.
        # Se histórico vazio, deixa em branco pra que o próprio diálogo
        # exiba a dica "vazio = usa o Histórico da planilha" e o operador
        # decida se quer texto fixo ou fallback dinâmico linha-a-linha.
        if historico:
            sugestao_hist = f"PAGAMENTO REF. A {historico}"
        else:
            sugestao_hist = ""
        regra_inicial = {
            "padrao": sugestao,
            "historico": sugestao_hist,
            "tipo": "fornecedor",
        }
        dlg = DialogoNovaRegra(
            self, regra_inicial, plano_contas=self.plano_contas,
        )
        emp = self._empresa_selecionada()
        dlg.title(f"Criar regra (fornecedor) — {emp['razao'][:50]} (empresa {emp['codi_emp']})")
        self.wait_window(dlg)
        if not dlg.regra:
            return

        # Garante que a regra seja tipo "fornecedor" mesmo que o diálogo não
        # exponha esse campo (compat com diálogos antigos)
        dlg.regra.setdefault("tipo", "fornecedor")
        if dlg.regra.get("tipo") != "fornecedor":
            dlg.regra["tipo"] = "fornecedor"

        regras = self._get_regras_empresa()
        regras.append(dlg.regra)
        self._set_regras_empresa(regras)
        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()
        self._atualiza_resumo()

    def _criar_lancamento_padrao(self) -> None:
        """Atalho: cria uma regra de taxa a partir do pendente OFX
        selecionado e a salva imediatamente nas regras da empresa atual."""
        if not self._exigir_empresa("criar regras de taxas"):
            return
        sel = self.tree_pend_o.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um lançamento na lista do OFX (lado direito) "
                "para criar uma regra a partir dele.",
            )
            return
        transacao = self.itens_pendentes_o[sel[0]]
        memo = transacao.descricao or ""
        documento = transacao.extras.get("documento", "") or ""
        # Dados enriquecidos por PDF (se houver) — mais confiáveis que memo
        fornecedor_pdf = (transacao.extras.get("fornecedor", "") or "").strip()
        cnpj_pdf = (transacao.extras.get("cnpj", "") or "").strip()

        # Ordem de preferência do padrão:
        # 1) Fornecedor enriquecido (nome real, ex: "COMERCIAL OESTE LTDA")
        # 2) CNPJ enriquecido (identificador único)
        # 3) Memo do OFX
        # 4) Documento do OFX
        sugestao = fornecedor_pdf or cnpj_pdf or memo.strip() or documento.strip()
        if not sugestao:
            messagebox.showwarning(
                "Sem dados",
                "O lançamento selecionado não tem fornecedor, memo nem "
                "documento — não dá pra gerar um padrão automático.",
            )
            return

        # Histórico contábil sugerido: se tem fornecedor enriquecido,
        # usa "PAGAMENTO REF. A <fornecedor>"; senão vazio.
        sugestao_hist = ""
        if fornecedor_pdf:
            sugestao_hist = f"PAGAMENTO REF. A {fornecedor_pdf}"

        # Pré-popula o banco: regra só vale pra esse banco específico, evita
        # falso positivo quando memos parecidos vêm de bancos diferentes.
        banco_origem = (transacao.extras.get("banco", "") or "").strip()
        regra_inicial = {
            "padrao": sugestao,
            "historico": sugestao_hist,
            "banco": banco_origem,
        }
        dlg = DialogoNovaRegra(
            self, regra_inicial, plano_contas=self.plano_contas,
        )
        emp = self._empresa_selecionada()
        dlg.title(f"Criar regra — {emp['razao'][:60]} (empresa {emp['codi_emp']})")
        self.wait_window(dlg)
        if not dlg.regra:
            return

        # Salva a nova regra e refaz os lançamentos contábeis
        regras = self._get_regras_empresa()
        regras.append(dlg.regra)
        self._set_regras_empresa(regras)
        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()
        self._atualiza_resumo()

    def _gerar_lancamentos_contabeis(self) -> None:
        """Classifica DUAS fontes contra as regras da EMPRESA ATUAL:
        - Pendentes OFX brutos → regras tipo memo
        - Pares conciliados sem Domínio → regras tipo fornecedor

        Deriva ``pendentes_ofx`` e ``pendentes_planilha`` (visíveis, sem os
        classificados manualmente) e marca os pares que viraram lançamento
        em ``self.ids_pares_classificados`` (usado pra ocultá-los na aba
        Comparação).

        Lançamentos manuais avulsos (tipo_regra=manual_planilha/manual_ofx)
        consomem a transação correspondente dos Pendentes.
        """
        regras = self._get_regras_empresa()
        # Candidatos a "fornecedor": pares P×O sem match no Domínio
        # E também pendentes da planilha (que não casaram com OFX) —
        # essas regras viram lançamento com banco='Caixa geral'.
        pares_sem_dominio = [
            p for p in self.pares_conciliados if p.dominio is None
        ]
        automaticos = gerar_lancamentos_contabeis(
            self.pendentes_ofx_brutos, regras, pares_sem_dominio,
            pendentes_planilha=self.pendentes_planilha_brutos,
        )
        # Remove os lançamentos cujo usuário marcou como ignorado (excluiu
        # ou editou — a versão editada vai como manual em lancamentos_manuais)
        if self.lancamentos_ignorados:
            automaticos = [
                l for l in automaticos
                if l.transacao_origem is None
                or id(l.transacao_origem) not in self.lancamentos_ignorados
            ]
        # Filtros de persistência dos manuais:
        # - manual (com par_origem): par precisa ainda existir
        # - manual_planilha: transação precisa ainda estar em pendentes_planilha_brutos
        #   (fonte da verdade — a lista visível é DERIVADA e vai perder a
        #   transação quando ela vira lançamento, causando bug de ela reaparecer)
        # - manual_ofx: transação precisa ainda estar em pendentes_ofx_brutos
        ids_pares_atuais = {id(p) for p in self.pares_conciliados}
        ids_pendentes_p = {id(t) for t in self.pendentes_planilha_brutos}
        ids_pendentes_o = {id(t) for t in self.pendentes_ofx_brutos}

        def _manual_vivo(l: LancamentoContabil) -> bool:
            if l.tipo_regra == "manual":
                return l.par_origem is None or id(l.par_origem) in ids_pares_atuais
            if l.tipo_regra == "manual_planilha":
                return (
                    l.transacao_origem is None
                    or id(l.transacao_origem) in ids_pendentes_p
                )
            if l.tipo_regra == "manual_ofx":
                return (
                    l.transacao_origem is None
                    or id(l.transacao_origem) in ids_pendentes_o
                )
            return True

        self.lancamentos_manuais = [l for l in self.lancamentos_manuais if _manual_vivo(l)]
        self.lancamentos_contabeis = automaticos + self.lancamentos_manuais

        # Deriva pendentes_ofx (visível) removendo os classificados por memo
        # OU por lançamento manual avulso do OFX
        ids_trans_ofx_classificadas = {
            id(l.transacao_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("memo", "manual_ofx") and l.transacao_origem is not None
        }
        # Também esconde os OFX que casaram com Domínio direto (ofx_ok):
        # a movimentação já está registrada contabilmente, nada a fazer.
        ids_ofx_no_dominio = {
            id_t for id_t, m in self.pendentes_ofx_dominio.items()
            if m.get("dominio") is not None
        }
        self.pendentes_ofx = [
            t for t in self.pendentes_ofx_brutos
            if id(t) not in ids_trans_ofx_classificadas
            and id(t) not in ids_ofx_no_dominio
        ]
        # Deriva pendentes_planilha (visível) a partir dos brutos, removendo:
        # - manual_planilha (lançamento avulso da planilha)
        # - fornecedor_planilha (regra fornecedor aplicada em pendente da planilha)
        # - pendentes que casaram com Domínio (caixa_ok — já lançados)
        ids_trans_p_classificadas = {
            id(l.transacao_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("manual_planilha", "fornecedor_planilha")
            and l.transacao_origem is not None
        }
        ids_pendentes_no_dominio = {
            id_t for id_t, m in self.pendentes_planilha_dominio.items()
            if m.get("dominio") is not None
        }
        self.pendentes_planilha = [
            t for t in self.pendentes_planilha_brutos
            if id(t) not in ids_trans_p_classificadas
            and id(t) not in ids_pendentes_no_dominio
        ]

        # Pares (sem Domínio) que viraram lançamento — pra ocultar em Comparação
        # (tanto por regra de fornecedor quanto por lançamento manual)
        self.ids_pares_classificados = {
            id(l.par_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("fornecedor", "manual") and l.par_origem is not None
        }
        if hasattr(self, "tree_lancamentos"):
            self._render_aba_lancamentos()

    def _render_aba_lancamentos(self) -> None:
        for item in self.tree_lancamentos.get_children():
            self.tree_lancamentos.delete(item)
        self.itens_lancamentos.clear()
        for l in self.lancamentos_contabeis:
            iid = self.tree_lancamentos.insert(
                "", "end",
                values=(
                    l.data.strftime("%d/%m/%Y"),
                    l.banco,
                    f"{l.valor:.2f}",
                    l.conta,
                    l.historico,
                    l.memo_original,
                    l.padrao_match,
                ),
            )
            self.itens_lancamentos[iid] = l
        idx = self.notebook.index(self._aba_lancamentos)
        self.notebook.tab(
            idx, text=f"Lançamentos contábeis ({len(self.lancamentos_contabeis)})",
        )
        # Atualiza label que indica qual empresa está ativa
        if hasattr(self, "lbl_lancamentos_empresa"):
            emp = self._empresa_selecionada()
            n_regras = len(self._get_regras_empresa())
            if emp:
                self.lbl_lancamentos_empresa.config(
                    text=(
                        f"Empresa ativa: {emp['razao']} (código {emp['codi_emp']}) "
                        f"— {n_regras} regra(s) cadastrada(s)"
                    ),
                )
            else:
                self.lbl_lancamentos_empresa.config(
                    text="⚠ Nenhuma empresa selecionada — clique em "
                         "'Selecionar empresa' para ativar regras.",
                )

    def _editar_lancamento_contabil(self) -> None:
        """Edita o lançamento contábil selecionado. Se for automático
        (vindo de regra), promove para manual — assim a edição persiste
        e não é sobrescrita no próximo recálculo."""
        sel = self.tree_lancamentos.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um lançamento na lista para editar.",
            )
            return
        lanc = self.itens_lancamentos.get(sel[0])
        if lanc is None:
            return

        dlg = DialogoEditarLancamento(
            self, lanc, plano_contas=self.plano_contas,
        )
        self.wait_window(dlg)
        if not dlg.resultado:
            return
        r = dlg.resultado

        eh_manual = lanc.tipo_regra in (
            "manual", "manual_planilha", "manual_ofx",
        )
        if eh_manual:
            # Edita o objeto manual existente — persiste em lancamentos_manuais
            lanc.data = r["data"]
            lanc.valor = r["valor"]
            lanc.banco = r["banco"]
            lanc.conta = r["conta"]
            lanc.historico = r["historico"]
        else:
            # Automático: marca o original como ignorado e cria uma cópia
            # manual com os novos valores (vai para lancamentos_manuais)
            if lanc.transacao_origem is not None:
                self.lancamentos_ignorados.add(id(lanc.transacao_origem))
            # Tipo do novo manual: baseado na origem
            if lanc.tipo_regra == "memo":
                novo_tipo = "manual_ofx"
            elif lanc.tipo_regra == "fornecedor":
                novo_tipo = "manual"
            elif lanc.tipo_regra == "fornecedor_planilha":
                novo_tipo = "manual_planilha"
            else:
                novo_tipo = "manual"
            novo = LancamentoContabil(
                data=r["data"],
                historico=r["historico"],
                valor=r["valor"],
                banco=r["banco"],
                memo_original=lanc.memo_original,
                padrao_match=f"{lanc.padrao_match} (editado)",
                conta=r["conta"],
                tipo_regra=novo_tipo,
                fornecedor=lanc.fornecedor,
                cnpj=lanc.cnpj,
                transacao_origem=lanc.transacao_origem,
                par_origem=lanc.par_origem,
            )
            self.lancamentos_manuais.append(novo)

        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()

    def _excluir_lancamento_contabil(self) -> None:
        """Exclui o lançamento contábil selecionado. Se for automático,
        marca a transação origem como ignorada para não voltar a aparecer.
        A transação volta pra Pendentes (do lado correspondente)."""
        sel = self.tree_lancamentos.selection()
        if not sel:
            messagebox.showinfo(
                "Sem seleção",
                "Selecione um lançamento na lista para excluir.",
            )
            return
        lanc = self.itens_lancamentos.get(sel[0])
        if lanc is None:
            return
        if not messagebox.askyesno(
            "Confirmar exclusão",
            f"Excluir este lançamento?\n\n"
            f"Histórico: {lanc.historico}\n"
            f"Valor: R$ {lanc.valor:.2f}\n"
            f"Conta: {lanc.conta or '(vazia)'}\n\n"
            "A transação volta para a aba Pendentes (se aplicável).",
        ):
            return

        eh_manual = lanc.tipo_regra in (
            "manual", "manual_planilha", "manual_ofx",
        )
        if eh_manual:
            # Remove dos manuais
            if lanc in self.lancamentos_manuais:
                self.lancamentos_manuais.remove(lanc)
        else:
            # Automático: marca a origem como ignorada
            if lanc.transacao_origem is not None:
                self.lancamentos_ignorados.add(id(lanc.transacao_origem))

        self._gerar_lancamentos_contabeis()
        self._redesenha_abas()

    # ----------------- Exportação para Excel -----------------

    def _sugere_nome_export(self, prefixo: str) -> str:
        """Gera nome sugerido para o arquivo .xlsx, incluindo empresa e data."""
        from datetime import datetime as _dt
        emp = self._empresa_selecionada() or {}
        codi = emp.get("codi_emp", "")
        hoje = _dt.now().strftime("%Y-%m-%d")
        # Só caracteres seguros no nome de arquivo
        base = f"{prefixo}_{codi}_{hoje}" if codi else f"{prefixo}_{hoje}"
        return "".join(c if c.isalnum() or c in "_-" else "_" for c in base) + ".xlsx"

    def _exportar_conciliados_dominio(self) -> None:
        """Exporta a aba Conciliados × Domínio para .xlsx.
        Inclui pares P×OFX triple-matched, pendentes de Caixa geral E
        pendentes OFX (sem planilha) que casaram com o Domínio."""
        # Fontes de dados exatamente como o render da aba
        pares_triple = [p for p in self.pares_conciliados if p.dominio is not None]
        caixa_dominio: list[tuple[Transacao, dict]] = []
        for t in self.pendentes_planilha_brutos:
            m = self.pendentes_planilha_dominio.get(id(t))
            if m and m.get("dominio") is not None:
                caixa_dominio.append((t, m))
        ofx_dominio: list[tuple[Transacao, dict]] = []
        for t in self.pendentes_ofx_brutos:
            m = self.pendentes_ofx_dominio.get(id(t))
            if m and m.get("dominio") is not None:
                ofx_dominio.append((t, m))

        if not pares_triple and not caixa_dominio and not ofx_dominio:
            messagebox.showinfo(
                "Sem dados",
                "Não há lançamentos conciliados com o Domínio para exportar.",
            )
            return

        caminho = filedialog.asksaveasfilename(
            title="Salvar Conciliados × Domínio",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialfile=self._sugere_nome_export("conciliados_dominio"),
        )
        if not caminho:
            return
        try:
            total = exportar_conciliados_dominio(
                caminho, pares_triple, caixa_dominio,
                pendentes_ofx_dominio=ofx_dominio,
            )
        except PermissionError:
            messagebox.showerror(
                "Arquivo em uso",
                f"Não foi possível gravar em:\n{caminho}\n\n"
                "Feche o arquivo se ele já está aberto no Excel e tente novamente.",
            )
            return
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))
            return
        messagebox.showinfo(
            "Exportação concluída",
            f"{total} linha(s) exportada(s) para:\n{caminho}",
        )

    def _exportar_pendentes(self) -> None:
        """Exporta a aba Pendentes para .xlsx (2 abas: planilha + OFX).
        Usa a lista VISÍVEL (sem os classificados por regra/manual)."""
        if not self.pendentes_planilha and not self.pendentes_ofx:
            messagebox.showinfo(
                "Sem dados",
                "Não há pendentes para exportar.",
            )
            return
        caminho = filedialog.asksaveasfilename(
            title="Salvar Pendentes",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialfile=self._sugere_nome_export("pendentes"),
        )
        if not caminho:
            return
        try:
            n_p, n_o = exportar_pendentes(
                caminho, self.pendentes_planilha, self.pendentes_ofx,
            )
        except PermissionError:
            messagebox.showerror(
                "Arquivo em uso",
                f"Não foi possível gravar em:\n{caminho}\n\n"
                "Feche o arquivo se ele já está aberto no Excel e tente novamente.",
            )
            return
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))
            return
        messagebox.showinfo(
            "Exportação concluída",
            f"Exportado para:\n{caminho}\n\n"
            f"• Pendentes Planilha: {n_p} linha(s)\n"
            f"• Pendentes OFX: {n_o} linha(s)",
        )

    def _exportar_pendencias_comparacao(self) -> None:
        """Exporta as pendências da aba Comparação (linhas amarelas +
        cinzas + laranjas — tudo que falta no Domínio) para .xlsx.
        Gera abas separadas por tipo pra manter estruturas coerentes."""
        # Amarelos: pares P×OFX sem Domínio E que ainda não viraram lançamento
        amarelos = [
            par for par in self.pares_conciliados
            if par.dominio is None and id(par) not in self.ids_pares_classificados
        ]
        # Cinzas: pendentes planilha sem match no Domínio E não classificados
        ids_p_classificadas = {
            id(l.transacao_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("manual_planilha", "fornecedor_planilha")
            and l.transacao_origem is not None
        }
        cinzas = [
            t for t in self.pendentes_planilha_brutos
            if id(t) not in ids_p_classificadas
            and not (
                (m := self.pendentes_planilha_dominio.get(id(t)))
                and m.get("dominio") is not None
            )
        ]
        # Laranjas: pendentes OFX sem match no Domínio E não classificados
        ids_o_classificadas = {
            id(l.transacao_origem) for l in self.lancamentos_contabeis
            if l.tipo_regra in ("memo", "manual_ofx")
            and l.transacao_origem is not None
        }
        laranjas = [
            t for t in self.pendentes_ofx_brutos
            if id(t) not in ids_o_classificadas
            and not (
                (m := self.pendentes_ofx_dominio.get(id(t)))
                and m.get("dominio") is not None
            )
        ]
        if not amarelos and not cinzas and not laranjas:
            messagebox.showinfo(
                "Sem pendências",
                "Não há linhas amarelas, cinzas ou laranjas para exportar.",
            )
            return
        caminho = filedialog.asksaveasfilename(
            title="Salvar pendências da Comparação",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialfile=self._sugere_nome_export("pendencias_comparacao"),
        )
        if not caminho:
            return
        try:
            n_a, n_c, n_l = exportar_pendencias_comparacao(
                caminho, amarelos, cinzas, laranjas,
            )
        except PermissionError:
            messagebox.showerror(
                "Arquivo em uso",
                f"Não foi possível gravar em:\n{caminho}\n\n"
                "Feche o arquivo se ele já está aberto no Excel e tente novamente.",
            )
            return
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))
            return
        partes = []
        if n_a:
            partes.append(f"• Amarelos (P×OFX): {n_a}")
        if n_c:
            partes.append(f"• Cinzas (Caixa geral): {n_c}")
        if n_l:
            partes.append(f"• Laranjas (OFX sem planilha): {n_l}")
        messagebox.showinfo(
            "Exportação concluída",
            f"Exportado para:\n{caminho}\n\n" + "\n".join(partes),
        )

    def _exportar_lancamentos_contabeis(self) -> None:
        """Exporta a aba Lançamentos contábeis para .xlsx."""
        if not self.lancamentos_contabeis:
            messagebox.showinfo(
                "Sem dados",
                "Não há lançamentos contábeis para exportar.",
            )
            return
        caminho = filedialog.asksaveasfilename(
            title="Salvar Lançamentos contábeis",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")],
            initialfile=self._sugere_nome_export("lancamentos_contabeis"),
        )
        if not caminho:
            return
        try:
            total = exportar_lancamentos_contabeis(caminho, self.lancamentos_contabeis)
        except PermissionError:
            messagebox.showerror(
                "Arquivo em uso",
                f"Não foi possível gravar em:\n{caminho}\n\n"
                "Feche o arquivo se ele já está aberto no Excel e tente novamente.",
            )
            return
        except Exception as e:
            messagebox.showerror("Erro ao exportar", str(e))
            return
        messagebox.showinfo(
            "Exportação concluída",
            f"{total} lançamento(s) exportado(s) para:\n{caminho}",
        )


if __name__ == "__main__":
    App().mainloop()
