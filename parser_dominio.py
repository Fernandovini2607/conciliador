"""Conexão ODBC com o ERP Domínio (SQL Anywhere 17) — padrão Janco.

Convenções
----------
* Credenciais ficam em ``data/dominio_config.json`` (não commitado, ver .gitignore).
* Toda conexão é **read-only por padrão** — o usuário ODBC do Domínio não tem
  GRANT de UPDATE. Qualquer escrita exige RPA, fora do escopo deste app.
* Schema padrão das tabelas é ``bethadba`` — sempre prefixar.

Gotchas SQL Anywhere
--------------------
* ``DISTINCT + ORDER BY`` exige usar **aliases do SELECT**, não nomes
  qualificados (erro -854 ao referenciar ``p.CODI_GRU``; use ``grupo`` se
  aliasou).
* ``SELECT TOP n`` é a sintaxe preferida para limitar resultados; alguns
  drivers/versões aceitam ``FIRST n`` ou ``LIMIT n``. Aqui tentamos os três.
"""

from __future__ import annotations

import json
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

import pyodbc

from parser_xlsx import CAMPOS_EXTRAS, Transacao, para_data, para_decimal


@dataclass
class ContaContabil:
    codigo: str           # código da conta (ex.: "1.01.001")
    descricao: str        # descrição/nome
    tipo: str = ""        # opcional: "A" analítica / "S" sintética

DOMINIO_CONFIG_PATH = Path(__file__).parent / "data" / "dominio_config.json"
SCHEMA_PADRAO = "bethadba"


# --------------------------------------------------------------- config ODBC

def load_odbc_config() -> dict[str, Any]:
    """Lê ``data/dominio_config.json``. Retorna ``{}`` se não existir."""
    if not DOMINIO_CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(DOMINIO_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_odbc_config(cfg: dict[str, Any]) -> None:
    """Grava credenciais em ``data/dominio_config.json``. Cria o diretório."""
    DOMINIO_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    DOMINIO_CONFIG_PATH.write_text(
        json.dumps(cfg, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def _build_conn_str(cfg: dict[str, Any]) -> str:
    dsn = cfg.get("dsn", "").strip()
    if not dsn:
        raise RuntimeError(
            "DSN ODBC não configurado. Use save_odbc_config() ou o diálogo do app."
        )
    return f"DSN={dsn};UID={cfg.get('usuario', '')};PWD={cfg.get('senha', '')}"


# ------------------------------------------------------------- conexão

@contextmanager
def connect_dominio(
    readonly: bool = True,
    cfg: dict[str, Any] | None = None,
    timeout: int = 10,
) -> Iterator[pyodbc.Connection]:
    """Context manager: abre a conexão, devolve, e fecha automaticamente.

    Use em scripts e operações pontuais::

        with connect_dominio(readonly=True) as conn:
            cols, rows = executar_query(conn, "SELECT ...")
    """
    cfg = cfg if cfg is not None else load_odbc_config()
    conn = pyodbc.connect(_build_conn_str(cfg), readonly=readonly, timeout=timeout)
    try:
        yield conn
    finally:
        conn.close()


def open_connection(
    readonly: bool = True,
    cfg: dict[str, Any] | None = None,
    timeout: int = 10,
) -> pyodbc.Connection:
    """Abre conexão sem context manager — use em GUIs que precisam mantê-la viva.

    O chamador é responsável por chamar ``.close()`` quando terminar.
    """
    cfg = cfg if cfg is not None else load_odbc_config()
    return pyodbc.connect(_build_conn_str(cfg), readonly=readonly, timeout=timeout)


# ------------------------------------------------------------- exploração

def listar_tabelas(
    conn: pyodbc.Connection,
    schema: str | None = SCHEMA_PADRAO,
) -> list[str]:
    """Lista tabelas e views do banco.

    Por padrão filtra ``schema='bethadba'``. Passe ``schema=None`` para ver tudo.
    Sempre filtra tabelas de sistema (``sys``, ``dbo``, ``SYS*``).
    """
    cursor = conn.cursor()
    nomes: set[str] = set()
    for tipo in ("TABLE", "VIEW"):
        try:
            iterador = (
                cursor.tables(schema=schema, tableType=tipo)
                if schema else cursor.tables(tableType=tipo)
            )
            for row in iterador:
                sch = (row.table_schem or "").strip()
                nome = row.table_name
                if sch.lower() in {"sys", "dbo"} or nome.lower().startswith("sys"):
                    continue
                nomes.add(f"{sch}.{nome}" if sch else nome)
        except pyodbc.Error:
            continue
    return sorted(nomes)


def listar_colunas(conn: pyodbc.Connection, tabela: str) -> list[str]:
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM {tabela} WHERE 1=0")
    return [c[0] for c in cursor.description]


def amostra(
    conn: pyodbc.Connection,
    tabela: str,
    n: int = 20,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Primeiras N linhas. Tenta TOP / FIRST / LIMIT (fallback entre dialetos)."""
    cursor = conn.cursor()
    for sql in (
        f"SELECT TOP {n} * FROM {tabela}",
        f"SELECT FIRST {n} * FROM {tabela}",
        f"SELECT * FROM {tabela} LIMIT {n}",
    ):
        try:
            cursor.execute(sql)
            colunas = [c[0] for c in cursor.description]
            linhas = [tuple(r) for r in cursor.fetchall()]
            return colunas, linhas
        except pyodbc.Error:
            continue
    raise RuntimeError(f"Não consegui obter amostra da tabela {tabela}.")


def executar_query(
    conn: pyodbc.Connection,
    sql: str,
    params: tuple[Any, ...] = (),
) -> tuple[list[str], list[tuple[Any, ...]]]:
    cursor = conn.cursor()
    cursor.execute(sql, params) if params else cursor.execute(sql)
    colunas = [c[0] for c in cursor.description]
    linhas = [tuple(r) for r in cursor.fetchall()]
    return colunas, linhas


# ---------------------------------------------------------- empresas

def listar_empresas(conn: pyodbc.Connection) -> list[dict[str, Any]]:
    """Lista empresas cadastradas em ``bethadba.geempre``.

    Como os nomes das colunas mudam entre versões do Domínio (RAZA_EMP /
    RAZAO_EMP / RAZAO_SOCIAL etc.), tenta variantes conhecidas e cai num
    fallback que inspeciona o schema da tabela.
    """
    cursor = conn.cursor()
    tentativas = [
        ("SELECT CODI_EMP, RAZA_EMP, CGCE_EMP FROM bethadba.geempre ORDER BY RAZA_EMP",
         ("codi_emp", "razao", "cnpj")),
        ("SELECT CODI_EMP, RAZAO_EMP, CGCE_EMP FROM bethadba.geempre ORDER BY RAZAO_EMP",
         ("codi_emp", "razao", "cnpj")),
        ("SELECT CODI_EMP, NOME_EMP, CGCE_EMP FROM bethadba.geempre ORDER BY NOME_EMP",
         ("codi_emp", "razao", "cnpj")),
    ]
    for sql, _ in tentativas:
        try:
            cursor.execute(sql)
            return [
                {"codi_emp": r[0], "razao": str(r[1] or "").strip(),
                 "cnpj": str(r[2] or "").strip()}
                for r in cursor.fetchall()
            ]
        except pyodbc.Error:
            continue

    # Fallback: SELECT * e detecta colunas
    cursor.execute("SELECT * FROM bethadba.geempre")
    colunas = [c[0].upper() for c in cursor.description]
    idx_emp = next((i for i, c in enumerate(colunas) if c == "CODI_EMP"), None)
    idx_raz = next(
        (i for i, c in enumerate(colunas) if "RAZ" in c or "NOME" in c),
        None,
    )
    idx_cnpj = next(
        (i for i, c in enumerate(colunas) if "CGC" in c or "CNPJ" in c),
        None,
    )
    if idx_emp is None:
        raise RuntimeError("Tabela bethadba.geempre não tem coluna CODI_EMP.")
    resultado = []
    for r in cursor.fetchall():
        resultado.append({
            "codi_emp": r[idx_emp],
            "razao": str(r[idx_raz] or "").strip() if idx_raz is not None else "",
            "cnpj": str(r[idx_cnpj] or "").strip() if idx_cnpj is not None else "",
        })
    resultado.sort(key=lambda e: e["razao"])
    return resultado


# --------------------------------------------------- extração de pagamentos

def _monta_select(
    tabela: str,
    colunas: list[str],
    where: str,
    codi_emp: int | None = None,
) -> tuple[str, tuple[Any, ...]]:
    """Monta SELECT com WHERE composto. Quando ``codi_emp`` é informado,
    injeta automaticamente ``CODI_EMP = ?`` no filtro (parâmetro pyodbc)."""
    sql = f"SELECT {', '.join(colunas)} FROM {tabela}"
    cond: list[str] = []
    params: list[Any] = []
    if where.strip():
        cond.append(f"({where.strip()})")
    if codi_emp is not None:
        cond.append("CODI_EMP = ?")
        params.append(codi_emp)
    if cond:
        sql += " WHERE " + " AND ".join(cond)
    return sql, tuple(params)


def extrair_pagamentos(
    conn: pyodbc.Connection,
    fonte: dict[str, Any],
    codi_emp: int | None = None,
) -> list[Transacao]:
    """Extrai pagamentos do Domínio conforme a configuração da fonte.

    ``fonte`` aceita dois modos:

    * ``{"modo": "tabela", "tabela": "bethadba.efentradas",
       "mapeamento": {data, valor, descricao}, "where": "...opcional"}``
    * ``{"modo": "sql", "sql": "SELECT ...",
       "mapeamento": {data, valor, descricao}}``

    Quando ``codi_emp`` é fornecido e o modo é "tabela", a função injeta
    ``CODI_EMP = ?`` no WHERE automaticamente. No modo "sql" o filtro deve
    estar dentro do SELECT escrito pelo usuário (o app só passa ``codi_emp``
    como parâmetro se o SQL contiver ``?``).
    """
    mapeamento: dict[str, str] = fonte["mapeamento"]
    obrigatorios = ("data", "valor", "data_emissao", "numero_nf", "cnpj", "fornecedor")
    faltando = [c for c in obrigatorios if not mapeamento.get(c)]
    if faltando:
        raise ValueError(f"Mapeamento incompleto no Domínio: {', '.join(faltando)}")

    extras_pedidos = {
        campo: mapeamento[campo]
        for campo in CAMPOS_EXTRAS
        if mapeamento.get(campo)
    }

    if fonte.get("modo") == "sql":
        sql = fonte["sql"]
        params: tuple[Any, ...] = (codi_emp,) if (codi_emp is not None and "?" in sql) else ()
        colunas, linhas = executar_query(conn, sql, params)
    else:
        cols_pedidas = [mapeamento["data"], mapeamento["valor"]]
        cols_pedidas.extend(extras_pedidos.values())
        # Remove duplicatas mantendo ordem
        vistos: set[str] = set()
        cols_unicas: list[str] = []
        for c in cols_pedidas:
            if c not in vistos:
                cols_unicas.append(c)
                vistos.add(c)
        sql, params = _monta_select(
            fonte["tabela"], cols_unicas, fonte.get("where", ""), codi_emp=codi_emp,
        )
        colunas, linhas = executar_query(conn, sql, params)

    try:
        i_data = colunas.index(mapeamento["data"])
        i_valor = colunas.index(mapeamento["valor"])
    except ValueError as e:
        raise ValueError(
            f"Coluna do mapeamento não encontrada no resultado da query: {e}"
        )
    indices_extras = {
        campo: colunas.index(col_nome)
        for campo, col_nome in extras_pedidos.items()
        if col_nome in colunas
    }

    # Detecta colunas opcionais de status e valor pago (vindas do SQL livre).
    # Aceita variações de nome para facilitar.
    def _achar(*nomes: str) -> int | None:
        upper = [c.upper() for c in colunas]
        for n in nomes:
            if n.upper() in upper:
                return upper.index(n.upper())
        return None

    i_status = _achar("status_parcela", "status")
    i_pago = _achar("valor_pago", "vpago", "total_pago")

    transacoes: list[Transacao] = []
    for linha in linhas:
        data = para_data(linha[i_data])
        valor = para_decimal(linha[i_valor])
        if data is None or valor is None:
            continue

        extras: dict[str, Any] = {}
        for campo, idx in indices_extras.items():
            celula = linha[idx]
            if celula is None or celula == "":
                continue
            if campo == "data_emissao":
                d = para_data(celula)
                if d is not None:
                    extras[campo] = d
            else:
                extras[campo] = str(celula).strip()

        if i_status is not None and linha[i_status] not in (None, ""):
            extras["status"] = str(linha[i_status]).strip()
        if i_pago is not None and linha[i_pago] not in (None, ""):
            v_pago = para_decimal(linha[i_pago])
            if v_pago is not None:
                extras["valor_pago"] = v_pago

        transacoes.append(Transacao(
            data=data, valor=abs(valor), descricao="",
            origem="dominio", extras=extras,
        ))
    return transacoes


# ---------------------------------------------------- plano de contas

def extrair_plano_contas(
    conn: pyodbc.Connection,
    fonte: dict[str, Any],
    codi_emp: int | None = None,
) -> list[ContaContabil]:
    """Extrai o plano de contas do Domínio conforme a configuração da fonte.

    ``fonte`` tem o mesmo modelo das outras fontes do app:
    - ``{"modo": "tabela", "tabela": "bethadba.ctcontas",
       "mapeamento": {codigo, descricao, tipo?}, "where": "...opcional"}``
    - ``{"modo": "sql", "sql": "SELECT ...",
       "mapeamento": {codigo, descricao, tipo?}}``
    """
    mapeamento: dict[str, str] = fonte["mapeamento"]
    obrigatorios = ("codigo", "descricao")
    faltando = [c for c in obrigatorios if not mapeamento.get(c)]
    if faltando:
        raise ValueError(
            f"Mapeamento incompleto no plano de contas: {', '.join(faltando)}"
        )

    if fonte.get("modo") == "sql":
        sql = fonte["sql"]
        params: tuple[Any, ...] = (codi_emp,) if (codi_emp is not None and "?" in sql) else ()
        colunas, linhas = executar_query(conn, sql, params)
    else:
        cols_pedidas = [mapeamento["codigo"], mapeamento["descricao"]]
        if mapeamento.get("tipo"):
            cols_pedidas.append(mapeamento["tipo"])
        vistos: set[str] = set()
        cols_unicas: list[str] = []
        for c in cols_pedidas:
            if c not in vistos:
                cols_unicas.append(c)
                vistos.add(c)
        sql, params = _monta_select(
            fonte["tabela"], cols_unicas, fonte.get("where", ""), codi_emp=codi_emp,
        )
        colunas, linhas = executar_query(conn, sql, params)

    try:
        i_cod = colunas.index(mapeamento["codigo"])
        i_desc = colunas.index(mapeamento["descricao"])
    except ValueError as e:
        raise ValueError(
            f"Coluna do mapeamento não encontrada no resultado: {e}"
        )
    i_tipo = None
    if mapeamento.get("tipo") and mapeamento["tipo"] in colunas:
        i_tipo = colunas.index(mapeamento["tipo"])
    else:
        # Auto-detecta coluna de tipo (analítica/sintética) por nome usual.
        # Domínio: CTRG_CTA / TIPO_CTA / ANAL_CTA / CONT_CTA — depende da
        # versão. Procuramos qualquer coluna cujo nome bate com esses
        # heurísticos.
        candidatos_nomes = (
            "TIPO_CTA", "TIPO_CTC", "CTRG_CTA", "CONT_CTA",
            "ANAL_CTA", "TIPO_CONTA", "ANALITICA", "TIPO",
        )
        cols_upper = [c.upper() for c in colunas]
        for nome in candidatos_nomes:
            if nome in cols_upper:
                i_tipo = cols_upper.index(nome)
                break

    contas: list[ContaContabil] = []
    for linha in linhas:
        cod = str(linha[i_cod] or "").strip() if linha[i_cod] is not None else ""
        desc = str(linha[i_desc] or "").strip() if linha[i_desc] is not None else ""
        tipo = (
            str(linha[i_tipo] or "").strip()
            if i_tipo is not None and linha[i_tipo] is not None else ""
        )
        if not cod:
            continue
        # Filtra: se temos info de tipo, mantém apenas contas ANALÍTICAS.
        # Aceita variações comuns: "A", "Analítica", "ANALITICA", etc.
        # Quando não há info de tipo, mantém tudo (comportamento legado).
        if i_tipo is not None:
            t_norm = tipo.upper()
            eh_analitica = (
                t_norm.startswith("A")        # "A", "Analítica"
                or t_norm in ("1", "true")    # alguns sistemas: 1/0
            )
            if not eh_analitica:
                continue
        contas.append(ContaContabil(codigo=cod, descricao=desc, tipo=tipo))
    return contas
