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
from pathlib import Path
from typing import Any, Iterator

import pyodbc

from parser_xlsx import Transacao, para_data, para_decimal

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
) -> tuple[list[str], list[tuple[Any, ...]]]:
    cursor = conn.cursor()
    cursor.execute(sql)
    colunas = [c[0] for c in cursor.description]
    linhas = [tuple(r) for r in cursor.fetchall()]
    return colunas, linhas


# --------------------------------------------------- extração de pagamentos

def _monta_select(tabela: str, colunas: list[str], where: str) -> str:
    sql = f"SELECT {', '.join(colunas)} FROM {tabela}"
    if where.strip():
        sql += f" WHERE {where.strip()}"
    return sql


def extrair_pagamentos(
    conn: pyodbc.Connection,
    fonte: dict[str, Any],
) -> list[Transacao]:
    """Extrai pagamentos do Domínio conforme a configuração da fonte.

    ``fonte`` aceita dois modos:

    * ``{"modo": "tabela", "tabela": "bethadba.efpagamentos",
       "mapeamento": {data, valor, descricao}, "where": "...opcional"}``
    * ``{"modo": "sql", "sql": "SELECT ...",
       "mapeamento": {data, valor, descricao}}``
    """
    mapeamento: dict[str, str] = fonte["mapeamento"]
    if fonte.get("modo") == "sql":
        colunas, linhas = executar_query(conn, fonte["sql"])
    else:
        cols_pedidas = [mapeamento["data"], mapeamento["valor"], mapeamento["descricao"]]
        sql = _monta_select(fonte["tabela"], cols_pedidas, fonte.get("where", ""))
        colunas, linhas = executar_query(conn, sql)

    try:
        i_data = colunas.index(mapeamento["data"])
        i_valor = colunas.index(mapeamento["valor"])
        i_desc = colunas.index(mapeamento["descricao"])
    except ValueError as e:
        raise ValueError(
            f"Coluna do mapeamento não encontrada no resultado da query: {e}"
        )

    transacoes: list[Transacao] = []
    for linha in linhas:
        data = para_data(linha[i_data])
        valor = para_decimal(linha[i_valor])
        if data is None or valor is None:
            continue
        descricao = "" if linha[i_desc] is None else str(linha[i_desc]).strip()
        transacoes.append(Transacao(
            data=data, valor=abs(valor), descricao=descricao, origem="dominio",
        ))
    return transacoes
