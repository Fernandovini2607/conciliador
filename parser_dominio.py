"""Conexão ODBC e extração de pagamentos do banco do Domínio (SQL Anywhere)."""

from __future__ import annotations

from typing import Any

import pyodbc

from parser_xlsx import Transacao, para_data, para_decimal


def conectar(dsn: str, usuario: str, senha: str, timeout: int = 10) -> pyodbc.Connection:
    conn_str = f"DSN={dsn};UID={usuario};PWD={senha}"
    return pyodbc.connect(conn_str, timeout=timeout)


def listar_tabelas(conn: pyodbc.Connection) -> list[str]:
    """Lista tabelas e views do banco, no formato 'schema.tabela' quando houver schema."""
    cursor = conn.cursor()
    nomes: set[str] = set()
    for tipo in ("TABLE", "VIEW"):
        try:
            for row in cursor.tables(tableType=tipo):
                schema = (row.table_schem or "").strip()
                nome = row.table_name
                # ignora tabelas de sistema do SQL Anywhere
                if schema.lower() in {"sys", "dbo"} or nome.lower().startswith("sys"):
                    continue
                nomes.add(f"{schema}.{nome}" if schema else nome)
        except pyodbc.Error:
            continue
    return sorted(nomes)


def listar_colunas(conn: pyodbc.Connection, tabela: str) -> list[str]:
    cursor = conn.cursor()
    cursor.execute(f"SELECT * FROM {tabela} WHERE 1=0")
    return [c[0] for c in cursor.description]


def amostra(
    conn: pyodbc.Connection, tabela: str, n: int = 20,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Primeiras N linhas de uma tabela. Tenta TOP, depois FIRST, depois LIMIT."""
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
    conn: pyodbc.Connection, sql: str,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    cursor = conn.cursor()
    cursor.execute(sql)
    colunas = [c[0] for c in cursor.description]
    linhas = [tuple(r) for r in cursor.fetchall()]
    return colunas, linhas


def _monta_select(tabela: str, colunas: list[str], where: str) -> str:
    sql = f"SELECT {', '.join(colunas)} FROM {tabela}"
    if where.strip():
        sql += f" WHERE {where.strip()}"
    return sql


def extrair_pagamentos(
    conn: pyodbc.Connection,
    fonte: dict[str, Any],
) -> list[Transacao]:
    """Extrai pagamentos do Domínio conforme a configuração.

    `fonte` aceita dois modos:
    - {"modo": "tabela", "tabela": "...", "mapeamento": {data,valor,descricao},
       "where": "...opcional"}
    - {"modo": "sql", "sql": "SELECT ...", "mapeamento": {data,valor,descricao}}
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
