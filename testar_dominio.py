"""Diagnóstico de conexão ODBC com o Domínio.

Rode com:
    python testar_dominio.py

Testa em ordem:
1. O arquivo data/dominio_config.json existe e tem JSON válido?
2. A DSN configurada (default 'Contabil') está listada no ODBC 64-bit?
3. Consegue abrir conexão com user+senha do config?
4. Consegue rodar um SELECT simples?

Mostra o erro exato em cada passo — assim dá pra saber onde tá travando.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


CONFIG_PATH = _base_dir() / "data" / "dominio_config.json"


def passo(n: int, titulo: str) -> None:
    print(f"\n[{n}] {titulo}")
    print("-" * 60)


def ok(msg: str) -> None:
    print(f"    [OK] {msg}")


def erro(msg: str) -> None:
    print(f"    [ERRO] {msg}")


def main() -> int:
    print("=" * 60)
    print("DIAGNOSTICO ODBC - Dominio Contabil")
    print("=" * 60)

    # 1) Existe pyodbc?
    passo(1, "Verificar biblioteca pyodbc")
    try:
        import pyodbc
        ok(f"pyodbc {pyodbc.version} carregado")
    except ImportError as e:
        erro(f"pyodbc nao instalado: {e}")
        return 1

    # 2) Existe o arquivo de credenciais?
    passo(2, "Verificar data/dominio_config.json")
    if not CONFIG_PATH.exists():
        erro(f"Arquivo nao encontrado: {CONFIG_PATH}")
        print()
        print("    Solucao: copie 'dominio_config.json.EXEMPLO' para")
        print("    'dominio_config.json' (mesmo diretorio) e edite as credenciais.")
        return 1
    ok(f"Arquivo existe: {CONFIG_PATH}")

    # 3) JSON valido?
    passo(3, "Ler JSON do arquivo")
    try:
        cred = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        erro(f"JSON invalido: {e}")
        print()
        print("    Verifique: aspas devem ser retas (\"), nao curvas,")
        print("    e sem virgulas sobrando no fim de linhas.")
        return 1
    dsn = cred.get("dsn", "")
    usuario = cred.get("usuario", "")
    senha = cred.get("senha", "")
    if not dsn:
        erro("Campo 'dsn' vazio no JSON.")
        return 1
    ok(f"DSN configurada: '{dsn}'")
    ok(f"Usuario: '{usuario}'")
    ok(f"Senha: {'*' * len(senha)} ({len(senha)} caracteres)")

    # 4) DSN esta listada no ODBC 64-bit?
    passo(4, "Listar DSNs disponiveis no ODBC 64-bit")
    dsn_encontrada = False
    try:
        dsns_sys = pyodbc.dataSources(pyodbc.SQL_FETCH_FIRST)
        # dataSources() sem argumento retorna dict com todas
        todas = pyodbc.dataSources()
        for nome, driver in todas.items():
            marca = "  <=== esta e a sua" if nome.lower() == dsn.lower() else ""
            print(f"    - {nome:30s} ({driver}){marca}")
            if nome.lower() == dsn.lower():
                dsn_encontrada = True
    except Exception as e:
        erro(f"Nao consegui listar DSNs: {e}")
        return 1

    if not dsn_encontrada:
        print()
        erro(f"DSN '{dsn}' NAO ESTA na lista acima.")
        print()
        print("    Causas comuns:")
        print("    - A DSN foi criada em 'Fontes de Dados ODBC (32 bits)',")
        print("      mas o app usa 64 bits. Recrie em 64 bits.")
        print("    - A DSN foi criada como 'DSN de Usuario' de OUTRA conta")
        print("      Windows. Crie como 'DSN de Sistema' pra todos verem.")
        print("    - O nome no arquivo JSON esta diferente do nome da DSN")
        print("      (verifique maiusculas/minusculas).")
        return 1

    ok(f"DSN '{dsn}' encontrada no ODBC 64-bit")

    # 5) Tentar conectar
    passo(5, "Abrir conexao ODBC (isso pode demorar 5-10s...)")
    conn_str = f"DSN={dsn};UID={usuario};PWD={senha}"
    try:
        conn = pyodbc.connect(conn_str, readonly=True, timeout=15)
    except pyodbc.Error as e:
        erro(f"Falha na conexao: {e}")
        print()
        sqlstate = e.args[0] if e.args else ""
        if sqlstate == "28000":
            print("    Codigo 28000 = usuario/senha invalidos.")
        elif sqlstate == "IM002":
            print("    Codigo IM002 = DSN nao encontrada em runtime.")
        elif sqlstate == "08001":
            print("    Codigo 08001 = servidor Dominio (SQL Anywhere) fora do ar.")
            print("    Verifique se o servico do Dominio esta rodando na maquina.")
        return 1
    ok("Conexao aberta com sucesso")

    # 6) Rodar um SELECT simples
    passo(6, "Rodar SELECT simples pra confirmar que ODBC responde")
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM bethadba.geempre")
        n = cur.fetchone()[0]
        ok(f"Query rodou. bethadba.geempre tem {n} empresa(s) cadastrada(s).")
    except pyodbc.Error as e:
        erro(f"Query falhou: {e}")
        print()
        print("    Talvez o usuario nao tem permissao de SELECT no schema")
        print("    'bethadba' do Dominio. Consulte o admin do Dominio.")
        conn.close()
        return 1

    conn.close()

    print()
    print("=" * 60)
    print("SUCESSO - Dominio esta conectando normalmente.")
    print("Se o Conciliador ainda nao conecta, feche e reabra o app.")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
