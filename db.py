"""Conexão com MariaDB — camada base pra ler/escrever as configurações
do app. Segue o padrão Janco (credenciais em ``data/db_config.json``,
gitignored).

Uso:

    from db import conexao

    with conexao() as conn:
        cur = conn.cursor()
        cur.execute("SELECT * FROM regra_taxa WHERE codi_emp = %s", (55,))
        for row in cur.fetchall():
            ...

Compatível com PyMySQL (pure Python — sem compilação C, instala liso no
Windows).
"""

from __future__ import annotations

import json
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

import pymysql
from pymysql.connections import Connection


def _base_dir() -> Path:
    """Diretório-base pra encontrar a pasta ``data/``.

    - Em modo dev (rodando ``python main.py``): pasta do arquivo .py.
    - Empacotado com PyInstaller (.exe): pasta do executável — permite
      que o usuário edite ``data/db_config.json`` ao lado do .exe.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


DB_CONFIG_PATH = _base_dir() / "data" / "db_config.json"

# Valores padrão pra facilitar setup inicial. NUNCA usar senha em branco
# em produção — o setup_db.py força o usuário a preencher.
DEFAULT_CONFIG: dict[str, Any] = {
    "host": "localhost",
    "port": 3306,
    "user": "root",
    "password": "",
    "database": "conciliador",
    "charset": "utf8mb4",
}


def load_db_config() -> dict[str, Any]:
    """Lê ``data/db_config.json``. Retorna default se não existir."""
    if not DB_CONFIG_PATH.exists():
        return dict(DEFAULT_CONFIG)
    try:
        cfg = json.loads(DB_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return dict(DEFAULT_CONFIG)
    # Preenche chaves faltantes com defaults
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    return cfg


def save_db_config(cfg: dict[str, Any]) -> None:
    """Grava credenciais em ``data/db_config.json``. Cria a pasta se preciso."""
    DB_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    DB_CONFIG_PATH.write_text(
        json.dumps(cfg, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def _build_kwargs(cfg: dict[str, Any], skip_db: bool = False) -> dict[str, Any]:
    """Monta kwargs pra pymysql.connect. Se ``skip_db=True``, omite o
    ``database`` — usado no setup pra criar o schema."""
    kwargs = {
        "host": cfg["host"],
        "port": int(cfg["port"]),
        "user": cfg["user"],
        "password": cfg["password"],
        "charset": cfg.get("charset", "utf8mb4"),
        "autocommit": False,
        "connect_timeout": 10,
    }
    if not skip_db:
        kwargs["database"] = cfg["database"]
    return kwargs


def get_connection(
    cfg: dict[str, Any] | None = None,
    skip_db: bool = False,
) -> Connection:
    """Abre uma conexão pymysql. O chamador é responsável por fechar
    (ou usar o context manager ``conexao()``)."""
    cfg = cfg if cfg is not None else load_db_config()
    return pymysql.connect(**_build_kwargs(cfg, skip_db=skip_db))


@contextmanager
def conexao(
    cfg: dict[str, Any] | None = None,
    skip_db: bool = False,
) -> Iterator[Connection]:
    """Context manager: abre, entrega, fecha automaticamente. Faz commit
    no fim se não houve exceção; rollback caso contrário."""
    conn = get_connection(cfg=cfg, skip_db=skip_db)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def testar_conexao(cfg: dict[str, Any] | None = None) -> tuple[bool, str]:
    """Devolve (ok, mensagem). Usado pelo setup pra validar credenciais."""
    try:
        with conexao(cfg=cfg) as conn:
            cur = conn.cursor()
            cur.execute("SELECT VERSION()")
            versao = cur.fetchone()[0]
        return True, f"Conectado — MariaDB/MySQL {versao}"
    except pymysql.err.OperationalError as e:
        return False, f"Falha ODBC: {e.args[1] if len(e.args) > 1 else e}"
    except Exception as e:
        return False, f"Erro: {e}"
