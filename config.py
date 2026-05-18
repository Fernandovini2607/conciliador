"""Persistência simples de configuração em config.json na pasta do projeto."""

import json
from pathlib import Path
from typing import Any

CONFIG_PATH = Path(__file__).parent / "config.json"


def carregar() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        return {}
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def salvar(dados: dict[str, Any]) -> None:
    CONFIG_PATH.write_text(
        json.dumps(dados, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
