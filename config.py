"""Persistência de configuração — usa MariaDB via db.py.

Mantém a MESMA API do config.json anterior (``carregar()`` retorna dict,
``salvar(dados)`` persiste) — assim o restante do código não precisou
mudar nada.

Como o config antigo era um dict aninhado (com listas por empresa,
mapeamentos por empresa, etc.), aqui a gente distribui em 3 tabelas:

- ``app_config``           → configs escalares (empresa ativa, fontes)
- ``regra_taxa``           → regras de classificação (lista por empresa)
- ``mapeamento_planilha``  → mapeamento de colunas (dict por empresa)

Chaves que vão pra ``app_config`` (formato JSON):
- dominio_empresa
- dominio_fonte_pagamentos
- dominio_fonte_plano_contas
- (qualquer outra chave desconhecida também vai pra cá)

Chaves que viram tabelas relacionais:
- regras_taxas_por_empresa       → tabela regra_taxa
- mapeamentos_planilha_por_empresa → tabela mapeamento_planilha
"""

from __future__ import annotations

import json
from typing import Any

from db import conexao

# Chaves do dict antigo que viram tabelas separadas — não vão pra app_config
CHAVES_RELACIONAIS = {
    "regras_taxas_por_empresa",
    "mapeamentos_planilha_por_empresa",
}

# ID do usuário logado — setado pelo main.py após login bem-sucedido.
# Quando None, o config.carregar() não injeta empresa_ativa (comportamento
# do config.json antigo global).
_usuario_atual_id: int | None = None


def set_usuario_atual(usuario_id: int | None) -> None:
    """Registra o usuário logado. A partir daqui, carregar()/salvar()
    lêem/escrevem a ``dominio_empresa`` na coluna ``empresa_ativa`` desse
    usuário — não mais no ``app_config`` global."""
    global _usuario_atual_id
    _usuario_atual_id = usuario_id


def get_usuario_atual_id() -> int | None:
    return _usuario_atual_id


def _decode_json(v: Any) -> Any:
    """PyMySQL retorna JSON como string em algumas versões — decodifica."""
    if isinstance(v, (str, bytes, bytearray)):
        try:
            return json.loads(v)
        except (TypeError, ValueError):
            return v
    return v


def carregar() -> dict[str, Any]:
    """Lê tudo do banco e devolve um dict no mesmo formato do config.json
    antigo. Se o banco não tem nada, devolve dict vazio.
    """
    dados: dict[str, Any] = {}

    try:
        with conexao() as conn:
            cur = conn.cursor()

            # 1) app_config → chaves escalares (fontes SQL, etc). Note que
            # 'dominio_empresa' NÃO vive mais aqui — vem de usuario.empresa_ativa
            cur.execute("SELECT chave, valor FROM app_config")
            for chave, valor in cur.fetchall():
                if chave == "dominio_empresa":
                    continue  # obsoleto no schema global, ignora
                dados[chave] = _decode_json(valor)

            # 1b) Injeta dominio_empresa do usuário logado (se houver)
            if _usuario_atual_id is not None:
                cur.execute(
                    "SELECT empresa_ativa FROM usuario WHERE id = %s",
                    (_usuario_atual_id,),
                )
                row = cur.fetchone()
                if row and row[0] is not None:
                    emp = _decode_json(row[0])
                    if emp:
                        dados["dominio_empresa"] = emp

            # 2) regra_taxa → regras_taxas_por_empresa
            cur.execute(
                """
                SELECT codi_emp, tipo, padrao, historico, conta, banco
                FROM regra_taxa
                ORDER BY codi_emp, ordem, id
                """
            )
            regras_por_emp: dict[str, list[dict[str, Any]]] = {}
            for codi_emp, tipo, padrao, historico, conta, banco in cur.fetchall():
                chave_emp = str(codi_emp)
                regra: dict[str, Any] = {
                    "tipo": tipo,
                    "padrao": padrao or "",
                    "historico": historico or "",
                    "conta": conta or "",
                }
                if banco:
                    regra["banco"] = banco
                regras_por_emp.setdefault(chave_emp, []).append(regra)
            if regras_por_emp:
                dados["regras_taxas_por_empresa"] = regras_por_emp

            # 3) mapeamento_planilha → mapeamentos_planilha_por_empresa
            cur.execute(
                """
                SELECT codi_emp, campo, nome_coluna
                FROM mapeamento_planilha
                ORDER BY codi_emp, campo
                """
            )
            mapa_por_emp: dict[str, dict[str, str]] = {}
            for codi_emp, campo, nome_coluna in cur.fetchall():
                chave_emp = str(codi_emp)
                mapa_por_emp.setdefault(chave_emp, {})[campo] = nome_coluna
            if mapa_por_emp:
                dados["mapeamentos_planilha_por_empresa"] = mapa_por_emp
    except Exception:
        # Banco fora do ar / não configurado → devolve vazio (comportamento
        # legado do config.json ausente). O caller trata como "primeira vez".
        return {}

    return dados


def salvar(dados: dict[str, Any]) -> None:
    """Persiste o dict no banco. Estratégia: substitui TUDO — assim a
    semântica é idêntica ao config.json antigo (escrever um dict novo
    apagava/sobrescrevia tudo)."""
    # Separa o que é escalar do que é tabela relacional
    escalares = {k: v for k, v in dados.items() if k not in CHAVES_RELACIONAIS}
    regras_por_emp = dados.get("regras_taxas_por_empresa", {}) or {}
    mapa_por_emp = dados.get("mapeamentos_planilha_por_empresa", {}) or {}

    # dominio_empresa NÃO vai pro app_config global — é por usuário. Extrai
    # antes pra atualizar usuario.empresa_ativa ao final.
    dominio_empresa = escalares.pop("dominio_empresa", None)

    with conexao() as conn:
        cur = conn.cursor()

        # --- app_config: apaga tudo e re-insere (globais só)
        cur.execute("DELETE FROM app_config")
        if escalares:
            for chave, valor in escalares.items():
                cur.execute(
                    "INSERT INTO app_config (chave, valor) VALUES (%s, %s)",
                    (chave, json.dumps(valor, ensure_ascii=False, default=str)),
                )

        # --- dominio_empresa: salva na coluna empresa_ativa do usuário logado
        if _usuario_atual_id is not None:
            cur.execute(
                "UPDATE usuario SET empresa_ativa = %s WHERE id = %s",
                (
                    json.dumps(dominio_empresa, ensure_ascii=False, default=str)
                    if dominio_empresa else None,
                    _usuario_atual_id,
                ),
            )

        # --- regra_taxa: apaga tudo e re-insere
        cur.execute("DELETE FROM regra_taxa")
        for codi_emp, regras in regras_por_emp.items():
            if not isinstance(regras, list):
                continue
            try:
                codi_int = int(codi_emp)
            except (TypeError, ValueError):
                continue
            for ordem, regra in enumerate(regras):
                if not isinstance(regra, dict):
                    continue
                cur.execute(
                    """
                    INSERT INTO regra_taxa
                      (codi_emp, tipo, padrao, historico, conta, banco, ordem)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    """,
                    (
                        codi_int,
                        (regra.get("tipo") or "memo")[:20],
                        (regra.get("padrao") or "")[:500],
                        (regra.get("historico") or "")[:500],
                        (regra.get("conta") or "")[:100],
                        (regra.get("banco") or None),
                        ordem,
                    ),
                )

        # --- mapeamento_planilha: apaga tudo e re-insere
        cur.execute("DELETE FROM mapeamento_planilha")
        for codi_emp, mapa in mapa_por_emp.items():
            if not isinstance(mapa, dict):
                continue
            try:
                codi_int = int(codi_emp)
            except (TypeError, ValueError):
                continue
            for campo, nome in mapa.items():
                if not nome:
                    continue
                cur.execute(
                    """
                    INSERT INTO mapeamento_planilha (codi_emp, campo, nome_coluna)
                    VALUES (%s, %s, %s)
                    """,
                    (codi_int, str(campo)[:50], str(nome)[:200]),
                )
