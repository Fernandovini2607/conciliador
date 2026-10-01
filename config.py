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

    # Lista de regras rejeitadas pelo banco (ENUM / VARCHAR / etc).
    # Reportada no stderr apos o commit, sem derrubar o save inteiro.
    regras_rejeitadas: list[dict[str, Any]] = []

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

        # --- regra_taxa: apaga tudo e re-insere.
        # UMA regra com conteudo invalido (tipo nao-ENUM, texto > VARCHAR,
        # etc) nao derruba as demais — try/except por INSERT. As que
        # falharam vao pra regras_rejeitadas pra reportar no stderr.
        cur.execute("DELETE FROM regra_taxa")
        TIPOS_VALIDOS = {"memo", "fornecedor"}
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
                tipo = (regra.get("tipo") or "memo").strip()
                # ENUM so aceita 'memo' ou 'fornecedor'. Qualquer outro
                # valor (ex: 'fornecedor_planilha', '', typo) seria
                # rejeitado pelo MariaDB — normaliza aqui.
                if tipo not in TIPOS_VALIDOS:
                    tipo = "fornecedor" if "fornec" in tipo.lower() else "memo"
                padrao = (regra.get("padrao") or "")[:500]
                historico = (regra.get("historico") or "")[:500]
                conta = (regra.get("conta") or "")[:100]
                banco = regra.get("banco") or None
                if banco is not None:
                    banco = str(banco)[:200]
                try:
                    cur.execute(
                        """
                        INSERT INTO regra_taxa
                          (codi_emp, tipo, padrao, historico, conta, banco, ordem)
                        VALUES (%s, %s, %s, %s, %s, %s, %s)
                        """,
                        (codi_int, tipo, padrao, historico, conta, banco, ordem),
                    )
                except Exception as e:  # noqa: BLE001
                    regras_rejeitadas.append({
                        "codi_emp": codi_int,
                        "ordem": ordem,
                        "regra": regra,
                        "erro": str(e),
                    })

        # --- mapeamento_planilha: apaga tudo e re-insere.
        # Mesmo tratamento isolado por insert pra uma linha ruim nao
        # derrubar as demais.
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
                try:
                    cur.execute(
                        """
                        INSERT INTO mapeamento_planilha (codi_emp, campo, nome_coluna)
                        VALUES (%s, %s, %s)
                        """,
                        (codi_int, str(campo)[:50], str(nome)[:200]),
                    )
                except Exception as e:  # noqa: BLE001
                    import sys as _sys
                    print(
                        f"[config.salvar] mapeamento rejeitado "
                        f"(codi_emp={codi_int}, campo={campo!r}): {e}",
                        file=_sys.stderr,
                    )

    # Fora do context manager — commit ja aconteceu. Logar rejeitadas no
    # stderr pra ficar visivel em modo debug (iniciar_debug.bat).
    if regras_rejeitadas:
        import sys as _sys
        print(
            f"[config.salvar] {len(regras_rejeitadas)} regra(s) rejeitada(s) "
            "pelo banco:",
            file=_sys.stderr,
        )
        for r in regras_rejeitadas:
            print(
                f"  codi_emp={r['codi_emp']} ordem={r['ordem']} "
                f"regra={r['regra']} erro={r['erro']}",
                file=_sys.stderr,
            )
