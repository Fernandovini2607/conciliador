"""Autenticação de usuários do app.

Modelo:
- Senhas armazenadas com PBKDF2-HMAC-SHA256, 200_000 iterações, salt
  aleatório de 16 bytes. Formato: "iteracoes$salt_hex$hash_hex".
- Login retorna dict do usuário (sem senha). Falha retorna None.
- Todas as operações de escrita usam transações via db.conexao().

Perfis:
- admin: pode gerenciar (criar/editar/desativar) outros usuários.
- operador: usa o app normalmente, muda a própria senha.
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
from typing import Any

from db import conexao

ITERACOES_PBKDF2 = 200_000
SALT_BYTES = 16
SENHA_MIN_LEN = 4


class LoginErro(Exception):
    """Erro de autenticação — mensagem já legível pro usuário."""


class UsuarioErro(Exception):
    """Erro em CRUD de usuário — mensagem já legível pro usuário."""


# -------------------------------------------------------------- hash

def hash_senha(senha: str) -> str:
    """Devolve o hash da senha no formato 'iteracoes$salt_hex$hash_hex'."""
    salt = os.urandom(SALT_BYTES)
    dk = hashlib.pbkdf2_hmac(
        "sha256", senha.encode("utf-8"), salt, ITERACOES_PBKDF2,
    )
    return f"{ITERACOES_PBKDF2}${salt.hex()}${dk.hex()}"


def verifica_senha(senha_plain: str, hash_stored: str) -> bool:
    """Compara em tempo constante. Se o hash está mal-formado, retorna False."""
    try:
        iters_str, salt_hex, hash_hex = hash_stored.split("$", 2)
        iters = int(iters_str)
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(hash_hex)
    except (ValueError, AttributeError):
        return False
    dk = hashlib.pbkdf2_hmac(
        "sha256", senha_plain.encode("utf-8"), salt, iters,
    )
    return hmac.compare_digest(dk, expected)


# -------------------------------------------------------------- helpers

def _row_para_usuario(row: tuple) -> dict[str, Any]:
    """Converte row do SELECT ... FROM usuario (sem senha_hash) para dict."""
    id_, username, nome, email, ativo, admin, empresa_ativa, ultimo_login, criado_em = row
    return {
        "id": id_,
        "username": username,
        "nome": nome or "",
        "email": email or "",
        "ativo": bool(ativo),
        "admin": bool(admin),
        "empresa_ativa": _decode_json(empresa_ativa),
        "ultimo_login": ultimo_login,
        "criado_em": criado_em,
    }


def _decode_json(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, (str, bytes, bytearray)):
        try:
            return json.loads(v)
        except (TypeError, ValueError):
            return None
    return v


_COLS_USUARIO = (
    "id, username, nome, email, ativo, admin, "
    "empresa_ativa, ultimo_login, criado_em"
)


# -------------------------------------------------------------- login

def login(username: str, senha: str) -> dict[str, Any]:
    """Valida credenciais. Devolve dict do usuário (sem hash) em sucesso.

    Levanta LoginErro com mensagem específica em caso de falha.
    """
    if not username or not senha:
        raise LoginErro("Informe usuário e senha.")

    with conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            f"SELECT {_COLS_USUARIO}, senha_hash FROM usuario "
            "WHERE username = %s",
            (username.strip(),),
        )
        row = cur.fetchone()
        if not row:
            raise LoginErro("Usuário ou senha inválidos.")

        senha_hash = row[-1]
        dados_row = row[:-1]
        dados = _row_para_usuario(dados_row)

        if not dados["ativo"]:
            raise LoginErro("Usuário desativado. Fale com o admin.")

        if not verifica_senha(senha, senha_hash):
            raise LoginErro("Usuário ou senha inválidos.")

        # Atualiza ultimo_login
        cur.execute(
            "UPDATE usuario SET ultimo_login = NOW() WHERE id = %s",
            (dados["id"],),
        )
    return dados


# -------------------------------------------------------------- CRUD

def existe_algum_usuario() -> bool:
    """Usado pelo bootstrap: se não tem ninguém, exige criar o primeiro admin."""
    try:
        with conexao() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM usuario")
            return cur.fetchone()[0] > 0
    except Exception:
        return False


def listar_usuarios(incluir_inativos: bool = True) -> list[dict[str, Any]]:
    with conexao() as conn:
        cur = conn.cursor()
        sql = f"SELECT {_COLS_USUARIO} FROM usuario"
        if not incluir_inativos:
            sql += " WHERE ativo = 1"
        sql += " ORDER BY ativo DESC, username"
        cur.execute(sql)
        return [_row_para_usuario(r) for r in cur.fetchall()]


def buscar_usuario(usuario_id: int) -> dict[str, Any] | None:
    with conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            f"SELECT {_COLS_USUARIO} FROM usuario WHERE id = %s",
            (usuario_id,),
        )
        row = cur.fetchone()
        return _row_para_usuario(row) if row else None


def criar_usuario(
    username: str,
    senha: str,
    nome: str = "",
    email: str = "",
    admin: bool = False,
) -> int:
    """Devolve o id do usuário criado."""
    username = (username or "").strip()
    if not username:
        raise UsuarioErro("Username é obrigatório.")
    if len(username) > 50:
        raise UsuarioErro("Username muito longo (máx. 50 caracteres).")
    if not senha or len(senha) < SENHA_MIN_LEN:
        raise UsuarioErro(f"Senha muito curta (mínimo {SENHA_MIN_LEN} caracteres).")

    with conexao() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id FROM usuario WHERE username = %s", (username,))
        if cur.fetchone():
            raise UsuarioErro(f"Já existe usuário com username '{username}'.")

        cur.execute(
            """
            INSERT INTO usuario (username, nome, email, senha_hash, admin)
            VALUES (%s, %s, %s, %s, %s)
            """,
            (
                username,
                (nome or "").strip()[:200],
                (email or "").strip()[:200],
                hash_senha(senha),
                1 if admin else 0,
            ),
        )
        return cur.lastrowid


def atualizar_usuario(
    usuario_id: int,
    nome: str | None = None,
    email: str | None = None,
    admin: bool | None = None,
    ativo: bool | None = None,
) -> None:
    """Atualiza campos passados. Não mexe em senha (use mudar_senha)."""
    campos: list[str] = []
    valores: list[Any] = []
    if nome is not None:
        campos.append("nome = %s")
        valores.append(nome.strip()[:200])
    if email is not None:
        campos.append("email = %s")
        valores.append(email.strip()[:200])
    if admin is not None:
        campos.append("admin = %s")
        valores.append(1 if admin else 0)
    if ativo is not None:
        campos.append("ativo = %s")
        valores.append(1 if ativo else 0)
    if not campos:
        return
    valores.append(usuario_id)
    with conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            f"UPDATE usuario SET {', '.join(campos)} WHERE id = %s",
            valores,
        )


def mudar_senha(usuario_id: int, senha_nova: str) -> None:
    if not senha_nova or len(senha_nova) < SENHA_MIN_LEN:
        raise UsuarioErro(f"Senha muito curta (mínimo {SENHA_MIN_LEN} caracteres).")
    with conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE usuario SET senha_hash = %s WHERE id = %s",
            (hash_senha(senha_nova), usuario_id),
        )


def get_empresa_ativa(usuario_id: int) -> dict[str, Any] | None:
    """Devolve o dict da empresa ativa do usuário, ou None."""
    with conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT empresa_ativa FROM usuario WHERE id = %s",
            (usuario_id,),
        )
        row = cur.fetchone()
        if not row:
            return None
        return _decode_json(row[0])


def set_empresa_ativa(
    usuario_id: int,
    empresa: dict[str, Any] | None,
) -> None:
    """Salva a empresa ativa do usuário (dict com codi_emp/razao/cnpj) ou None."""
    valor = json.dumps(empresa, ensure_ascii=False, default=str) if empresa else None
    with conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            "UPDATE usuario SET empresa_ativa = %s WHERE id = %s",
            (valor, usuario_id),
        )
