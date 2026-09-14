"""Setup inicial do banco MariaDB. Rode uma vez:

    python setup_db.py

O que faz:
1. Pergunta credenciais (host, usuário, senha, nome do banco) - ou reusa
   as que já estão salvas em ``data/db_config.json``.
2. Testa conexão.
3. Cria o banco (CREATE DATABASE IF NOT EXISTS).
4. Cria as 3 tabelas do app.
5. Salva as credenciais em ``data/db_config.json``.
6. Migra o ``config.json`` existente pro banco (se houver).

Pode ser re-executado com segurança - CREATE IF NOT EXISTS não sobrescreve
tabelas nem dados.
"""

from __future__ import annotations

import getpass
import json
from pathlib import Path

import pymysql

import db

CONFIG_JSON_PATH = Path(__file__).parent / "config.json"

DDL = [
    # Tabela chave/valor pra configs globais (empresa ativa, fontes, etc)
    """
    CREATE TABLE IF NOT EXISTS app_config (
        chave         VARCHAR(100) PRIMARY KEY,
        valor         JSON,
        atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
                      ON UPDATE CURRENT_TIMESTAMP
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
      COLLATE=utf8mb4_unicode_ci
    """,
    # Regras de classificação contábil (memo/fornecedor), uma linha por regra
    """
    CREATE TABLE IF NOT EXISTS regra_taxa (
        id         INT AUTO_INCREMENT PRIMARY KEY,
        codi_emp   INT NOT NULL,
        tipo       ENUM('memo', 'fornecedor') NOT NULL,
        padrao     VARCHAR(500) NOT NULL,
        historico  VARCHAR(500),
        conta      VARCHAR(100),
        banco      VARCHAR(200),
        ordem      INT NOT NULL DEFAULT 0,
        criado_em  DATETIME DEFAULT CURRENT_TIMESTAMP,
        INDEX idx_empresa (codi_emp, ordem)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
      COLLATE=utf8mb4_unicode_ci
    """,
    # Mapeamento de colunas da planilha por empresa
    """
    CREATE TABLE IF NOT EXISTS mapeamento_planilha (
        id           INT AUTO_INCREMENT PRIMARY KEY,
        codi_emp     INT NOT NULL,
        campo        VARCHAR(50) NOT NULL,
        nome_coluna  VARCHAR(200) NOT NULL,
        UNIQUE KEY uniq_emp_campo (codi_emp, campo)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
      COLLATE=utf8mb4_unicode_ci
    """,
    # Usuários do app (login com senha via PBKDF2). Cada um tem sua
    # empresa_ativa individual — não confundir com regras/mapeamentos que
    # continuam por empresa (compartilhados).
    """
    CREATE TABLE IF NOT EXISTS usuario (
        id            INT AUTO_INCREMENT PRIMARY KEY,
        username      VARCHAR(50) NOT NULL UNIQUE,
        nome          VARCHAR(200),
        email         VARCHAR(200),
        senha_hash    VARCHAR(500) NOT NULL,
        ativo         TINYINT(1) NOT NULL DEFAULT 1,
        admin         TINYINT(1) NOT NULL DEFAULT 0,
        empresa_ativa JSON,
        ultimo_login  DATETIME NULL,
        criado_em     DATETIME DEFAULT CURRENT_TIMESTAMP,
        INDEX idx_ativo (ativo)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
      COLLATE=utf8mb4_unicode_ci
    """,
]


def pergunta(pergunta: str, default: str = "") -> str:
    sufixo = f" [{default}]" if default else ""
    resp = input(f"{pergunta}{sufixo}: ").strip()
    return resp or default


def pergunta_senha() -> str:
    return getpass.getpass("Senha do MariaDB (não aparece na tela): ")


def coletar_credenciais(atuais: dict) -> dict:
    print("=" * 60)
    print("Setup do MariaDB - Conciliador")
    print("=" * 60)
    print("Confirme (ou altere) as credenciais de acesso ao MariaDB.")
    print()
    cfg = dict(atuais)
    cfg["host"] = pergunta("Host", cfg.get("host", "localhost"))
    cfg["port"] = int(pergunta("Porta", str(cfg.get("port", 3306))))
    cfg["user"] = pergunta("Usuário", cfg.get("user", "root"))
    senha_nova = pergunta_senha()
    if senha_nova:
        cfg["password"] = senha_nova
    cfg["database"] = pergunta("Nome do banco", cfg.get("database", "conciliador"))
    return cfg


def criar_banco(cfg: dict) -> None:
    """Cria o database se ainda não existir. Conecta SEM database para
    poder rodar o CREATE DATABASE."""
    print(f"\n-> Criando banco '{cfg['database']}' (se ainda não existir)...")
    with db.conexao(cfg=cfg, skip_db=True) as conn:
        cur = conn.cursor()
        cur.execute(
            f"CREATE DATABASE IF NOT EXISTS `{cfg['database']}` "
            "DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        )
    print(f"  [OK]Banco pronto.")


def criar_tabelas(cfg: dict) -> None:
    print("\n-> Criando tabelas (se ainda não existirem)...")
    with db.conexao(cfg=cfg) as conn:
        cur = conn.cursor()
        for ddl in DDL:
            cur.execute(ddl)
    print("  [OK] Tabelas prontas: app_config, regra_taxa, "
          "mapeamento_planilha, usuario")


def migrar_config_json() -> int:
    """Se ``config.json`` existir e o banco estiver vazio, migra os dados.
    Devolve quantidade total de itens migrados (0 = não migrou nada)."""
    if not CONFIG_JSON_PATH.exists():
        print("\n-> Migração: nenhum config.json antigo encontrado - nada a fazer.")
        return 0

    try:
        dados = json.loads(CONFIG_JSON_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        print(f"\n[AVISO]config.json existe mas não pôde ser lido: {e}")
        return 0

    if not dados:
        print("\n-> Migração: config.json está vazio - nada a fazer.")
        return 0

    # Checa se o banco já tem dados (evita sobrescrever config existente)
    with db.conexao() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM app_config")
        n_app = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM regra_taxa")
        n_reg = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM mapeamento_planilha")
        n_map = cur.fetchone()[0]
    total_existente = n_app + n_reg + n_map
    if total_existente > 0:
        print(
            f"\n-> Migração: banco já contém {total_existente} registros - "
            "config.json antigo NÃO será importado (segurança)."
        )
        return 0

    print("\n-> Migração: importando config.json antigo pro banco...")
    # Importa via a nova config.py (que sabe fazer o upsert direito)
    import config
    config.salvar(dados)

    # Renomeia pra .bak pra não confundir depois
    bak = CONFIG_JSON_PATH.with_suffix(".json.bak")
    CONFIG_JSON_PATH.rename(bak)
    print(f"  [OK]Migração feita. Arquivo antigo salvo em: {bak.name}")
    return 1  # sinaliza que migrou


def bootstrap_admin() -> None:
    """Se ainda não existe nenhum usuário, cadastra o primeiro admin
    interativamente. Idempotente — pode rodar de novo, só age se vazio."""
    import auth

    if auth.existe_algum_usuario():
        with db.conexao() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM usuario")
            n = cur.fetchone()[0]
        print(f"\n-> {n} usuario(s) ja cadastrado(s). Pulando bootstrap.")
        return

    print()
    print("=" * 60)
    print("Cadastro do PRIMEIRO usuario administrador")
    print("=" * 60)
    print("Sem admin ninguem consegue entrar no app. Cadastre agora.")
    print()

    username = ""
    while not username:
        username = pergunta("Username").strip()
        if not username:
            print("  [ERRO] Username obrigatorio.")

    nome = pergunta("Nome completo (opcional)")
    email = pergunta("Email (opcional)")

    while True:
        senha = pergunta_senha()
        if len(senha) < auth.SENHA_MIN_LEN:
            print(f"  [ERRO] Senha muito curta (min {auth.SENHA_MIN_LEN}).")
            continue
        senha2 = getpass.getpass("Confirme a senha: ")
        if senha != senha2:
            print("  [ERRO] Senhas nao conferem, tente de novo.")
            continue
        break

    try:
        uid = auth.criar_usuario(username, senha, nome=nome, email=email, admin=True)
    except auth.UsuarioErro as e:
        print(f"  [ERRO] {e}")
        return

    print(f"  [OK] Admin '{username}' criado (id={uid}).")


def migrar_empresa_ativa_para_admin() -> None:
    """Se o app_config ainda tem 'dominio_empresa' (config antigo global),
    move pra empresa_ativa do primeiro admin cadastrado."""
    import auth

    with db.conexao() as conn:
        cur = conn.cursor()
        cur.execute(
            "SELECT valor FROM app_config WHERE chave = 'dominio_empresa'"
        )
        row = cur.fetchone()
        if not row:
            return
        try:
            emp = row[0] if isinstance(row[0], dict) else json.loads(row[0])
        except (TypeError, ValueError, json.JSONDecodeError):
            return
        if not isinstance(emp, dict) or emp.get("codi_emp") is None:
            return

        cur.execute(
            "SELECT id, username FROM usuario "
            "WHERE ativo = 1 AND admin = 1 ORDER BY id LIMIT 1"
        )
        admin_row = cur.fetchone()
        if not admin_row:
            print("  [AVISO] Nao ha admin ativo pra receber a empresa migrada.")
            return

    auth.set_empresa_ativa(admin_row[0], emp)
    with db.conexao() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM app_config WHERE chave = 'dominio_empresa'")
    print(
        f"  [OK] dominio_empresa migrada pro admin '{admin_row[1]}' "
        f"(empresa {emp.get('codi_emp')})."
    )


def main() -> int:
    print()
    atuais = db.load_db_config()
    cfg = coletar_credenciais(atuais)

    print("\n-> Testando conexão (sem selecionar banco)...")
    try:
        conn_teste = pymysql.connect(**db._build_kwargs(cfg, skip_db=True))
        cur = conn_teste.cursor()
        cur.execute("SELECT VERSION()")
        versao = cur.fetchone()[0]
        conn_teste.close()
        print(f"  [OK]Conectado - MariaDB/MySQL {versao}")
    except Exception as e:
        print(f"  [ERRO]Falha na conexão: {e}")
        print("\nVerifique se o MariaDB está rodando e as credenciais estão certas.")
        return 1

    try:
        criar_banco(cfg)
        criar_tabelas(cfg)
    except Exception as e:
        print(f"\nX Erro ao criar banco/tabelas: {e}")
        return 1

    # Salva as credenciais só depois que tudo deu certo
    db.save_db_config(cfg)
    print(f"\n-> Credenciais salvas em: {db.DB_CONFIG_PATH}")

    # Migração (se aplicável)
    try:
        migrar_config_json()
    except Exception as e:
        print(f"\n[AVISO]Migração do config.json falhou: {e}")
        print("O banco está configurado; ajuste manualmente se necessário.")

    # Bootstrap do primeiro admin (interativo se ainda não tiver ninguém)
    try:
        bootstrap_admin()
    except KeyboardInterrupt:
        print("\n[AVISO] Bootstrap cancelado. Rode setup_db.py de novo pra cadastrar.")
        return 1
    except Exception as e:
        print(f"\n[AVISO] Falha no bootstrap: {e}")

    # Move dominio_empresa antigo (global) pra empresa_ativa do primeiro admin
    try:
        migrar_empresa_ativa_para_admin()
    except Exception as e:
        print(f"\n[AVISO] Migracao da empresa ativa falhou: {e}")

    print("\n" + "=" * 60)
    print("Setup concluido. Ja pode rodar o app normal (python main.py).")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
