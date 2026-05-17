# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — BANCO DE DADOS (PostgreSQL)
# =====================================================================

import os
import psycopg2
from psycopg2.extras import RealDictCursor

DATABASE_URL = os.environ.get("DATABASE_URL")


def _conn():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)


def inicializar_banco():
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS empresas (
                    id          SERIAL PRIMARY KEY,
                    nome        TEXT    NOT NULL,
                    email       TEXT    NOT NULL UNIQUE,
                    senha_hash  TEXT    NOT NULL,
                    vencimento  TEXT    NOT NULL,
                    chave_aes   TEXT    NOT NULL,
                    ativa       INTEGER NOT NULL DEFAULT 1,
                    criado_em   TEXT    NOT NULL DEFAULT (NOW()::TEXT)
                );
                CREATE TABLE IF NOT EXISTS maquinas (
                    id          SERIAL PRIMARY KEY,
                    empresa_id  INTEGER NOT NULL,
                    fingerprint TEXT    NOT NULL,
                    registrado  TEXT    NOT NULL DEFAULT (NOW()::TEXT),
                    UNIQUE(empresa_id, fingerprint),
                    FOREIGN KEY (empresa_id) REFERENCES empresas(id)
                );
                CREATE TABLE IF NOT EXISTS acessos (
                    id          SERIAL PRIMARY KEY,
                    empresa_id  INTEGER NOT NULL,
                    fingerprint TEXT    NOT NULL,
                    versao      TEXT,
                    quando      TEXT    NOT NULL DEFAULT (NOW()::TEXT),
                    FOREIGN KEY (empresa_id) REFERENCES empresas(id)
                );
            """)
        con.commit()


def cadastrar_empresa(nome, email, senha_hash, vencimento, chave_aes):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute(
                "INSERT INTO empresas (nome, email, senha_hash, vencimento, chave_aes) VALUES (%s,%s,%s,%s,%s)",
                (nome, email, senha_hash, vencimento, chave_aes)
            )
        con.commit()


def buscar_empresa(email):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("SELECT * FROM empresas WHERE email=%s AND ativa=1", (email,))
            row = cur.fetchone()
    return dict(row) if row else None


def listar_empresas():
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("SELECT * FROM empresas ORDER BY id")
            rows = cur.fetchall()
    return [dict(r) for r in rows]


def registrar_maquina(empresa_id, fingerprint):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute(
                "INSERT INTO maquinas (empresa_id, fingerprint) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                (empresa_id, fingerprint)
            )
        con.commit()


def listar_maquinas(empresa_id):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("SELECT * FROM maquinas WHERE empresa_id=%s", (empresa_id,))
            rows = cur.fetchall()
    return [dict(r) for r in rows]


def remover_maquina(empresa_id, fingerprint):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("DELETE FROM maquinas WHERE empresa_id=%s AND fingerprint=%s", (empresa_id, fingerprint))
        con.commit()


def renovar_licenca(email, novo_vencimento):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("UPDATE empresas SET vencimento=%s WHERE email=%s", (novo_vencimento, email))
        con.commit()


def bloquear_empresa(email):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("UPDATE empresas SET ativa=0 WHERE email=%s", (email,))
        con.commit()


def registrar_acesso(empresa_id, fingerprint, versao):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute(
                "INSERT INTO acessos (empresa_id, fingerprint, versao) VALUES (%s,%s,%s)",
                (empresa_id, fingerprint, versao)
            )
        con.commit()


def buscar_proxima_tarefa(empresa_id):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                SELECT * FROM tarefas 
                WHERE empresa_id=%s AND status='pendente'
                ORDER BY criado_em ASC LIMIT 1
            """, (empresa_id,))
            row = cur.fetchone()
    return dict(row) if row else None


def atualizar_status_tarefa(tarefa_id, status, observacao=""):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                UPDATE tarefas SET status=%s, observacao=%s, atualizado_em=NOW()::TEXT
                WHERE id=%s
            """, (status, observacao, tarefa_id))
        con.commit()


def criar_tabela_tarefas():
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS tarefas (
                    id           SERIAL PRIMARY KEY,
                    empresa_id   INTEGER NOT NULL,
                    modulo       TEXT    NOT NULL,
                    parametros   TEXT    DEFAULT '{}',
                    status       TEXT    NOT NULL DEFAULT 'pendente',
                    observacao   TEXT    DEFAULT '',
                    criado_em    TEXT    NOT NULL DEFAULT (NOW()::TEXT),
                    atualizado_em TEXT   DEFAULT '',
                    FOREIGN KEY (empresa_id) REFERENCES empresas(id)
                );
            """)
        con.commit()
