# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — BANCO DE DADOS (SQLite)
# =====================================================================

import sqlite3
from datetime import datetime

DB_PATH = "rwa.db"


def _conn():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def inicializar_banco():
    with _conn() as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS empresas (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                nome        TEXT    NOT NULL,
                email       TEXT    NOT NULL UNIQUE,
                senha_hash  TEXT    NOT NULL,
                vencimento  TEXT    NOT NULL,
                chave_aes   TEXT    NOT NULL,
                ativa       INTEGER NOT NULL DEFAULT 1,
                criado_em   TEXT    NOT NULL DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS maquinas (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                empresa_id  INTEGER NOT NULL,
                fingerprint TEXT    NOT NULL,
                registrado  TEXT    NOT NULL DEFAULT (datetime('now')),
                UNIQUE(empresa_id, fingerprint),
                FOREIGN KEY (empresa_id) REFERENCES empresas(id)
            );

            CREATE TABLE IF NOT EXISTS acessos (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                empresa_id  INTEGER NOT NULL,
                fingerprint TEXT    NOT NULL,
                versao      TEXT,
                quando      TEXT    NOT NULL DEFAULT (datetime('now')),
                FOREIGN KEY (empresa_id) REFERENCES empresas(id)
            );
        """)


def cadastrar_empresa(nome, email, senha_hash, vencimento, chave_aes):
    with _conn() as con:
        con.execute(
            "INSERT INTO empresas (nome, email, senha_hash, vencimento, chave_aes) "
            "VALUES (?, ?, ?, ?, ?)",
            (nome, email, senha_hash, vencimento, chave_aes)
        )
    print(f"[DB] Empresa cadastrada: {nome} <{email}>")


def buscar_empresa(email):
    with _conn() as con:
        row = con.execute(
            "SELECT * FROM empresas WHERE email=? AND ativa=1", (email,)
        ).fetchone()
    return dict(row) if row else None


def listar_empresas():
    with _conn() as con:
        rows = con.execute("SELECT * FROM empresas ORDER BY id").fetchall()
    if not rows:
        print("  Nenhuma empresa cadastrada.")
        return []
    for r in rows:
        r = dict(r)
        status = "ATIVA" if r["ativa"] else "BLOQUEADA"
        print(f"  [{r['id']}] {r['nome']} | {r['email']} | venc: {r['vencimento']} | {status}")
    return [dict(r) for r in rows]


def registrar_maquina(empresa_id, fingerprint):
    with _conn() as con:
        con.execute(
            "INSERT OR IGNORE INTO maquinas (empresa_id, fingerprint) VALUES (?, ?)",
            (empresa_id, fingerprint)
        )


def listar_maquinas(empresa_id):
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM maquinas WHERE empresa_id=?", (empresa_id,)
        ).fetchall()
    if not rows:
        print("  Nenhuma máquina registrada.")
    for r in rows:
        r = dict(r)
        print(f"  FP: {r['fingerprint']} | desde: {r['registrado']}")
    return [dict(r) for r in rows]


def remover_maquina(empresa_id, fingerprint):
    with _conn() as con:
        con.execute(
            "DELETE FROM maquinas WHERE empresa_id=? AND fingerprint=?",
            (empresa_id, fingerprint)
        )
    print(f"[DB] Máquina removida: {fingerprint}")


def renovar_licenca(email, novo_vencimento):
    with _conn() as con:
        con.execute(
            "UPDATE empresas SET vencimento=? WHERE email=?",
            (novo_vencimento, email)
        )
    print(f"[DB] Licença renovada: {email} → {novo_vencimento}")


def bloquear_empresa(email):
    with _conn() as con:
        con.execute("UPDATE empresas SET ativa=0 WHERE email=?", (email,))
    print(f"[DB] Empresa bloqueada: {email}")


def registrar_acesso(empresa_id, fingerprint, versao):
    with _conn() as con:
        con.execute(
            "INSERT INTO acessos (empresa_id, fingerprint, versao) VALUES (?, ?, ?)",
            (empresa_id, fingerprint, versao)
        )
