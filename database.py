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
                    criado_em   TEXT    NOT NULL DEFAULT ((NOW() AT TIME ZONE 'America/Sao_Paulo')::TEXT)
                );
                CREATE TABLE IF NOT EXISTS maquinas (
                    id          SERIAL PRIMARY KEY,
                    empresa_id  INTEGER NOT NULL,
                    fingerprint TEXT    NOT NULL,
                    registrado  TEXT    NOT NULL DEFAULT ((NOW() AT TIME ZONE 'America/Sao_Paulo')::TEXT),
                    UNIQUE(empresa_id, fingerprint),
                    FOREIGN KEY (empresa_id) REFERENCES empresas(id)
                );
                CREATE TABLE IF NOT EXISTS acessos (
                    id          SERIAL PRIMARY KEY,
                    empresa_id  INTEGER NOT NULL,
                    fingerprint TEXT    NOT NULL,
                    versao      TEXT,
                    quando      TEXT    NOT NULL DEFAULT ((NOW() AT TIME ZONE 'America/Sao_Paulo')::TEXT),
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


def buscar_empresa_sem_senha(email):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("SELECT * FROM empresas WHERE email=%s AND ativa=1", (email,))
            row = cur.fetchone()
    return dict(row) if row else None


def definir_senha(email, senha_hash):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("UPDATE empresas SET senha_hash=%s WHERE email=%s", (senha_hash, email))
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
                    criado_em    TEXT    NOT NULL DEFAULT ((NOW() AT TIME ZONE 'America/Sao_Paulo')::TEXT),
                    atualizado_em TEXT   DEFAULT '',
                    FOREIGN KEY (empresa_id) REFERENCES empresas(id)
                );
            """)
        con.commit()


def criar_tarefa(empresa_id, modulo, parametros):
    import json
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute(
                "INSERT INTO tarefas (empresa_id, modulo, parametros, status) VALUES (%s,%s,%s,'pendente')",
                (empresa_id, modulo, json.dumps(parametros))
            )
        con.commit()


def criar_tarefa_agendada(empresa_id, modulo, agendado_para):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute(
                "INSERT INTO tarefas (empresa_id, modulo, parametros, status, agendado_para) "
                "VALUES (%s,%s,'{}','agendado',%s)",
                (empresa_id, modulo, agendado_para)
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
            cur.execute(
                "UPDATE tarefas SET status=%s, observacao=%s WHERE id=%s",
                (status, observacao, tarefa_id)
            )
        con.commit()


def cancelar_tarefa(empresa_id, tarefa_id):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute(
                "UPDATE tarefas SET status='cancelado', observacao='Cancelado pelo usuário' WHERE id=%s AND empresa_id=%s AND status IN ('pendente','em_execucao','agendado')",
                (tarefa_id, empresa_id)
            )
            linhas = cur.rowcount
        con.commit()
    return linhas > 0


def buscar_status_tarefa(tarefa_id):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("SELECT status FROM tarefas WHERE id=%s", (tarefa_id,))
            row = cur.fetchone()
    return dict(row)["status"] if row else None


def buscar_historico(empresa_id, limite=10):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                SELECT id, modulo, status, criado_em as quando, observacao
                FROM tarefas WHERE empresa_id=%s
                ORDER BY criado_em DESC LIMIT %s
            """, (empresa_id, limite))
            rows = cur.fetchall()
    return [dict(r) for r in rows]


# ── CONFERÊNCIAS ────────────────────────────────────────────────────

def criar_tabela_conferencias():
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS conferencias (
                    id            SERIAL PRIMARY KEY,
                    empresa_id    INTEGER NOT NULL,
                    modulo        TEXT    NOT NULL,
                    competencia   TEXT    NOT NULL,
                    gerado_em     TEXT    NOT NULL DEFAULT ((NOW() AT TIME ZONE 'America/Sao_Paulo')::TEXT),
                    total         INTEGER NOT NULL DEFAULT 0,
                    total_ok      INTEGER NOT NULL DEFAULT 0,
                    total_div     INTEGER NOT NULL DEFAULT 0,
                    total_pend    INTEGER NOT NULL DEFAULT 0,
                    total_parc    INTEGER NOT NULL DEFAULT 0,
                    total_sem_mov INTEGER NOT NULL DEFAULT 0,
                    resultados    TEXT    NOT NULL DEFAULT '[]',
                    sem_movimento TEXT    NOT NULL DEFAULT '[]',
                    FOREIGN KEY (empresa_id) REFERENCES empresas(id)
                );
            """)
        con.commit()


def salvar_conferencia(empresa_id, modulo, competencia, resultados, sem_movimento):
    import json
    total     = len(resultados)
    total_ok  = sum(1 for r in resultados if r.get("status") == "OK")
    total_div = sum(1 for r in resultados if r.get("status") == "DIVERGENTE")
    total_pend= sum(1 for r in resultados if r.get("status") == "PENDENCIA")
    total_parc= sum(1 for r in resultados if r.get("status") == "PARCIAL")
    total_sm  = len(sem_movimento)
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                INSERT INTO conferencias
                    (empresa_id, modulo, competencia, total, total_ok, total_div,
                     total_pend, total_parc, total_sem_mov, resultados, sem_movimento)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                empresa_id, modulo, competencia,
                total, total_ok, total_div, total_pend, total_parc, total_sm,
                json.dumps(resultados, ensure_ascii=False),
                json.dumps(sem_movimento, ensure_ascii=False),
            ))
        con.commit()


def buscar_conferencia(empresa_id, modulo, competencia=None):
    import json
    with _conn() as con:
        with con.cursor() as cur:
            if competencia:
                cur.execute("""
                    SELECT * FROM conferencias
                    WHERE empresa_id=%s AND modulo=%s AND competencia=%s
                    ORDER BY gerado_em DESC LIMIT 1
                """, (empresa_id, modulo, competencia))
            else:
                cur.execute("""
                    SELECT * FROM conferencias
                    WHERE empresa_id=%s AND modulo=%s
                    ORDER BY gerado_em DESC LIMIT 1
                """, (empresa_id, modulo))
            row = cur.fetchone()
    if not row:
        return None
    r = dict(row)
    r["resultados"]    = json.loads(r["resultados"])
    r["sem_movimento"] = json.loads(r["sem_movimento"])
    return r


def listar_competencias_conferencia(empresa_id, modulo):
    with _conn() as con:
        with con.cursor() as cur:
            cur.execute("""
                SELECT DISTINCT competencia, gerado_em
                FROM conferencias
                WHERE empresa_id=%s AND modulo=%s
                ORDER BY competencia DESC
            """, (empresa_id, modulo))
            rows = cur.fetchall()
    return [dict(r) for r in rows]
