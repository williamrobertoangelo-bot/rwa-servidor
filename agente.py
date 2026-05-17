# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — AGENTE LOCAL
# Roda em segundo plano. Consulta servidor. Executa comandos.
# =====================================================================

import os
import sys
import time
import uuid
import socket
import hashlib
import platform
import requests
import json
import subprocess
import threading
import logging
from datetime import datetime
from pathlib import Path

# ── Configuração ───────────────────────────────────────────────────
SERVIDOR_URL    = "https://web-production-31152.up.railway.app"
INTERVALO_SEG   = 1        # consulta a cada 1 segundo
PASTA_BASE      = Path("C:/RWA_AUTOMACOES")
PASTA_CONFIG    = PASTA_BASE / "CONFIG"
PASTA_SENHAS    = PASTA_BASE / "SENHAS"
PASTA_LOGS      = PASTA_BASE / "LOGS"
ARQUIVO_CRED    = PASTA_CONFIG / "credenciais.json"
ARQUIVO_LOG     = PASTA_LOGS  / "agente.log"

# ── Logging ────────────────────────────────────────────────────────
PASTA_LOGS.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(ARQUIVO_LOG, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger("RWA_AGENTE")


# ── Fingerprint da máquina ─────────────────────────────────────────

def gerar_fingerprint() -> str:
    dados = "|".join([
        socket.gethostname(),
        getpass_user(),
        platform.system(),
        platform.node(),
        str(uuid.getnode()),
    ])
    return hashlib.sha256(dados.encode()).hexdigest()


def getpass_user() -> str:
    try:
        import getpass
        return getpass.getuser()
    except Exception:
        return "desconhecido"


# ── Credenciais salvas localmente ──────────────────────────────────

def carregar_credenciais():
    if ARQUIVO_CRED.exists():
        with open(ARQUIVO_CRED, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def salvar_credenciais(email: str, token: str):
    PASTA_CONFIG.mkdir(parents=True, exist_ok=True)
    with open(ARQUIVO_CRED, "w", encoding="utf-8") as f:
        json.dump({"email": email, "token": token}, f)


# ── Comunicação com servidor ───────────────────────────────────────

def verificar_tarefa(email: str, fingerprint: str):
    try:
        resp = requests.post(
            f"{SERVIDOR_URL}/agente/tarefa",
            json={"email": email, "fingerprint": fingerprint},
            timeout=5
        )
        if resp.status_code == 200:
            return resp.json()
    except Exception as e:
        log.warning(f"Falha ao consultar servidor: {e}")
    return None


def enviar_status(email: str, fingerprint: str, tarefa_id: str, status: str, obs: str = ""):
    try:
        requests.post(
            f"{SERVIDOR_URL}/agente/status",
            json={
                "email":      email,
                "fingerprint": fingerprint,
                "tarefa_id":  tarefa_id,
                "status":     status,
                "observacao": obs,
            },
            timeout=5
        )
    except Exception as e:
        log.warning(f"Falha ao enviar status: {e}")


# ── Executor de tarefas ────────────────────────────────────────────

def executar_tarefa(tarefa: dict, email: str, fingerprint: str):
    tarefa_id = tarefa.get("id", "")
    modulo    = tarefa.get("modulo", "")
    parametros = tarefa.get("parametros", {})

    log.info(f"[TAREFA] Iniciando: {modulo} | ID: {tarefa_id}")
    enviar_status(email, fingerprint, tarefa_id, "em_execucao")

    try:
        # Aqui vai conectar com a automação real futuramente
        # Por enquanto retorna sucesso simulado
        log.info(f"[TAREFA] Módulo: {modulo} | Parâmetros: {parametros}")
        time.sleep(2)  # placeholder
        log.info(f"[TAREFA] Concluída: {modulo}")
        enviar_status(email, fingerprint, tarefa_id, "concluido")

    except Exception as e:
        log.error(f"[TAREFA] Erro: {e}")
        enviar_status(email, fingerprint, tarefa_id, "erro", str(e))


# ── Loop principal ─────────────────────────────────────────────────

def loop_principal(email: str, fingerprint: str):
    log.info(f"[AGENTE] Iniciado. Email: {email} | FP: {fingerprint[:16]}...")
    log.info(f"[AGENTE] Consultando servidor a cada {INTERVALO_SEG}s")

    tarefa_em_execucao = False

    while True:
        try:
            if not tarefa_em_execucao:
                resultado = verificar_tarefa(email, fingerprint)
                if resultado and resultado.get("tarefa"):
                    tarefa = resultado["tarefa"]
                    tarefa_em_execucao = True
                    t = threading.Thread(
                        target=_executar_e_liberar,
                        args=(tarefa, email, fingerprint),
                        daemon=True
                    )
                    t.start()
        except Exception as e:
            log.error(f"[AGENTE] Erro no loop: {e}")

        time.sleep(INTERVALO_SEG)


def _executar_e_liberar(tarefa, email, fingerprint):
    global tarefa_em_execucao
    try:
        executar_tarefa(tarefa, email, fingerprint)
    finally:
        tarefa_em_execucao = False


# ── Entrada ────────────────────────────────────────────────────────

def main():
    # Cria estrutura de pastas padrão
    for pasta in [PASTA_BASE, PASTA_CONFIG, PASTA_SENHAS, PASTA_LOGS]:
        pasta.mkdir(parents=True, exist_ok=True)

    fingerprint = gerar_fingerprint()
    cred = carregar_credenciais()

    if not cred:
        log.error("[AGENTE] Sem credenciais. Execute o setup primeiro.")
        sys.exit(1)

    email = cred["email"]
    loop_principal(email, fingerprint)


if __name__ == "__main__":
    main()
