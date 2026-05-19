from pathlib import Path
# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — SERVIDOR
# =====================================================================

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel
from datetime import datetime
import database
import auth
import os

app = FastAPI(title="RWA Tecnologia Operacional")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

if Path("static").exists():
    app.mount("/static", StaticFiles(directory="static"), name="static")

database.inicializar_banco()


# ── Models ─────────────────────────────────────────────────────────

class LoginRequest(BaseModel):
    email:    str
    senha:    str

class PrimeiroAcessoRequest(BaseModel):
    email:    str
    senha:    str

class AgenteLoginRequest(BaseModel):
    email:       str
    senha:       str
    fingerprint: str = ""
    versao:      str = "1.0"

class TarefaRequest(BaseModel):
    email:       str
    fingerprint: str

class StatusRequest(BaseModel):
    email:       str
    fingerprint: str
    tarefa_id:   str
    status:      str
    observacao:  str = ""

class ExecutarRequest(BaseModel):
    email:       str
    modulo:      str

class AgendarRequest(BaseModel):
    email:         str
    modulo:        str
    agendado_para: str

class CancelarRequest(BaseModel):
    email:     str
    tarefa_id: int

class RegistrarMaquinaRequest(BaseModel):
    email:       str
    fingerprint: str

class StatusTarefaRequest(BaseModel):
    email:       str
    fingerprint: str
    tarefa_id:   str


# ── Portal (web — sem fingerprint) ─────────────────────────────────

@app.get("/")
def root():
    return FileResponse("static/index.html")


@app.post("/auth/login")
def login(req: LoginRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"status": "erro", "mensagem": "Email não cadastrado."}

    if not auth.verificar_senha(req.senha, empresa["senha_hash"]):
        return {"status": "erro", "mensagem": "Senha incorreta."}

    hoje = datetime.now().date()
    venc = datetime.strptime(empresa["vencimento"], "%Y-%m-%d").date()
    if hoje > venc:
        return {"status": "erro", "mensagem": "Licença vencida. Entre em contato com a RWA."}

    return {
        "status":     "ok",
        "mensagem":   "Acesso autorizado.",
        "cliente":    empresa["nome"],
        "vencimento": empresa["vencimento"],
        "email":      req.email,
    }


@app.post("/auth/primeiro-acesso")
def primeiro_acesso(req: PrimeiroAcessoRequest):
    empresa = database.buscar_empresa_sem_senha(req.email)
    if not empresa:
        return {"status": "erro", "mensagem": "Email não encontrado. Entre em contato com a RWA."}

    if empresa.get("senha_hash"):
        return {"status": "erro", "mensagem": "Esse email já possui senha. Use a tela de login."}

    senha_hash = auth.hash_senha(req.senha)
    database.definir_senha(req.email, senha_hash)

    return {
        "status":     "ok",
        "mensagem":   "Senha criada com sucesso.",
        "cliente":    empresa["nome"],
        "vencimento": empresa["vencimento"],
        "email":      req.email,
    }


@app.post("/portal/executar")
def portal_executar(req: ExecutarRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}

    database.criar_tarefa(empresa["id"], req.modulo, {})
    return {"ok": True}


@app.post("/portal/agendar")
def portal_agendar(req: AgendarRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}

    database.criar_tarefa_agendada(empresa["id"], req.modulo, req.agendado_para)
    return {"ok": True}


@app.post("/portal/cancelar")
def portal_cancelar(req: CancelarRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}

    ok = database.cancelar_tarefa(empresa["id"], req.tarefa_id)
    if not ok:
        return {"ok": False, "erro": "Tarefa não localizada ou não cancelável."}

    return {"ok": True}


@app.get("/portal/historico")
def portal_historico(email: str):
    try:
        empresa = database.buscar_empresa(email)
        if not empresa:
            return {"historico": []}
        historico = database.buscar_historico(empresa["id"])
        return {"historico": historico}
    except Exception:
        return {"historico": []}


# ── Agente (com fingerprint) ────────────────────────────────────────

@app.post("/agente/login")
def agente_login(req: AgenteLoginRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"status": "erro", "mensagem": "Email não cadastrado."}

    if not auth.verificar_senha(req.senha, empresa["senha_hash"]):
        return {"status": "erro", "mensagem": "Senha incorreta."}

    hoje = datetime.now().date()
    venc = datetime.strptime(empresa["vencimento"], "%Y-%m-%d").date()
    if hoje > venc:
        return {"status": "erro", "mensagem": "Licença vencida. Entre em contato com a RWA."}

    return {
        "status":     "ok",
        "mensagem":   "Acesso autorizado.",
        "cliente":    empresa["nome"],
        "vencimento": empresa["vencimento"],
    }


@app.post("/agente/tarefa")
def agente_tarefa(req: TarefaRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"tarefa": None}

    maquinas = database.listar_maquinas(empresa["id"])
    fps = [m["fingerprint"] for m in maquinas]
    if req.fingerprint not in fps:
        return {"tarefa": None, "erro": "maquina_nao_autorizada"}

    tarefa = database.buscar_proxima_tarefa(empresa["id"])
    return {"tarefa": tarefa}


@app.post("/agente/status")
def agente_status(req: StatusRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False}

    database.atualizar_status_tarefa(req.tarefa_id, req.status, req.observacao)
    return {"ok": True}


@app.post("/agente/status_tarefa")
def agente_status_tarefa(req: StatusTarefaRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"status": ""}

    status = database.buscar_status_tarefa(req.tarefa_id)
    return {"status": status or ""}


@app.post("/admin/registrar-maquina")
def admin_registrar_maquina(req: RegistrarMaquinaRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}
    database.registrar_maquina(empresa["id"], req.fingerprint)
    return {"ok": True, "mensagem": f"Máquina registrada para {empresa['nome']}"}


@app.get("/health")
def health():
    return {"status": "online", "sistema": "RWA Tecnologia Operacional"}


# ── SPA Fallback (deve ser a última rota) ──────────────────────────

@app.get("/{full_path:path}")
def spa_fallback(full_path: str):
    return FileResponse("static/index.html")
