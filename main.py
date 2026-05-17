# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — SERVIDOR DE AUTENTICAÇÃO
# FastAPI + SQLite
# =====================================================================

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from datetime import datetime
import database
import auth

app = FastAPI(title="RWA Tecnologia Operacional — Servidor")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

database.inicializar_banco()


# ── Modelos ────────────────────────────────────────────────────────

class LoginRequest(BaseModel):
    email:       str
    senha:       str
    fingerprint: str
    versao:      str = "1.0"

class LoginResponse(BaseModel):
    status:     str
    mensagem:   str
    chave_aes:  str = ""
    cliente:    str = ""
    vencimento: str = ""


# ── Endpoint de login ──────────────────────────────────────────────

@app.post("/auth/login", response_model=LoginResponse)
def login(req: LoginRequest):

    # 1. Busca empresa
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return LoginResponse(status="erro", mensagem="Email não cadastrado.")

    # 2. Verifica senha
    if not auth.verificar_senha(req.senha, empresa["senha_hash"]):
        return LoginResponse(status="erro", mensagem="Senha incorreta.")

    # 3. Verifica vencimento
    hoje = datetime.now().date()
    venc = datetime.strptime(empresa["vencimento"], "%Y-%m-%d").date()
    if hoje > venc:
        return LoginResponse(status="erro", mensagem="Licença vencida. Entre em contato com a RWA.")

    # 4. Verifica máquina
    maquinas = database.listar_maquinas(empresa["id"])
    fps = [m["fingerprint"] for m in maquinas]

    if req.fingerprint not in fps:
        if len(fps) == 0:
            # Primeira máquina — registra automaticamente
            database.registrar_maquina(empresa["id"], req.fingerprint)
        else:
            # Máquina não autorizada — notifica
            auth.notificar_maquina_nao_autorizada(
                empresa["nome"], req.email, req.fingerprint
            )
            return LoginResponse(
                status="erro",
                mensagem="Máquina não autorizada. Entre em contato com a RWA."
            )

    # 5. Registra acesso
    database.registrar_acesso(empresa["id"], req.fingerprint, req.versao)

    return LoginResponse(
        status="ok",
        mensagem="Acesso autorizado.",
        chave_aes=empresa["chave_aes"],
        cliente=empresa["nome"],
        vencimento=empresa["vencimento"],
    )


# ── Health check ───────────────────────────────────────────────────

@app.get("/")
def root():
    return {"status": "online", "sistema": "RWA Tecnologia Operacional"}


# ── Modelos agente ─────────────────────────────────────────────────

class TarefaRequest(BaseModel):
    email:       str
    fingerprint: str

class StatusRequest(BaseModel):
    email:       str
    fingerprint: str
    tarefa_id:   str
    status:      str
    observacao:  str = ""


# ── Endpoints do agente ────────────────────────────────────────────

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
