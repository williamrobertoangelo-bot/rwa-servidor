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
import uuid
import time


def _formatar_documento(doc: str) -> str:
    d = ''.join(c for c in (doc or '') if c.isdigit())
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return doc or '—'


def _enviar_email_cadastro_senha(empresa: dict):
    api_key = os.environ.get("RESEND_API_KEY", "").strip()
    if not api_key:
        print("[EMAIL] RESEND_API_KEY não configurada.")
        return

    nome      = empresa.get("nome", "—")
    email     = empresa.get("email", "—")
    documento = _formatar_documento(empresa.get("documento", ""))
    telefone  = empresa.get("telefone", "") or "—"
    venc_raw  = empresa.get("vencimento", "")
    try:
        y, m, d  = venc_raw.split("-")
        venc_fmt = f"{d}/{m}/{y}"
    except Exception:
        venc_fmt = venc_raw

    from urllib.parse import quote
    link = f"https://rwasolucoes.com.br/primeiro-acesso?email={quote(email)}"

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"/></head>
<body style="font-family:Arial,sans-serif;background:#f4f4f4;padding:20px;">
<div style="max-width:480px;margin:0 auto;background:#fff;border-radius:6px;padding:24px;border:1px solid #e0e0e0;">
  <div style="background:#0F1B2D;padding:16px;border-radius:4px;margin-bottom:20px;">
    <div style="color:#7FB3E0;font-size:11px;font-weight:bold;letter-spacing:2px;">RWA SOLUÇÕES</div>
    <div style="color:#fff;font-size:16px;font-weight:bold;margin-top:6px;">Cadastro de senha</div>
    <div style="color:#B5D4F4;font-size:12px;margin-top:4px;">Ative seu acesso à plataforma RWA</div>
  </div>
  <p style="font-size:14px;color:#1a1a1a;">Olá, <strong>{nome}</strong>. Sua licença foi aprovada. Cadastre sua senha para ativar o acesso.</p>
  <table width="100%" style="border-collapse:collapse;font-size:13px;margin-bottom:20px;">
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#6B6B6B;">Titular</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">{nome}</td></tr>
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#6B6B6B;">CNPJ/CPF</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">{documento}</td></tr>
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#6B6B6B;">E-mail</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">{email}</td></tr>
    <tr><td style="padding:8px 0;border-bottom:1px solid #eee;color:#6B6B6B;">Telefone</td><td style="padding:8px 0;border-bottom:1px solid #eee;text-align:right;font-weight:bold;">{telefone}</td></tr>
    <tr><td style="padding:8px 0;color:#6B6B6B;">Válida até</td><td style="padding:8px 0;text-align:right;font-weight:bold;">{venc_fmt}</td></tr>
  </table>
  <div style="text-align:center;margin-bottom:16px;">
    <a href="{link}" style="background:#4f46e5;color:#fff;padding:12px 28px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:14px;">Cadastrar minha senha &rarr;</a>
  </div>
  <div style="background:#E6EEFF;border-left:3px solid #4f46e5;padding:10px 12px;font-size:12px;color:#1a1a1a;margin-bottom:16px;">
    Este link é pessoal e intransferível. Após cadastrar, acesse sempre por <strong>rwasolucoes.com.br</strong>
  </div>
  <div style="font-size:11px;color:#888;border-top:1px solid #eee;padding-top:10px;">
    <strong style="color:#0F1B2D;">RWA Soluções</strong> — Automação fiscal para escritórios contábeis
  </div>
</div>
</body>
</html>"""

    try:
        import resend
        resend.api_key = api_key
        resend.Emails.send({
            "from": "RWA Soluções <noreply@rwasolucoes.com.br>",
            "to": [email],
            "subject": "RWA Soluções — Cadastro de senha",
            "html": html,
        })
        print(f"[EMAIL] Enviado via Resend para {email}")
    except Exception as e:
        print(f"[EMAIL] Erro Resend: {e}")


app = FastAPI(title="RWA Tecnologia Operacional")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

if Path("static").exists():
    app.mount("/static", StaticFiles(directory="static"), name="static")

database.inicializar_banco()
database.criar_tabela_tarefas()
database.criar_tabela_conferencias()

_tokens_launcher: dict = {}


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

class GerarTokenLauncherRequest(BaseModel):
    email: str

class ValidarTokenLauncherRequest(BaseModel):
    token:       str
    fingerprint: str


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
        "documento":  empresa.get("documento", ""),
        "telefone":   empresa.get("telefone", ""),
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
        "documento":  empresa.get("documento", ""),
        "telefone":   empresa.get("telefone", ""),
    }


class CadastrarEmpresaRequest(BaseModel):
    nome:       str
    email:      str
    documento:  str = ""
    telefone:   str = ""
    vencimento: str


_ADMIN_EMAIL = "rwaautomacoes@gmail.com"
_ADMIN_SENHA = "Rwa001130@"


class AdminLoginRequest(BaseModel):
    email: str
    senha: str


@app.post("/admin/login")
def admin_login(req: AdminLoginRequest):
    if req.email != _ADMIN_EMAIL or req.senha != _ADMIN_SENHA:
        return {"ok": False, "erro": "Credenciais inválidas."}
    return {"ok": True}


@app.post("/admin/cadastrar-empresa")
def admin_cadastrar_empresa(req: CadastrarEmpresaRequest):
    import secrets
    chave = secrets.token_hex(32).upper()
    try:
        database.cadastrar_empresa(req.nome, req.email, None, req.vencimento, chave, req.documento, req.telefone)
    except Exception as e:
        err = str(e)
        if "unique" in err.lower() or "already exists" in err.lower() or "duplicate" in err.lower():
            return {"ok": False, "erro": "Este e-mail já está cadastrado no sistema."}
        return {"ok": False, "erro": "Erro ao cadastrar. Verifique os dados e tente novamente."}
    empresa = database.buscar_empresa_sem_senha(req.email)
    if empresa:
        try:
            print(f"[EMAIL] Disparando email para {req.email}")
            _enviar_email_cadastro_senha(empresa)
        except Exception as e:
            print(f"[EMAIL] Erro ao enviar email: {e}")
    return {"ok": True, "mensagem": f"Cliente {req.nome} cadastrado e email enviado."}


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

    if req.fingerprint:
        maquinas = database.listar_maquinas(empresa["id"])
        fps = [m["fingerprint"] for m in maquinas]
        if req.fingerprint not in fps:
            if len(fps) == 0:
                database.registrar_maquina(empresa["id"], req.fingerprint)
            else:
                return {"status": "erro", "mensagem": "Licenca ja vinculada a outra maquina. Entre em contato com a RWA."}

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


# ── Token launcher ──────────────────────────────────────────────────

@app.post("/auth/gerar-token-launcher")
def auth_gerar_token_launcher(req: GerarTokenLauncherRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}
    token = uuid.uuid4().hex
    _tokens_launcher[token] = {
        "email":      req.email,
        "cliente":    empresa["nome"],
        "expires_at": time.time() + 60,
        "usado":      False,
    }
    return {"ok": True, "token": token}


@app.post("/auth/validar-token-launcher")
def auth_validar_token_launcher(req: ValidarTokenLauncherRequest):
    dado = _tokens_launcher.get(req.token)
    if not dado:
        return {"ok": False, "erro": "Token inválido."}
    if dado["usado"]:
        return {"ok": False, "erro": "Token já utilizado."}
    if time.time() > dado["expires_at"]:
        del _tokens_launcher[req.token]
        return {"ok": False, "erro": "Token expirado."}
    empresa = database.buscar_empresa(dado["email"])
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}
    maquinas = database.listar_maquinas(empresa["id"])
    fps = [m["fingerprint"] for m in maquinas]
    if req.fingerprint not in fps:
        return {"ok": False, "erro": "Máquina não autorizada."}
    dado["usado"] = True
    return {"ok": True, "email": dado["email"], "cliente": dado["cliente"]}


# ── Conferência ─────────────────────────────────────────────────────

class ConferenciaRequest(BaseModel):
    email:         str
    modulo:        str
    competencia:   str
    resultados:    list
    sem_movimento: list


@app.post("/agente/conferencia")
def agente_conferencia(req: ConferenciaRequest):
    empresa = database.buscar_empresa(req.email)
    if not empresa:
        return {"ok": False, "erro": "Empresa não encontrada."}

    database.salvar_conferencia(
        empresa["id"], req.modulo, req.competencia,
        req.resultados, req.sem_movimento
    )
    return {"ok": True}


@app.get("/portal/conferencia")
def portal_conferencia(email: str, modulo: str, competencia: str = None):
    try:
        empresa = database.buscar_empresa(email)
        if not empresa:
            return {"ok": False, "erro": "Empresa não encontrada."}
        dados = database.buscar_conferencia(empresa["id"], modulo, competencia)
        competencias = database.listar_competencias_conferencia(empresa["id"], modulo)
        return {"ok": True, "dados": dados, "competencias": competencias}
    except Exception as e:
        return {"ok": False, "erro": str(e)}


# ── SPA Fallback (deve ser a última rota) ──────────────────────────

@app.get("/{full_path:path}")
def spa_fallback(full_path: str):
    return FileResponse("static/index.html")
