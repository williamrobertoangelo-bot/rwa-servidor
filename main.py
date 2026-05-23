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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


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

    html = f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html><head><meta charset="utf-8"/></head>
<body style="margin:0;padding:0;background-color:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f4;padding:20px 0;">
<tr><td align="center">
<table width="480" cellpadding="0" cellspacing="0" border="0" style="max-width:480px;background-color:#ffffff;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
<tr><td style="background-color:#0F1B2D;padding:20px 18px;">
  <div style="margin-bottom:14px;">
    <table cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background:#1e1e2e;border:1px solid rgba(99,102,241,0.4);border-radius:50%;width:38px;height:38px;text-align:center;vertical-align:middle;">
        <span style="font-size:18px;font-weight:900;color:#4f46e5;letter-spacing:2px;">&#8801;</span>
      </td>
    </tr></table>
  </div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#7FB3E0;letter-spacing:1.5px;font-weight:bold;margin-bottom:8px;">RWA SOLUÇÕES</div>
  <div style="display:inline-block;background:#1a3a6b;color:#7FB3E0;font-family:Arial,Helvetica,sans-serif;font-size:10px;padding:3px 8px;border-radius:3px;font-weight:bold;letter-spacing:0.5px;margin-bottom:10px;">CADASTRO DE SENHA</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:17px;color:#FFFFFF;font-weight:bold;line-height:1.3;margin-top:4px;">Cadastro de senha</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#B5D4F4;margin-top:2px;">Ative seu acesso à plataforma RWA</div>
</td></tr>
<tr><td style="padding:18px;">
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.6;margin-bottom:4px;">Olá, <strong>{nome}</strong>. Sua licença foi aprovada. Para ativação, realize o cadastro de sua senha.</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6B6B6B;margin-bottom:16px;">(Senha simples ou com caracteres especiais.)</div>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
    <tr><td style="padding:14px 0 4px 0;font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#6B6B6B;letter-spacing:1px;text-transform:uppercase;font-weight:bold;">Dados do registro</td></tr>
    <tr><td style="padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
        <tr><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;width:55%;">Titular</td><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{nome}</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">CNPJ/CPF</td><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{documento}</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">E-mail</td><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{email}</td></tr>
        <tr><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Telefone</td><td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{telefone}</td></tr>
        <tr><td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Válida até</td><td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{venc_fmt}</td></tr>
      </table>
    </td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:20px;border-collapse:collapse;">
    <tr><td align="center"><a href="{link}" style="display:inline-block;background:#4f46e5;color:#fff;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;padding:13px 32px;border-radius:8px;text-decoration:none;letter-spacing:0.02em;">Cadastrar minha senha &#8594;</a></td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;border-collapse:collapse;">
    <tr><td style="background-color:#E6EEFF;border-left:3px solid #4f46e5;padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#1a1a1a;line-height:1.5;">Este link é pessoal e intransferível. Após cadastrar sua senha, acesse sempre por <strong>rwasolucoes.com.br</strong></td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;border-top:1px solid #EFEFEF;">
    <tr><td style="padding-top:10px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#888;"><span style="color:#0F1B2D;font-weight:bold;">RWA Soluções</span><br/><span style="font-size:11px;">Automação fiscal para escritórios contábeis</span></td></tr>
  </table>
</td></tr>
</table>
</td></tr></table>
</body></html>"""

    try:
        import urllib.request, json as _json
        payload = _json.dumps({
            "from": "RWA Soluções <noreply@rwasolucoes.com.br>",
            "to": [email],
            "subject": "RWA Soluções — Cadastro de senha",
            "html": html,
        }).encode("utf-8")
        req = urllib.request.Request(
            "https://api.resend.com/emails",
            data=payload,
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req) as resp:
            print(f"[EMAIL] Enviado via Resend para {email} — status {resp.status}")
    except urllib.error.HTTPError as e:
        corpo = e.read().decode("utf-8")
        print(f"[EMAIL] Erro Resend {e.code}: {corpo}")
    except Exception as e:
        print(f"[EMAIL] Erro Resend: {e}")

    html = f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html><head><meta charset="utf-8"/></head>
<body style="margin:0;padding:0;background-color:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f4;padding:20px 0;">
<tr><td align="center">
<table width="480" cellpadding="0" cellspacing="0" border="0" style="max-width:480px;background-color:#ffffff;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
<tr><td style="background-color:#0F1B2D;padding:20px 18px;">
  <div style="margin-bottom:14px;">
    <table cellpadding="0" cellspacing="0" border="0"><tr>
      <td style="background:#1e1e2e;border:1px solid rgba(99,102,241,0.4);border-radius:50%;width:38px;height:38px;text-align:center;vertical-align:middle;">
        <span style="font-size:18px;font-weight:900;color:#4f46e5;letter-spacing:2px;">≡</span>
      </td>
    </tr></table>
  </div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#7FB3E0;letter-spacing:1.5px;font-weight:bold;margin-bottom:8px;">RWA SOLUÇÕES</div>
  <div style="display:inline-block;background:#1a3a6b;color:#7FB3E0;font-family:Arial,Helvetica,sans-serif;font-size:10px;padding:3px 8px;border-radius:3px;font-weight:bold;letter-spacing:0.5px;margin-bottom:10px;">CADASTRO DE SENHA</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:17px;color:#FFFFFF;font-weight:bold;line-height:1.3;margin-top:4px;">Cadastro de senha</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#B5D4F4;margin-top:2px;">Ative seu acesso à plataforma RWA</div>
</td></tr>
<tr><td style="padding:18px;">
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.6;margin-bottom:4px;">Olá, <strong>{nome}</strong>. Sua licença foi aprovada. Para ativação, realize o cadastro de sua senha.</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#6B6B6B;margin-bottom:16px;">(Senha simples ou com caracteres especiais.)</div>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
    <tr><td style="padding:14px 0 4px 0;font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#6B6B6B;letter-spacing:1px;text-transform:uppercase;font-weight:bold;">Dados do registro</td></tr>
    <tr><td style="padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;width:55%;">Titular</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{nome}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">CNPJ/CPF</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{documento}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">E-mail</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{email}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Telefone</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{telefone}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Válida até</td>
          <td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{venc_fmt}</td>
        </tr>
      </table>
    </td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:20px;border-collapse:collapse;">
    <tr><td align="center">
      <a href="{link}" style="display:inline-block;background:#4f46e5;color:#fff;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;padding:13px 32px;border-radius:8px;text-decoration:none;letter-spacing:0.02em;">Cadastrar minha senha &#8594;</a>
    </td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;border-collapse:collapse;">
    <tr><td style="background-color:#E6EEFF;border-left:3px solid #4f46e5;padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#1a1a1a;line-height:1.5;">
      Este link é pessoal e intransferível. Após cadastrar sua senha, acesse sempre por <strong>rwasolucoes.com.br</strong>
    </td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;border-top:1px solid #EFEFEF;">
    <tr><td style="padding-top:10px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#888;">
      <span style="color:#0F1B2D;font-weight:bold;">RWA Soluções</span><br/>
      <span style="font-size:11px;">Automação fiscal para escritórios contábeis</span>
    </td></tr>
  </table>
</td></tr>
</table>
</td></tr></table>
</body></html>"""

    plain = f"Olá, {nome}!\n\nSua licença foi aprovada. Acesse o link abaixo para cadastrar sua senha:\n{link}\n\nDados do registro:\nCNPJ/CPF: {documento}\nE-mail: {email}\nVálida até: {venc_fmt}\n\nRWA Soluções — rwasolucoes.com.br"

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "RWA Soluções — Cadastro de senha"
        msg["From"]    = conta
        msg["To"]      = email
        msg.attach(MIMEText(plain, "plain"))
        msg.attach(MIMEText(html, "html"))
        srv = smtplib.SMTP("smtp.gmail.com", 587)
        srv.starttls()
        srv.login(conta, senha)
        srv.sendmail(conta, email, msg.as_string())
        srv.quit()
        print(f"[EMAIL] Email cadastro senha enviado para {email}")
    except Exception as e:
        print(f"[EMAIL] Erro ao enviar cadastro senha: {e}")
    conta  = os.environ.get("RWA_EMAIL_CONTA", "").strip()
    senha  = os.environ.get("RWA_EMAIL_SENHA_APP", "").strip()
    if not conta or not senha:
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

    html = f"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html><head><meta charset="utf-8"/></head>
<body style="margin:0;padding:0;background-color:#f4f4f4;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f4;padding:20px 0;">
<tr><td align="center">
<table width="480" cellpadding="0" cellspacing="0" border="0" style="max-width:480px;background-color:#ffffff;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">
<tr><td style="background-color:#0F1B2D;padding:20px 18px;">
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#7FB3E0;letter-spacing:1.5px;font-weight:bold;margin-bottom:8px;">RWA SOLUÇÕES</div>
  <div style="display:inline-block;background-color:#1F6B43;color:#fff;font-family:Arial,Helvetica,sans-serif;font-size:10px;padding:3px 8px;border-radius:3px;font-weight:bold;letter-spacing:0.5px;margin-bottom:10px;">NOVO ACESSO</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:17px;color:#FFFFFF;font-weight:bold;line-height:1.3;margin-top:4px;">Bem-vindo à plataforma RWA</div>
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#B5D4F4;margin-top:2px;">Sua conta foi ativada com sucesso</div>
</td></tr>
<tr><td style="padding:18px;">
  <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#1F6B43;line-height:1.5;margin-bottom:12px;font-weight:bold;">✓ Acesso liberado. Sua licença está ativa e pronta para uso.</div>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
    <tr><td style="padding:14px 0 4px 0;font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#6B6B6B;letter-spacing:1px;text-transform:uppercase;font-weight:bold;">Dados do registro</td></tr>
    <tr><td style="padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;width:55%;">Titular</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{nome}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">CNPJ/CPF</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{documento}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">E-mail</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{email}</td>
        </tr>
        <tr>
          <td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Telefone</td>
          <td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{telefone}</td>
        </tr>
      </table>
    </td></tr>
    <tr><td style="padding:14px 0 4px 0;font-family:Arial,Helvetica,sans-serif;font-size:10px;color:#6B6B6B;letter-spacing:1px;text-transform:uppercase;font-weight:bold;">Licença</td></tr>
    <tr><td style="padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
        <tr>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Status</td>
          <td style="padding:8px 0;border-bottom:1px solid #EFEFEF;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1F6B43;text-align:right;font-weight:bold;">ATIVA</td>
        </tr>
        <tr>
          <td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#6B6B6B;">Válida até</td>
          <td style="padding:8px 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1a1a1a;text-align:right;font-weight:bold;">{venc_fmt}</td>
        </tr>
      </table>
    </td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;border-collapse:collapse;">
    <tr><td style="background-color:#EAF3DE;border-left:3px solid #3B6D11;padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#27500A;line-height:1.5;">
      Acesse o portal em <strong>rwasolucoes.com.br</strong> e baixe o instalador para começar a usar as automações.
    </td></tr>
  </table>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;border-top:1px solid #EFEFEF;">
    <tr><td style="padding-top:10px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#888;">
      <span style="color:#0F1B2D;font-weight:bold;">RWA Soluções</span><br/>
      <span style="font-size:11px;">Automação fiscal para escritórios contábeis</span>
    </td></tr>
  </table>
</td></tr>
</table>
</td></tr></table>
</body></html>"""

    plain = f"Bem-vindo, {nome}!\n\nSua licença está ativa.\nCNPJ/CPF: {documento}\nE-mail: {email}\nVálida até: {venc_fmt}\n\nAcesse: rwasolucoes.com.br\n\nRWA Soluções"

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = "RWA Soluções — Acesso liberado"
        msg["From"]    = conta
        msg["To"]      = email
        msg.attach(MIMEText(plain, "plain"))
        msg.attach(MIMEText(html, "html"))
        srv = smtplib.SMTP("smtp.gmail.com", 587)
        srv.starttls()
        srv.login(conta, senha)
        srv.sendmail(conta, email, msg.as_string())
        srv.quit()
    except Exception as e:
        print(f"[EMAIL] Erro ao enviar boas-vindas: {e}")

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
            print(f"[EMAIL] Erro ao enviar cadastro senha: {e}")
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

    # Valida/registra fingerprint automaticamente
    if req.fingerprint:
        maquinas = database.listar_maquinas(empresa["id"])
        fps = [m["fingerprint"] for m in maquinas]
        if req.fingerprint not in fps:
            if len(fps) == 0:
                # Primeiro acesso — registra automaticamente
                database.registrar_maquina(empresa["id"], req.fingerprint)
            else:
                # Já existe fingerprint de outra máquina — bloqueia
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
