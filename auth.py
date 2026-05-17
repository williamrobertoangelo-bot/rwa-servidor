# =====================================================================
# RWA TECNOLOGIA OPERACIONAL — AUTENTICAÇÃO
# =====================================================================

import hashlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# ── Configure aqui o email de alerta ──────────────────────────────
_EMAIL_CONTA   = "rwaautomacoes@gmail.com"
_EMAIL_SENHA   = "SUA_SENHA_APP_AQUI"        # senha de app Gmail
_EMAIL_DESTINO = "rwaautomacoes@gmail.com"


def hash_senha(senha: str) -> str:
    return hashlib.sha256(senha.encode()).hexdigest()


def verificar_senha(senha_digitada: str, senha_hash: str) -> bool:
    return hash_senha(senha_digitada) == senha_hash


def notificar_maquina_nao_autorizada(nome_empresa: str, email: str, fingerprint: str):
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = f"[RWA ALERTA] Máquina não autorizada — {nome_empresa}"
        msg["From"]    = _EMAIL_CONTA
        msg["To"]      = _EMAIL_DESTINO

        corpo = (
            f"ALERTA DE SEGURANÇA — RWA Tecnologia Operacional\n"
            f"{'─' * 45}\n\n"
            f"Empresa    : {nome_empresa}\n"
            f"Email      : {email}\n"
            f"Fingerprint: {fingerprint}\n"
            f"Data/Hora  : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
            f"Máquina não está na lista de autorizadas.\n"
            f"Se for legítima, adicione via painel admin.\n\n"
            f"RWA Tecnologia Operacional"
        )

        msg.attach(MIMEText(corpo, "plain"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(_EMAIL_CONTA, _EMAIL_SENHA)
            smtp.sendmail(_EMAIL_CONTA, _EMAIL_DESTINO, msg.as_string())

    except Exception as e:
        print(f"[EMAIL] Falha ao notificar: {e}")
