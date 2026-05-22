# -*- coding: utf-8 -*-
"""
RWA Tecnologia Operacional — Launcher v2.0
Interface visual. Envia comandos ao servidor. Agente executa.
"""

import os
import sys
import json
import time
import socket
import uuid
import hashlib
import getpass
import platform
import subprocess
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
import pystray
from PIL import Image, ImageDraw
import urllib.request
import urllib.error

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Versão e servidor ───────────────────────────────────────────────
_VERSAO       = "2.0"
_SERVIDOR_URL = "https://web-production-31152.up.railway.app"
_INTERVALO_S  = 3

_SCRIPT_DIR = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, "frozen", False) else __file__
))
if not os.path.isdir(_SCRIPT_DIR):
    _SCRIPT_DIR = os.getcwd()

_PASTA_CONFIG  = os.path.join(os.environ.get("LOCALAPPDATA", _SCRIPT_DIR), "RWA_AUTOMACOES", "config")
_PASTA_LOGS    = os.path.join(os.environ.get("LOCALAPPDATA", _SCRIPT_DIR), "RWA_AUTOMACOES", "logs")
_ARQUIVO_CRED  = os.path.join(_PASTA_CONFIG, "credenciais.json")
_ARQUIVO_PATH  = os.path.join(_PASTA_CONFIG, "paths.json")
_ARQUIVO_SINAL = os.path.join(_PASTA_CONFIG, "parar.signal")
_ARQUIVO_LOCK  = os.path.join(_PASTA_CONFIG, "launcher.lock")
_ARQUIVO_LOG_LAUNCHER = os.path.join(_PASTA_LOGS, "launcher_debug.log")
os.makedirs(_PASTA_CONFIG, exist_ok=True)
os.makedirs(_PASTA_LOGS, exist_ok=True)


def _log_launcher(msg: str):
    try:
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f")[:-3]
        linha = f"{agora} [PID {os.getpid()}] {msg}\n"
        with open(_ARQUIVO_LOG_LAUNCHER, "a", encoding="utf-8") as f:
            f.write(linha)
    except Exception:
        pass

# ── Cores ────────────────────────────────────────────────────────────
C_BG      = "#1e1e2e"
C_BG2     = "#13131f"
C_CARD    = "#252538"
C_HOV     = "#2e2e4a"
C_ACCENT  = "#4f46e5"
C_ACCENT2 = "#6366f1"
C_TEXT    = "#ffffff"
C_TEXT2   = "#aaaacc"
C_TEXT3   = "#6666aa"
C_DIM     = "#333355"
C_GREEN   = "#22c55e"
C_RED     = "#ef4444"
C_YELLOW  = "#f59e0b"
C_SEP     = "#2a2a3e"

# ── Módulos ──────────────────────────────────────────────────────────
_MODULOS = [
    ("sao_luis",             "STM — São Luís",              "XML, Relatório, Talão, Guia e consultas automáticas."),
    ("padrao_nacional",      "NFS-e — Padrão Nacional",     "XML, Relatório Gerencial e detalhamento."),
    ("conferencia_sao_luis", "Conferência STM São Luís",    "Conferência inteligente e validação de relatórios."),
    ("conferencia_pn",       "Conferência Padrão Nacional", "Conferência inteligente e validação de relatórios."),
]

# ── Estado global ────────────────────────────────────────────────────
_root          = None
_tray_icon     = None
_EMAIL_SESSAO  = ""
_NOME_CLIENTE  = ""
_FINGERPRINT   = ""
_loop_ativo    = False


# ─────────────────────────────────────────────────────────────────────
# FINGERPRINT
# ─────────────────────────────────────────────────────────────────────

def _gerar_fingerprint() -> str:
    def _disco():
        try:
            r = subprocess.check_output(
                "wmic diskdrive get serialnumber",
                shell=True, stderr=subprocess.DEVNULL)
            ls = [l.strip() for l in r.decode(errors="ignore").splitlines()
                  if l.strip() and "SerialNumber" not in l]
            return ls[0] if ls else "N/A"
        except Exception:
            return "N/A"

    def _bios():
        try:
            r = subprocess.check_output(
                "wmic csproduct get uuid",
                shell=True, stderr=subprocess.DEVNULL)
            ls = [l.strip() for l in r.decode(errors="ignore").splitlines()
                  if l.strip() and "UUID" not in l]
            return ls[0] if ls else "N/A"
        except Exception:
            return "N/A"

    base = "|".join([
        socket.gethostname(),
        getpass.getuser(),
        platform.platform(),
        str(uuid.getnode()),
        _disco(),
        _bios(),
    ])
    return hashlib.sha256(base.encode()).hexdigest().upper()[:32]


# ─────────────────────────────────────────────────────────────────────
# CREDENCIAIS E PATHS
# ─────────────────────────────────────────────────────────────────────

def _salvar_credenciais(email: str, senha: str):
    with open(_ARQUIVO_CRED, "w", encoding="utf-8") as f:
        json.dump({"email": email, "senha": senha}, f)


def _carregar_credenciais() -> dict:
    try:
        if os.path.exists(_ARQUIVO_CRED):
            with open(_ARQUIVO_CRED, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _salvar_paths(paths: dict):
    with open(_ARQUIVO_PATH, "w", encoding="utf-8") as f:
        json.dump(paths, f, ensure_ascii=False, indent=2)


def _carregar_paths() -> dict:
    try:
        if os.path.exists(_ARQUIVO_PATH):
            with open(_ARQUIVO_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {
        "senhas_sl":     "",
        "senhas_pn":     "",
        "pasta_conf_sl": "",
        "pasta_emp_sl":  "",
        "pasta_conf_pn": "",
        "pasta_emp_pn":  "",
    }


# ─────────────────────────────────────────────────────────────────────
# HTTP
# ─────────────────────────────────────────────────────────────────────

def _post(endpoint: str, payload: dict) -> dict:
    url  = _SERVIDOR_URL + endpoint
    data = json.dumps(payload).encode("utf-8")
    req  = urllib.request.Request(
        url, data=data, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=8) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return {"status": "erro", "mensagem": f"HTTP {e.code}"}
    except Exception as e:
        return {"status": "erro", "mensagem": str(e)}


def _get(endpoint: str) -> dict:
    try:
        with urllib.request.urlopen(_SERVIDOR_URL + endpoint, timeout=8) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:
        return {"erro": str(e)}


def _chamar_login(email: str, senha: str) -> dict:
    return _post("/agente/login", {
        "email":       email.strip().lower(),
        "senha":       senha,
        "fingerprint": _FINGERPRINT,
        "versao":      _VERSAO,
    })


def _chamar_executar(modulo: str) -> dict:
    return _post("/portal/executar", {"email": _EMAIL_SESSAO, "modulo": modulo})


def _chamar_agendar(modulo: str, agendado_para: str) -> dict:
    return _post("/portal/agendar", {
        "email":         _EMAIL_SESSAO,
        "modulo":        modulo,
        "agendado_para": agendado_para,
    })


def _chamar_historico() -> list:
    r = _get(f"/portal/historico?email={_EMAIL_SESSAO}")
    return r.get("historico", [])


def _chamar_cancelar(tarefa_id: int) -> dict:
    return _post("/portal/cancelar", {"email": _EMAIL_SESSAO, "tarefa_id": tarefa_id})


def _sinalizar_parar():
    """Grava arquivo de sinal — agente detecta e mata o Chrome."""
    try:
        with open(_ARQUIVO_SINAL, "w") as f:
            f.write("parar")
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────
# HELPERS VISUAIS
# ─────────────────────────────────────────────────────────────────────

def _centralizar(win, w: int, h: int):
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")


def _barra_topo(win):
    tk.Frame(win, bg=C_ACCENT, height=4).pack(fill="x")


def _rodape(win, texto="RWA Tecnologia Operacional"):
    tk.Label(win, text=texto, font=("Arial", 8),
             bg=C_BG, fg=C_TEXT3).pack(side="bottom", pady=8)


def _logo_bloco(parent, bg=C_BG, size_rwa=30, size_sub=10, size_tag=9):
    frame = tk.Frame(parent, bg=bg)
    frame.pack(pady=(20, 4))
    # RWA com A em índigo
    linha = tk.Frame(frame, bg=bg)
    linha.pack()
    tk.Label(linha, text="RW", font=("Arial", size_rwa, "bold"),
             bg=bg, fg=C_TEXT).pack(side="left")
    tk.Label(linha, text="A",  font=("Arial", size_rwa, "bold"),
             bg=bg, fg=C_ACCENT2).pack(side="left")
    # TECNOLOGIA OPERACIONAL
    tk.Label(frame, text="TECNOLOGIA OPERACIONAL",
             font=("Arial", size_sub, "bold"),
             bg=bg, fg=C_ACCENT2).pack(pady=(3, 0))
    # Tagline
    tk.Label(frame, text="Estabilidade. Resultado. Conformidade.",
             font=("Arial", size_tag),
             bg=bg, fg=C_TEXT2).pack(pady=(2, 0))
    return frame


# ─────────────────────────────────────────────────────────────────────
# SYSTEM TRAY
# ─────────────────────────────────────────────────────────────────────

def _criar_imagem_tray():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d   = ImageDraw.Draw(img)
    d.ellipse([4, 4, 60, 60], fill="#4f46e5")
    d.rectangle([18, 20, 24, 44], fill="white")
    d.rectangle([29, 20, 35, 44], fill="white")
    d.rectangle([40, 20, 46, 44], fill="white")
    return img


def _on_tray_abrir(icon, item):
    if _root:
        _root.after(0, _root.deiconify)


def _on_tray_sair(icon, item):
    global _loop_ativo
    _loop_ativo = False
    if _tray_icon:
        _tray_icon.stop()
    if _root:
        _root.after(0, _root.destroy)


def _iniciar_tray():
    global _tray_icon
    menu = pystray.Menu(
        pystray.MenuItem("Abrir RWA", _on_tray_abrir),
        pystray.MenuItem("Sair",      _on_tray_sair),
    )
    _tray_icon = pystray.Icon(
        "RWA", _criar_imagem_tray(), "RWA Tecnologia Operacional", menu)
    threading.Thread(target=_tray_icon.run, daemon=True).start()


# ─────────────────────────────────────────────────────────────────────
# SPLASH
# ─────────────────────────────────────────────────────────────────────




# ─────────────────────────────────────────────────────────────────────
# TELA LOGIN
# ─────────────────────────────────────────────────────────────────────

def _tela_login(email_salvo: str = ""):
    global _EMAIL_SESSAO, _NOME_CLIENTE

    win = tk.Tk()
    win.title("RWA — Acesso")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.configure(bg=C_BG)
    _centralizar(win, 400, 440)

    _barra_topo(win)
    _logo_bloco(win, size_rwa=30, size_sub=10, size_tag=9)

    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(16, 0))

    form = tk.Frame(win, bg=C_BG)
    form.pack(padx=36, pady=(16, 0), fill="x")

    # E-mail
    tk.Label(form, text="E-MAIL", font=("Arial", 8, "bold"),
             bg=C_BG, fg=C_TEXT3, anchor="w").pack(fill="x")
    entry_email = tk.Entry(form, font=("Arial", 11),
                           bg=C_CARD, fg=C_TEXT, insertbackground=C_TEXT,
                           relief="flat", bd=0)
    entry_email.pack(fill="x", ipady=9, pady=(3, 12))
    if email_salvo:
        entry_email.insert(0, email_salvo)

    # Senha
    tk.Label(form, text="SENHA", font=("Arial", 8, "bold"),
             bg=C_BG, fg=C_TEXT3, anchor="w").pack(fill="x")
    frame_pw = tk.Frame(form, bg=C_CARD)
    frame_pw.pack(fill="x", pady=(3, 0))
    entry_senha = tk.Entry(frame_pw, font=("Arial", 11), show="*",
                           bg=C_CARD, fg=C_TEXT, insertbackground=C_TEXT,
                           relief="flat", bd=0)
    entry_senha.pack(side="left", fill="x", expand=True, ipady=9, padx=(6, 0))

    _pw = {"v": False}
    def _toggle():
        _pw["v"] = not _pw["v"]
        entry_senha.config(show="" if _pw["v"] else "*")
    tk.Button(frame_pw, text="👁", font=("Arial", 10),
              bg=C_CARD, fg=C_TEXT3, relief="flat", bd=0,
              cursor="hand2", command=_toggle).pack(side="right", padx=6)

    lbl_erro = tk.Label(win, text="", font=("Arial", 9),
                        bg=C_BG, fg=C_RED)
    lbl_erro.pack(pady=(12, 0))

    resultado = {"ok": False}

    def _entrar(event=None):
        email = entry_email.get().strip().lower()
        senha = entry_senha.get()
        if not email or not senha:
            lbl_erro.config(text="Preencha e-mail e senha.")
            return
        lbl_erro.config(text="Verificando...", fg=C_TEXT3)
        win.update()

        resp = _chamar_login(email, senha)
        if resp.get("status") == "ok":
            global _EMAIL_SESSAO, _NOME_CLIENTE
            _EMAIL_SESSAO  = email
            _NOME_CLIENTE  = resp.get("cliente", "")
            _salvar_credenciais(email, senha)
            resultado["ok"] = True
            win.destroy()
        else:
            lbl_erro.config(
                text=resp.get("mensagem", "Erro ao conectar."),
                fg=C_RED)
            entry_senha.delete(0, tk.END)
            entry_senha.focus_set()

    tk.Frame(win, bg=C_BG, height=4).pack()
    tk.Button(win, text="Entrar na Plataforma",
              font=("Arial", 11, "bold"),
              bg=C_ACCENT, fg=C_TEXT, relief="flat",
              padx=0, pady=11, cursor="hand2",
              command=_entrar).pack(fill="x", padx=36)

    entry_senha.bind("<Return>", _entrar)
    entry_email.bind("<Return>", lambda e: entry_senha.focus_set())

    (entry_senha if email_salvo else entry_email).focus_set()

    _rodape(win)
    win.mainloop()

    if not resultado["ok"]:
        sys.exit(0)


# ─────────────────────────────────────────────────────────────────────
# TELA CONFIGURAÇÕES
# ─────────────────────────────────────────────────────────────────────

def _tela_config():
    paths = _carregar_paths()

    win = tk.Toplevel(_root)
    win.title("RWA — Configurações")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.configure(bg=C_BG)
    _centralizar(win, 560, 620)
    win.grab_set()

    # ── Topo fixo ────────────────────────────────────────────────────
    _barra_topo(win)

    tk.Label(win, text="Configurações de caminhos",
             font=("Arial", 13, "bold"), bg=C_BG, fg=C_TEXT).pack(pady=(16, 2))
    tk.Label(win, text="Planilhas de senhas e pastas usadas pelas automações.",
             font=("Arial", 9), bg=C_BG, fg=C_TEXT2).pack()

    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=30, pady=(12, 0))

    # ── Área scrollável ───────────────────────────────────────────────
    canvas_frame = tk.Frame(win, bg=C_BG)
    canvas_frame.pack(fill="both", expand=True, padx=0, pady=0)

    canvas = tk.Canvas(canvas_frame, bg=C_BG, highlightthickness=0, bd=0)
    scrollbar = tk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas, bg=C_BG)
    inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_configure(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(inner_id, width=canvas.winfo_width())

    inner.bind("<Configure>", _on_configure)
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(inner_id, width=e.width))
    win.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

    frame = tk.Frame(inner, bg=C_BG)
    frame.pack(fill="x", padx=30, pady=(10, 16))

    vars_path = {
        "senhas_sl":     tk.StringVar(value=paths.get("senhas_sl",     "")),
        "senhas_pn":     tk.StringVar(value=paths.get("senhas_pn",     "")),
        "pasta_conf_sl": tk.StringVar(value=paths.get("pasta_conf_sl", "")),
        "pasta_emp_sl":  tk.StringVar(value=paths.get("pasta_emp_sl",  "")),
        "pasta_conf_pn": tk.StringVar(value=paths.get("pasta_conf_pn", "")),
        "pasta_emp_pn":  tk.StringVar(value=paths.get("pasta_emp_pn",  "")),
    }

    # label, chave, tipo, dica
    campos = [
        ("senhas_sl",     "Planilha SENHAS SAO LUIS (.xlsx)",        "arquivo",
         "Ex: C:/rwa-agente/SENHAS SAO LUIS.xlsx"),
        ("senhas_pn",     "Planilha SENHAS PADRAO NACIONAL (.xlsx)", "arquivo",
         "Ex: C:/rwa-agente/SENHAS PADRAO NACIONAL.xlsx"),
        ("pasta_conf_sl", "Pasta CONFERÊNCIAS — São Luís",           "pasta",
         "Pasta onde ficam os PDFs Domínio (Serviços, Simples) e o Gerencial STM."),
        ("pasta_emp_sl",  "Pasta EMPRESAS — São Luís",               "pasta",
         "Pasta raiz com subpastas de cada empresa (ex: 001 - EMPRESA LTDA)."),
        ("pasta_conf_pn", "Pasta CONFERÊNCIAS — Padrão Nacional",    "pasta",
         "Pasta onde ficam os PDFs Domínio e o Gerencial/CONCLUSAO do PN."),
        ("pasta_emp_pn",  "Pasta EMPRESAS — Padrão Nacional",        "pasta",
         "Pasta raiz com subpastas de cada empresa do Padrão Nacional."),
    ]

    for chave, label, tipo, dica in campos:
        tk.Label(frame, text=label, font=("Arial", 9, "bold"),
                 bg=C_BG, fg=C_TEXT3, anchor="w").pack(fill="x", pady=(14, 1))

        tk.Label(frame, text=dica, font=("Arial", 8),
                 bg=C_BG, fg=C_TEXT3, anchor="w").pack(fill="x", pady=(0, 3))

        row = tk.Frame(frame, bg=C_BG)
        row.pack(fill="x")

        tk.Entry(row, textvariable=vars_path[chave],
                 font=("Arial", 9), bg=C_CARD, fg=C_TEXT,
                 insertbackground=C_TEXT, relief="flat", bd=0
                 ).pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 8))

        def _browse(ch=chave, tp=tipo):
            if tp == "arquivo":
                p = filedialog.askopenfilename(
                    parent=win,
                    title=f"Selecionar {ch}",
                    filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
            else:
                p = filedialog.askdirectory(parent=win, title="Selecionar pasta")
            if p:
                vars_path[ch].set(p)

        tk.Button(row, text="...", font=("Arial", 9, "bold"),
                  bg=C_DIM, fg=C_TEXT2, relief="flat",
                  width=4, pady=5, cursor="hand2",
                  command=_browse).pack(side="right")

    # ── Rodapé fixo ───────────────────────────────────────────────────
    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=30)

    rodape_frame = tk.Frame(win, bg=C_BG)
    rodape_frame.pack(fill="x", padx=30, pady=10)

    lbl_ok = tk.Label(rodape_frame, text="", font=("Arial", 9),
                      bg=C_BG, fg=C_GREEN, anchor="w")
    lbl_ok.pack(side="left")

    def _salvar():
        novo = {ch: vars_path[ch].get().strip() for ch in vars_path}
        _salvar_paths(novo)
        lbl_ok.config(text="✓ Salvo.")
        win.after(1400, win.destroy)

    tk.Button(rodape_frame, text="Salvar configurações",
              font=("Arial", 10, "bold"),
              bg=C_ACCENT, fg=C_TEXT, relief="flat",
              padx=20, pady=6, cursor="hand2",
              command=_salvar).pack(side="right")

    win.wait_window()


# ─────────────────────────────────────────────────────────────────────
# TELA AGENDAMENTO
# ─────────────────────────────────────────────────────────────────────

def _tela_agendamento():
    resultado = {"modulo": None, "dt": None}

    win = tk.Toplevel(_root)
    win.title("RWA — Agendar")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.configure(bg=C_BG)
    _centralizar(win, 440, 420)
    win.grab_set()

    # Barra topo
    tk.Frame(win, bg=C_ACCENT, height=4).pack(fill="x")

    # Título
    tk.Label(win, text="Agendar execução",
             font=("Arial", 13, "bold"), bg=C_BG, fg=C_TEXT).pack(pady=(18, 2))
    tk.Label(win, text="Escolha o módulo, data e hora.",
             font=("Arial", 9), bg=C_BG, fg=C_TEXT2).pack()

    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(14, 0))

    frame = tk.Frame(win, bg=C_BG)
    frame.pack(padx=36, pady=(14, 0), fill="x")

    # Módulo
    tk.Label(frame, text="MÓDULO", font=("Arial", 8, "bold"),
             bg=C_BG, fg=C_ACCENT2, anchor="w").pack(fill="x", pady=(0, 4))

    mod_var = tk.StringVar(value=_MODULOS[0][0])
    for mod_id, mod_nome, _ in _MODULOS:
        tk.Radiobutton(frame, text=mod_nome,
                       variable=mod_var, value=mod_id,
                       font=("Arial", 10), bg=C_BG, fg=C_TEXT2,
                       selectcolor=C_CARD, activebackground=C_BG,
                       activeforeground=C_TEXT,
                       highlightthickness=0).pack(anchor="w", pady=1)

    tk.Frame(frame, bg=C_SEP, height=1).pack(fill="x", pady=(12, 12))

    # Data e hora — dois campos lado a lado, alinhados
    dt_row = tk.Frame(frame, bg=C_BG)
    dt_row.pack(fill="x")

    # Coluna Data
    col_d = tk.Frame(dt_row, bg=C_BG)
    col_d.pack(side="left", padx=(0, 16))
    tk.Label(col_d, text="DATA", font=("Arial", 8, "bold"),
             bg=C_BG, fg=C_ACCENT2).pack(anchor="w")
    f_data = tk.Frame(col_d, bg=C_CARD)
    f_data.pack(anchor="w", pady=(3, 0))
    entry_data = tk.Entry(f_data, font=("Courier", 11), width=12,
                          bg=C_CARD, fg=C_TEXT, insertbackground=C_TEXT,
                          relief="flat", bd=0)
    entry_data.pack(ipady=8, padx=10)
    entry_data.insert(0, "dd/mm/aaaa")
    entry_data.config(fg=C_TEXT3)

    def _data_focus_in(e):
        if entry_data.get() == "dd/mm/aaaa":
            entry_data.delete(0, "end")
            entry_data.config(fg=C_TEXT)
    def _data_focus_out(e):
        if not entry_data.get().strip():
            entry_data.insert(0, "dd/mm/aaaa")
            entry_data.config(fg=C_TEXT3)
    entry_data.bind("<FocusIn>",  _data_focus_in)
    entry_data.bind("<FocusOut>", _data_focus_out)

    def _mascara_data(e):
        v = ''.join(c for c in entry_data.get() if c.isdigit())
        if len(v) > 2: v = v[:2] + '/' + v[2:]
        if len(v) > 5: v = v[:5] + '/' + v[5:9]
        entry_data.delete(0, "end")
        entry_data.insert(0, v)
        entry_data.config(fg=C_TEXT)
    entry_data.bind("<KeyRelease>", _mascara_data)

    # Coluna Hora
    col_h = tk.Frame(dt_row, bg=C_BG)
    col_h.pack(side="left")
    tk.Label(col_h, text="HORA", font=("Arial", 8, "bold"),
             bg=C_BG, fg=C_ACCENT2).pack(anchor="w")
    f_hora = tk.Frame(col_h, bg=C_CARD)
    f_hora.pack(anchor="w", pady=(3, 0))
    entry_hora = tk.Entry(f_hora, font=("Courier", 11), width=7,
                          bg=C_CARD, fg=C_TEXT, insertbackground=C_TEXT,
                          relief="flat", bd=0)
    entry_hora.pack(ipady=8, padx=10)
    entry_hora.insert(0, "HH:MM")
    entry_hora.config(fg=C_TEXT3)

    def _hora_focus_in(e):
        if entry_hora.get() == "HH:MM":
            entry_hora.delete(0, "end")
            entry_hora.config(fg=C_TEXT)
    def _hora_focus_out(e):
        if not entry_hora.get().strip():
            entry_hora.insert(0, "HH:MM")
            entry_hora.config(fg=C_TEXT3)
    entry_hora.bind("<FocusIn>",  _hora_focus_in)
    entry_hora.bind("<FocusOut>", _hora_focus_out)

    def _mascara_hora(e):
        v = ''.join(c for c in entry_hora.get() if c.isdigit())
        if len(v) > 2: v = v[:2] + ':' + v[2:4]
        entry_hora.delete(0, "end")
        entry_hora.insert(0, v)
        entry_hora.config(fg=C_TEXT)
    entry_hora.bind("<KeyRelease>", _mascara_hora)

    # Erro
    lbl_erro = tk.Label(win, text="", font=("Arial", 9),
                        bg=C_BG, fg=C_RED)
    lbl_erro.pack(pady=(12, 0))

    def _confirmar(event=None):
        data_val = entry_data.get().strip()
        hora_val = entry_hora.get().strip()
        if data_val == "dd/mm/aaaa" or hora_val == "HH:MM":
            lbl_erro.config(text="Preencha data e hora.")
            return
        try:
            dt = datetime.strptime(f"{data_val} {hora_val}", "%d/%m/%Y %H:%M")
            if dt <= datetime.now():
                lbl_erro.config(text="A data/hora deve ser no futuro.")
                return
        except ValueError:
            lbl_erro.config(text="Formato inválido. Use dd/mm/aaaa e HH:MM")
            return
        resultado["modulo"] = mod_var.get()
        resultado["dt"]     = dt
        win.destroy()

    def _cancelar():
        win.destroy()

    # Botões — centralizados, mesma largura
    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(4, 0))

    frame_btn = tk.Frame(win, bg=C_BG)
    frame_btn.pack(pady=16)
    tk.Button(frame_btn, text="Agendar",
              font=("Arial", 10, "bold"),
              bg=C_ACCENT, fg=C_TEXT, relief="flat",
              width=14, pady=7, cursor="hand2",
              command=_confirmar).pack(side="left", padx=(0, 8))
    tk.Button(frame_btn, text="Cancelar",
              font=("Arial", 10),
              bg=C_DIM, fg=C_TEXT2, relief="flat",
              width=10, pady=7, cursor="hand2",
              command=_cancelar).pack(side="left")

    entry_data.bind("<Return>", lambda e: entry_hora.focus_set())
    entry_hora.bind("<Return>", _confirmar)

    _rodape(win)
    win.wait_window()
    return resultado["modulo"], resultado["dt"]

def _janela_contagem_regressiva(modulo: str, dt_agendado: datetime) -> bool:
    nomes    = {m[0]: m[1] for m in _MODULOS}
    resultado = {"executar": False}

    win = tk.Toplevel(_root)
    win.title("RWA — Aguardando agendamento")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.configure(bg=C_BG)
    _centralizar(win, 440, 320)

    _barra_topo(win)

    tk.Label(win, text="✅  Agendamento confirmado",
             font=("Arial", 13, "bold"), bg=C_BG, fg=C_GREEN).pack(pady=(20, 4))
    tk.Label(win, text=nomes.get(modulo, modulo),
             font=("Arial", 11), bg=C_BG, fg=C_TEXT2).pack()
    tk.Label(win, text=dt_agendado.strftime("%d/%m/%Y  às  %H:%M"),
             font=("Arial", 10), bg=C_BG, fg=C_TEXT3).pack(pady=(3, 0))

    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(16, 0))

    tk.Label(win, text="Iniciando em:", font=("Arial", 10),
             bg=C_BG, fg=C_TEXT3).pack(pady=(14, 2))

    lbl_cnt = tk.Label(win, text="--:--:--",
                       font=("Arial", 30, "bold"), bg=C_BG, fg=C_ACCENT2)
    lbl_cnt.pack()

    def _cancelar():
        resultado["executar"] = False
        win.destroy()

    tk.Button(win, text="Cancelar agendamento",
              font=("Arial", 9), bg=C_DIM, fg=C_TEXT2, relief="flat",
              padx=14, pady=5, cursor="hand2",
              command=_cancelar).pack(pady=(16, 0))

    _rodape(win)

    def _tick():
        if not win.winfo_exists():
            return
        restante = (dt_agendado - datetime.now()).total_seconds()
        if restante <= 0:
            resultado["executar"] = True
            win.destroy()
            return
        h = int(restante) // 3600
        m = (int(restante) % 3600) // 60
        s = int(restante) % 60
        lbl_cnt.config(text=f"{h:02d}:{m:02d}:{s:02d}")
        win.after(1000, _tick)

    win.after(400, _tick)
    win.wait_window()
    return resultado["executar"]


# ─────────────────────────────────────────────────────────────────────
# EXECUTAR / AGENDAR
# ─────────────────────────────────────────────────────────────────────

# Módulos que suportam parada forçada
_MODULOS_COM_PARAR = {"sao_luis", "padrao_nacional"}

# Referências globais para botões — permitem troca executar ↔ parar
_btns_executar = {}
_tarefa_ativa  = {"id": None, "modulo": None}


def _executar_modulo(modulo: str, lbl_status: tk.Label):
    lbl_status.config(text="Enviando...", fg=C_YELLOW)
    if _root:
        _root.update()

    def _run():
        resp = _chamar_executar(modulo)
        if resp.get("ok"):
            lbl_status.config(text="✓ Comando enviado.", fg=C_GREEN)
        else:
            lbl_status.config(
                text=f"✗ {resp.get('erro', 'Erro ao enviar.')}", fg=C_RED)
        if _root and _root.winfo_exists():
            _root.after(4000, lambda: lbl_status.config(text=""))

    threading.Thread(target=_run, daemon=True).start()


def _parar_execucao():
    """Sinaliza parada: grava arquivo + cancela tarefa no servidor."""
    _sinalizar_parar()
    tid = _tarefa_ativa.get("id")
    if tid:
        threading.Thread(target=_chamar_cancelar, args=(tid,), daemon=True).start()


def _agendar_modulo():
    modulo, dt = _tela_agendamento()
    if not modulo or not dt:
        return
    resp = _chamar_agendar(modulo, dt.strftime("%Y-%m-%d %H:%M"))
    if resp.get("ok"):
        executar = _janela_contagem_regressiva(modulo, dt)
        if executar:
            _chamar_executar(modulo)
    else:
        messagebox.showerror("RWA", resp.get("erro", "Erro ao agendar."))


# ─────────────────────────────────────────────────────────────────────
# PAINEL PRINCIPAL
# ─────────────────────────────────────────────────────────────────────

def _painel_principal():
    global _root, _loop_ativo

    _root = tk.Tk()
    _root.title("RWA Tecnologia Operacional")
    _root.resizable(False, False)
    _root.configure(bg=C_BG)
    _centralizar(_root, 560, 660)
    _root.protocol("WM_DELETE_WINDOW", lambda: _root.withdraw())

    def _snapshot_janela(origem: str):
        try:
            _log_launcher(
                f"[VIGIA] origem={origem} state={_root.state()!r} "
                f"exists={_root.winfo_exists()} viewable={_root.winfo_viewable()} "
                f"pos=({_root.winfo_x()},{_root.winfo_y()}) "
                f"tamanho=({_root.winfo_width()}x{_root.winfo_height()})"
            )
        except Exception:
            _log_launcher(f"[VIGIA][ERRO] falha ao capturar estado origem={origem}")
            _log_launcher(traceback.format_exc())

    def _tk_exception_auditada(exc, val, tb):
        _log_launcher("[TK_EXCEPTION] exception em callback Tkinter")
        _log_launcher("".join(traceback.format_exception(exc, val, tb)))

    _root.report_callback_exception = _tk_exception_auditada

    _log_launcher("[ROTA] entrou em _painel_principal; root criado/configurado")
    _snapshot_janela("apos_root_configurado")

    _log_launcher("[ROTA] forçando exibicao inicial da janela")
    try:
        _root.deiconify()
        _root.lift()
        _root.attributes("-topmost", True)
        _root.focus_force()
        _root.update_idletasks()
        _root.update()
        _root.attributes("-topmost", False)
        _snapshot_janela("apos_forcar_visual_inicial")
    except Exception:
        _log_launcher("[ROTA][ERRO] falha ao forcar exibicao inicial")
        _log_launcher(traceback.format_exc())

    _log_launcher("[ROTA] iniciando barra topo")

    # Barra topo índigo
    tk.Frame(_root, bg=C_ACCENT, height=4).pack(fill="x")
    _log_launcher("[ROTA] barra topo montada; iniciando header")

    # Header
    header = tk.Frame(_root, bg=C_BG2)
    header.pack(fill="x")

    h_logo = tk.Frame(header, bg=C_BG2)
    h_logo.pack(side="left", padx=20, pady=14)
    linha_rwa = tk.Frame(h_logo, bg=C_BG2)
    linha_rwa.pack(anchor="w")
    tk.Label(linha_rwa, text="RW", font=("Arial", 20, "bold"),
             bg=C_BG2, fg=C_TEXT).pack(side="left")
    tk.Label(linha_rwa, text="A",  font=("Arial", 20, "bold"),
             bg=C_BG2, fg=C_ACCENT2).pack(side="left")
    tk.Label(h_logo, text="TECNOLOGIA OPERACIONAL",
             font=("Arial", 8, "bold"), bg=C_BG2, fg=C_ACCENT2).pack(anchor="w")

    h_right = tk.Frame(header, bg=C_BG2)
    h_right.pack(side="right", padx=16, pady=14)
    tk.Label(h_right, text=_NOME_CLIENTE,
             font=("Arial", 10, "bold"), bg=C_BG2, fg=C_TEXT).pack(anchor="e")
    tk.Label(h_right, text=_EMAIL_SESSAO,
             font=("Arial", 8), bg=C_BG2, fg=C_TEXT3).pack(anchor="e")

    _log_launcher("[ROTA] header montado; iniciando corpo")
    # Corpo
    corpo = tk.Frame(_root, bg=C_BG)
    corpo.pack(fill="both", expand=True, padx=20, pady=(16, 0))

    _log_launcher("[ROTA] corpo montado; iniciando linha topo")
    # Linha topo: título + botões
    topo = tk.Frame(corpo, bg=C_BG)
    topo.pack(fill="x", pady=(0, 12))

    tk.Label(topo, text="AUTOMAÇÕES DISPONÍVEIS",
             font=("Arial", 9, "bold"), bg=C_BG, fg=C_TEXT3).pack(side="left")

    tk.Button(topo, text="⏱  Agendar",
              font=("Arial", 9, "bold"),
              bg=C_DIM, fg=C_ACCENT2, relief="flat",
              padx=12, pady=5, cursor="hand2",
              command=_agendar_modulo).pack(side="right", padx=(6, 0))

    tk.Button(topo, text="⚙  Configurações",
              font=("Arial", 9, "bold"),
              bg=C_DIM, fg=C_TEXT2, relief="flat",
              padx=12, pady=5, cursor="hand2",
              command=_tela_config).pack(side="right")

    _log_launcher("[ROTA] linha topo montada; iniciando cards")
    # Cards dos módulos - altura fixa para uniformidade
    _status_labels = {}
    CARD_H = 92

    for mod_id, mod_nome, mod_desc in _MODULOS:
        _log_launcher(f"[ROTA][CARD] iniciando card mod_id={mod_id}")
        try:
            card = tk.Frame(corpo, bg=C_CARD, bd=0, height=CARD_H)
            card.pack(fill="x", pady=(0, 10))
            card.pack_propagate(False)

            inner = tk.Frame(card, bg=C_CARD)
            inner.place(relx=0, rely=0, relwidth=1, relheight=1)

            tk.Label(inner, text=mod_nome, font=("Arial", 11, "bold"),
                     bg=C_CARD, fg=C_TEXT, anchor="w").place(x=16, y=14)

            tk.Label(inner, text=mod_desc, font=("Arial", 9),
                     bg=C_CARD, fg=C_TEXT2, anchor="w").place(x=16, y=40)

            lbl_st = tk.Label(inner, text="",
                              font=("Arial", 9), bg=C_CARD, fg=C_TEXT3)
            lbl_st.place(x=16, y=64)
            _status_labels[mod_id] = lbl_st

            btn = tk.Button(inner, text="▶  Executar agora",
                            font=("Arial", 9, "bold"),
                            bg=C_ACCENT, fg=C_TEXT, relief="flat",
                            padx=14, pady=5, cursor="hand2",
                            command=lambda m=mod_id, lbl=lbl_st: _executar_modulo(m, lbl))
            btn.place(relx=1.0, rely=1.0, anchor="se", x=-16, y=-10)
            _btns_executar[mod_id] = btn
            _log_launcher(f"[ROTA][CARD] card concluido mod_id={mod_id}")
        except Exception:
            _log_launcher(f"[ROTA][CARD][ERRO] falha ao montar card mod_id={mod_id}")
            _log_launcher(traceback.format_exc())
            raise

    _log_launcher("[ROTA] cards montados; iniciando separador/ultima execucao")
    _snapshot_janela("apos_cards")

    # Separador + última execução
    tk.Frame(corpo, bg=C_SEP, height=1).pack(fill="x", pady=(12, 10))
    _log_launcher("[ROTA] separador montado")

    tk.Label(corpo, text="ÚLTIMA EXECUÇÃO",
             font=("Arial", 9, "bold"), bg=C_BG, fg=C_TEXT3,
             anchor="w").pack(fill="x")

    lbl_ultima = tk.Label(corpo, text="Carregando...",
                          font=("Arial", 10), bg=C_BG, fg=C_TEXT2,
                          anchor="w")
    lbl_ultima.pack(fill="x", pady=(4, 0))
    _log_launcher("[ROTA] ultima execucao montada; iniciando rodape")

    # Rodapé
    tk.Frame(_root, bg=C_SEP, height=1).pack(fill="x", padx=20, pady=(14, 0))
    rod = tk.Frame(_root, bg=C_BG)
    rod.pack(fill="x", padx=20, pady=(6, 12))
    tk.Label(rod, text="Estabilidade. Resultado. Conformidade.",
             font=("Arial", 8), bg=C_BG, fg=C_TEXT3).pack(side="left")
    tk.Label(rod, text=f"v{_VERSAO}",
             font=("Arial", 8), bg=C_BG, fg=C_DIM).pack(side="right")
    _log_launcher("[ROTA] rodape montado; preparando polling")
    _snapshot_janela("apos_rodape")

    # Polling última execução
    nomes_mod = {m[0]: m[1] for m in _MODULOS}

    def _atualizar():
        try:
            hist = _chamar_historico()
            if hist:
                ult  = hist[0]
                nome = nomes_mod.get(ult["modulo"], ult["modulo"])
                when = ult.get("quando", "")[:16].replace("T", " ")
                st   = ult.get("status", "")
                cor  = (C_GREEN  if st in ("concluido", "em_execucao") else
                        C_RED    if st == "erro"      else
                        C_YELLOW if st == "cancelado" else C_TEXT2)
                lbl_ultima.config(
                    text=f"{nome}  —  {when}  —  {st.upper()}",
                    fg=cor)

                # Verifica se há execução em andamento (só SL e PN)
                em_exec_mod = None
                em_exec_id  = None
                for h in hist:
                    if h.get("status") == "em_execucao" and h.get("modulo") in _MODULOS_COM_PARAR:
                        em_exec_mod = h["modulo"]
                        em_exec_id  = h["id"]
                        break

                # Atualiza referência de tarefa ativa
                _tarefa_ativa["id"]    = em_exec_id
                _tarefa_ativa["modulo"] = em_exec_mod

                # Troca botões conforme estado
                for mod_id, btn in _btns_executar.items():
                    if mod_id == em_exec_mod:
                        # Este módulo está rodando — mostra Parar
                        btn.config(
                            text="■  Parar",
                            bg=C_RED,
                            command=_parar_execucao)
                        _status_labels[mod_id].config(
                            text="● Em execução...", fg=C_YELLOW)
                    elif em_exec_mod and mod_id in _MODULOS_COM_PARAR:
                        # Outro módulo com parar — desabilita enquanto tem execução
                        btn.config(
                            text="▶  Executar agora",
                            bg="#333355",
                            command=lambda: None)
                    else:
                        # Nenhum rodando ou conferência — botão normal
                        btn.config(
                            text="▶  Executar agora",
                            bg=C_ACCENT,
                            command=lambda m=mod_id, lbl=_status_labels[mod_id]: _executar_modulo(m, lbl))
                        if not em_exec_mod:
                            _status_labels[mod_id].config(text="")

        except Exception:
            _log_launcher("[ATUALIZAR][ERRO] exception dentro do polling de historico")
            _log_launcher(traceback.format_exc())
        if _loop_ativo and _root.winfo_exists():
            _root.after(_INTERVALO_S * 1000, _atualizar)

    def _vigiar_janela():
        try:
            _snapshot_janela("watchdog")
            if _loop_ativo and _root.winfo_exists():
                _root.after(1000, _vigiar_janela)
        except Exception:
            _log_launcher("[VIGIA][ERRO] exception no watchdog da janela")
            _log_launcher(traceback.format_exc())

    _log_launcher("[ROTA] polling definido; forçando visual final")
    try:
        _root.deiconify()
        _root.lift()
        _root.attributes("-topmost", True)
        _root.focus_force()
        _root.update_idletasks()
        _root.update()
        _root.attributes("-topmost", False)
        _snapshot_janela("apos_forcar_visual_final")
    except Exception:
        _log_launcher("[ROTA][ERRO] falha ao forcar visual final")
        _log_launcher(traceback.format_exc())

    _log_launcher("[ROTA] antes de ativar loop")
    _loop_ativo = True
    _log_launcher("[ROTA] antes de registrar after historico")
    _root.after(2000, _atualizar)
    _log_launcher("[ROTA] after historico registrado")
    _log_launcher("[ROTA] antes de registrar watchdog")
    _root.after(1000, _vigiar_janela)
    _log_launcher("[ROTA] watchdog registrado")
    _snapshot_janela("antes_mainloop")
    _log_launcher("[ROTA] entrando no mainloop")
    _root.mainloop()
    _log_launcher("[ROTA] mainloop saiu")
    _snapshot_janela("depois_mainloop")


# ─────────────────────────────────────────────────────────────────────
# TOKEN PORTAL
# ─────────────────────────────────────────────────────────────────────

def _extrair_token_argv() -> str:
    _log_launcher(f"[TOKEN] argv recebido={sys.argv!r}")
    for arg in sys.argv[1:]:
        if "token=" in arg:
            token = arg.split("token=")[-1].split("&")[0].strip()
            _log_launcher(f"[TOKEN] token extraido tamanho={len(token)}")
            return token
    _log_launcher("[TOKEN] nenhum token encontrado")
    return ""


def _adquirir_lock_launcher() -> bool:
    try:
        _log_launcher(f"[LOCK] tentando adquirir lock={_ARQUIVO_LOCK}")
        if os.path.exists(_ARQUIVO_LOCK):
            idade = time.time() - os.path.getmtime(_ARQUIVO_LOCK)
            try:
                with open(_ARQUIVO_LOCK, "r", encoding="utf-8") as f:
                    pid_lock = f.read().strip()
            except Exception:
                pid_lock = ""
            _log_launcher(f"[LOCK] lock existente idade={idade:.1f}s pid_gravado={pid_lock!r}")
            if idade > 60:
                os.remove(_ARQUIVO_LOCK)
                _log_launcher("[LOCK] lock antigo removido")
        fd = os.open(_ARQUIVO_LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        _log_launcher("[LOCK] adquirido com sucesso")
        return True
    except FileExistsError:
        _log_launcher("[LOCK] bloqueado por outra instancia")
        return False
    except Exception as e:
        _log_launcher(f"[LOCK] erro ao adquirir lock, liberando execucao por seguranca: {e!r}")
        return True


def _liberar_lock_launcher():
    try:
        if os.path.exists(_ARQUIVO_LOCK):
            os.remove(_ARQUIVO_LOCK)
            _log_launcher("[LOCK] liberado")
    except Exception as e:
        _log_launcher(f"[LOCK] erro ao liberar: {e!r}")


def _tela_bloqueio(mensagem: str = ""):
    win = tk.Tk()
    win.title("RWA")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.configure(bg=C_BG)
    _centralizar(win, 400, 260)
    _barra_topo(win)
    _logo_bloco(win, size_rwa=24, size_sub=9, size_tag=8)
    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(12, 0))
    tk.Label(win,
             text=mensagem or "Este programa deve ser aberto pelo portal.",
             font=("Arial", 10), bg=C_BG, fg=C_TEXT2,
             wraplength=320, justify="center").pack(pady=(16, 4))
    tk.Label(win, text="rwasolucoes.com.br",
             font=("Arial", 10, "bold"), bg=C_BG, fg=C_ACCENT2).pack()
    _rodape(win)
    win.after(4000, win.destroy)
    win.mainloop()


def _validar_token_portal(token: str) -> bool:
    global _EMAIL_SESSAO, _NOME_CLIENTE
    resp = _post("/auth/validar-token-launcher", {
        "token":       token,
        "fingerprint": _FINGERPRINT,
    })
    if resp.get("ok"):
        _EMAIL_SESSAO = resp.get("email", "")
        _NOME_CLIENTE = resp.get("cliente", "")
        return True
    _tela_bloqueio(mensagem=resp.get("erro", "Token inválido ou expirado."))
    return False


# ─────────────────────────────────────────────────────────────────────
# PRIMEIRO ACESSO (chamado pelo instalador com --primeiro-acesso)
# ─────────────────────────────────────────────────────────────────────

def _tela_primeiro_acesso():
    """Tela de login para primeiro acesso via instalador."""
    win = tk.Tk()
    win.title("RWA")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    win.configure(bg=C_BG)
    _centralizar(win, 400, 380)
    _barra_topo(win)
    _logo_bloco(win, size_rwa=24, size_sub=9, size_tag=8)
    tk.Frame(win, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(12, 0))

    tk.Label(win, text="Primeiro acesso — informe suas credenciais",
             font=("Arial", 9), bg=C_BG, fg=C_TEXT2,
             wraplength=320, justify="center").pack(pady=(14, 10))

    frm = tk.Frame(win, bg=C_BG)
    frm.pack(padx=36, fill="x")

    tk.Label(frm, text="E-mail", font=("Arial", 9), bg=C_BG, fg=C_TEXT2, anchor="w").pack(fill="x")
    ent_email = tk.Entry(frm, font=("Arial", 11), bg=C_CARD, fg=C_TEXT, insertbackground=C_TEXT,
                         relief="flat", bd=6)
    ent_email.pack(fill="x", pady=(2, 10))

    tk.Label(frm, text="Senha", font=("Arial", 9), bg=C_BG, fg=C_TEXT2, anchor="w").pack(fill="x")
    ent_senha = tk.Entry(frm, font=("Arial", 11), bg=C_CARD, fg=C_TEXT, insertbackground=C_TEXT,
                         relief="flat", bd=6, show="●")
    ent_senha.pack(fill="x", pady=(2, 0))

    lbl_status = tk.Label(win, text="", font=("Arial", 9), bg=C_BG, fg=C_YELLOW,
                          wraplength=320, justify="center")
    lbl_status.pack(pady=(10, 0))

    def _confirmar(event=None):
        email = ent_email.get().strip()
        senha = ent_senha.get().strip()
        if not email or not senha:
            lbl_status.config(text="Preencha e-mail e senha.", fg=C_RED)
            return

        lbl_status.config(text="Validando credenciais...", fg=C_YELLOW)
        win.update()

        resp = _post("/agente/login", {
            "email":       email.lower(),
            "senha":       senha,
            "fingerprint": _FINGERPRINT,
            "versao":      _VERSAO,
        })

        if resp.get("status") != "ok":
            lbl_status.config(text=resp.get("mensagem", "Credenciais inválidas."), fg=C_RED)
            return

        # Salva credenciais
        os.makedirs(_PASTA_CONFIG, exist_ok=True)
        with open(_ARQUIVO_CRED, "w", encoding="utf-8") as f:
            json.dump({"email": email.lower(), "senha": senha}, f)

        # Inicia agente em segundo plano
        if getattr(sys, "frozen", False):
            agente_path = os.path.join(_SCRIPT_DIR, "RWA_AGENTE.exe")
        else:
            agente_path = os.path.join(_SCRIPT_DIR, "RWA_AGENTE.py")

        if os.path.exists(agente_path):
            if getattr(sys, "frozen", False):
                subprocess.Popen(
                    [agente_path],
                    cwd=_SCRIPT_DIR,
                    creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
                )
            else:
                subprocess.Popen(
                    [sys.executable, agente_path],
                    cwd=_SCRIPT_DIR,
                    creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
                )

        # Mensagem final customizada RWA
        win.destroy()
        fim = tk.Tk()
        fim.title("RWA")
        fim.resizable(False, False)
        fim.attributes("-topmost", True)
        fim.configure(bg=C_BG)
        _centralizar(fim, 400, 260)
        _barra_topo(fim)
        _logo_bloco(fim, size_rwa=24, size_sub=9, size_tag=8)
        tk.Frame(fim, bg=C_SEP, height=1).pack(fill="x", padx=36, pady=(12, 0))
        tk.Label(fim, text="Instalação concluída com sucesso.",
                 font=("Arial", 11, "bold"), bg=C_BG, fg=C_GREEN).pack(pady=(16, 4))
        tk.Label(fim, text="O agente RWA está rodando em segundo plano.\nAcesse o portal para usar o sistema.",
                 font=("Arial", 9), bg=C_BG, fg=C_TEXT2,
                 wraplength=320, justify="center").pack(pady=(4, 4))
        tk.Label(fim, text="rwasolucoes.com.br",
                 font=("Arial", 10, "bold"), bg=C_BG, fg=C_ACCENT2).pack()
        _rodape(fim)
        fim.after(5000, fim.destroy)
        fim.mainloop()

    btn = tk.Button(frm, text="Confirmar", font=("Arial", 10, "bold"),
                    bg=C_ACCENT, fg=C_TEXT, relief="flat", bd=0,
                    activebackground=C_ACCENT2, activeforeground=C_TEXT,
                    cursor="hand2", command=_confirmar)
    btn.pack(fill="x", pady=(14, 0), ipady=8)
    ent_senha.bind("<Return>", _confirmar)

    _rodape(win)
    win.mainloop()


# ─────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────

def main():
    global _FINGERPRINT, _EMAIL_SESSAO, _NOME_CLIENTE

    _log_launcher("=" * 70)
    _log_launcher("[MAIN] inicio do launcher")
    _log_launcher(f"[MAIN] argv={sys.argv!r}")
    _log_launcher(f"[MAIN] frozen={getattr(sys, 'frozen', False)} executable={sys.executable!r}")
    _log_launcher(f"[MAIN] script_dir={_SCRIPT_DIR!r} cwd={os.getcwd()!r}")
    _log_launcher(f"[MAIN] arquivo_log={_ARQUIVO_LOG_LAUNCHER!r}")

    _FINGERPRINT = _gerar_fingerprint()
    _log_launcher(f"[MAIN] fingerprint={_FINGERPRINT}")

    # 0. Primeiro acesso via instalador
    if "--primeiro-acesso" in sys.argv:
        _log_launcher("[MAIN] modo primeiro acesso")
        _tela_primeiro_acesso()
        _log_launcher("[MAIN] fim primeiro acesso")
        return

    # 1. Execucao direta via agente
    _ARGS_MODULO = {"stm": "sao_luis", "pn": "padrao_nacional", "conf_sl": "conferencia_sao_luis", "conf_pn": "conferencia_pn"}
    _arg = next((a for a in sys.argv[1:] if a in _ARGS_MODULO), None)
    if _arg:
        _log_launcher(f"[MAIN] modo execucao direta via agente arg={_arg}")
        import importlib
        _modulo_nome = {"stm": "Sao_Luis", "pn": "Padrao_nacional", "conf_sl": "conferencias_sao_luis", "conf_pn": "conferencias_pn"}
        mod = importlib.import_module(_modulo_nome[_arg])
        _log_launcher(f"[MAIN] modulo importado={_modulo_nome[_arg]}")
        mod.main()
        _log_launcher(f"[MAIN] fim execucao direta arg={_arg}")
        return

    # 2. Verifica token (rwa://abrir?token=XYZ)
    token = _extrair_token_argv()
    if not token:
        _log_launcher("[MAIN] sem token; abrindo tela de bloqueio")
        _tela_bloqueio()
        _log_launcher("[MAIN] fim por ausencia de token")
        return

    _log_launcher("[MAIN] modo portal/protocolo detectado")
    if not _adquirir_lock_launcher():
        _log_launcher("[MAIN] encerrando: lock nao adquirido")
        return

    try:
        # 2. Dispara validação em background e exibe splash simultaneamente
        _resultado = {"ok": None, "email": "", "cliente": "", "erro": ""}

        def _bg():
            _log_launcher("[TOKEN] iniciando validacao no servidor")
            resp = _post("/auth/validar-token-launcher", {
                "token":       token,
                "fingerprint": _FINGERPRINT,
            })
            _resultado["ok"]      = resp.get("ok", False)
            _resultado["email"]   = resp.get("email", "")
            _resultado["cliente"] = resp.get("cliente", "")
            _resultado["erro"]    = resp.get("erro", "Token inválido ou expirado.")
            _log_launcher(f"[TOKEN] resposta validacao ok={_resultado['ok']} email={_resultado['email']!r} cliente={_resultado['cliente']!r} erro={_resultado['erro']!r}")

        t = threading.Thread(target=_bg, daemon=True)
        t.start()
        _log_launcher("[MAIN] iniciando validacao sem splash")
        t.join(timeout=10)
        _log_launcher(f"[TOKEN] thread validacao viva apos join={t.is_alive()}")

        # 3. Checa resultado da validação
        if not _resultado["ok"]:
            _log_launcher(f"[MAIN] token recusado; abrindo bloqueio erro={_resultado['erro']!r}")
            _tela_bloqueio(mensagem=_resultado["erro"])
            return

        _EMAIL_SESSAO = _resultado["email"]
        _NOME_CLIENTE = _resultado["cliente"]
        _log_launcher(f"[MAIN] token ok; email={_EMAIL_SESSAO!r} cliente={_NOME_CLIENTE!r}")

        # 4. System tray
        _log_launcher("[MAIN] iniciando system tray")
        _iniciar_tray()
        _log_launcher("[MAIN] system tray iniciado")

        # 5. Painel principal
        _log_launcher("[MAIN] chamando painel principal")
        _painel_principal()
        _log_launcher("[MAIN] painel principal encerrado")
    finally:
        _log_launcher("[MAIN] finally; liberando lock")
        _liberar_lock_launcher()


if __name__ == "__main__":
    try:
        main()
        _log_launcher("[MAIN] encerrado sem exception")
    except SystemExit as e:
        _log_launcher(f"[MAIN] SystemExit code={getattr(e, 'code', None)!r}")
        raise
    except Exception:
        _log_launcher("[ERRO_FATAL] exception nao tratada no launcher")
        _log_launcher(traceback.format_exc())
        raise
