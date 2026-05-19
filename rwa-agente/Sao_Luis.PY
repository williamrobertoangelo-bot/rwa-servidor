import os
import base64
import re
import time
import shutil
import traceback
import tempfile
import calendar
import threading
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from urllib.parse import urljoin
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import sys as _sys

import openpyxl
import requests
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, InvalidSessionIdException

try:
    import pyautogui
except ImportError:
    pyautogui = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from PIL import Image
except ImportError:
    Image = None




# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACAO GLOBAL
# ─────────────────────────────────────────────────────────────────────────────

URL_LOGIN     = "https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/login/login.jsf"
if getattr(_sys, 'frozen', False):
    PASTA_COMANDO = os.path.dirname(_sys.executable)
else:
    PASTA_COMANDO = os.path.dirname(os.path.abspath(__file__))
PASTA_BASE    = PASTA_COMANDO  # Relatorios (CONCLUSAO, GERENCIAL) salvos na pasta do script
PASTA_DIAG_TECNICO = os.path.join(PASTA_BASE, "DIAGNOSTICO TECNICO")
SENHAS_XLSX   = os.path.join(PASTA_COMANDO, "SENHAS SAO LUIS.xlsx")

hoje        = datetime.today()
competencia = hoje - relativedelta(months=1)
PERIODO     = competencia.strftime("%m/%Y")
PASTA_MES   = competencia.strftime("%m%Y")

# Periodo dd/mm/aaaa para o Talao Fiscal
DIA_INICIO = competencia.replace(day=1).strftime("%d/%m/%Y")
ultimo_dia = calendar.monthrange(competencia.year, competencia.month)[1]
DIA_FIM    = competencia.replace(day=ultimo_dia).strftime("%d/%m/%Y")

WAIT               = 30
WAIT_INSTABILIDADE = 28

_log_diagnostico = []  # Registro de erros para resumo final

# Checkboxes de situacao fiscal (Prestados e Tomados)
SITUACAO_FISCAL = {
    "j_id184:j_id234:0": (True,  "Emitida/Nao Paga"),
    "j_id184:j_id234:1": (True,  "Paga"),
    "j_id184:j_id234:2": (True,  "Estornada"),
    "j_id184:j_id234:3": (False, "Cancelada"),
    "j_id184:j_id234:4": (True,  "Isenta"),
    "j_id184:j_id234:5": (False, "Aguardando Conf. Cancelamento"),
    "j_id184:j_id234:6": (True,  "Pagamento Receita Federal"),
    "j_id184:j_id234:7": (True,  "Substituida"),
}

# Regimes que habilitam o Talao Fiscal
REGIMES_COM_TALAO = {"LUCRO_REAL", "LUCRO_PRESUMIDO"}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE TEXTO / REGIME
# ─────────────────────────────────────────────────────────────────────────────

def limpar_texto_planilha(valor):
    if valor is None:
        return ""
    texto = str(valor)
    texto = texto.replace("\xa0", " ").replace("\u200b", "")
    texto = texto.replace("\r", "").replace("\n", "")
    return texto.strip()


def normalizar_regime(regime):
    texto = limpar_texto_planilha(regime).upper()
    for origem, destino in [
        ("Á","A"),("À","A"),("Â","A"),("Ã","A"),
        ("É","E"),("Ê","E"),("Í","I"),
        ("Ó","O"),("Ô","O"),("Õ","O"),("Ú","U"),("Ç","C"),
    ]:
        texto = texto.replace(origem, destino)
    texto = re.sub(r"[^A-Z0-9]+", "_", texto).strip("_")
    if texto in ["SIMPLES", "SIMPLES_NACIONAL", "SN"]:
        return "SIMPLES"
    if texto in ["LUCRO_REAL", "REAL"]:
        return "LUCRO_REAL"
    if texto in ["LUCRO_PRESUMIDO", "PRESUMIDO", "NORMAL"]:
        return "LUCRO_PRESUMIDO"
    return texto


def regime_tem_talao(regime):
    """Retorna True se o regime habilita o Talao Fiscal."""
    return normalizar_regime(regime) in REGIMES_COM_TALAO


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE PASTA / NOME DE ARQUIVO
# ─────────────────────────────────────────────────────────────────────────────

def pasta_empresa(nome, salvamento):
    caminho = os.path.join(salvamento, nome, PASTA_MES)
    os.makedirs(caminho, exist_ok=True)
    return caminho


def pasta_interna(pasta_mes, codigo):
    """
    Retorna a pasta de TOMADOS.
    A pasta só é criada quando houver arquivo tomado para salvar.
    """
    return os.path.join(pasta_mes, f"tomados_{codigo}_{PASTA_MES}")


def nome_pdf_prestado(codigo):
    return f"servico_prestado_saoluis_{codigo}_{PASTA_MES}.pdf"

def nome_pdf_tomado(codigo):
    return f"servico_tomado_saoluis_{codigo}_{PASTA_MES}.pdf"

def nome_pdf_talao_nf(codigo):
    return f"talao_fiscal_{codigo}_{PASTA_MES}.pdf"

def nome_pdf_iss(codigo):
    return f"ISS_{codigo}_{PASTA_MES}.pdf"

def nome_xml_prestado(codigo):
    return f"servico_prestado_saoluis_{codigo}_{PASTA_MES}.xml"

def nome_xml_tomado(codigo):
    return f"servico_tomado_saoluis_{codigo}_{PASTA_MES}.xml"


# ─────────────────────────────────────────────────────────────────────────────
# LEITURA DA PLANILHA DE SENHAS
# Colunas: CODIGO | APELIDO | LOGIN | SENHA | REGIME
# ─────────────────────────────────────────────────────────────────────────────

def ler_empresas():
    wb = openpyxl.load_workbook(SENHAS_XLSX)
    ws = wb.active
    lista = []

    def normalizar_cab(valor):
        texto = limpar_texto_planilha(valor).upper()
        for origem, destino in [
            ("Á","A"),("À","A"),("Â","A"),("Ã","A"),
            ("É","E"),("Ê","E"),("Í","I"),
            ("Ó","O"),("Ô","O"),("Õ","O"),("Ú","U"),("Ç","C"),
        ]:
            texto = texto.replace(origem, destino)
        return texto

    cabecalhos = [normalizar_cab(cel.value) for cel in ws[1]]

    def indice(*nomes):
        for nome in nomes:
            if nome in cabecalhos:
                return cabecalhos.index(nome)
        return None

    idx_codigo        = indice("CODIGO", "CODIGO") or 0
    idx_empresa       = indice("APELIDO", "EMPRESA", "NOME") or 1
    idx_login         = indice("LOGIN") or 2
    idx_senha         = indice("SENHA") or 3
    idx_regime        = indice("REGIME") or 4
    idx_salvamento    = indice("PASTA")
    idx_tomados       = indice("TOMADOS")
    idx_xml_prestados = indice("XML PRESTADOS", "XML_PRESTADOS", "XML PRESTADO", "XML")

    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) <= idx_codigo or not row[idx_codigo]:
            continue

        codigo            = limpar_texto_planilha(row[idx_codigo])
        empresa           = limpar_texto_planilha(row[idx_empresa])    if len(row) > idx_empresa    else ""
        login             = limpar_texto_planilha(row[idx_login])      if len(row) > idx_login      else ""
        senha             = limpar_texto_planilha(row[idx_senha])      if len(row) > idx_senha      else ""
        regime_original   = limpar_texto_planilha(row[idx_regime])     if len(row) > idx_regime     else ""
        regime            = normalizar_regime(regime_original)
        nome_pasta        = f"{codigo}-{empresa}"
        salvamento        = limpar_texto_planilha(row[idx_salvamento]) if (idx_salvamento    is not None and len(row) > idx_salvamento)    else ""
        tomados_raw       = limpar_texto_planilha(row[idx_tomados])    if (idx_tomados       is not None and len(row) > idx_tomados)       else "SIM"
        tomados           = "NAO" if tomados_raw.upper() in ("NAO", "NÃO", "N", "NO") else "SIM"
        xml_prestados_raw = limpar_texto_planilha(row[idx_xml_prestados]) if (idx_xml_prestados is not None and len(row) > idx_xml_prestados) else "SIM"
        xml_prestados     = "NAO" if xml_prestados_raw.upper() in ("NAO", "NÃO", "N", "NO") else "SIM"

        lista.append({
            "codigo":          codigo,
            "empresa":         empresa,
            "nome":            nome_pasta,
            "login":           login,
            "senha":           senha,
            "regime":          regime,
            "regime_original": regime_original,
            "salvamento":      salvamento,
            "tomados":         tomados,
            "xml_prestados":   xml_prestados,
        })

    return lista


# ─────────────────────────────────────────────────────────────────────────────
# DRIVER CHROME
# ─────────────────────────────────────────────────────────────────────────────

def configurar_driver(pasta_download):
    options = Options()
    perfil_temp = tempfile.mkdtemp(prefix="ChromePerfilSTM_")

    options.add_argument(f"--user-data-dir={perfil_temp}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-session-crashed-bubble")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])

    prefs = {
    "download.default_directory":               pasta_download,
    "download.prompt_for_download":              False,
    "download.directory_upgrade":                True,
    "safebrowsing.enabled":                      False,
    "safebrowsing.disable_download_protection":  True,
    "plugins.always_open_pdf_externally":        True,
    "profile.default_content_setting_values.automatic_downloads": 1,
    "profile.default_content_settings.popups":  0,

    "credentials_enable_service": False,
    "profile.password_manager_enabled": False,
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    driver._perfil_temp_stm = perfil_temp
    driver._alto_volume_stm = False
    driver._itbi_stm        = False
    driver._selecionar_empresa_stm = False
    driver.set_page_load_timeout(WAIT_INSTABILIDADE)
    driver.set_script_timeout(WAIT_INSTABILIDADE)
    driver.maximize_window()
    time.sleep(1)
    driver.get("about:blank")
    time.sleep(1)
    return driver


# ─────────────────────────────────────────────────────────────────────────────
# UTILITARIOS GERAIS
# ─────────────────────────────────────────────────────────────────────────────

def forcar_fechar_driver(driver, timeout_s=8):
    """
    Fecha o driver com timeout real via thread.
    Se driver.quit() travar (renderer morto), mata o processo ChromeDriver + Chrome
    forcadamente via taskkill. Garante que a automacao nunca trava no fechamento.
    """
    if driver is None:
        return

    resultado = [False]

    def _quit():
        try:
            driver.quit()
        except Exception:
            pass
        resultado[0] = True

    t = threading.Thread(target=_quit, daemon=True)
    t.start()
    t.join(timeout=timeout_s)

    if not resultado[0]:
        print(f"  [DRIVER] driver.quit() nao respondeu em {timeout_s}s — encerrando Chrome forcadamente...")
        try:
            pid = driver.service.process.pid
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True, timeout=5
            )
            print(f"  [DRIVER] Processo Chrome (PID {pid}) encerrado via taskkill.")
        except Exception as _e_kill:
            print(f"  [DRIVER] Falha no taskkill: {_e_kill}")


class PortalInstabilidadeException(Exception):
    pass


class PortalBloqueioAceoException(Exception):
    """Levantada quando ACEO007-016 é detectado — exige fechar browser e aguardar."""
    pass


def tela_aparentemente_em_branco(driver):
    try:
        ready = driver.execute_script("return document.readyState")
    except Exception:
        ready = ""
    try:
        body = driver.find_element(By.TAG_NAME, "body").text or ""
    except Exception:
        body = ""
    try:
        html_len = driver.execute_script(
            "return document.body ? document.body.innerText.length : 0")
    except Exception:
        html_len = len(body)

    texto = re.sub(r"\s+", " ", body).strip()
    if ready in ["complete", "interactive"] and len(texto) < 20 and int(html_len or 0) < 50:
        return True
    return False


_RENDERER_ERROS = (
    "receiving message from renderer",
    "chrome not reachable",
    "unable to connect to renderer",
    "session deleted",
    "no such window",
)


def aguardar_elemento_essencial(driver, condicao, descricao, timeout=WAIT_INSTABILIDADE):
    try:
        return WebDriverWait(driver, timeout).until(condicao)
    except Exception as e:
        _msg = str(e).lower()
        if any(r in _msg for r in _RENDERER_ERROS):
            raise PortalInstabilidadeException(
                f"Renderer morto ao aguardar: {descricao}") from e
        if isinstance(e, TimeoutException):
            if tela_aparentemente_em_branco(driver):
                raise PortalInstabilidadeException(
                    f"Portal em branco ao aguardar: {descricao} ({timeout}s)") from e
        raise


import tempfile
from PIL import Image

def tirar_print(driver, caminho):
    """
    Print profissional:
    - Captura topo (Prestado/Tomado)
    - Captura região do resultado
    - Junta em UMA imagem final
    - Temporários ficam fora da pasta (não polui)
    """

    zoom_original = None

    try:
        if os.path.exists(caminho):
            os.remove(caminho)

        if Image is None:
            raise Exception("Pillow não disponível")

        # 🔹 cria arquivos temporários fora da pasta da empresa
        temp_dir = tempfile.gettempdir()
        caminho_topo = os.path.join(temp_dir, "stm_topo_tmp.png")
        caminho_base = os.path.join(temp_dir, "stm_base_tmp.png")

        # limpa se existir
        for arq in [caminho_topo, caminho_base]:
            if os.path.exists(arq):
                os.remove(arq)

        # 🔹 garante foco e maximização
        driver.switch_to.window(driver.current_window_handle)
        driver.maximize_window()
        time.sleep(0.5)

        # 🔹 zoom reduzido
        try:
            zoom_original = driver.execute_script("return document.body.style.zoom || ''")
            driver.execute_script("document.body.style.zoom='70%'")
            time.sleep(1)
        except:
            pass

        # ==========================
        # PRINT 1 — TOPO
        # ==========================
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)
        driver.save_screenshot(caminho_topo)

        # ==========================
        # PRINT 2 — RESULTADO
        # ==========================
        achou = False
        try:
            achou = driver.execute_script("""
                var texto = 'Nenhum resultado foi encontrado';
                var els = document.querySelectorAll('body *');
                for (var i=0;i<els.length;i++){
                    if ((els[i].innerText || '').includes(texto)){
                        els[i].scrollIntoView({block:'center'});
                        return true;
                    }
                }
                return false;
            """)
        except:
            pass

        if not achou:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        time.sleep(1)
        driver.save_screenshot(caminho_base)

        # ==========================
        # JUNÇÃO FINAL
        # ==========================
        img_topo = Image.open(caminho_topo).convert("RGB")
        img_base = Image.open(caminho_base).convert("RGB")

        largura = max(img_topo.width, img_base.width)
        altura = img_topo.height + img_base.height

        final = Image.new("RGB", (largura, altura), (255, 255, 255))
        final.paste(img_topo, (0, 0))
        final.paste(img_base, (0, img_topo.height))

        final.save(caminho)

        # 🔹 remove temporários
        for arq in [caminho_topo, caminho_base]:
            try:
                if os.path.exists(arq):
                    os.remove(arq)
            except:
                pass

        print(f"    [PRINT] Evidência FINAL (única) salva: {caminho}")

    except Exception as e:
        print(f"    [ERRO] Print combinado falhou: {e}")
        try:
            driver.save_screenshot(caminho)
            print(f"    [PRINT] Fallback simples salvo: {caminho}")
        except:
            pass

    finally:
        # reset visual
        try:
            if zoom_original:
                driver.execute_script("document.body.style.zoom=arguments[0];", zoom_original)
            else:
                driver.execute_script("document.body.style.zoom='100%'")
            driver.execute_script("window.scrollTo(0,0)")
        except:
            pass



def salvar_print_simples(driver, caminho, descricao="Evidência"):
    """
    Salva print simples da tela atual, sem juntar imagens.
    Usado para evidências específicas do Talão Fiscal.
    """
    try:
        pasta_print = os.path.dirname(caminho)
        if pasta_print:
            os.makedirs(pasta_print, exist_ok=True)
        if os.path.exists(caminho):
            os.remove(caminho)
        driver.save_screenshot(caminho)
        print(f"    [PRINT] {descricao} salva: {caminho}")
        return True
    except Exception as e:
        print(f"    [AVISO] Não foi possível salvar print ({descricao}): {e}")
        return False


def salvar_print_talao_fiscal(driver, pasta, codigo):
    """
    Salva evidência do Talão Fiscal com nome padronizado da rotina.
    Não interfere no PDF do Talão Fiscal.
    """
    caminho = os.path.join(pasta, f"talao_fiscal_{codigo}_{PASTA_MES}.png")
    salvar_print_simples(driver, caminho, "Evidência do Talão Fiscal")
    return caminho

def caminho_diagnostico_erros():
    os.makedirs(PASTA_DIAG_TECNICO, exist_ok=True)
    return os.path.join(PASTA_DIAG_TECNICO, f"DIAGNOSTICO_ERROS_SAO_LUIS_{PASTA_MES}.txt")


def _extrair_elemento_selenium(erro):
    """Extrai o ID/descricao do elemento que o Selenium tentava localizar."""
    try:
        tb = traceback.format_exc()
        # Extrai descricao da PortalInstabilidadeException (ex: "ao aguardar: botao Consultar Prestado")
        match = re.search(r'ao aguardar[:\s]+(.+?)\s*\(', str(erro))
        if match:
            return match.group(1).strip()
        # Extrai By.ID ou By.XPATH do traceback
        match = re.search(r'By\.(?:ID|XPATH|NAME|CLASS_NAME)[,\s]+"([^"]+)"', tb)
        if match:
            return f'By ID/XPATH: "{match.group(1)}"'
        # Extrai ultimo argumento de descricao do aguardar_elemento_essencial
        match = re.search(r'aguardar_elemento_essencial.*?["\']([^"\']{5,80})["\']', tb)
        if match:
            return match.group(1)
    except Exception:
        pass
    return None


def _capturar_msgs_erro_portal(driver):
    """Captura mensagens de erro visiveis no portal — filtra menu padrao."""
    msgs = []
    try:
        elementos = driver.find_elements(By.XPATH,
            "//*[contains(@class,'erro') or contains(@class,'error') "
            "or contains(@class,'alert') or contains(@class,'mensagem') "
            "or contains(@class,'warning') or contains(@id,'error') "
            "or contains(@id,'msg')]")
        for el in elementos:
            t = (el.text or "").strip()
            if t and 5 < len(t) < 500 and t not in msgs:
                msgs.append(t)
    except Exception:
        pass

    # Fallback: procura termos criticos no body
    if not msgs:
        try:
            body = driver.find_element(By.TAG_NAME, "body").text or ""
            for termo in ["ORA-", "ACEO", "paginaErro", "Exception", "java.lang"]:
                if termo in body:
                    idx = body.find(termo)
                    trecho = body[max(0, idx - 30):idx + 250].strip()
                    trecho = re.sub(r"\s+", " ", trecho)
                    msgs.append(trecho)
                    break
        except Exception:
            pass

    return msgs[:3]


def _filtrar_traceback(tb_texto):
    """Mantem apenas linhas do proprio script — remove stacktrace do chromedriver."""
    linhas_out = []
    linhas = tb_texto.split("\n")
    script_kw = "automacao_saoluis"
    capturar_codigo = False

    for linha in linhas:
        # Linha de excecao final — sempre inclui
        if re.match(r'^[A-Za-z][\w\.]*(?:Exception|Error|Warning)', linha.strip()):
            linhas_out.append(f"  {linha.strip()}")
            capturar_codigo = False
            continue
        # "The above exception..." — inclui
        if "The above exception" in linha:
            linhas_out.append(f"  {linha.strip()}")
            continue
        # Linha File do proprio script
        if 'File "' in linha and script_kw in linha:
            match = re.search(r'line (\d+), in (\S+)', linha)
            if match:
                linhas_out.append(f"  → {match.group(2)} (linha {match.group(1)})")
            capturar_codigo = True
            continue
        # Linha de codigo executada (logo apos File do script)
        if capturar_codigo and linha.strip() and not linha.strip().startswith("File"):
            linhas_out.append(f"      {linha.strip()[:120]}")
            capturar_codigo = False
            continue
        # Pula tudo do chromedriver
        if "chromedriver!" in linha or "GetHandleVerifier" in linha:
            capturar_codigo = False
            continue
        capturar_codigo = False

    return "\n".join(linhas_out) if linhas_out else "  (traceback nao extraido)"


def registrar_diagnostico(emp, etapa, erro=None, driver=None, pasta=None,
                           tentativa=1, fase=None, status_atual=None, tempo_decorrido=None,
                           decisao=None):
    caminho = caminho_diagnostico_erros()
    L = 90
    linhas = []

    linhas.append("=" * L)
    linhas.append(f"Data/Hora  : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    linhas.append(f"Empresa    : {emp.get('codigo','')} - {emp.get('empresa', emp.get('nome',''))}")
    linhas.append(f"Regime     : {emp.get('regime', 'NAO INFORMADO')}")
    linhas.append(f"Tentativa  : {tentativa}")
    linhas.append(f"Fase       : {fase or etapa}")
    linhas.append(f"Decisao    : {decisao or 'NAO INFORMADA'}")
    linhas.append("-" * L)

    if erro is not None:
        linhas.append(f"Erro tipo  : {type(erro).__name__}")
        linhas.append(f"Erro msg   : {str(erro)}")

    # Status ate o erro
    linhas.append("-" * L)
    linhas.append("Status ate o erro:")
    if status_atual:
        for campo, val in status_atual.items():
            v = val if val is not None else "N/A"
            linhas.append(f"  {campo:<18}: {v}")
    else:
        linhas.append("  (nao disponivel)")

    if tempo_decorrido is not None:
        m = int(tempo_decorrido) // 60
        s = int(tempo_decorrido) % 60
        linhas.append(f"Tempo ate o erro   : {m}min{s:02d}s")

    # Elemento Selenium + URL + titulo
    linhas.append("-" * L)
    elemento = _extrair_elemento_selenium(erro)
    if elemento:
        linhas.append(f"Selenium buscava  : {elemento}")

    if driver is not None:
        try:
            url = driver.current_url
            pagina = url.split("/")[-1] if "/" in url else url
            linhas.append(f"Pagina no erro    : {pagina}")
            linhas.append(f"URL completa      : {url}")
        except Exception:
            linhas.append("URL no erro       : nao capturada")
        try:
            linhas.append(f"Titulo            : {driver.title}")
        except Exception:
            linhas.append("Titulo            : nao capturado")

        # Mensagens de erro visiveis no portal
        linhas.append("-" * L)
        linhas.append("Erro portal (visivel):")
        msgs = _capturar_msgs_erro_portal(driver)
        if msgs:
            for msg in msgs:
                linhas.append(f"  >> {msg}")
        else:
            linhas.append("  [nenhuma mensagem de erro encontrada na pagina]")

    # Traceback filtrado
    linhas.append("-" * L)
    linhas.append("Traceback (apenas script):")
    if erro is not None:
        linhas.append(_filtrar_traceback(traceback.format_exc()))

    if pasta:
        linhas.append("-" * L)
        linhas.append(f"Pasta      : {pasta}")

    linhas.append("=" * L)
    linhas.append("")

    # Registra para resumo final
    _log_diagnostico.append({
        "empresa":   emp.get("empresa", emp.get("nome", "")),
        "codigo":    emp.get("codigo", ""),
        "fase":      fase or etapa,
        "erro_tipo": type(erro).__name__ if erro else "",
        "tentativa": tentativa,
        "decisao":   decisao or "NAO INFORMADA",
    })

    try:
        with open(caminho, "a", encoding="utf-8") as f:
            f.write("\n".join(linhas))
        print(f"    [DIAGNOSTICO] Erro registrado em: {caminho}")
    except Exception as e:
        print(f"    [AVISO] Nao foi possivel salvar diagnostico: {e}")


def finalizar_diagnostico():
    """Escreve bloco de resumo final no arquivo de diagnostico."""
    if not _log_diagnostico:
        return
    caminho = caminho_diagnostico_erros()
    L = 90
    linhas = []
    linhas.append("")
    linhas.append("=" * L)
    linhas.append("  RESUMO FINAL DE ERROS")
    linhas.append("=" * L)
    linhas.append(f"  Total de registros : {len(_log_diagnostico)}")
    linhas.append("")

    por_fase = {}
    for entry in _log_diagnostico:
        f = entry["fase"]
        por_fase.setdefault(f, []).append(str(entry["codigo"]))
    linhas.append("  Por fase:")
    for f, codigos in sorted(por_fase.items(), key=lambda x: -len(x[1])):
        linhas.append(f"    {f:<30}: {len(codigos)}x  (cod: {', '.join(codigos)})")

    linhas.append("")
    por_tipo = {}
    for entry in _log_diagnostico:
        t = entry["erro_tipo"]
        por_tipo[t] = por_tipo.get(t, 0) + 1
    linhas.append("  Por tipo de erro:")
    for tipo, qtd in sorted(por_tipo.items(), key=lambda x: -x[1]):
        linhas.append(f"    {tipo:<40}: {qtd}x")

    linhas.append("")
    por_decisao = {}
    for entry in _log_diagnostico:
        d = entry.get("decisao", "NAO INFORMADA")
        por_decisao.setdefault(d, []).append(str(entry["codigo"]))
    linhas.append("  Por decisao:")
    for d, codigos in sorted(por_decisao.items(), key=lambda x: -len(x[1])):
        linhas.append(f"    {d:<50}: {len(codigos)}x  (cod: {', '.join(codigos)})")

    linhas.append("=" * L)
    linhas.append("")

    try:
        with open(caminho, "a", encoding="utf-8") as f:
            f.write("\n".join(linhas))
        print(f"\n[DIAGNOSTICO] Resumo final gravado em: {caminho}")
    except Exception as e:
        print(f"  [AVISO] Nao foi possivel salvar resumo diagnostico: {e}")


def salvar_relatorio_parcial(resultados):
    try:
        if resultados:
            gerar_relatorio(resultados)
            print("  [RELATORIO] Parcial atualizado.")
    except Exception as e:
        print(f"  [AVISO] Nao foi possivel salvar relatorio parcial: {e}")


def aguardar_download(pasta, extensao, timeout=120, arquivos_antes=None):
    """
    Aguarda um arquivo NOVO aparecer na pasta com a extensao informada.
    Se detectar .crdownload, aguarda ele sumir (download ativo) sem limite fixo.
    Se nao aparecer nada no timeout, retorna False (instabilidade).
    """
    arquivos_antes = set(arquivos_antes or [])
    inicio = time.time()
    crdownload_detectado = False

    while True:
        # Verifica .crdownload ativo
        crdownloads = [f for f in os.listdir(pasta) if f.endswith(".crdownload")]
        if crdownloads:
            crdownload_detectado = True
            time.sleep(1)
            continue

        # Se tinha .crdownload e sumiu — download concluiu
        if crdownload_detectado:
            arquivos = [f for f in os.listdir(pasta)
                        if f.lower().endswith(extensao)
                        and f not in arquivos_antes
                        and not f.endswith(".crdownload")]
            if arquivos:
                return True

        # Sem .crdownload — verifica se arquivo ja chegou direto
        arquivos = [f for f in os.listdir(pasta)
                    if f.lower().endswith(extensao)
                    and f not in arquivos_antes
                    and not f.endswith(".crdownload")]
        if arquivos:
            return True

        # Timeout apenas se nunca detectou .crdownload
        if not crdownload_detectado and time.time() - inicio > timeout:
            return False

        time.sleep(1)


def renomear_ultimo_download(pasta, novo_nome, extensao=None):
    arquivos = [os.path.join(pasta, f) for f in os.listdir(pasta)
                if os.path.isfile(os.path.join(pasta, f)) and not f.endswith(".crdownload")]
    if extensao:
        arquivos = [a for a in arquivos if a.lower().endswith(extensao)]
    if not arquivos:
        return
    mais_recente = max(arquivos, key=os.path.getmtime)
    destino = os.path.join(pasta, novo_nome)
    if os.path.exists(destino):
        os.remove(destino)
    shutil.move(mais_recente, destino)


def renomear_download_novo(pasta, novo_nome, extensao, arquivos_antes):
    """
    Renomeia APENAS arquivo novo surgido depois do clique de download.
    Evita pegar PDF/ZIP/XML antigo da pasta e salvar com nome errado.
    """
    arquivos_antes = set(arquivos_antes or [])
    candidatos = []

    for nome in os.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        if not os.path.isfile(caminho):
            continue
        if nome.endswith(".crdownload"):
            continue
        if extensao and not nome.lower().endswith(extensao.lower()):
            continue
        if nome in arquivos_antes:
            continue
        candidatos.append(caminho)

    if not candidatos:
        return False

    mais_recente = max(candidatos, key=os.path.getmtime)
    destino = os.path.join(pasta, novo_nome)

    if os.path.abspath(mais_recente) == os.path.abspath(destino):
        return True

    if os.path.exists(destino):
        os.remove(destino)

    shutil.move(mais_recente, destino)
    return True


def _br_money_to_float(valor):
    try:
        return float(str(valor).replace(".", "").replace(",", "."))
    except Exception:
        return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# ANALISE PDF GERENCIAL
# ─────────────────────────────────────────────────────────────────────────────

def analisar_pdf_gerencial(pasta, arquivo_pdf):
    resumo = {"qtd_notas": "", "valor_total": "", "iss_proprio": "", "iss_retido": ""}

    caminho_pdf = os.path.join(pasta, arquivo_pdf)
    if pdfplumber is None:
        print("    [AVISO] pdfplumber nao instalado. Analise gerencial nao sera preenchida.")
        return resumo
    if not os.path.exists(caminho_pdf):
        print(f"    [AVISO] PDF nao encontrado: {arquivo_pdf}")
        return resumo

    qtd = 0
    valor_total = 0.0
    iss_proprio = 0.0
    iss_retido  = 0.0

    TIPOS_NOTA = ("NFSe", "Escriturada", "NFSe(RPS)", "Avulsa", "Manual")

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                for linha in texto.split("\n"):
                    if not any(tipo in linha for tipo in TIPOS_NOTA):
                        continue
                    partes = linha.split()
                    if len(partes) < 8:
                        continue
                    qtd += 1
                    valores_monetarios = [
                        p for p in partes if "," in p and any(ch.isdigit() for ch in p)
                    ]
                    if valores_monetarios:
                        valor_total += _br_money_to_float(valores_monetarios[0])
                        iss_valor    = _br_money_to_float(valores_monetarios[-1])
                    else:
                        iss_valor = 0.0
                    linha_upper = linha.upper()
                    if "PROPRIO" in linha_upper or "PRÓPRIO" in linha_upper:
                        iss_proprio += iss_valor
                    elif "RETIDO" in linha_upper:
                        iss_retido += iss_valor

        resumo["qtd_notas"]   = qtd
        resumo["valor_total"] = valor_total
        resumo["iss_proprio"] = iss_proprio
        resumo["iss_retido"]  = iss_retido
        print(f"    [OK] PDF analisado: {qtd} nota(s), total {valor_total:.2f}")
        return resumo

    except Exception as e:
        print(f"    [AVISO] Nao foi possivel analisar PDF gerencial: {e}")
        return resumo


# ─────────────────────────────────────────────────────────────────────────────
# LOGIN
# ─────────────────────────────────────────────────────────────────────────────

BLOQUEIO_LOGIN = "BLOQUEADO"  # Sinal especial retornado quando portal bloqueia login

def _tentar_login_uma_vez(driver, login, senha):
    """
    Executa uma tentativa de login e retorna:
      True          — login OK
      False         — credenciais incorretas
      BLOQUEIO_LOGIN — portal bloqueou (ACEO007-016), aguardar e retentar
    """
    try:
        driver.get(URL_LOGIN)
    except TimeoutException as e:
        raise PortalInstabilidadeException(
            f"Portal nao carregou a tela de login em {WAIT_INSTABILIDADE}s") from e

    # Aguarda campo de login aparecer (substitui sleep fixo de 4s)
    campo_user = aguardar_elemento_essencial(driver,
        EC.presence_of_element_located(
            (By.XPATH, "(//input[@type='text' or not(@type)])[1]")),
        "campo de login")

    try:
        texto_inicial = driver.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        texto_inicial = ""

    # Detecta bloqueio ANTES de preencher (mensagem ja na tela de login)
    if "aceo007-016" in texto_inicial or "logins sucessivos" in texto_inicial:
        return BLOQUEIO_LOGIN

    campo_user.click()
    campo_user.clear()
    time.sleep(0.2)
    for c in str(login):
        campo_user.send_keys(c)
        time.sleep(0.05)
    time.sleep(0.3)

    campo_pass = driver.find_element(By.XPATH, "//input[@type='password']")
    campo_pass.click()
    campo_pass.clear()
    time.sleep(0.2)
    for c in str(senha):
        campo_pass.send_keys(c)
        time.sleep(0.05)
    time.sleep(0.3)

    btn = driver.find_element(By.XPATH,
        "//button[contains(text(),'Entrar')] | //input[@value='Entrar']")
    btn.click()
    time.sleep(3)

    try:
        texto_pagina = driver.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        texto_pagina = driver.page_source.lower()

    # Detecta bloqueio APOS clicar em Entrar
    if "aceo007-016" in texto_pagina or "logins sucessivos" in texto_pagina:
        return BLOQUEIO_LOGIN

    # Erros catalogados — retorna a mensagem exata do portal
    if "aceo007-008" in texto_pagina or "usuário não existe" in texto_pagina or "usuario nao existe" in texto_pagina:
        return "ACEO007-008: Este usuário não existe na prefeitura."
    if "aceo007-001" in texto_pagina or "senha incorreta" in texto_pagina:
        return "ACEO007-001: Senha incorreta. Por favor, tente novamente."

    # Outros erros de credencial conhecidos (fallback genérico)
    mensagens_erro_genericas = [
        "login incorreto", "usuário incorreto", "usuario incorreto",
        "usuário ou senha", "usuario ou senha", "login ou senha", "credenciais",
        "inválido", "invalido", "não confere", "nao confere",
    ]
    if any(msg in texto_pagina for msg in mensagens_erro_genericas):
        # Tenta capturar o texto exato da mensagem de erro do portal
        try:
            msgs = driver.find_elements(By.XPATH,
                "//*[contains(@class,'erro') or contains(@class,'error') or contains(@class,'alert') or contains(@class,'mensagem')]")
            for m in msgs:
                t = (m.text or "").strip()
                if t and len(t) < 300:
                    return f"ERRO_LOGIN: {t}"
        except Exception:
            pass
        return "ERRO_LOGIN: Credenciais inválidas (erro não catalogado)."

    return True


def fazer_login(driver, login, senha):
    """
    Realiza login com tratamento de bloqueio do portal (ACEO007-016).
    Se bloqueado: levanta PortalBloqueioAceoException para que o main()
    feche o browser, aguarde com browser fechado e retente limpo.
    Retorna True (ok) ou string com mensagem de erro para o relatorio.
    """
    print("  [LOGIN] Acessando portal...")
    resultado = _tentar_login_uma_vez(driver, login, senha)

    if resultado == BLOQUEIO_LOGIN:
        print("  [BLOQUEIO] Portal bloqueou login (ACEO007-016) — fechando browser para espera limpa.")
        raise PortalBloqueioAceoException("ACEO007-016: logins consecutivos bloqueados.")

    # Erro de login — qualquer string retornada e a mensagem exata para o relatorio
    if isinstance(resultado, str) and resultado != BLOQUEIO_LOGIN:
        print(f"    [LOGIN] ERRO — {resultado}")
        return resultado

    print("    [OK] Login realizado com sucesso.")
    return True


# ─────────────────────────────────────────────────────────────────────────────
# NAVEGACAO — PRESTADOS E TOMADOS (PDF)
# ─────────────────────────────────────────────────────────────────────────────

def _selecionar_empresa_stm(driver, cnpj):
    """
    CAMINHO EXCECAO — SELECIONAR EMPRESA (detectado 03/05/2026)
    Exclusivo para empresas cujo portal exige selecao de empresa
    apos clicar em PRESTADOS E TOMADOS.
    Seta flag driver._selecionar_empresa_stm = True para uso posterior no XML.
    """
    print("  [PDF] Caminho: EXCECAO — SELECIONAR EMPRESA")
    print("    [SELECIONAR EMPRESA] Preenchendo CNPJ...")

    # Seta flag para identificar esse caminho nas etapas seguintes (XML)
    driver._selecionar_empresa_stm = True

    # 1. Localizar campo CNPJ
    campo_cnpj = None
    try:
        campo_cnpj = aguardar_elemento_essencial(driver,
            EC.presence_of_element_located(
                (By.XPATH,
                 "//input[contains(@id,'cnpj') or contains(@name,'cnpj') or "
                 "contains(@id,'CNPJ') or contains(@name,'CNPJ')]")),
            "campo CNPJ selecao empresa", timeout=15)
    except Exception:
        pass

    if campo_cnpj is None:
        try:
            campo_cnpj = aguardar_elemento_essencial(driver,
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@maxlength='18']")),
                "campo CNPJ fallback maxlength", timeout=10)
        except Exception:
            pass

    if campo_cnpj is None:
        campo_cnpj = aguardar_elemento_essencial(driver,
            EC.presence_of_element_located(
                (By.XPATH, "//input[@type='text'][1]")),
            "campo CNPJ fallback primeiro input", timeout=10)

    campo_cnpj.click()
    driver.execute_script("arguments[0].value = '';", campo_cnpj)
    campo_cnpj.clear()
    time.sleep(0.2)
    for c in str(cnpj):
        campo_cnpj.send_keys(c)
        time.sleep(0.04)
    time.sleep(0.5)
    print(f"    [OK] CNPJ preenchido: {cnpj}")

    # 2. Clicar em Consultar
    btn_consultar = aguardar_elemento_essencial(driver,
        EC.element_to_be_clickable(
            (By.XPATH,
             "//*[contains(@value,'Consultar')] | "
             "//a[normalize-space(.)='Consultar'] | "
             "//button[normalize-space(.)='Consultar'] | "
             "//*[contains(text(),'Consultar')]")),
        "botao Consultar selecao empresa", timeout=15)
    driver.execute_script("arguments[0].click();", btn_consultar)
    time.sleep(3)
    print("    [OK] Consulta enviada. Aguardando resultado...")

    # 3. Clicar no checkmark (botao Acoes) — primeira linha de resultado
    btn_acoes = None
    for _xpath_acoes in [
        "//tbody/tr[1]/td[last()]//input[@type='image' or @type='submit' or @type='button']",
        "//tbody/tr[1]/td[last()]//a",
        "//tbody/tr[1]/td[last()]//button",
        "//input[@type='image'][1]",
        "//a[contains(@title,'Selecionar') or contains(@onclick,'selecionar')][1]",
    ]:
        try:
            btn_acoes = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.XPATH, _xpath_acoes)))
            print("    [OK] Botao Acoes localizado.")
            break
        except Exception:
            continue

    if btn_acoes is None:
        raise PortalInstabilidadeException(
            "Botao Acoes (checkmark) nao localizado na tela Selecionar Empresa")

    driver.execute_script("arguments[0].click();", btn_acoes)
    time.sleep(2)
    print("    [OK] Checkmark clicado — aguardando botao de confirmacao...")

    # 4. Clicar em "Confirmar Selecao de Empresa" (aparece apos checkmark)
    try:
        btn_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 "//*[contains(@value,'Confirmar') and contains(@value,'Empresa')] | "
                 "//*[contains(text(),'Confirmar') and contains(text(),'Empresa')] | "
                 "//*[contains(@value,'Confirmar Sele')] | "
                 "//*[contains(text(),'Confirmar Sele')]")))
        driver.execute_script("arguments[0].click();", btn_confirmar)
        time.sleep(3)
        print("    [OK] Confirmacao de empresa concluida.")
    except Exception:
        print("    [AVISO] Botao 'Confirmar Selecao de Empresa' nao apareceu — prosseguindo.")

    # 5. Verificar que saiu da tela de selecao
    try:
        url_final = driver.current_url.lower()
        if "selecionarempresa" in url_final:
            raise PortalInstabilidadeException(
                "Ainda na tela de selecao de empresa apos confirmacao — portal instavel")
        if "bemvindo" in url_final:
            raise PortalInstabilidadeException(
                "Redirecionado para bemVindo apos selecao de empresa — portal instavel")
    except PortalInstabilidadeException:
        raise
    except Exception:
        pass

    print("    [OK] Empresa confirmada. Tela de Prestados e Tomados pronta.")


def navegar_prestados_tomados(driver, cnpj=""):
    print("  [PDF] Acessando Prestados e Tomados...")

    # ── Detecta caminho ────────────────────────────────────────────────────────
    _caminho_itbi = False
    try:
        menu = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[contains(.,'ITBI')] | //a[contains(.,'ITBI')]")
            )
        )
        print("  [PDF] Caminho: EXCECAO — ITBI")
        _caminho_itbi = True
        driver._itbi_stm = True
    except Exception:
        print("  [PDF] Caminho: PADRAO — Alvara de Funcionamento")
        menu = aguardar_elemento_essencial(driver,
            EC.presence_of_element_located(
                (By.XPATH, "//button[contains(.,'Alvar')] | //a[contains(.,'Alvar')]")),
            "menu Alvara de Funcionamento")

    driver.execute_script("arguments[0].click();", menu)
    time.sleep(2)

    nfe = aguardar_elemento_essencial(driver, EC.visibility_of_element_located(
        (By.XPATH, "//a[contains(@onclick,'Nota Fiscal Eletronica')]")),
        "menu Nota Fiscal Eletronica")
    driver.execute_script("arguments[0].click();", nfe)
    print("    [OK] Nota Fiscal Eletronica aberta.")
    time.sleep(4)

    pt = aguardar_elemento_essencial(driver, EC.visibility_of_element_located(
        (By.XPATH, "//a[contains(normalize-space(.),'PRESTADOS E TOMADOS')]")),
        "tela Prestados e Tomados")
    driver.execute_script("arguments[0].click();", pt)
    time.sleep(3)

    # ── Detecta excecao: portal exige Selecionar Empresa apos PRESTADOS E TOMADOS ──
    _url_pos_pt = driver.current_url.lower()

    if "selecionarempresa" in _url_pos_pt:
        _selecionar_empresa_stm(driver, cnpj)
    else:
        print("    [OK] Tela Prestados e Tomados aberta.")

    if _caminho_itbi:
        print("    [INFO] Acesso via ITBI — XML nao disponivel no portal STM. Consultar pelo Portal Nacional.")

    # Verifica se o portal realmente navegou — portal travado permanece em bemVindo.jsf
    try:
        url_atual = driver.current_url.lower()
        if "bemvindo.jsf" in url_atual:
            raise PortalInstabilidadeException(
                "Portal travado em bemVindo.jsf apos navegacao — pagina nao carregou"
            )
    except PortalInstabilidadeException:
        raise
    except Exception:
        pass


def ajustar_checkboxes(driver, contexto="PDF"):
    print(f"  [{contexto}] Conferindo filtros de situacao fiscal...")
    time.sleep(2)
    for cb_id, (deve_marcar, nome) in SITUACAO_FISCAL.items():
        try:
            cb = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, cb_id)))
            marcado = cb.is_selected()
            if deve_marcar and not marcado:
                driver.execute_script("arguments[0].click();", cb)
                print(f"    [OK] Marcado   : {nome}")
                time.sleep(0.3)
            elif not deve_marcar and marcado:
                driver.execute_script("arguments[0].click();", cb)
                print(f"    [OK] Desmarcado: {nome}")
                time.sleep(0.3)
        except Exception:
            print(f"    [AVISO] Campo nao localizado: {nome}")


def preencher_periodo(driver):
    mes = competencia.strftime("%m")
    ano = competencia.strftime("%Y")
    print(f"  [PDF] Periodo: {mes}/{ano}")
    try:
        campos_mes = driver.find_elements(By.XPATH, "//input[@type='text'][@maxlength='2']")
        campos_ano = driver.find_elements(By.XPATH, "//input[@type='text'][@maxlength='4']")
        if len(campos_mes) >= 2 and len(campos_ano) >= 2:
            for campo in [campos_mes[0], campos_mes[1]]:
                driver.execute_script("arguments[0].value = '';", campo)
                campo.send_keys(mes)
                time.sleep(0.2)
            for campo in [campos_ano[0], campos_ano[1]]:
                driver.execute_script("arguments[0].value = '';", campo)
                campo.send_keys(ano)
                time.sleep(0.2)
        print("    [OK] Periodo preenchido.")
    except Exception as e:
        print(f"    [ERRO] Falha ao preencher periodo: {e}")


def selecionar_notas_prestadas(driver):
    """Radio 'Notas Prestadas' — ID: j_id184:j_id190:0"""
    try:
        radio = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "j_id184:j_id190:0")))
        if not radio.is_selected():
            driver.execute_script("arguments[0].click();", radio)
            time.sleep(0.5)
        print("    [OK] Notas Prestadas selecionado.")
    except Exception as e:
        print(f"    [AVISO] Radio Prestadas nao confirmado: {e}")


def selecionar_notas_tomadas(driver):
    """Radio 'Notas Tomadas' — ID: j_id184:j_id190:1"""
    try:
        radio = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "j_id184:j_id190:1")))
        driver.execute_script("arguments[0].click();", radio)
        time.sleep(0.5)
        print("    [OK] Notas Tomadas selecionado.")
    except Exception as e:
        print(f"    [AVISO] Radio Tomadas nao confirmado: {e}")


def _fechar_janela_extra(driver):
    """Fecha aba extra aberta pelo Imprimir e retorna para a principal."""
    if len(driver.window_handles) > 1:
        try:
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(1)
            driver.close()
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])
        except Exception:
            try:
                driver.switch_to.window(driver.window_handles[0])
            except Exception:
                pass


def detectar_total_paginas(driver):
    """Detecta total de paginas apos consulta. Retorna int ou 0 se nao detectar."""
    try:
        paginas = driver.find_elements(
            By.CSS_SELECTOR, "td.rich-datascr-inact, td.rich-datascr-act")
        numeros = []
        for p in paginas:
            t = p.text.strip()
            if t.isdigit():
                numeros.append(int(t))
        if numeros:
            return max(numeros)
    except Exception:
        pass
    return 0


def _tentar_imprimir_pdf(driver, pasta, nome_final, label):
    """
    Tenta imprimir e salvar um PDF uma vez.
    Deleta o arquivo alvo antes de cada tentativa.
    Retorna True se salvo com sucesso, False caso contrário.
    """
    # Log do radio selecionado (apenas informativo)
    try:
        radio_prest = driver.find_element(By.ID, "j_id184:j_id190:0")
        tipo_radio = "Prestadas" if radio_prest.is_selected() else "Tomadas"
        print(f"    [INFO] Radio selecionado: Notas {tipo_radio}")
    except Exception:
        pass

    # Deleta arquivo alvo antes do download
    arquivo_alvo = os.path.join(pasta, nome_final)
    if os.path.exists(arquivo_alvo):
        os.remove(arquivo_alvo)
        print(f"    [INFO] Arquivo anterior removido: {nome_final}")

    arquivos_antes = os.listdir(pasta)

    btn_imp = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
        (By.XPATH, "//*[contains(text(),'Imprimir Relat')] | //*[contains(@value,'Imprimir Relat')]")),
        f"botao Imprimir Relatorio {label}")
    btn_imp.click()
    time.sleep(2)
    _fechar_janela_extra(driver)

    if aguardar_download(pasta, ".pdf", timeout=10, arquivos_antes=arquivos_antes):
        if renomear_download_novo(pasta, nome_final, ".pdf", arquivos_antes):
            print(f"    [OK] PDF {label} salvo: {nome_final}")
            return True

        print(f"    [AVISO] PDF {label} apareceu, mas nao foi possivel identificar arquivo novo para renomear.")
        return False

    return False


def _baixar_pdf_alto_volume_requests(driver, pasta, nome_final, label):
    """
    Alto volume (total_pags >= 10): reproduz o POST do botao "Imprimir Relatorio"
    diretamente via requests, usando os cookies da sessao Selenium.

    Botao confirmado pelo diagnostico:
      id/name : j_id184:j_id399
      form    : j_id184
      target  : _blank
      onclick : jsfcljs(document.forms['j_id184'],'j_id184:j_id399,j_id184:j_id399','_blank')
      ViewState: j_id6 (valor dinamico — sempre coletado do DOM)

    O portal retorna o PDF diretamente no corpo da resposta (Content-Type: application/pdf).
    """
    URL_BASE = "https://stm.semfaz.saoluis.ma.gov.br"
    URL_RELATORIO = (
        URL_BASE
        + "/sistematributario/jsp/relatorioServicosPrestadosTomados"
        + "/relatorioServicosPrestadosTomadosFiltro.jsf"
    )

    # ── 1. Coletar cookies da sessao Selenium ────────────────────────────────
    sessao = requests.Session()
    for cookie in driver.get_cookies():
        sessao.cookies.set(cookie["name"], cookie["value"])

    # ── 2. Coletar todos os inputs do formulario j_id184 ─────────────────────
    dados = {}
    try:
        inputs = driver.find_elements(
            By.XPATH,
            "//form[@id='j_id184']//input | //form[@id='j_id184']//select | //form[@id='j_id184']//textarea"
        )
        for el in inputs:
            tipo  = (el.get_attribute("type") or "text").lower()
            nome  = el.get_attribute("name") or ""
            valor = el.get_attribute("value") or ""
            if not nome:
                continue
            if tipo == "hidden":
                dados[nome] = valor
            elif tipo in ("text", "textarea"):
                dados[nome] = el.get_attribute("value") or ""
            elif tipo == "checkbox":
                if el.is_selected():
                    dados[nome] = valor
            elif tipo == "radio":
                if el.is_selected():
                    dados[nome] = valor
            elif el.tag_name == "select":
                # pega option selecionado
                try:
                    selecionado = el.find_element(By.XPATH, ".//option[@selected]")
                    dados[nome] = selecionado.get_attribute("value") or ""
                except Exception:
                    dados[nome] = valor
    except Exception as e:
        print(f"    [ALTO VOLUME] Aviso ao coletar inputs: {e}")

    # ── 3. Adicionar campo do botao Imprimir (evento JSF) ────────────────────
    dados["j_id184"]         = "j_id184"
    dados["j_id184:j_id399"] = "j_id184:j_id399"

    # ViewState dinamico — pegar do DOM (pode ter mudado apos Consultar)
    try:
        vs = driver.find_element(
            By.XPATH,
            "//form[@id='j_id184']//input[@name='javax.faces.ViewState']"
        )
        dados["javax.faces.ViewState"] = vs.get_attribute("value") or "j_id6"
    except Exception:
        dados["javax.faces.ViewState"] = "j_id6"

    print(f"    [ALTO VOLUME] POST para o portal ({len(dados)} campos)...")

    # ── 4. Enviar POST e salvar PDF ──────────────────────────────────────────
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": URL_RELATORIO,
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    }

    resposta = sessao.post(URL_RELATORIO, data=dados, headers=headers, timeout=120, stream=True)

    content_type = resposta.headers.get("Content-Type", "")
    print(f"    [ALTO VOLUME] Resposta: HTTP {resposta.status_code} | Content-Type: {content_type}")

    if resposta.status_code == 200 and "pdf" in content_type.lower():
        arquivo_alvo = os.path.join(pasta, nome_final)
        if os.path.exists(arquivo_alvo):
            os.remove(arquivo_alvo)
        tamanho = 0
        with open(arquivo_alvo, "wb") as f:
            for chunk in resposta.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)
                    tamanho += len(chunk)
        tamanho_kb = tamanho // 1024
        print(f"    [OK] PDF {label} salvo (alto volume via requests): {nome_final} ({tamanho_kb} KB)")
        return "PDF"

    # Nao veio PDF — salva diagnostico HTML para analise
    diag_path = os.path.join(pasta, f"diagnostico_pdf_requests_{label.lower()}.html")
    try:
        with open(diag_path, "wb") as f:
            f.write(resposta.content)
        print(f"    [ALTO VOLUME] Resposta nao era PDF. Diagnostico salvo: {diag_path}")
    except Exception:
        pass

    raise Exception(
        f"Resposta nao era PDF: HTTP {resposta.status_code} | Content-Type: {content_type}"
    )


def _consultar_e_imprimir(driver, pasta, nome_final, label, codigo=None, pasta_print=None):
    """
    Clica em Consultar, verifica 'sem movimento', clica em Imprimir Relatorio
    e aguarda o PDF. 2 tentativas por arquivo. Se falhar lança InstabilidadeDownload.
    pasta_print: onde salvar prints de evidência (sem_movimento). Default = pasta.
    Retorna: 'PDF' | 'Sem movimento' | 'Erro'
    """
    _pasta_print = pasta_print if pasta_print else pasta
    btn = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
        (By.ID, "j_id184:j_id285")),
        f"botao Consultar {label}")
    driver.execute_script("arguments[0].click();", btn)
    print(f"    [OK] Consulta {label} enviada. Aguardando resultado...")
    time.sleep(5)

    # Aguarda modal "Carregando..." sumir — alto volume como CONSTROEM exibe esse modal
    # Para empresas normais o modal nao aparece e a espera passa imediatamente
    try:
        WebDriverWait(driver, 120).until_not(
            lambda d: "carregando" in (d.find_element(By.TAG_NAME, "body").text or "").lower()
        )
        print(f"    [OK] Portal concluiu carregamento do resultado.")
    except Exception:
        pass  # Se timeout ou erro, segue normalmente

    body = driver.find_element(By.TAG_NAME, "body").text
    if "nenhum" in body.lower() and "resultado" in body.lower():
        print(f"    [INFO] Sem movimento — {label}.")
        tirar_print(driver, os.path.join(_pasta_print, f"sem_movimento_{label.lower()}.png"))
        return "Sem movimento"

    # Detecta alto volume — download direto via requests (POST autenticado)
    total_pags = detectar_total_paginas(driver)
    if total_pags >= 10:
        driver._alto_volume_stm = True
        print(f"    [ALTO VOLUME] {total_pags} paginas detectadas — empresa com volume acima do limite padrao.")
        print(f"    [ALTO VOLUME] Iniciando download via POST autenticado...")
        try:
            return _baixar_pdf_alto_volume_requests(driver, pasta, nome_final, label)
        except Exception as e:
            print(f"    [ALTO VOLUME] Falha: {e}")
        return "Erro"

    # Tentativa 1
    if _tentar_imprimir_pdf(driver, pasta, nome_final, label):
        return "PDF"

    # PDF nao chegou — aguarda 30s (pode estar atrasado no browser)
    print(f"    [AVISO] PDF {label} nao chegou. Aguardando 30s...")
    _arquivos_antes_espera = set(os.listdir(pasta))
    time.sleep(30)

    # Verifica se o PDF chegou durante a espera antes de retentar
    if renomear_download_novo(pasta, nome_final, ".pdf", _arquivos_antes_espera):
        print(f"    [OK] PDF {label} chegou durante a espera: {nome_final}")
        return "PDF"

    # Nao chegou na espera — tenta imprimir de novo
    print(f"    [INFO] PDF nao chegou na espera. Retentando impressao...")
    if _tentar_imprimir_pdf(driver, pasta, nome_final, label):
        return "PDF"

    # Ambas falharam → instabilidade
    print(f"    [INSTABILIDADE] PDF {label} nao gerado apos 2 tentativas. Empresa para fila.")
    raise PortalInstabilidadeException(f"PDF {label} nao gerado apos 2 tentativas")


def baixar_pdfs_prestado_e_tomado(driver, pasta, pasta_int, codigo, cnpj="", baixar_tomado=True):
    """
    Navega para Prestados e Tomados, baixa Prestado e (opcionalmente) Tomado.
    PDF prestado vai para pasta_int. PDF tomado vai para pasta_int se baixar_tomado=True.
    Retorna: (status_prestado, status_tomado)
    """
    print("  [PDF] Iniciando download Prestado + Tomado...")
    try:
        navegar_prestados_tomados(driver, cnpj=cnpj)
        selecionar_notas_prestadas(driver)
        ajustar_checkboxes(driver, "PDF PRESTADO")
        time.sleep(1)
        preencher_periodo(driver)
        time.sleep(1)

        status_prestado = _consultar_e_imprimir(
            driver, pasta, nome_pdf_prestado(codigo), "Prestado",
            codigo=codigo, pasta_print=pasta)

        if status_prestado == "PDF":
            pass

        # ── CONSULTA AGUARDANDO CONFIRMACAO CANCELAMENTO (somente Prestadas) ──
        try:
            print("  [PDF] Consultando Aguardando Confirmacao Cancelamento (Prestadas)...")
            _nome_agc = f"Aguardando_Confirmacao_Cancelamento_{PASTA_MES}"
            for _cb_id in SITUACAO_FISCAL:
                try:
                    _cb = driver.find_element(By.ID, _cb_id)
                    _marcado = _cb.is_selected()
                    if _cb_id == "j_id184:j_id234:5":
                        if not _marcado:
                            driver.execute_script("arguments[0].click();", _cb)
                    else:
                        if _marcado:
                            driver.execute_script("arguments[0].click();", _cb)
                except Exception:
                    pass
            time.sleep(1)
            _btn_agc = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
                (By.ID, "j_id184:j_id285")), "botao Consultar AGC")
            driver.execute_script("arguments[0].click();", _btn_agc)
            time.sleep(5)
            try:
                WebDriverWait(driver, 120).until_not(
                    lambda d: "carregando" in (d.find_element(By.TAG_NAME, "body").text or "").lower()
                )
            except Exception:
                pass
            _body_agc = driver.find_element(By.TAG_NAME, "body").text
            if "nenhum" in _body_agc.lower() and "resultado" in _body_agc.lower():
                print("  [AGC] Nenhum resultado — salvando print.")
                tirar_print(driver, os.path.join(pasta, f"{_nome_agc}.png"))
            else:
                print("  [AGC] Resultado encontrado — baixando PDF.")
                try:
                    _consultar_e_imprimir(driver, pasta, f"{_nome_agc}.pdf",
                                          "AGC", codigo=codigo, pasta_print=pasta)
                except Exception as _e_agc:
                    print(f"  [AGC] Falha ao baixar PDF AGC: {_e_agc}")
            # Restaura checkboxes para o padrao coletivo
            ajustar_checkboxes(driver, "PDF PRESTADO RESTAURADO")
            time.sleep(1)
        except Exception as _e_agc_outer:
            print(f"  [AGC] Erro na consulta AGC (nao impacta fluxo): {_e_agc_outer}")
        # ── FIM CONSULTA AGUARDANDO CONFIRMACAO CANCELAMENTO ──

        if not baixar_tomado:
            print("  [INFO] Tomado dispensado (TOMADOS=NAO na planilha).")
            return status_prestado, "DISPENSADO"

        print("  [PDF] Alternando para Notas Tomadas...")
        selecionar_notas_tomadas(driver)
        time.sleep(1)
        ajustar_checkboxes(driver, "PDF TOMADO")
        time.sleep(1)
        preencher_periodo(driver)
        time.sleep(1)

        status_tomado = _consultar_e_imprimir(
            driver, pasta, nome_pdf_tomado(codigo), "Tomado",
            codigo=codigo, pasta_print=pasta_int)

        if status_tomado == "PDF":
            os.makedirs(pasta_int, exist_ok=True)
            _src = os.path.join(pasta, nome_pdf_tomado(codigo))
            _dst = os.path.join(pasta_int, nome_pdf_tomado(codigo))
            if os.path.exists(_src):
                shutil.move(_src, _dst)

        return status_prestado, status_tomado

    except PortalInstabilidadeException:
        raise
    except Exception as e:
        _e_str = str(e) + traceback.format_exc()
        _ASSINATURAS_RENDERER = [
            "receiving message from renderer",
            "GetHandleVerifier",
            "chrome not reachable",
            "unable to connect to renderer",
            "session deleted",
            "no such window",
            "no such session",
            "disconnected: not connected to DevTools",
        ]
        if any(sig in _e_str for sig in _ASSINATURAS_RENDERER):
            raise PortalInstabilidadeException(
                f"Renderer morto durante PDF Prestado+Tomado: {type(e).__name__}"
            ) from e
        print(f"    [ERRO] Falha no fluxo PDF Prestado+Tomado: {e}")
        tirar_print(driver, os.path.join(pasta, "erro_pdf.png"))
        return "Erro", "Erro"


# ─────────────────────────────────────────────────────────────────────────────
# DOWNLOAD TALAO FISCAL
# Regras:
#   - Executa somente se regime != SIMPLES
#   - Executa somente se houver movimento em Prestado OU Tomado
#   - Consulta 1 (Guia Recolhimento) -> ISS_[cod]_[mmaaaa].pdf
#   - Consulta 2 (Nota Fiscal)       -> talao_fiscal_[cod]_[mmaaaa].pdf
#   - Sem resultado em qualquer consulta -> print de evidencia
# ─────────────────────────────────────────────────────────────────────────────

def navegar_talao_fiscal(driver):
    print("  [TALAO] Preparando acesso ao Talao Fiscal...")

    # Passo 1 — volta pro inicio e recria contexto JSF
    driver.get("https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/login/bemVindo.jsf")
    time.sleep(3)

    # Passo 2 — Contribuinte + Nota Fiscal (mesmo helper que funciona no XML)
    abrir_menu_contribuinte_nota_fiscal(driver, "Talao")

    # Passo 3 — Consulta de Taloes Fiscais
    talao = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH,
             "//a[contains(normalize-space(.),'Consulta de Tal')] | "
             "//li/a[contains(normalize-space(.),'Consulta de Tal')]"))
    )
    driver.execute_script("arguments[0].click();", talao)
    time.sleep(3)

    # Passo 4 — valida chegada
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "form1:botaoConsultar"))
    )
    print("    [OK] Tela do Talao Fiscal carregada corretamente.")
def preencher_periodo_talao(driver):
    """
    Preenche Periodo Referencia do Talao Fiscal.
    IDs confirmados: form1:j_id201InputDate (inicio) | form1:j_id203InputDate (fim)
    """
    print(f"  [TALAO] Preenchendo periodo: {DIA_INICIO} a {DIA_FIM}")
    try:
        aguardar_elemento_essencial(driver, EC.presence_of_element_located(
            (By.ID, "form1:j_id201InputDate")), "campo data inicio Talao")
        aguardar_elemento_essencial(driver, EC.presence_of_element_located(
            (By.ID, "form1:j_id203InputDate")), "campo data fim Talao")

        ok = driver.execute_script("""
            function setValor(id, valor) {
                var el = document.getElementById(id);
                if (!el) return false;
                el.removeAttribute('readonly');
                el.removeAttribute('disabled');
                el.value = valor;
                el.setAttribute('value', valor);
                el.dispatchEvent(new Event('input',  {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
                el.dispatchEvent(new Event('blur',   {bubbles: true}));
                return true;
            }
            var a = setValor('form1:j_id201InputDate', arguments[0]);
            var b = setValor('form1:j_id203InputDate', arguments[1]);
            setValor('form1:j_id201InputCurrentDate', arguments[2]);
            setValor('form1:j_id203InputCurrentDate', arguments[2]);
            return a && b;
        """, DIA_INICIO, DIA_FIM, competencia.strftime("%m/%Y"))

        time.sleep(0.8)
        if not ok:
            raise Exception("Campos de data do Talao nao encontrados pelos IDs")
        print(f"    [OK] Periodo preenchido: {DIA_INICIO} a {DIA_FIM}")

    except Exception as e:
        print(f"    [ERRO] Falha ao preencher periodo do Talao: {e}")
        raise


def selecionar_documento_talao(driver, tipo):
    """
    Seleciona radio Documento na tela de Taloes Fiscais.
    tipo='NOTA' -> form1:radioTipoDocumento:0
    tipo='GUIA' -> form1:radioTipoDocumento:1
    """
    radio_id    = "form1:radioTipoDocumento:0" if tipo == "NOTA" else "form1:radioTipoDocumento:1"
    texto_alvo  = "Nota Fiscal" if tipo == "NOTA" else "Guia de Recolhimento"
    try:
        radio = aguardar_elemento_essencial(driver, EC.presence_of_element_located(
            (By.ID, radio_id)), f"radio {texto_alvo}")
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", radio)
        time.sleep(0.2)
        if not radio.is_selected():
            driver.execute_script("arguments[0].click();", radio)
            time.sleep(0.8)
        print(f"    [OK] Documento = {texto_alvo} selecionado.")
        return True
    except Exception as e:
        print(f"    [ERRO] Nao foi possivel selecionar {texto_alvo}: {e}")
        return False


def clicar_consultar_talao(driver, label):
    """Clica no botao Consultar da tela de Taloes — ID: form1:botaoConsultar"""
    try:
        btn = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
            (By.ID, "form1:botaoConsultar")),
            f"botao Consultar Talao — {label}")
    except Exception:
        btn = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
            (By.XPATH,
             "//a[contains(normalize-space(.),'Consultar')] | //button[contains(normalize-space(.),'Consultar')]")),
            f"botao Consultar Talao (fallback) — {label}")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.2)
    driver.execute_script("arguments[0].click();", btn)
    print(f"    [OK] Consultar clicado ({label}). Aguardando resultado...")
    if label == "Nota Fiscal" and getattr(driver, "_alto_volume_stm", False):
        time.sleep(1)
    else:
        time.sleep(5)


def tela_sem_resultado(driver):
    """Retorna True se a tela exibe mensagem de 'nenhum resultado'."""
    body = driver.find_element(By.TAG_NAME, "body").text
    return "nenhum" in body.lower() and "resultado" in body.lower()




def _nome_pelo_radio_prestado_tomado(driver, codigo):
    """
    Lê o radio marcado na tela de Prestados/Tomados e retorna o nome correto do PDF.
    IDs: j_id184:j_id190:0 = Prestadas | j_id184:j_id190:1 = Tomadas
    """
    try:
        radio_prestado = driver.find_element(By.ID, "j_id184:j_id190:0")
        if radio_prestado.is_selected():
            return nome_pdf_prestado(codigo), "Prestado"
    except Exception:
        pass
    try:
        radio_tomado = driver.find_element(By.ID, "j_id184:j_id190:1")
        if radio_tomado.is_selected():
            return nome_pdf_tomado(codigo), "Tomado"
    except Exception:
        pass
    return None, None


def _nome_pelo_radio_xml(driver, codigo):
    """
    Lê o radio marcado na tela de XML e retorna o nome correto do arquivo.
    IDs: form1:j_id206:0 = Emitidos/Prestador | form1:j_id206:1 = Recebidas/Tomador
    """
    try:
        radio_emit = driver.find_element(By.ID, "form1:j_id206:0")
        if radio_emit.is_selected():
            return nome_xml_prestado(codigo), "XML Prestado"
    except Exception:
        pass
    try:
        radio_receb = driver.find_element(By.ID, "form1:j_id206:1")
        if radio_receb.is_selected():
            return nome_xml_tomado(codigo), "XML Tomado"
    except Exception:
        pass
    return None, None


def _nome_pelo_radio_talao(driver):
    """
    Lê o radio marcado na tela de Talão e retorna o nome correto do PDF.
    IDs: form1:radioTipoDocumento:0 = Nota Fiscal | form1:radioTipoDocumento:1 = Guia ISS
    """
    try:
        radio_nf = driver.find_element(By.ID, "form1:radioTipoDocumento:0")
        if radio_nf.is_selected():
            return "NOTA"
    except Exception:
        pass
    try:
        radio_guia = driver.find_element(By.ID, "form1:radioTipoDocumento:1")
        if radio_guia.is_selected():
            return "GUIA"
    except Exception:
        pass
    return None

def _imprimir_e_salvar_talao(driver, pasta, nome_final, label, codigo="", pasta_print=None):
    """
    Localiza e clica no botao 'Imprimir Talao', aguarda o PDF e renomeia.
    pasta_print: onde salvar prints de evidência. Default = pasta.
    Retorna True se salvo com sucesso, False caso contrario.
    ID confirmado pelo diagnostico: form1:botaoGuia (onclick _blank)
    """
    _pasta_print = pasta_print if pasta_print else pasta
    # Tenta pelo ID confirmado primeiro
    botao = None
    try:
        botao = driver.find_element(By.ID, "form1:botaoGuia")
    except Exception:
        pass

    if botao is None:
        # Fallback por texto
        botoes = driver.find_elements(By.XPATH,
            "//*[contains(normalize-space(.),'Imprimir Talao')] | "
            "//*[contains(normalize-space(.),'Imprimir Tal\u00e3o')]")
        if botoes:
            botao = botoes[-1]

    if botao is None:
        print(f"    [AVISO] Botao de impressao nao localizado ({label}).")
        return False

    # Valida pelo radio qual documento está selecionado antes de baixar
    tipo_radio = _nome_pelo_radio_talao(driver)
    if tipo_radio == "NOTA" and "talao_fiscal" not in nome_final:
        nome_final = f"talao_fiscal_{codigo}_{PASTA_MES}.pdf"
        print(f"    [VALIDACAO] Radio confirma: Nota Fiscal. Nome corrigido: {nome_final}")
    elif tipo_radio == "GUIA" and "ISS_" not in nome_final:
        nome_final = f"ISS_{codigo}_{PASTA_MES}.pdf"
        print(f"    [VALIDACAO] Radio confirma: Guia ISS. Nome corrigido: {nome_final}")

    def _uma_tentativa_talao():
        """Executa uma tentativa de download do talão. Retorna True se salvo."""
        # Deleta arquivo alvo antes do download
        arquivo_alvo = os.path.join(pasta, nome_final)
        if os.path.exists(arquivo_alvo):
            os.remove(arquivo_alvo)

        arquivos_antes = os.listdir(pasta)
        handles_antes = set(driver.window_handles)
        driver.execute_script("arguments[0].click();", botao)
        time.sleep(3)

        # Detecta TPA na tela principal (sem nova aba)
        try:
            corpo_principal = driver.find_element(By.TAG_NAME, "body").text
            if "NFEO011-003" in corpo_principal or "pagamento antecipado" in corpo_principal.lower():
                print(f"    [TALAO] Bloqueio do portal (tela principal): nota com pagamento antecipado pendente.")
                salvar_print_talao_fiscal(driver, _pasta_print, codigo)
                raise ValueError("TPA")
        except ValueError:
            raise
        except Exception:
            pass

        # Aba nova: diagnóstico + detecta erro TPA
        handles_depois = set(driver.window_handles)
        novas_abas = handles_depois - handles_antes
        if novas_abas:
            try:
                driver.switch_to.window(list(novas_abas)[0])
                time.sleep(2)
                url_aba   = driver.current_url
                corpo_aba = driver.find_element(By.TAG_NAME, "body").text
                print(f"    [DIAG] Aba talao — URL: {url_aba[:100]}")
                # Detecta erro de infraestrutura do portal (alto volume)
                if "paginaErro.jsf" in url_aba and getattr(driver, "_alto_volume_stm", False):
                    print(f"    [TALAO] Portal retornou erro de infraestrutura (paginaErro.jsf). Abortando aguardo.")
                    driver.close()
                    time.sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                    return False
                # Detecta erro TPA
                if "NFEO011-003" in corpo_aba or "pagamento antecipado" in corpo_aba.lower():
                    print(f"    [TALAO] Bloqueio do portal (nova aba): nota com pagamento antecipado pendente.")
                    salvar_print_talao_fiscal(driver, _pasta_print, codigo)
                    driver.close()
                    time.sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                    raise ValueError("TPA")
                driver.close()
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[0])
            except ValueError:
                raise
            except Exception:
                try:
                    driver.switch_to.window(driver.window_handles[0])
                except Exception:
                    pass

        timeout_talao = 15 if getattr(driver, "_alto_volume_stm", False) else 60
        if getattr(driver, "_alto_volume_stm", False):
            print(f"    [TALAO] Alto volume detectado na empresa — timeout reduzido para {timeout_talao}s.")

        if aguardar_download(pasta, ".pdf", timeout=timeout_talao, arquivos_antes=arquivos_antes):
            if renomear_download_novo(pasta, nome_final, ".pdf", arquivos_antes):
                print(f"    [OK] Talao salvo ({label}): {nome_final}")
                return True

            print(f"    [AVISO] Talao {label} apareceu, mas nao foi possivel identificar arquivo novo para renomear.")
            return False
        return False

    # Tentativa unica
    try:
        if _uma_tentativa_talao():
            return True
    except ValueError:
        # Erro TPA — não é instabilidade, é bloqueio do portal
        return False

    # Falhou — instabilidade de portal (volume grande), sem reprocessamento
    print(f"    [INSTABILIDADE] Talao {label} nao gerado apos 2 tentativas.")
    raise PortalInstabilidadeException(f"Talao {label} nao gerado apos 2 tentativas")


def baixar_talao(driver, pasta, codigo, emp=None, executar_guia=True, pasta_print=None):
    """
    Executa as consultas do Talao Fiscal na mesma tela:
      1. Guia ISS     -> ISS_[cod]_[mmaaaa].pdf           (so se executar_guia=True)
      2. Nota Fiscal  -> talao_fiscal_[cod]_[mmaaaa].pdf  (sempre, qualquer regime)

    - Sem resultado em qualquer consulta: tira print de evidencia.
    - pasta_print: onde salvar prints. Default = pasta.
    - Retorna dict com status individual de cada consulta.
    """
    _pasta_print = pasta_print if pasta_print else pasta
    print("  [TALAO] Iniciando Talao Fiscal...")
    if executar_guia:
        print("  [TALAO] Guia ISS habilitada (regime com tributacao propria).")
    else:
        print("  [TALAO] Guia ISS dispensada (Simples Nacional).")

    resultado = {"nf": None, "guia": None}

    try:
        if getattr(driver, "_selecionar_empresa_stm", False):
            print("  [TALAO] Caminho: EXCECAO — SELECIONAR EMPRESA")
            voltar_bemvindo_stm(driver, "Talao excecao")
        navegar_talao_fiscal(driver)
        preencher_periodo_talao(driver)

        # ── CONSULTA 1: GUIA ISS (so regime != SIMPLES) ──────────────────────
        if not executar_guia:
            print("  [TALAO] Guia ISS pulada — regime Simples Nacional.")
            resultado["guia"] = "Recolhimento via PGDAS"
        else:
            print("  [TALAO] Consulta 1 — Guia ISS")
            selecionar_documento_talao(driver, "GUIA")
            time.sleep(1)
            clicar_consultar_talao(driver, "Guia ISS")

            if tela_sem_resultado(driver):
                print("    [INFO] Sem resultado — Guia ISS.")
                tirar_print(driver, os.path.join(_pasta_print, "sem_resultado_iss.png"))
                resultado["guia"] = "SEM GUIA"
            else:
                salvo = _imprimir_e_salvar_talao(
                    driver, pasta, nome_pdf_iss(codigo), "Guia ISS",
                    pasta_print=_pasta_print)
                resultado["guia"] = "PDF" if salvo else "Erro"

        # ── CONSULTA 2: NOTA FISCAL (qualquer regime) ────────────────────────
        print("  [TALAO] Consulta 2 — Nota Fiscal")
        selecionar_documento_talao(driver, "NOTA")
        time.sleep(1)
        clicar_consultar_talao(driver, "Nota Fiscal")

        if tela_sem_resultado(driver):
            print("    [INFO] Sem resultado — Nota Fiscal.")
            salvar_print_talao_fiscal(driver, _pasta_print, codigo)
            resultado["nf"] = "SEM GUIA"
            resultado["obs_talao"] = "Talão NF sem resultado no portal. Evidência salva na pasta da empresa."
        else:
            try:
                salvo = _imprimir_e_salvar_talao(
                    driver, pasta, nome_pdf_talao_nf(codigo), "Nota Fiscal",
                    codigo=codigo, pasta_print=_pasta_print)
                resultado["nf"] = "PDF" if salvo else "Erro"
                if not salvo:
                    resultado["obs_talao"] = "Talão NF bloqueado pelo portal: nota fiscal sujeita a pagamento antecipado. Evidência salva na pasta da empresa."
            except PortalInstabilidadeException as e:
                if "Talao" in str(e) and "nao gerado apos 2 tentativas" in str(e):
                    resultado["nf"] = "✗"
                    resultado["obs_talao"] = "Portal instável em três tentativas — limite de infraestrutura da prefeitura atingido, volume de notas excede a capacidade do servidor municipal. Emitir talão manualmente."
                    print(f"  [TALAO] NF instavel por volume.")
                else:
                    raise

        print(f"  [TALAO] Concluido — Guia ISS: {resultado['guia']} | NF: {resultado['nf']}")
        return resultado

    except PortalInstabilidadeException:
        raise
    except Exception as e:
        _e_str = str(e) + traceback.format_exc()
        _ASSINATURAS_RENDERER = [
            "receiving message from renderer",
            "GetHandleVerifier",
            "chrome not reachable",
            "unable to connect to renderer",
            "session deleted",
            "no such window",
            "no such session",
            "disconnected: not connected to DevTools",
        ]
        if any(sig in _e_str for sig in _ASSINATURAS_RENDERER):
            raise PortalInstabilidadeException(
                f"Renderer morto durante Talao Fiscal: {type(e).__name__}"
            ) from e
        print(f"    [ERRO] Falha no fluxo Talao Fiscal: {e}")
        tirar_print(driver, os.path.join(_pasta_print, "erro_talao.png"))
        _emp = emp if emp else {"nome": codigo, "codigo": codigo}
        registrar_diagnostico(_emp, "TALAO FISCAL", e, driver, pasta,
                               tentativa=1, fase="TALAO_FISCAL")
        return {"nf": "Erro", "guia": "Erro", "obs_talao": "Erro no fluxo do Talão Fiscal. Verificar diagnóstico e print na pasta da empresa."}


# ─────────────────────────────────────────────────────────────────────────────
# DOWNLOAD XML
# ─────────────────────────────────────────────────────────────────────────────

URL_BEMVINDO_STM = "https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/login/bemVindo.jsf"


def voltar_bemvindo_stm(driver, contexto=""):
    """
    Volta para a tela inicial autenticada do STM.
    Usado nas empresas com fluxo especial 'Selecionar Empresa', porque depois do PDF
    a tela atual é relatorioServicosPrestadosTomadosFiltro.jsf e o submenu especial
    não fica disponível corretamente ali.
    """
    label = f" ({contexto})" if contexto else ""
    print(f"  [NAV] Voltando para bemVindo.jsf{label}...")
    driver.get(URL_BEMVINDO_STM)
    aguardar_elemento_essencial(driver,
        EC.presence_of_element_located((By.TAG_NAME, "body")),
        f"bemVindo STM{label}", timeout=20)
    time.sleep(2)


def abrir_menu_contribuinte_nota_fiscal(driver, contexto=""):
    """
    Abre Contribuinte -> Nota Fiscal partindo do bemVindo.jsf.
    Mantém o fluxo normal intacto e serve apenas para rotas especiais.
    """
    label = f" {contexto}" if contexto else ""

    try:
        btn_contribuinte = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
            (By.XPATH,
             "//button[contains(normalize-space(.),'Contribuinte')] | "
             "//a[contains(@class,'dropdown-toggle') and contains(normalize-space(.),'Contribuinte')]")),
            f"botao Contribuinte{label}", timeout=20)
        driver.execute_script("arguments[0].click();", btn_contribuinte)
        print(f"    [NAV] Menu Contribuinte aberto{label}.")
        time.sleep(1.5)
    except Exception:
        btn_contribuinte = aguardar_elemento_essencial(driver, EC.presence_of_element_located(
            (By.XPATH, "//*[normalize-space(text())='Contribuinte']")),
            f"botao Contribuinte fallback{label}", timeout=20)
        driver.execute_script("arguments[0].click();", btn_contribuinte)
        print(f"    [NAV] Menu Contribuinte aberto{label} (fallback).")
        time.sleep(1.5)

    try:
        nota_fiscal = driver.find_element(By.ID, "j_id4:j_id27:0:j_id40:2:j_id53")
        driver.execute_script("arguments[0].click();", nota_fiscal)
        print(f"    [NAV] Menu Nota Fiscal acessado{label}.")
        time.sleep(1.5)
    except Exception:
        nota_fiscal = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
            (By.XPATH,
             "//a[contains(normalize-space(.),'Nota Fiscal')] | "
             "//li/a[contains(normalize-space(.),'Nota Fiscal')]")),
            f"submenu Nota Fiscal fallback{label}", timeout=20)
        driver.execute_script("arguments[0].click();", nota_fiscal)
        print(f"    [NAV] Menu Nota Fiscal acessado{label} (fallback).")
        time.sleep(1.5)

def navegar_exportar_xml(driver):
    print("  [XML] Acessando tela Exportar XML...")
    xml = aguardar_elemento_essencial(driver, EC.presence_of_element_located(
        (By.ID, "j_id4:j_id27:0:j_id40:2:j_id53:4:j_id56")),
        "menu Exportar XML")
    driver.execute_script("arguments[0].click();", xml)
    time.sleep(3)
    print("    [OK] Tela Exportar XML aberta.")


def preencher_periodo_xml(driver):
    mes    = competencia.strftime("%m")
    ano    = competencia.strftime("%Y")
    periodo = f"{mes}/{ano}"
    print(f"  [XML] Competencia: {periodo}")
    try:
        aguardar_elemento_essencial(driver,
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")),
            "campos de periodo XML")
        driver.execute_script("""
            var campos = document.querySelectorAll("input[type='text']");
            campos[2].click(); campos[2].value = arguments[0];
            campos[2].dispatchEvent(new KeyboardEvent('keyup'));
            campos[3].click(); campos[3].value = arguments[0];
            campos[3].dispatchEvent(new KeyboardEvent('keyup'));
        """, periodo)
        print("    [OK] Periodo XML preenchido.")
    except Exception as e:
        print(f"    [ERRO] Falha ao preencher periodo XML: {e}")


def selecionar_xml_emitidos(driver):
    """Radio Emitidos (Prestador) — ID: form1:j_id206:0"""
    try:
        radio = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "form1:j_id206:0")))
        if not radio.is_selected():
            driver.execute_script("arguments[0].click();", radio)
            time.sleep(0.5)
        print("    [OK] XML Emitidos (Prestador) selecionado.")
    except Exception as e:
        print(f"    [AVISO] Radio Emitidos nao confirmado: {e}")


def selecionar_xml_recebidos(driver):
    """Radio Recebidas (Tomador) — ID: form1:j_id206:1"""
    try:
        radio = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "form1:j_id206:1")))
        driver.execute_script("arguments[0].click();", radio)
        time.sleep(0.5)
        print("    [OK] XML Recebidas (Tomador) selecionado.")
    except Exception as e:
        print(f"    [AVISO] Radio Recebidos nao confirmado: {e}")


def extrair_evento_jsf(onclick):
    if not onclick:
        return None, None
    m = re.search(r"jsfcljs\(document\.forms\['([^']+)'\]\s*,\s*'([^']+)'", onclick)
    if not m:
        return None, None
    form_id = m.group(1)
    pares   = m.group(2)
    campo, valor = pares.split(",", 1) if "," in pares else (pares, pares)
    return form_id, (campo.strip(), valor.strip())


def resposta_parece_xml(resp):
    content_type     = (resp.headers.get("Content-Type") or "").lower()
    conteudo_inicio  = resp.content[:300].lower()
    return (
        "xml" in content_type
        or b"<?xml"  in conteudo_inicio
        or b"<nfse"  in conteudo_inicio
        or b"<nfs"   in conteudo_inicio
    )


def _arquivo_parece_xml(caminho: str) -> bool:
    """Verifica se um arquivo salvo no disco é realmente XML (não HTML do portal)."""
    try:
        with open(caminho, "rb") as f:
            inicio = f.read(300).lower()
        return (
            b"<?xml"  in inicio
            or b"<nfse"  in inicio
            or b"<nfs"   in inicio
        )
    except Exception:
        return False


def criar_sessao_requests(driver):
    sess = requests.Session()
    for c in driver.get_cookies():
        try:
            sess.cookies.set(c["name"], c["value"],
                             domain=c.get("domain"), path=c.get("path", "/"))
        except Exception:
            sess.cookies.set(c["name"], c["value"])
    sess.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent;"),
        "Referer":    driver.current_url,
        "Origin":     "https://stm.semfaz.saoluis.ma.gov.br",
    })
    return sess


def coletar_campos_form(driver, form_id="form1"):
    dados = {}
    campos = driver.find_elements(By.XPATH, f"//form[@id='{form_id}']//input")
    for campo in campos:
        try:
            nome = campo.get_attribute("name")
            if not nome:
                continue
            tipo = (campo.get_attribute("type") or "").lower()
            if tipo in ["checkbox", "radio"]:
                if campo.is_selected():
                    dados[nome] = campo.get_attribute("value") or "on"
            else:
                dados[nome] = campo.get_attribute("value") or ""
        except Exception:
            continue
    dados[form_id] = form_id
    return dados


def baixar_xml_individual_requests(driver, pasta, nome_final):
    print("    [XML] Download individual via sessao autenticada...")

    botoes = driver.find_elements(By.XPATH,
        "//*[contains(@onclick,'dtNotaFiscal') and "
        "(contains(@onclick,'j_id254') or contains(@onclick,'XML') or contains(@title,'XML'))]")
    if not botoes:
        botoes = driver.find_elements(By.XPATH,
            "//*[contains(@onclick,'dtNotaFiscal') and contains(@onclick,'form1')]")
    if not botoes:
        print("    [AVISO] Botao de XML individual nao localizado.")
        return False

    botao   = botoes[0]
    onclick = botao.get_attribute("onclick") or ""
    form_id, evento = extrair_evento_jsf(onclick)
    if not form_id or not evento:
        print(f"    [AVISO] Evento JSF nao identificado: {onclick[:160]}")
        return False

    campo_evento, valor_evento = evento
    action = driver.execute_script("""
        var f = document.getElementById(arguments[0]);
        if (!f) return "";
        return f.getAttribute("action") || window.location.href;
    """, form_id)

    url_post = urljoin(driver.current_url, action)
    dados    = coletar_campos_form(driver, form_id)
    dados[campo_evento] = valor_evento
    sess = criar_sessao_requests(driver)

    try:
        resp = sess.post(url_post, data=dados, timeout=60, allow_redirects=True)
    except Exception as e:
        print(f"    [ERRO] Falha no download direto do XML: {e}")
        return False

    if resp.status_code != 200:
        return False

    caminho_xml = os.path.join(pasta, nome_final)
    if resposta_parece_xml(resp):
        if os.path.exists(caminho_xml):
            os.remove(caminho_xml)
        with open(caminho_xml, "wb") as f:
            f.write(resp.content)
        print(f"    [OK] XML individual salvo: {nome_final}")
        return True

    diag = os.path.join(pasta, "diagnostico_xml_requests.html")
    try:
        with open(diag, "wb") as f:
            f.write(resp.content)
        print(f"    [AVISO] Servidor nao retornou XML. Diagnostico salvo: {diag}")
    except Exception:
        pass
    return False


def baixar_xml_lista_requests(driver, pasta, nome_final):
    """
    Baixa o XML/ZIP da lista via POST autenticado, sem depender do download do navegador.
    Usado principalmente na tela especial exportarNotasXmlFiltro.jsf.
    """
    print("    [XML] Tentando download da lista via sessao autenticada...")

    botoes = driver.find_elements(By.XPATH,
        "//*[contains(normalize-space(.),'Baixar Notas da Lista em XML') or "
        "contains(normalize-space(.),'Baixar Notas')]"
    )
    if not botoes:
        print("    [AVISO] Botao Baixar Notas da Lista em XML nao localizado para requests.")
        return False

    botao = botoes[0]
    onclick = botao.get_attribute("onclick") or ""
    form_id, evento = extrair_evento_jsf(onclick)

    # Se o onclick nao veio claro, tenta descobrir o form pai e o name/id do botao.
    if not form_id:
        try:
            form_id = driver.execute_script("""
                var el = arguments[0];
                while (el && el.tagName && el.tagName.toLowerCase() !== 'form') {
                    el = el.parentElement;
                }
                return el ? (el.id || el.name || '') : '';
            """, botao)
        except Exception:
            form_id = "form1"

    if not form_id:
        form_id = "form1"

    action = driver.execute_script("""
        var f = document.getElementById(arguments[0]);
        if (!f) return "";
        return f.getAttribute("action") || window.location.href;
    """, form_id)

    url_post = urljoin(driver.current_url, action or driver.current_url)
    dados = coletar_campos_form(driver, form_id)

    if evento:
        campo_evento, valor_evento = evento
        dados[campo_evento] = valor_evento
    else:
        nome_botao = botao.get_attribute("name") or botao.get_attribute("id") or ""
        valor_botao = botao.get_attribute("value") or nome_botao
        if nome_botao:
            dados[nome_botao] = valor_botao

    sess = criar_sessao_requests(driver)

    try:
        resp = sess.post(url_post, data=dados, timeout=120, allow_redirects=True)
    except Exception as e:
        print(f"    [ERRO] Falha no POST do XML lista: {e}")
        return False

    if resp.status_code != 200:
        print(f"    [AVISO] POST XML lista retornou HTTP {resp.status_code}")
        return False

    content_type = (resp.headers.get("Content-Type") or "").lower()
    conteudo = resp.content or b""

    nome_zip = nome_final.replace(".xml", ".zip")
    caminho_zip = os.path.join(pasta, nome_zip)
    caminho_xml = os.path.join(pasta, nome_final)

    if "zip" in content_type or conteudo[:2] == b"PK":
        if os.path.exists(caminho_zip):
            os.remove(caminho_zip)
        with open(caminho_zip, "wb") as f:
            f.write(conteudo)
        try:
            import io as _io, re as _re
            _buf = _io.BytesIO()
            with zipfile.ZipFile(caminho_zip, "r") as _zin, zipfile.ZipFile(_buf, "w", zipfile.ZIP_DEFLATED) as _zout:
                for _item in _zin.infolist():
                    _dados = _zin.read(_item.filename)
                    _nome_limpo = _re.sub(r'\s+\d{2}:\d{2}:\d{2}', '', _item.filename)
                    _item.filename = _nome_limpo
                    _zout.writestr(_item, _dados)
            with open(caminho_zip, "wb") as _fw:
                _fw.write(_buf.getvalue())
        except Exception as _es:
            print(f"    [AVISO] Sanitizacao ZIP requests falhou (nao impacta): {_es}")
        print(f"    [OK] ZIP XML lista salvo via requests: {nome_zip}")
        return "ZIP"

    if resposta_parece_xml(resp):
        if os.path.exists(caminho_xml):
            os.remove(caminho_xml)
        with open(caminho_xml, "wb") as f:
            f.write(conteudo)
        print(f"    [OK] XML lista salvo via requests: {nome_final}")
        return "XML"

    diag = os.path.join(pasta, "diagnostico_xml_lista_requests.html")
    try:
        with open(diag, "wb") as f:
            f.write(conteudo)
        print(f"    [AVISO] Servidor nao retornou XML/ZIP da lista. Diagnostico salvo: {diag}")
    except Exception:
        pass
    return False


def confirmar_manter_edge(tentativas=6):
    if pyautogui is None:
        print("    [AVISO] pyautogui nao instalado. Confirmacao manual nao disponivel.")
        return False
    try:
        time.sleep(1)
        for _ in range(tentativas):
            pyautogui.press("enter")
            time.sleep(0.7)
            pyautogui.press("space")
            time.sleep(0.7)
        return True
    except Exception as e:
        print(f"    [AVISO] Falha na confirmacao manual: {e}")
        return False


def _executar_download_xml(driver, pasta, nome_final):
    """
    Verifica sem movimento, baixa lista ZIP ou XML individual.
    2 tentativas por arquivo. Se falhar lança InstabilidadeDownload.
    Retorna: 'ZIP' | 'XML' | 'Sem movimento' | 'Erro'
    """
    body = driver.find_element(By.TAG_NAME, "body").text
    if "nenhum" in body.lower() and "resultado" in body.lower():
        print("    [INFO] Sem movimento no periodo para XML.")
        return "Sem movimento"

    botoes_lista = driver.find_elements(By.XPATH,
        "//*[contains(text(),'Baixar Notas da Lista em XML')]")
    if botoes_lista:
        nome_zip = nome_final.replace(".xml", ".zip")

        def _tentar_zip():
            # Deleta arquivo alvo antes do download
            alvo = os.path.join(pasta, nome_zip)
            if os.path.exists(alvo):
                os.remove(alvo)
            arquivos_antes = os.listdir(pasta)
            driver.execute_script("arguments[0].click();", botoes_lista[0])
            print("    [XML] Download da lista iniciado em ZIP.")
            time.sleep(5)
            if aguardar_download(pasta, ".zip", arquivos_antes=arquivos_antes):
                if renomear_download_novo(pasta, nome_zip, ".zip", arquivos_antes):
                    # Valida que o arquivo é realmente um ZIP (começa com PK)
                    _alvo_zip = os.path.join(pasta, nome_zip)
                    try:
                        with open(_alvo_zip, "rb") as _fz:
                            _magic = _fz.read(2)
                    except Exception:
                        _magic = b""
                    if _magic == b"PK":
                        try:
                            import io as _io, re as _re
                            _buf = _io.BytesIO()
                            with zipfile.ZipFile(_alvo_zip, "r") as _zin, zipfile.ZipFile(_buf, "w", zipfile.ZIP_DEFLATED) as _zout:
                                for _item in _zin.infolist():
                                    _dados = _zin.read(_item.filename)
                                    _nome_limpo = _re.sub(r'\s+\d{2}:\d{2}:\d{2}', '', _item.filename)
                                    _item.filename = _nome_limpo
                                    _zout.writestr(_item, _dados)
                            with open(_alvo_zip, "wb") as _fw:
                                _fw.write(_buf.getvalue())
                        except Exception as _es:
                            print(f"    [AVISO] Sanitizacao ZIP browser falhou (nao impacta): {_es}")
                        print(f"    [OK] ZIP salvo: {nome_zip}")
                        return True
                    else:
                        print(f"    [AVISO] Arquivo nao e ZIP valido (portal retornou HTML). Descartando.")
                        try:
                            os.remove(_alvo_zip)
                        except Exception:
                            pass
                        return False

                print(f"    [AVISO] ZIP apareceu, mas nao foi possivel identificar arquivo novo para renomear: {nome_zip}")
                return False
            return False

        # Tentativa 1
        if _tentar_zip():
            return "ZIP"
        # Tentativa 2
        print(f"    [AVISO] ZIP nao chegou. Aguardando 30s e retentando...")
        time.sleep(30)
        if _tentar_zip():
            return "ZIP"

        print(f"    [INSTABILIDADE] ZIP nao gerado apos 2 tentativas.")
        raise PortalInstabilidadeException(f"ZIP XML nao gerado apos 2 tentativas")
    else:
        print("    [INFO] Uma nota localizada. Download individual do XML.")
        if baixar_xml_individual_requests(driver, pasta, nome_final):
            return "XML"

        raise PortalInstabilidadeException("Download XML individual falhou — reprocessar")


def _baixar_xml_excecao_selecionar_empresa(driver, pasta, pasta_int, codigo, pular_prestado=False, pular_tomado=False):
    """
    CAMINHO EXCECAO — XML SELECIONAR EMPRESA
    Exclusivo para empresas que passaram pelo fluxo 'Selecionar Empresa'.

    Baixa os dois XMLs na mesma tela especial:
      - Emitidos (Prestador)  -> XML Prestado
      - Recebidas (Tomador)   -> XML Tomado

    Mantem o navegar_exportar_xml original intacto para as 94 empresas.
    Qualquer falha vira PortalInstabilidadeException para cair no reprocessamento.
    """
    print("  [XML] Caminho: EXCECAO — SELECIONAR EMPRESA")

    def _abrir_tela_xml_excecao():
        # Depois do PDF a tela costuma estar em relatorioServicosPrestadosTomadosFiltro.jsf.
        # O menu especial correto das 6 empresas deve ser acessado a partir do bemVindo.jsf.
        voltar_bemvindo_stm(driver, "XML excecao")
        abrir_menu_contribuinte_nota_fiscal(driver, "XML excecao")

        # Nas empresas especiais o submenu tem 7 itens; XML é o índice 6.
        ID_XML_EXCECAO = "j_id4:j_id27:0:j_id40:2:j_id53:6:j_id56"
        try:
            xml_item = aguardar_elemento_essencial(driver,
                EC.presence_of_element_located((By.ID, ID_XML_EXCECAO)),
                "menu Exportar NFSe para XML excecao", timeout=15)
            driver.execute_script("arguments[0].click();", xml_item)
        except Exception:
            # Fallback por texto, caso o ID varie.
            xml_item = aguardar_elemento_essencial(driver,
                EC.presence_of_element_located(
                    (By.XPATH,
                     "//*[contains(normalize-space(.),'Exportar NFSe para XML') or "
                     "contains(normalize-space(.),'Exportar NFS') or "
                     "contains(normalize-space(.),'XML')]")),
                "menu Exportar NFSe para XML excecao fallback", timeout=15)
            driver.execute_script("arguments[0].click();", xml_item)

        time.sleep(3)
        print("    [OK] Tela Exportar XML (excecao) aberta.")

        # Confirma tela correta — apenas informativo (URL JSF pode variar entre sessoes).
        url_atual = driver.current_url.lower()
        if "exportarnotas" not in url_atual and "xml" not in url_atual:
            body_txt = ""
            try:
                body_txt = driver.find_element(By.TAG_NAME, "body").text.lower()
            except Exception:
                pass
            if "período ref" not in body_txt and "periodo ref" not in body_txt and len(body_txt) < 50:
                raise PortalInstabilidadeException(
                    f"XML excecao: pagina em branco ou inesperada apos clicar Exportar XML. URL: {driver.current_url}")
            print(f"    [AVISO] URL nao reconhecida, mas pagina tem conteudo — prosseguindo. URL: {driver.current_url}")

    def _preencher_periodo_xml_excecao():
        # Preencher periodo MM/YYYY nos campos Período Ref. Inicial e Final.
        # Nesta tela existem 4 inputs de texto visíveis: CPF/CNPJ, Razão Social,
        # Período Ref. Inicial e Período Ref. Final. Os 2 primeiros são travados.
        periodo = competencia.strftime("%m/%Y")

        campos_periodo = driver.execute_script("""
            var inputs = Array.from(document.querySelectorAll(
                "input[type='text'], input:not([type])"
            ));

            function visivel(el) {
                var r = el.getBoundingClientRect();
                var st = window.getComputedStyle(el);
                return r.width > 0 && r.height > 0 &&
                       st.display !== 'none' && st.visibility !== 'hidden';
            }

            var editaveis = inputs.filter(function(el) {
                if (!visivel(el)) return false;
                if (el.disabled || el.readOnly) return false;

                var id = (el.id || '').toLowerCase();
                var name = (el.name || '').toLowerCase();
                var value = (el.value || '').trim();
                var ph = (el.getAttribute('placeholder') || '').trim();

                return id.indexOf('period') >= 0 ||
                       name.indexOf('period') >= 0 ||
                       value === '' || value.indexOf('_') >= 0 ||
                       ph.indexOf('_') >= 0;
            });

            return editaveis.slice(-2);
        """)

        if not campos_periodo or len(campos_periodo) < 2:
            raise PortalInstabilidadeException("Campos de periodo XML excecao nao localizados")

        for campo in campos_periodo[:2]:
            driver.execute_script("""
                var el = arguments[0];
                var valor = arguments[1];
                el.scrollIntoView({block:'center'});
                el.focus();
                el.value = valor;
                el.setAttribute('value', valor);
                el.dispatchEvent(new Event('input',  {bubbles:true}));
                el.dispatchEvent(new Event('keyup',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.dispatchEvent(new Event('blur',   {bubbles:true}));
            """, campo, periodo)
            time.sleep(0.2)

        print(f"    [OK] Periodo XML excecao preenchido: {periodo}")

    def _selecionar_radio_xml_excecao(tipo):
        """tipo='prestado' => Emitidos | tipo='tomado' => Recebidas"""
        alvo = "Emitidos" if tipo == "prestado" else "Recebidas"
        indice = 0 if tipo == "prestado" else 1

        radio = driver.execute_script("""
            var alvo = arguments[0].toLowerCase();
            var indice = arguments[1];
            var radios = Array.from(document.querySelectorAll("input[type='radio']"));

            function txt(el) {
                if (!el) return '';
                return (el.innerText || el.textContent || '').toLowerCase();
            }

            for (var i = 0; i < radios.length; i++) {
                var r = radios[i];
                var id = r.id || '';
                var label = id ? document.querySelector("label[for='" + id.replace(/'/g, "\\'") + "']") : null;
                var parent = r.parentElement;
                var prox = r.nextElementSibling;
                var bloco = txt(label) + ' ' + txt(parent) + ' ' + txt(prox);
                if (bloco.indexOf(alvo) >= 0) {
                    return r;
                }
            }

            return radios.length > indice ? radios[indice] : null;
        """, alvo, indice)

        if radio is None:
            raise PortalInstabilidadeException(f"Radio XML excecao nao localizado: {alvo}")

        driver.execute_script("""
            var el = arguments[0];
            el.scrollIntoView({block:'center'});
            if (!el.checked) {
                el.click();
                el.checked = true;
                el.dispatchEvent(new Event('click',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
            }
        """, radio)
        time.sleep(0.8)
        print(f"    [OK] Radio {alvo} confirmado.")

    def _consultar_xml_excecao(label):
        btn = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
            (By.XPATH,
             "//*[@id='form1:botaoConsultar'] | "
             "//*[contains(@value,'Consultar') or normalize-space(.)='Consultar']")),
            f"botao Consultar XML excecao {label}", timeout=20)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(5)
        print(f"    [OK] Consulta XML excecao enviada ({label}).")

    def _baixar_resultado_xml_excecao(nome_final_xml, label, tipo):
        print(f"    [XML] Baixando {label} via _executar_download_xml...")
        _resultado = _executar_download_xml(driver, pasta, nome_final_xml)
        if tipo == "tomado" and _resultado in ("ZIP", "XML"):
            _nome_zip = nome_final_xml.replace(".xml", ".zip")
            for _arq in [_nome_zip, nome_final_xml]:
                _src = os.path.join(pasta, _arq)
                if os.path.exists(_src):
                    shutil.move(_src, os.path.join(pasta_int, _arq))
                    break
        return _resultado

    def _processar_tipo_xml_excecao(tipo, nome_final_xml, pular):
        label = "Prestado" if tipo == "prestado" else "Tomado"
        if pular:
            print(f"  [INFO] XML excecao {label} dispensado (sem movimento no PDF).")
            return "Sem movimento"

        # A tela mantém os campos após a consulta. Repreenche por segurança,
        # alterna o rádio, consulta e baixa o resultado do tipo selecionado.
        _preencher_periodo_xml_excecao()
        _selecionar_radio_xml_excecao(tipo)
        _consultar_xml_excecao(label)
        return _baixar_resultado_xml_excecao(nome_final_xml, label, tipo)

    try:
        _abrir_tela_xml_excecao()

        status_prestado = _processar_tipo_xml_excecao(
            "prestado", nome_xml_prestado(codigo), pular_prestado)

        status_tomado = _processar_tipo_xml_excecao(
            "tomado", nome_xml_tomado(codigo), pular_tomado)

        return status_prestado, status_tomado

    except PortalInstabilidadeException:
        raise
    except Exception as e:
        raise PortalInstabilidadeException(
            f"Falha no XML excecao Selecionar Empresa: {type(e).__name__}: {e}"
        ) from e

def baixar_xmls_prestado_e_tomado(driver, pasta, pasta_int, codigo,
                                   pular_prestado=False, pular_tomado=False):
    """
    Navega para Exportar XML uma vez, baixa Prestado e Tomado na mesma tela.
    XML Prestado fica em pasta (raiz). XML Tomado move para pasta_int.
    Retorna: (status_prestado, status_tomado)
    """
    print("  [XML] Iniciando download XML Prestado + Tomado...")

    # Empresas com excecao SELECIONAR EMPRESA: XML via caminho diferente
    if getattr(driver, "_selecionar_empresa_stm", False):
        return _baixar_xml_excecao_selecionar_empresa(
            driver, pasta, pasta_int, codigo,
            pular_prestado=pular_prestado,
            pular_tomado=pular_tomado)

    print("  [XML] Caminho: PADRAO")
    try:
        navegar_exportar_xml(driver)
        preencher_periodo_xml(driver)

        # ── XML Prestado — fica na raiz (para importação do sistema) ─────────
        if pular_prestado:
            status_prestado = "Sem movimento"
            print("  [INFO] XML Prestado dispensado (sem movimento no PDF).")
        else:
            selecionar_xml_emitidos(driver)
            time.sleep(1)
            btn = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
                (By.ID, "form1:botaoConsultar")), "botao Consultar XML Prestado")
            driver.execute_script("arguments[0].click();", btn)
            print("    [OK] Consulta XML Prestado enviada. Aguardando resultado...")
            time.sleep(5)
            status_prestado = _executar_download_xml(
                driver, pasta, nome_xml_prestado(codigo))
            if status_prestado == "Sem movimento":
                status_prestado = "Nao disponibilizado"
                print("  [INFO] XML Prestado: portal nao disponibiliza — PDF tem dados.")

        # ── XML Tomado — move para pasta_int ─────────────────────────────────
        if pular_tomado:
            status_tomado = "Sem movimento"
            print("  [INFO] XML Tomado dispensado (sem movimento no PDF).")
        else:
            print("  [XML] Alternando para Recebidas (Tomador)...")
            selecionar_xml_recebidos(driver)
            time.sleep(1)
            btn = aguardar_elemento_essencial(driver, EC.element_to_be_clickable(
                (By.ID, "form1:botaoConsultar")), "botao Consultar XML Tomado")
            driver.execute_script("arguments[0].click();", btn)
            print("    [OK] Consulta XML Tomado enviada. Aguardando resultado...")
            time.sleep(5)
            status_tomado = _executar_download_xml(
                driver, pasta, nome_xml_tomado(codigo))
            if status_tomado == "Sem movimento":
                status_tomado = "Nao disponibilizado"
                print("  [INFO] XML Tomado: portal nao disponibiliza — PDF tem dados.")
            elif status_tomado in ("ZIP", "XML"):
                _nome_tom = nome_xml_tomado(codigo)
                _nome_zip_tom = _nome_tom.replace(".xml", ".zip")
                for _arq_tom in [_nome_zip_tom, _nome_tom]:
                    _src_tom = os.path.join(pasta, _arq_tom)
                    if os.path.exists(_src_tom):
                        shutil.move(_src_tom, os.path.join(pasta_int, _arq_tom))
                        break

        return status_prestado, status_tomado

    except PortalInstabilidadeException:
        raise
    except Exception as e:
        _e_str = str(e) + traceback.format_exc()
        _ASSINATURAS_RENDERER = [
            "receiving message from renderer",
            "GetHandleVerifier",
            "chrome not reachable",
            "unable to connect to renderer",
            "session deleted",
            "no such window",
            "no such session",
            "disconnected: not connected to DevTools",
        ]
        if any(sig in _e_str for sig in _ASSINATURAS_RENDERER):
            raise PortalInstabilidadeException(
                f"Renderer morto durante XML Prestado+Tomado: {type(e).__name__}"
            ) from e
        raise PortalInstabilidadeException(f"Falha no fluxo XML Prestado+Tomado: {type(e).__name__}")


# ─────────────────────────────────────────────────────────────────────────────
# RELATORIO GERENCIAL
# ─────────────────────────────────────────────────────────────────────────────

def _icone_status(valor):
    if valor is None:
        return "N/A"
    ICONES = {
        "PDF":            "✓",
        "ZIP":            "✓",
        "XML":            "✓",
        "Sem movimento":  "⚠ Sem movimento",
        "Sem resultado":  "Sem talão no portal",
        "Erro":           "✗",
        "DISPENSADO":     "DISPENSADO",
        "Recolhimento via PGDAS": "Recolhimento via PGDAS",
        "Nao disponibilizado": "Nao disponibilizado",
        "SEM GUIA":       "Sem talão no portal",
        "Portal Nacional":"Portal Nacional",
    }
    return ICONES.get(str(valor), str(valor))


def gerar_detalhamento_stm(pasta, codigo, nome_empresa, regime):
    """
    Gera o detalhamento de NFS-e do STM São Luís por empresa.
    Lê o ZIP Prestado em memória, extrai dados nota a nota via extrair_dados_xml_bytes,
    e chama gerar_detalhamento_nfse com o mesmo layout do Padrão Nacional.
    """
    _zip_prest = os.path.join(pasta, nome_xml_prestado(codigo).replace(".xml", ".zip"))
    _xmls_emitidas = []
    if os.path.exists(_zip_prest):
        try:
            with zipfile.ZipFile(_zip_prest, "r") as _z:
                for _nome in sorted(_z.namelist()):
                    if not _nome.lower().endswith(".xml"):
                        continue
                    try:
                        _conteudo = _z.read(_nome)
                        _dados = extrair_dados_xml_bytes(_conteudo, status="Regular")
                        if _dados:
                            _xmls_emitidas.append(_dados)
                    except Exception:
                        continue
        except Exception as _e:
            print(f"  [DETALH] Erro ao ler ZIP Prestado para detalhamento: {_e}")

    if not _xmls_emitidas:
        print(f"  [DETALH] Sem XMLs para gerar detalhamento: {nome_empresa}")
        return

    try:
        gerar_detalhamento_nfse(
            pasta_download=pasta,
            nome_empresa=nome_empresa,
            competencia=PASTA_MES,
            xmls_emitidas=_xmls_emitidas,
            xmls_recebidas=[],
        )
        print(f"  [DETALH] Detalhamento gerado: {nome_empresa}")
    except Exception as _e:
        print(f"  [DETALH] Erro ao gerar detalhamento: {_e}")


def extrair_retencoes_zip(caminho_zip):
    """
    Lê todos os XMLs do ZIP do Prestado em memória e soma os campos de retenção federal.
    Retorna dict com totais. Campos sem valor ficam como None (vazio no gerencial).
    Alíquotas PIS e COFINS: exibe a predominante (moda). Se houver divergência,
    sinaliza via ret_aliq_pis_divergente / ret_aliq_cofins_divergente = True.
    """
    import xml.etree.ElementTree as _ET
    NS = "http://www.sped.fazenda.gov.br/nfse"
    def _n(tag): return f"{{{NS}}}{tag}"
    def _f(el): return float(el.text.strip()) if el is not None and el.text and el.text.strip() else 0.0

    totais = {
        "ret_base_pis":     0.0,
        "ret_valor_pis":    0.0,
        "ret_base_cofins":  0.0,
        "ret_valor_cofins": 0.0,
        "ret_irrf":         0.0,
        "ret_csll":         0.0,
    }
    _aliq_pis_lista    = []
    _aliq_cofins_lista = []
    tem_dados = False
    try:
        with zipfile.ZipFile(caminho_zip, "r") as _z:
            for _nome in _z.namelist():
                if not _nome.lower().endswith(".xml"):
                    continue
                try:
                    _conteudo = _z.read(_nome)
                    _texto = _conteudo.decode("windows-1252", errors="ignore")
                    _raiz = _ET.fromstring(_texto.encode("utf-8"))
                    _piscofins = _raiz.find(f".//{_n('piscofins')}")
                    if _piscofins is not None:
                        totais["ret_base_pis"]     += _f(_piscofins.find(_n("vBCPisCofins")))
                        totais["ret_valor_pis"]    += _f(_piscofins.find(_n("vPis")))
                        totais["ret_base_cofins"]  += _f(_piscofins.find(_n("vBCPisCofins")))
                        totais["ret_valor_cofins"] += _f(_piscofins.find(_n("vCofins")))
                        _ap = _f(_piscofins.find(_n("pAliqPis")))
                        _ac = _f(_piscofins.find(_n("pAliqCofins")))
                        if _ap: _aliq_pis_lista.append(round(_ap, 4))
                        if _ac: _aliq_cofins_lista.append(round(_ac, 4))
                        tem_dados = True
                    _tribfed = _raiz.find(f".//{_n('tribFed')}")
                    if _tribfed is not None:
                        totais["ret_irrf"] += _f(_tribfed.find(_n("vRetIRRF")))
                        totais["ret_csll"] += _f(_tribfed.find(_n("vRetCSLL")))
                        tem_dados = True
                except Exception:
                    continue
    except Exception:
        pass
    if not tem_dados:
        return {}

    resultado = {k: round(v, 2) if v != 0.0 else None for k, v in totais.items()}

    if _aliq_pis_lista:
        _moda_pis = max(set(_aliq_pis_lista), key=_aliq_pis_lista.count)
        resultado["ret_aliq_pis"] = _moda_pis
        resultado["ret_aliq_pis_divergente"] = len(set(_aliq_pis_lista)) > 1
    else:
        resultado["ret_aliq_pis"] = None
        resultado["ret_aliq_pis_divergente"] = False

    if _aliq_cofins_lista:
        _moda_cofins = max(set(_aliq_cofins_lista), key=_aliq_cofins_lista.count)
        resultado["ret_aliq_cofins"] = _moda_cofins
        resultado["ret_aliq_cofins_divergente"] = len(set(_aliq_cofins_lista)) > 1
    else:
        resultado["ret_aliq_cofins"] = None
        resultado["ret_aliq_cofins_divergente"] = False

    return resultado


def gerar_relatorio(resultados):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gerencial"

    azul       = PatternFill("solid", fgColor="1F4E79")
    azul_claro = PatternFill("solid", fgColor="D9E1F2")
    branco     = PatternFill("solid", fgColor="FFFFFF")
    verde      = PatternFill("solid", fgColor="E2F0D9")
    amarelo    = PatternFill("solid", fgColor="FFF2CC")
    vermelho   = PatternFill("solid", fgColor="FCE4D6")
    cinza      = PatternFill("solid", fgColor="EBEBEB")
    laranja    = PatternFill("solid", fgColor="F4B942")

    fonte_cab   = Font(bold=True, color="FFFFFF", size=11)
    fonte_dados = Font(size=10)
    centro   = Alignment(horizontal="center", vertical="center")
    esquerda = Alignment(horizontal="left",   vertical="center")
    borda    = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )

    cabecalhos = [
        "Código", "Empresa", "Regime",
        "Qtd Prestado", "Valor Total", "ISS Próprio", "ISS Retido",
        "PDF Prestado", "XML Prestado",
        "BASE PIS", "ALÍQ. PIS", "VALOR PIS",
        "BASE COFINS", "ALÍQ. COFINS", "VALOR COFINS",
        "IRRF", "CSLL",
        "Qtd Tomado", "Valor Total Tomado", "ISS Próprio Tomado", "ISS Retido Tomado",
        "PDF Tomado", "XML Tomado",
        "Talão NF", "Guia ISS",
        "LOGIN / SENHA", "OBSERVAÇÃO",
    ]
    larguras = [12, 40, 18, 14, 18, 18, 18, 16, 16,
                16, 14, 16, 16, 14, 16, 16, 16,
                14, 18, 18, 18, 16, 16, 16, 16, 20, 60]

    # ── Linha 1: subtítulo RETENÇÕES FEDERAIS em J1:Q1 ───────────────────────
    ws.merge_cells("J1:Q1")
    cel_sub = ws["J1"]
    cel_sub.value     = "RETENÇÕES FEDERAIS"
    cel_sub.fill      = PatternFill("solid", fgColor="2E75B6")
    cel_sub.font      = Font(bold=True, color="FFFFFF", size=11)
    cel_sub.alignment = centro
    cel_sub.border    = borda
    ws.row_dimensions[1].height = 20

    # ── Linha 2: cabeçalhos ───────────────────────────────────────────────────
    for col, (cab, larg) in enumerate(zip(cabecalhos, larguras), 1):
        cel = ws.cell(row=2, column=col, value=cab)
        cel.fill      = azul
        cel.font      = fonte_cab
        cel.alignment = centro
        cel.border    = borda
        ws.column_dimensions[cel.column_letter].width = larg
    ws.row_dimensions[2].height = 24

    # ── Linhas de dados a partir da linha 3 ──────────────────────────────────
    for i, r in enumerate(resultados, 3):
        base_fill = azul_claro if i % 2 == 0 else branco

        senha_incorreta = str(r.get("login_senha", "")).strip() == "INCORRETO"
        sem_prest = r.get("pdf_prestado") in ("Sem movimento", "Erro")
        sem_tom   = r.get("pdf_tomado")   in ("Sem movimento", "Erro")

        if senha_incorreta:
            qtd_prestado       = "Senha Incorreta"
            valor_total        = "Senha Incorreta"
            iss_proprio        = "Senha Incorreta"
            iss_retido         = "Senha Incorreta"
            qtd_tomado         = "Senha Incorreta"
            valor_total_tomado = "Senha Incorreta"
            iss_proprio_tomado = "Senha Incorreta"
            iss_retido_tomado  = "Senha Incorreta"
        else:
            qtd_prestado       = r.get("qtd_notas",          "") if not sem_prest else "⚠ Sem movimento"
            valor_total        = r.get("valor_total",         "") if not sem_prest else "⚠ Sem movimento"
            iss_proprio        = r.get("iss_proprio",         "") if not sem_prest else "⚠ Sem movimento"
            iss_retido         = r.get("iss_retido",          "") if not sem_prest else "⚠ Sem movimento"
            qtd_tomado         = r.get("qtd_tomado",          "") if not sem_tom else ""
            valor_total_tomado = r.get("valor_total_tomado",  "") if not sem_tom else ""
            iss_proprio_tomado = r.get("iss_proprio_tomado",  "") if not sem_tom else ""
            iss_retido_tomado  = r.get("iss_retido_tomado",   "") if not sem_tom else ""

        talao_nf   = r.get("talao_nf")
        talao_guia = r.get("talao_guia")

        if senha_incorreta:
            valores = [
                r.get("codigo",  ""),
                r.get("empresa", ""),
                r.get("regime",  ""),
                "Senha Incorreta", "Senha Incorreta", "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta", "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta", "Senha Incorreta",
                "Senha Incorreta",
                r.get("observacao", ""),
            ]
        else:
            _obs_raw   = r.get("observacao", "") or ""
            _pdf_prest = r.get("pdf_prestado", "")
            if _pdf_prest == "Sem movimento":
                obs_valor = "⚠ Sem movimento"
            elif _pdf_prest == "PDF" and not _obs_raw:
                obs_valor = "✓"
            else:
                obs_valor = _obs_raw

            _obs_div = []
            if r.get("ret_aliq_pis_div"):
                _obs_div.append("Alíquota PIS divergente — conferir detalhamento.")
            if r.get("ret_aliq_cofins_div"):
                _obs_div.append("Alíquota COFINS divergente — conferir detalhamento.")
            if _obs_div:
                _obs_div_txt = " | ".join(_obs_div)
                obs_valor = (_obs_div_txt + " | " + obs_valor) if obs_valor and obs_valor != "✓" else _obs_div_txt

            valores = [
                r.get("codigo",      ""),
                r.get("empresa",     ""),
                r.get("regime",      ""),
                qtd_prestado,
                valor_total,
                iss_proprio,
                iss_retido,
                _icone_status(r.get("pdf_prestado")),
                "✓" if r.get("pdf_prestado") == "PDF" else "⚠ Sem movimento",
                r.get("ret_base_pis"),
                r.get("ret_aliq_pis"),
                r.get("ret_valor_pis"),
                r.get("ret_base_cofins"),
                r.get("ret_aliq_cofins"),
                r.get("ret_valor_cofins"),
                r.get("ret_irrf"),
                r.get("ret_csll"),
                qtd_tomado,
                valor_total_tomado,
                iss_proprio_tomado,
                iss_retido_tomado,
                _icone_status(r.get("pdf_tomado")),
                _icone_status(r.get("xml_tomado")),
                _icone_status(talao_nf),
                _icone_status(talao_guia),
                r.get("login_senha", ""),
                obs_valor,
            ]

        for col, val in enumerate(valores, 1):
            cel = ws.cell(row=i, column=col, value=val)
            cel.fill      = base_fill
            cel.font      = fonte_dados
            cel.border    = borda
            cel.alignment = esquerda if col in [2, 27] else centro

            # Valores monetários prestado/tomado
            if col in [5, 6, 7, 19, 20, 21] and isinstance(val, (int, float)):
                cel.number_format = 'R$ #,##0.00'

            # Retenções federais J-Q (cols 10-17)
            if col in range(10, 18) and isinstance(val, (int, float)):
                if col in [11, 14]:
                    cel.number_format = '#,##0.00'
                else:
                    cel.number_format = 'R$ #,##0.00'

            # Alíquota PIS divergente — col 11 (K) em amarelo
            if col == 11 and r.get("ret_aliq_pis_div"):
                cel.fill = amarelo
                cel.font = Font(size=10, bold=True, color="7F6000")

            # Alíquota COFINS divergente — col 14 (N) em amarelo
            if col == 14 and r.get("ret_aliq_cofins_div"):
                cel.fill = amarelo
                cel.font = Font(size=10, bold=True, color="7F6000")

            # Colunas de status (PDF/XML/Talão) — cols 8,9,22,23,24,25
            if col in [8, 9, 22, 23, 24, 25]:
                texto = str(val) if val is not None else ""
                if texto == "✗" or "Erro" in texto:
                    cel.fill = vermelho
                    cel.font = Font(size=12, bold=True, color="C00000")
                elif "Sem movimento" in texto or "movimento" in texto:
                    cel.fill = amarelo
                    cel.font = Font(size=10, bold=True, color="7F6000")
                elif "Sem resultado" in texto or "resultado" in texto:
                    cel.fill = amarelo
                    cel.font = Font(size=10, bold=True, color="7F6000")
                elif texto in ("N/A", "Não disponibilizado"):
                    cel.fill = cinza
                    cel.font = Font(size=10, color="888888")
                elif texto == "Portal Nacional":
                    cel.fill = azul_claro
                    cel.font = Font(size=10, bold=True, color="1F4E79")
                elif "✓" in texto:
                    cel.fill = verde
                    cel.font = Font(size=12, bold=True, color="00B050")

            # Colunas D-G sem movimento
            if col in [4, 5, 6, 7] and str(val) == "⚠ Sem movimento":
                cel.fill = amarelo
                cel.font = Font(size=10, bold=True, color="7F6000")

            # Senha Incorreta cols 4-26
            if col in range(4, 27) and str(val) == "Senha Incorreta":
                cel.fill = vermelho
                cel.font = Font(size=10, bold=False, color="FF0000")

            # Login/Senha col 26
            if col == 26 and str(val) != "Senha Incorreta":
                if str(val) == "INCORRETO":
                    cel.value = "Senha Incorreta"
                    cel.fill = vermelho
                    cel.font = Font(size=12, bold=False, color="FF0000")
                else:
                    cel.fill = verde
                    cel.font = Font(size=12, bold=True, color="00B050")

            # Observação col 27
            if col == 27:
                texto = str(val) if val is not None else ""
                if texto == "✓":
                    cel.fill = verde
                    cel.font = Font(size=12, bold=True, color="00B050")
                    cel.alignment = centro
                elif "Sem movimento" in texto:
                    cel.fill = amarelo
                    cel.font = Font(size=10, bold=True, color="7F6000")
                elif "Portal Nacional" in texto:
                    cel.fill = azul_claro
                    cel.font = Font(size=10, color="1F4E79")
                elif texto and texto not in ("None", ""):
                    cel.fill = vermelho
                    cel.font = Font(size=10, bold=True, color="C00000")

        ws.row_dimensions[i].height = 20

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{ws.cell(row=2, column=len(cabecalhos)).column_letter}2"

    nome_arquivo = f"Relatorio_Gerencial_Sao_Luis_{PASTA_MES}.xlsx"
    caminho = os.path.join(PASTA_BASE, nome_arquivo)
    if os.path.exists(caminho):
        try:
            os.remove(caminho)
        except Exception:
            pass  # arquivo em uso — openpyxl sobrescreve direto

    wb.save(caminho)
    print(f"\n[RELATORIO] Planilha salva em:\n  {caminho}")


# ─────────────────────────────────────────────────────────────────────────────
# RELATORIO DE CONCLUSAO
# ─────────────────────────────────────────────────────────────────────────────

def gerar_relatorio_conclusao(resultados, tempo_total_s=0, tempos_por_empresa=None, total_original=None):
    tempos_por_empresa = tempos_por_empresa or {}
    agora   = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total   = len(resultados)
    if total_original is None:
        total_original = total

    # Nome do arquivo — não sobrescreve, sufixo timestamp se já existir
    caminho = os.path.join(PASTA_BASE, f"CONCLUSAO_SAO_LUIS_{PASTA_MES}.txt")

    def _fmt_tempo(s):
        s = int(s or 0)
        return f"{s // 60}min{s % 60:02d}s"

    def _status(r, campo):
        val = r.get(campo)
        if val is None:
            return "N/A"
        v = str(val).strip()
        if v in ["PDF", "ZIP", "XML", "✓"]:
            return "✓"
        if v in ["Sem movimento", "Sem resultado", "DISPENSADO", "SEM GUIA", "Recolhimento via PGDAS", "Nao disponibilizado"]:
            return v
        if v == "Erro":
            return "✗"
        return v

    colunas = [
        ("PDF Prestado", "pdf_prestado"),
        ("XML Prestado", "xml_prestado"),
        ("PDF Tomado",   "pdf_tomado"),
        ("XML Tomado",   "xml_tomado"),
        ("Talao NF",     "talao_nf"),
        ("Guia ISS",     "talao_guia"),
    ]

    # ── Contagens ────────────────────────────────────────────────────────────
    total_mins = int(tempo_total_s) // 60
    total_segs = int(tempo_total_s) % 60
    media_s    = int(tempo_total_s / total) if total else 0

    login_incorreto_lista = [r for r in resultados
                             if str(r.get("login_senha","")).strip() == "INCORRETO"]

    portal_instavel_lista = [r for r in resultados
                             if str(r.get("login_senha","")).strip() != "INCORRETO"
                             and ("INSTABILIDADE" in str(r.get("observacao","")).upper()
                                  or "REPROCESS"     in str(r.get("observacao","")).upper())]

    sem_mov_lista         = [r for r in resultados
                             if str(r.get("pdf_prestado","")) == "Sem movimento"
                             and str(r.get("login_senha","")) != "INCORRETO"
                             and r not in portal_instavel_lista]

    talao_manual_lista    = [r for r in resultados
                             if "EMITIR TALAO MANUALMENTE" in str(r.get("observacao","")).upper()]

    guia_manual_lista     = [r for r in resultados
                             if str(r.get("talao_guia","")).strip() == "Erro"
                             and str(r.get("login_senha","")).strip() != "INCORRETO"]

    reprocess_lista       = [r for r in resultados
                             if "REPROCESS" in str(r.get("observacao","")).upper()]

    nova_tentativa_lista  = [r for r in reprocess_lista
                             if any(str(r.get(c,"")).strip() == "Erro" for _, c in colunas)]

    ok_lista              = [r for r in resultados
                             if str(r.get("pdf_prestado","")).strip() == "PDF"
                             and str(r.get("login_senha","")).strip() != "INCORRETO"
                             and r not in portal_instavel_lista]

    sem_pasta_lista       = [r for r in resultados
                             if "PASTA NAO INFORMADA" in str(r.get("observacao","")).upper()]

    com_erros             = total_original - total

    linhas = []

    # ── CABEÇALHO ────────────────────────────────────────────────────────────
    linhas.append("=" * 80)
    linhas.append("  SERVICOS SAO LUIS  |  RELATORIO DE CONCLUSAO")
    linhas.append(f"  Competencia : {PERIODO}")
    linhas.append(f"  Gerado em   : {agora}")
    linhas.append(f"  Tempo total : {total_mins}min{total_segs:02d}s")
    linhas.append("=" * 80)

    # ── RESUMO ───────────────────────────────────────────────────────────────
    _sucesso_total = len(ok_lista) + len(sem_mov_lista)
    linhas.append(f"RESUMO - Media/empresa    : {_fmt_tempo(media_s)}")
    linhas.append("-" * 60)
    linhas.append(f"  Total de empresas                        : {total}")
    linhas.append(f"  Processadas com sucesso com movimento    : {len(ok_lista)}")
    linhas.append(f"  Processadas com sucesso sem movimento    : {len(sem_mov_lista)}")
    linhas.append(f"  Pendente processamento manual            : {len(portal_instavel_lista)}")
    linhas.append(f"  Login incorreto                          : {len(login_incorreto_lista)}")
    linhas.append("-" * 60)
    linhas.append("Observacoes:")
    linhas.append(f"  Emissao de talao com instabilidade       : {len(talao_manual_lista)}")
    linhas.append(f"  Emissao guia ISS com instabilidade       : {len(guia_manual_lista)}")
    linhas.append(f"  Pasta nao informada                      : {len(sem_pasta_lista)}")
    linhas.append(f"  Com erros                                : {com_erros}")
    linhas.append("")

    # ── EMPRESAS OK ──────────────────────────────────────────────────────────
    if ok_lista:
        linhas.append("=" * 80)
        linhas.append(f"EMPRESAS PROCESSADAS COM SUCESSO ({len(ok_lista)})")
        linhas.append("=" * 80)
        for r in ok_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            t    = tempos_por_empresa.get(cod, 0)
            linhas.append("")
            linhas.append(f"  Empresa  : {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")
            linhas.append(f"  Regime   : {r.get('regime','')}")
            for label, campo in colunas:
                linhas.append(f"  {label:<18}: {_status(r, campo)}")
            obs = r.get("observacao","")
            if obs:
                linhas.append(f"  Observacao: {obs}")
            if t:
                linhas.append(f"  Tempo    : {_fmt_tempo(t)}")
            linhas.append("  " + "-" * 60)

    # ── EMPRESAS SEM MOVIMENTO ───────────────────────────────────────────────
    if sem_mov_lista:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append(f"EMPRESAS PROCESSADAS COM SUCESSO SEM MOVIMENTO NA COMPETENCIA ({len(sem_mov_lista)})")
        linhas.append("=" * 80)
        _partes_sm = [f"  {r.get('codigo','')} - {r.get('empresa','')} - CNPJ: {r.get('cnpj','')}" for r in sem_mov_lista]
        _max_sm    = max(len(p) for p in _partes_sm) if _partes_sm else 0
        for r, parte in zip(sem_mov_lista, _partes_sm):
            cod       = r.get("codigo","")
            t         = tempos_por_empresa.get(cod, 0)
            tempo_str = f"  -  {_fmt_tempo(t)}" if t else ""
            linhas.append(f"{parte:<{_max_sm}}  |  Observacao: Evidencia sem movimento salvo na pasta{tempo_str}")

    # ── EMPRESAS COM LOGIN INCORRETO ─────────────────────────────────────────
    if login_incorreto_lista:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append(f"EMPRESAS COM LOGIN INCORRETO ({len(login_incorreto_lista)})")
        linhas.append("=" * 80)
        for r in login_incorreto_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            t    = tempos_por_empresa.get(cod, 0)
            linhas.append("")
            linhas.append(f"  Empresa  : {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")
            linhas.append(f"  Regime   : {r.get('regime','NAO INFORMADO')}")
            linhas.append(f"  Motivo   : Login/senha incorretos")
            if t:
                linhas.append(f"  Tempo    : {_fmt_tempo(t)}")
            linhas.append("  " + "-" * 60)
    else:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append("NENHUM LOGIN INCORRETO REGISTRADO.")
        linhas.append("=" * 80)

    # ── EMPRESAS PORTAL INSTAVEL ─────────────────────────────────────────────
    if portal_instavel_lista:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append(f"EMPRESAS PORTAL INSTAVEL ({len(portal_instavel_lista)})")
        linhas.append("=" * 80)
        for r in portal_instavel_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            t    = tempos_por_empresa.get(cod, 0)
            linhas.append("")
            linhas.append(f"  Empresa  : {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")
            linhas.append(f"  Regime   : {r.get('regime','NAO INFORMADO')}")
            for label, campo in colunas:
                st = _status(r, campo)
                if st == "✗":
                    linhas.append(f"  {label:<18}: BLOQUEIO/ERRO")
                else:
                    linhas.append(f"  {label:<18}: {st}")
            obs = r.get("observacao","")
            if obs:
                linhas.append(f"  Motivo   : {obs}")
            if t:
                linhas.append(f"  Tempo    : {_fmt_tempo(t)}")
            linhas.append("  " + "-" * 60)
    else:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append("NENHUMA INSTABILIDADE DE PORTAL REGISTRADA.")
        linhas.append("=" * 80)

    # ── TALAO MANUAL ─────────────────────────────────────────────────────────
    if talao_manual_lista:
        linhas.append("")
        _titulo_talao = f"EMPRESAS COM NECESSIDADE DE EMISSAO DE TALAO FISCAL MANUALMENTE DEVIDO INSTABILIDADE DO STM ({len(talao_manual_lista)})"
        linhas.append("=" * len(_titulo_talao))
        linhas.append(_titulo_talao)
        linhas.append("=" * len(_titulo_talao))
        for r in talao_manual_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            linhas.append(f"  {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")

    # ── GUIA ISS MANUAL ──────────────────────────────────────────────────────
    if guia_manual_lista:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append(f"GUIA ISS — EMITIR MANUALMENTE ({len(guia_manual_lista)})")
        linhas.append("=" * 80)
        for r in guia_manual_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            linhas.append(f"  {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")

    # ── REPROCESSADAS ─────────────────────────────────────────────────────────
    if reprocess_lista:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append(f"EMPRESAS REPROCESSADAS ({len(reprocess_lista)})")
        linhas.append("=" * 80)
        for r in reprocess_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            tem_erro = any(str(r.get(c,"")).strip() == "Erro" for _, c in colunas)
            situacao = "Reprocessada — bloqueio persistiu, reprocessamento manual necessario" if tem_erro \
                       else "Concluida com sucesso no reprocessamento"
            linhas.append("")
            linhas.append(f"  Empresa   : {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")
            linhas.append(f"  Situacao  : {situacao}")
            linhas.append("  " + "-" * 60)

    # ── SEM PASTA ──────────────────────────────────────────────
    if sem_pasta_lista:
        linhas.append("")
        linhas.append("=" * 80)
        linhas.append(f"EMPRESAS SEM PASTA INFORMADA ({len(sem_pasta_lista)})")
        linhas.append("=" * 80)
        linhas.append("  ATENCAO: Preencher a coluna PASTA na planilha SENHAS.xlsx")
        linhas.append("-" * 60)
        for r in sem_pasta_lista:
            cod  = r.get("codigo","")
            cnpj = r.get("cnpj","")
            linhas.append(f"  {cod} - {r.get('empresa','')} - CNPJ: {cnpj}")
        linhas.append("-" * 60)

    # ── RODAPÉ ────────────────────────────────────────────────────────────────
    linhas.append("")
    linhas.append("=" * 80)
    linhas.append("  FIM DO RELATORIO DE CONCLUSAO  |  STM Sao Luis")
    linhas.append("=" * 80)

    try:
        with open(caminho, "w", encoding="utf-8") as f:
            f.write("\n".join(linhas))
        print(f"\n[CONCLUSAO] Relatorio de conclusao salvo em:\n  {caminho}")
    except Exception as e:
        print(f"  [AVISO] Nao foi possivel salvar relatorio de conclusao: {e}")





# ─────────────────────────────────────────────────────────────────────────────
# DETALHAMENTO POR NOTA — funções extraídas do Padrão Nacional
# ─────────────────────────────────────────────────────────────────────────────

def extrair_dados_xml_bytes(conteudo_xml, numero_fallback="", status=""):
    """
    Extrai todos os campos do detalhamento a partir dos bytes do XML,
    enquanto ainda está na memória — antes de salvar e antes de zipar.
    Usa regex igual ao extrair_numero_nfse_do_xml, sem depender de namespace.
    """
    NS = "http://www.sped.fazenda.gov.br/nfse"

    def _n(tag): return f"{{{NS}}}{tag}"

    try:
        texto = conteudo_xml.decode("utf-8-sig", errors="ignore")
        # Fix encoding declaration conflict (STM São Luís declara WINDOWS-1252 mas grava UTF-8)
        texto = re.sub(r'encoding=["\']WINDOWS-1252["\']', 'encoding="utf-8"', texto, flags=re.IGNORECASE)
    except Exception:
        texto = str(conteudo_xml)

    dados = {
        "numero": numero_fallback,
        "cnpj": "", "cliente": "", "emissao": "",
        "cnae": "", "valor_nfse": 0.0, "deducao": 0.0, "valor_liq": 0.0,
        "valor_servico": 0.0, "iss": "", "aliquota": 0.0, "valor_iss": 0.0,
        "tp_ret_issqn": "", "base_calculo": 0.0, "chave": "",
        "status": status,
        # Tributação federal
        "op_simples":      "",
        "v_ret_cp":        0.0,
        "v_ret_irrf":      0.0,
        "v_ret_csll":      0.0,
        "v_bc_pis_cofins": 0.0,
        "p_aliq_pis":      0.0,
        "v_pis":           0.0,
        "p_aliq_cofins":   0.0,
        "v_cofins":        0.0,
        # Emitente (prestador) — usado na aba TOMADO
        "emit_cnpj": "",
        "emit_nome": "",
    }

    try:
        raiz = ET.fromstring(texto.encode("utf-8"))

        def _txt(el): return (el.text or "").strip() if el is not None else ""

        # NUMERO
        el = raiz.find(f".//{_n('nNFSe')}")
        if el is not None and _txt(el).isdigit():
            dados["numero"] = str(int(_txt(el)))

        # CHAVE — Id do infNFSe, só os dígitos após "NFS"
        inf_nfse = raiz.find(_n("infNFSe"))
        if inf_nfse is not None:
            id_val = inf_nfse.get("Id", "")
            m_chave = re.match(r"NFS(\d+)", id_val)
            if m_chave:
                dados["chave"] = m_chave.group(1)

        # EMISSAO — dhProc ou dhEmi, só a data dd.mm.aaaa
        for tag in ["dhProc", "dhEmi"]:
            el = raiz.find(f".//{_n(tag)}")
            if el is not None and _txt(el):
                raw = _txt(el)[:10]
                try:
                    from datetime import datetime as _dt
                    dados["emissao"] = _dt.strptime(raw, "%Y-%m-%d").strftime("%d.%m.%Y")
                except Exception:
                    dados["emissao"] = raw
                break

        # TOMADOR — CPF ou CNPJ + xNome dentro de <toma>
        toma = raiz.find(f".//{_n('toma')}")
        if toma is not None:
            cpf  = toma.find(_n("CPF"))
            cnpj = toma.find(_n("CNPJ"))
            dados["cnpj"]    = _txt(cpf) if cpf is not None else _txt(cnpj)
            dados["cliente"] = _txt(toma.find(_n("xNome")))

        # EMITENTE — CNPJ + xNome dentro de <emit> (usado na aba TOMADO)
        emit = raiz.find(f".//{_n('emit')}")
        if emit is not None:
            dados["emit_cnpj"] = _txt(emit.find(_n("CNPJ")))
            dados["emit_nome"] = _txt(emit.find(_n("xNome")))

        # CNAE — cTribNac: só o número inicial
        el_cnae = raiz.find(f".//{_n('cTribNac')}")
        if el_cnae is None:
            el_cnae = raiz.find(f".//{_n('xTribNac')}")
        if el_cnae is not None:
            m = re.match(r"(\d+)", _txt(el_cnae))
            dados["cnae"] = m.group(1) if m else _txt(el_cnae)

        # vServ — DPS/infDPS/valores/vServPrest/vServ
        # G (valor_nfse) e I (valor_servico) ambos vem de <vServ>
        v_serv_el = raiz.find(f".//{_n('vServPrest')}/{_n('vServ')}")
        if v_serv_el is not None and _txt(v_serv_el):
            dados["valor_servico"] = float(_txt(v_serv_el))
            dados["valor_nfse"]    = float(_txt(v_serv_el))  # G = <vServ>

        # VALORES bloco infNFSe — vBC, vLiq, pAliqAplic, vISSQN
        inf    = raiz.find(_n("infNFSe")) or raiz
        valores = inf.find(_n("valores"))
        if valores is not None:
            vbc   = valores.find(_n("vBC"))
            vliq  = valores.find(_n("vLiq"))
            paliq = valores.find(_n("pAliqAplic"))
            viss  = valores.find(_n("vISSQN"))
            # K (base_calculo) = <vBC>
            if vbc is not None and _txt(vbc):
                dados["base_calculo"]  = float(_txt(vbc))
            # vLiq (valor liquido)
            if vliq is not None and _txt(vliq):
                dados["valor_liq"] = float(_txt(vliq))
            if paliq is not None and _txt(paliq): dados["aliquota"]  = float(_txt(paliq))
            if viss  is not None and _txt(viss):  dados["valor_iss"] = float(_txt(viss))

        # ISS — tribISSQN: 1=NORMAL / 2=RETIDO
        trib = raiz.find(f".//{_n('tribISSQN')}")
        if trib is not None:
            dados["iss"] = "NORMAL" if _txt(trib) == "1" else "RETIDO"

        # tpRetISSQN — 1=Não Retido / 2=Retido
        tp_ret = raiz.find(f".//{_n('tpRetISSQN')}")
        if tp_ret is not None:
            dados["tp_ret_issqn"] = "Não Retido" if _txt(tp_ret) == "1" else "Retido"

        # opSimpNac — regime Simples Nacional
        op_sn = raiz.find(f".//{_n('opSimpNac')}")
        if op_sn is not None:
            dados["op_simples"] = _txt(op_sn)

        # tribFed — retenções federais
        trib_fed = raiz.find(f".//{_n('tribFed')}")
        if trib_fed is not None:
            def _f(tag):
                el = trib_fed.find(_n(tag))
                return float(_txt(el)) if el is not None and _txt(el) else 0.0
            dados["v_ret_cp"]   = _f("vRetCP")
            dados["v_ret_irrf"] = _f("vRetIRRF")
            dados["v_ret_csll"] = _f("vRetCSLL")
            piscofins = trib_fed.find(_n("piscofins"))
            if piscofins is not None:
                def _fp(tag):
                    el = piscofins.find(_n(tag))
                    return float(_txt(el)) if el is not None and _txt(el) else 0.0
                dados["v_bc_pis_cofins"] = _fp("vBCPisCofins")
                dados["p_aliq_pis"]      = _fp("pAliqPis")
                dados["v_pis"]           = _fp("vPis")
                dados["p_aliq_cofins"]   = _fp("pAliqCofins")
                dados["v_cofins"]        = _fp("vCofins")

    except Exception as _e_xml:
        print(f"  [ERRO DETALHAMENTO] Falha ao extrair XML: {_e_xml}")

    dados["valor_nfse"]    = round(dados["valor_nfse"], 2)
    dados["valor_liq"]     = round(dados["valor_liq"], 2)
    dados["valor_servico"] = round(dados["valor_servico"], 2)
    dados["valor_iss"]     = round(dados["valor_iss"], 2)
    dados["base_calculo"]  = round(dados["base_calculo"], 2)
    dados["v_ret_cp"]        = round(dados["v_ret_cp"],        2)
    dados["v_ret_irrf"]      = round(dados["v_ret_irrf"],      2)
    dados["v_ret_csll"]      = round(dados["v_ret_csll"],      2)
    dados["v_bc_pis_cofins"] = round(dados["v_bc_pis_cofins"], 2)
    dados["p_aliq_pis"]      = round(dados["p_aliq_pis"],      4)
    dados["v_pis"]           = round(dados["v_pis"],           2)
    dados["p_aliq_cofins"]   = round(dados["p_aliq_cofins"],   4)
    dados["v_cofins"]        = round(dados["v_cofins"],        2)
    return dados


def baixar_arquivo_direto(driver, url, pasta, extensao, nome_padrao, timeout=60):
    conteudo = baixar_conteudo_direto(driver, url, pasta, nome_debug=nome_padrao, timeout=timeout)
    return salvar_conteudo_arquivo(conteudo, pasta, extensao, nome_padrao)

def contar_arquivos(pasta, extensao):
    return len([
        f for f in os.listdir(pasta)
        if f.lower().endswith(extensao.lower())
        and not f.lower().endswith(".crdownload")
        and not f.lower().endswith(".tmp")
    ])


def obter_total_registros(driver):
    texto = driver.find_element(By.TAG_NAME, "body").text
    match = re.search(r"Total\s+de\s+(\d+)\s+registros?", texto, re.IGNORECASE)
    if match:
        return int(match.group(1))

    html = driver.page_source
    match = re.search(r"Total\s+de\s+(\d+)\s+registros?", html, re.IGNORECASE)
    if match:
        return int(match.group(1))

    return None


def salvar_print_tela(driver, pasta, codigo, nome):
    nome_limpo = limpar_nome(nome)
    caminho_print = os.path.join(
        pasta,
        f"PRINT_TOTAL_REGISTROS_{codigo}_{nome_limpo}_{COMPETENCIA}.png"
    )

    try:
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1.2)

            total_width = driver.execute_script(
                "return Math.max(document.body.scrollWidth, document.documentElement.scrollWidth);"
            )
            total_height = driver.execute_script(
                "return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);"
            )

            total_width = max(1366, int(total_width or 1366))
            total_height = max(768, min(int(total_height or 768), 12000))

            driver.set_window_size(total_width, total_height)
            time.sleep(1.2)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.8)
        except Exception:
            pass

        driver.save_screenshot(caminho_print)
        print(f"OK - print salvo: {caminho_print}")

        try:
            driver.maximize_window()
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.8)
        except Exception:
            pass

        return caminho_print
    except Exception as e:
        print(f"AVISO - não conseguiu salvar print: {e}")
        return ""


def esperar_tabela_atualizar(driver, texto_antigo=None, timeout=20):
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            linhas = driver.find_elements(By.XPATH, "//table//tbody//tr")
            texto_atual = driver.find_element(By.TAG_NAME, "body").text
            if linhas and (texto_antigo is None or texto_atual != texto_antigo):
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


# FUNCAO DA V18E VALIDADA - NAO RECRIADA
def capturar_links_menu(driver, wait, linha):
    try:
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.3)
    except Exception:
        pass

    botao_opcoes = linha.find_element(By.CLASS_NAME, "icone-trigger")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_opcoes)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", botao_opcoes)
    time.sleep(1.2)

    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//a[contains(@href, '/EmissorNacional/Notas/Download/')]")
        )
    )

    links_visiveis = driver.execute_script("""
        const anchors = Array.from(document.querySelectorAll('a[href*="/EmissorNacional/Notas/Download/"]'));
        return anchors
            .filter(a => {
                const r = a.getBoundingClientRect();
                const s = window.getComputedStyle(a);
                return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
            })
            .map(a => a.href);
    """)

    url_xml = None
    url_pdf = None

    for href in links_visiveis:
        if "/Notas/Download/NFSe/" in href:
            url_xml = href
        if "/Notas/Download/DANFSe/" in href:
            url_pdf = href

    if not url_xml or not url_pdf:
        todos_links = driver.execute_script("""
            return Array.from(document.querySelectorAll('a[href*="/EmissorNacional/Notas/Download/"]'))
                .map(a => a.href);
        """)

        for href in reversed(todos_links):
            if not url_xml and "/Notas/Download/NFSe/" in href:
                url_xml = href
            if not url_pdf and "/Notas/Download/DANFSe/" in href:
                url_pdf = href

    if not url_xml:
        raise Exception("Não consegui capturar o link XML da linha.")

    if not url_pdf:
        raise Exception("Não consegui capturar o link PDF da linha.")

    print(f"OK - capturou link XML: {url_xml}")
    print(f"OK - capturou link PDF: {url_pdf}")

    return url_xml, url_pdf


# FUNCAO DA V18E VALIDADA - NAO RECRIADA
def ir_para_proxima_pagina(driver, pagina_atual):
    proxima = pagina_atual + 1
    texto_antes = driver.find_element(By.TAG_NAME, "body").text

    # Busca pelo href exato pg=N — evita ambiguidade com outros botões fa-angle-right
    xpaths = [
        f"//a[contains(@href,'pg={proxima}')]",
        f"//a[normalize-space()='{proxima}']",
        f"//button[normalize-space()='{proxima}']",
        f"//*[self::a or self::button or self::span][normalize-space()='{proxima}' and not(contains(@class,'disabled'))]",
    ]

    for xp in xpaths:
        try:
            candidatos = driver.find_elements(By.XPATH, xp)
            for el in candidatos:
                classe = (el.get_attribute("class") or "").lower()
                if el.is_displayed() and el.is_enabled() and "disabled" not in classe:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
                    time.sleep(0.3)
                    driver.execute_script("arguments[0].click();", el)
                    print(f"OK - indo para página {proxima}")
                    esperar_tabela_atualizar(driver, texto_antigo=texto_antes, timeout=20)
                    time.sleep(2)
                    return True
        except Exception:
            pass

    return False


# =====================================================================
# CERTIFICADO DIGITAL - BASE V14 VALIDADA
# =====================================================================

def detectar_lang():
    if os.path.exists(r"C:\Program Files\Tesseract-OCR\tessdata\por.traineddata"):
        return "por"
    return "eng"


def ocr_linha(img_linha, lang):
    try:
        img_linha = img_linha.convert("L")
        img_linha = ImageOps.invert(img_linha)
        img_linha = ImageEnhance.Contrast(img_linha).enhance(2.0)
        w, h = img_linha.size
        img_linha = img_linha.resize((w * 2, h * 2))
        txt = pytesseract.image_to_string(img_linha, lang=lang, config="--psm 7")
        return normalizar(re.sub(r"\s+", " ", txt).strip())
    except Exception:
        return ""


def capturar_linhas_ocr(lang):
    resultado = {1: "", 2: "", 3: ""}
    try:
        img_full = ImageGrab.grab(bbox=(OCR_LEFT, OCR_TOP, OCR_RIGHT, OCR_BOTTOM))
        for n in range(1, 4):
            y0 = (n - 1) * ALTURA_LINHA_OCR
            y1 = n * ALTURA_LINHA_OCR
            img_l = img_full.crop((0, y0, img_full.width, y1))
            resultado[n] = ocr_linha(img_l, lang)
    except Exception as e:
        print(f"[OCR] Erro: {e}")
    return resultado



def _assinatura_tela_certificado():
    """Captura uma assinatura leve da área onde o popup/lista de certificados aparece."""
    try:
        img = ImageGrab.grab(bbox=(300, 170, 900, 430)).convert("L")
        img = img.resize((60, 26))
        if hasattr(img, "get_flattened_data"):
            return list(img.get_flattened_data())
        return list(img.getdata())
    except Exception:
        return None


def _diferenca_assinatura(a, b):
    if not a or not b or len(a) != len(b):
        return 0
    return sum(abs(x - y) for x, y in zip(a, b)) / len(a)


def aguardar_popup_certificado_inteligente(timeout=25):
    """
    Espera inteligente do popup de certificado.
    Não usa espera fixa de 7s: aguarda mudança real na região da tela
    onde a janela/lista de certificados aparece.
    """
    print("[CERT] Aguardando popup do certificado aparecer...")
    base = _assinatura_tela_certificado()
    inicio = time.time()

    while time.time() - inicio < timeout:
        atual = _assinatura_tela_certificado()
        diff = _diferenca_assinatura(base, atual)

        if diff > 6:
            time.sleep(0.8)
            print(f"[CERT] Popup detectado por mudança de tela. diff={diff:.2f}")
            return True

        time.sleep(0.25)

    raise Exception("Popup de certificado não apareceu dentro do tempo limite.")

def thread_selecionar_certificado(tema, lang):
    tema_norm = normalizar(tema)
    tema_busca = tema_norm[:15].strip()

    aguardar_popup_certificado_inteligente(timeout=25)

    # FOCO DA LISTA — sequência original V5 restaurada.
    # Popup nativo do Windows: Selenium não enxerga. O clique físico foca a lista,
    # CTRL+HOME volta para o topo, e o segundo clique garante foco antes do loop.
    time.sleep(0.5)
    pyautogui.click(LINHA_X[1], LINHA_Y[1])
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "Home")
    time.sleep(0.5)
    pyautogui.click(LINHA_X[1], LINHA_Y[1])
    time.sleep(0.3)

    print(f"[CERT] Buscando: '{tema_busca}'")

    encontrado = False
    linha_encontrada = None

    linhas_anteriores = None
    repeticoes = 0

    for i in range(MAX_ROLAGENS):
        linhas = capturar_linhas_ocr(lang)

        # DIAGNÓSTICO — mostra o que o OCR leu
        print(f"[OCR] Rolagem {i}: {linhas}")

        # DETECÇÃO DE REPETIÇÃO (FIM DA LISTA)
        if linhas == linhas_anteriores:
            repeticoes += 1
        else:
            repeticoes = 0
            linhas_anteriores = linhas

        if repeticoes >= 2:
            print("[CERT] FIM DA LISTA DETECTADO — certificado não localizado")
            pyautogui.press("esc")
            return False

        for num, txt in linhas.items():
            if parece_certificado_correto(txt, tema_busca):
                print(f"[CERT] ENCONTRADO rolagem {i} linha {num}: '{txt[:60]}'")
                encontrado = True
                linha_encontrada = num
                break

        if encontrado:
            break

        pyautogui.press("down")
        time.sleep(0.15)

    if not encontrado:
        print(f"[CERT] ERRO: certificado não encontrado.")
        pyautogui.press("esc")
        return False

    x_clique = LINHA_X[linha_encontrada]
    y_clique = LINHA_Y[linha_encontrada]
    print(f"[CERT] Clicando linha {linha_encontrada} X={x_clique}, Y={y_clique}...")
    pyautogui.click(x_clique, y_clique)
    time.sleep(0.5)
    pyautogui.click(x_clique, y_clique)
    time.sleep(0.3)

    print(f"[CERT] Clicando OK ({OK_X}, {OK_Y})...")
    pyautogui.click(OK_X, OK_Y)
    time.sleep(3)

    print("[CERT] OK clicado! Certificado selecionado.")
    return True

def clicar_botao_certificado(driver, wait):
    xpaths = [
        "//*[self::button or self::a][contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'certificado')]",
        "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'acesso com certificado')]",
    ]

    for xpath in xpaths:
        try:
            el = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            time.sleep(0.5)
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            driver.execute_script("arguments[0].click();", el)
            print("OK - botão certificado clicado.")
            return True
        except Exception:
            pass

    clicou = driver.execute_script("""
        const els = Array.from(document.querySelectorAll('button,a,img,div,span'));
        const alvo = els.find(el => {
            const txt = ((el.innerText||'')+(el.alt||'')+(el.title||'')+(el.ariaLabel||'')).toLowerCase();
            const r = el.getBoundingClientRect();
            const s = window.getComputedStyle(el);
            return r.width>0&&r.height>0&&s.visibility!=='hidden'&&s.display!=='none'&&txt.includes('certificado');
        });
        if(alvo){alvo.scrollIntoView({block:'center'});alvo.click();return true;}
        return false;
    """)

    if clicou:
        print("OK - botão certificado clicado via JS.")
        return True

    raise Exception("Botão de certificado não localizado.")


def login_com_senha(driver, wait, login, senha):
    print("Rota: LOGIN E SENHA")

    campo_login = wait.until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='CPF/CNPJ']"))
    )
    campo_login.click()

    for letra in str(login):
        campo_login.send_keys(letra)
        time.sleep(0.05)

    campo_senha = wait.until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
    )
    campo_senha.click()

    for letra in str(senha):
        campo_senha.send_keys(letra)
        time.sleep(0.05)

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Entrar')]"))
    ).click()

    time.sleep(3)

    # Detecta mensagens de login inválido ANTES de esperar o site estabilizar
    try:
        texto_pagina = driver.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        texto_pagina = driver.page_source.lower()

    mensagens_login_invalido = [
        "usuário informado deve ser um cpf",
        "usuario informado deve ser um cpf",
        "cpf(11 dígitos) ou cnpj(14 dígitos)",
        "cpf(11 digitos) ou cnpj(14 digitos)",
        "senha incorreta",
        "login incorreto",
        "usuário ou senha",
        "usuario ou senha",
        "login ou senha",
        "credenciais inválidas",
        "credenciais invalidas",
        "não confere",
        "nao confere",
        "inválido",
        "invalido",
        "acesso negado",
    ]

    if any(msg in texto_pagina for msg in mensagens_login_invalido):
        raise LoginInvalidoException(f"Login/senha inválido para o usuário '{login}'.")

    esperar_site_estavel(driver, wait, timeout=35, descricao="pós-login por senha")
    print("OK - login por senha concluído.")


def login_com_certificado(driver, wait, tema, lang):
    print("Rota: CERTIFICADO DIGITAL")

    if not tema or str(tema).strip().upper() == "X":
        raise CertificadoNaoLocalizadoException("TEMA_CERTIFICADO vazio ou X — campo nao preenchido na planilha.")

    resultado_cert = {"ok": None}

    def executar_selecao_certificado():
        try:
            resultado_cert["ok"] = thread_selecionar_certificado(tema, lang)
        except Exception as e:
            resultado_cert["ok"] = False
            resultado_cert["erro"] = str(e)

    t = threading.Thread(
        target=executar_selecao_certificado,
        daemon=True
    )
    t.start()

    try:
        clicar_botao_certificado(driver, wait)
    except Exception:
        print("OK - janela de certificado aberta; aguardando seleção segura.")

    print("Aguardando seleção do certificado...")
    t.join(timeout=120)

    if resultado_cert.get("ok") is False:
        raise CertificadoNaoLocalizadoException(
            f"CERTIFICADO NÃO LOCALIZADO: {resultado_cert.get('erro', tema)}"
        )

    if resultado_cert.get("ok") is not True:
        pyautogui.press("esc")
        raise CertificadoNaoLocalizadoException(
            "CERTIFICADO NÃO LOCALIZADO: seleção não retornou confirmação dentro do tempo limite."
        )

    esperar_site_estavel(driver, wait, timeout=45, descricao="pós-login por certificado")
    print("OK - login por certificado concluído.")

def extrair_numero_nfse_da_linha(linha):
    """
    Extrai o número da NFS-e diretamente do texto da linha da tabela.
    A coluna de número da nota contém apenas dígitos — pega o primeiro
    elemento da linha que seja somente numérico.
    Retorna o número como string, ou None se não encontrar.
    """
    try:
        celulas = linha.find_elements(By.TAG_NAME, "td")
        for cel in celulas:
            texto = (cel.text or "").strip()
            if texto.isdigit():
                return texto
    except Exception:
        pass
    return None


def processar_empresa_pos_login(driver, wait, codigo, nome, pasta_download):
    total_registros = None
    erros_empresa = []
    qtd_canceladas = 0
    qtd_analise = 0
    qtd_pdf_execucao = 0
    qtd_xml_execucao = 0
    xmls_execucao = []
    nfse_processadas = set()  # evita baixar mesma nota duas vezes (contingência)

    esperar_clicavel_seguro(
        driver,
        By.XPATH,
        "//a[contains(@href, '/EmissorNacional/Notas/Emitidas')]",
        timeout=40,
        descricao="menu NFS-e Emitidas"
    ).click()

    print("OK - entrou em Emitidas")
    esperar_site_estavel(driver, wait, timeout=35, descricao="tela de NFS-e Emitidas")

    campo_inicio = wait.until(EC.presence_of_element_located((By.ID, "datainicio")))
    campo_inicio.clear()
    campo_inicio.send_keys(DATA_INICIAL)

    campo_fim = wait.until(EC.presence_of_element_located((By.ID, "datafim")))
    campo_fim.clear()
    campo_fim.send_keys(DATA_FINAL)

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Filtrar')]"))
    ).click()

    print("OK - filtrou")

    # Detecção rápida de "Nenhum registro encontrado" — não espera timeout cheio
    inicio_espera = time.time()
    nenhum_registro = False
    while time.time() - inicio_espera < 20:
        try:
            texto = driver.find_element(By.TAG_NAME, "body").text
            if "nenhum registro" in texto.lower():
                nenhum_registro = True
                break
            linhas = driver.find_elements(By.XPATH, "//table//tbody//tr")
            if linhas:
                break
        except Exception:
            pass
        time.sleep(0.5)

    esperar_site_estavel(driver, wait, timeout=20, descricao="resultado do filtro")

    total_registros = obter_total_registros(driver)

    if nenhum_registro or total_registros == 0:
        total_registros = 0
        print("Empresa sem movimento: nenhum download será realizado.")
        salvar_print_tela(driver, pasta_download, codigo, f"{nome}_emitidas")
        return total_registros, qtd_canceladas, qtd_analise, erros_empresa, qtd_pdf_execucao, qtd_xml_execucao, xmls_execucao

    print(f"Total de registros no portal: {total_registros}")
    salvar_print_tela(driver, pasta_download, codigo, f"{nome}_emitidas")

    pagina_atual = 1

    while True:
        linhas = driver.find_elements(By.XPATH, "//table//tbody//tr")
        print(f"\nPágina {pagina_atual} - notas na tela: {len(linhas)}")

        for i in range(len(linhas)):
            print(f"\nProcessando página {pagina_atual}, nota {i + 1}")

            try:
                linhas = driver.find_elements(By.XPATH, "//table//tbody//tr")
                linha = linhas[i]

                # Detecta status da nota — 4 tipos pelo HTML confirmado via DevTools
                # Normal: data-situacao="P100_GERADA"
                # Cancelada: tb-cancelada no src do img
                # Cancelamento em análise: tb-pendente.svg no src do img
                # Substituída: tb-subs.svg / data-situacao="P105_NFSE_SUBSTITUIDA"
                status_nota = ""
                try:
                    html_linha = linha.get_attribute("outerHTML") or ""
                    html_lower = html_linha.lower()
                    if "tb-cancelada" in html_lower or "nfs-e cancelada" in html_lower:
                        status_nota = "Cancelada"
                    elif "tb-pendente" in html_lower or "cancelamento sob análise" in html_lower or "cancelamento sob analise" in html_lower:
                        status_nota = "Cancelamento em análise"
                    elif "tb-subs" in html_lower or "p105_nfse_substituida" in html_lower or "substituição" in html_lower or "substituicao" in html_lower:
                        status_nota = "Substituída"
                except Exception:
                    pass

                if not status_nota:
                    try:
                        status_nota = driver.execute_script("""
                            const linha = arguments[0];
                            const imgs = linha.querySelectorAll('img');
                            for (const img of imgs) {
                                const src = (img.getAttribute('src') || '').toLowerCase();
                                const title = (img.getAttribute('title') || '').toLowerCase();
                                const orig = (img.getAttribute('data-original-title') || '').toLowerCase();
                                if (src.includes('tb-cancelada') || title.includes('cancelada') || orig.includes('cancelada')) return 'Cancelada';
                                if (src.includes('tb-pendente') || title.includes('análise') || orig.includes('análise') || orig.includes('analise')) return 'Cancelamento em análise';
                                if (src.includes('tb-subs') || title.includes('substitui') || orig.includes('substitui')) return 'Substituída';
                            }
                            return '';
                        """, linha) or ""
                    except Exception:
                        pass

                if status_nota == "Cancelada":
                    qtd_canceladas += 1
                    print(f"  [CANCELADA] Baixando XML para detalhamento.")
                elif status_nota == "Cancelamento em análise":
                    qtd_analise += 1
                    print(f"  [EM ANÁLISE] Baixando XML para detalhamento.")
                elif status_nota == "Substituída":
                    print(f"  [SUBSTITUÍDA] Baixando XML para detalhamento.")

                url_xml, url_pdf = capturar_links_menu(driver, wait, linha)

                conteudo_xml = baixar_conteudo_direto(
                    driver=driver,
                    url=url_xml,
                    pasta=pasta_download,
                    nome_debug=f"XML_p{pagina_atual}_{i + 1}"
                )

                num_nfse = extrair_numero_nfse_do_xml(conteudo_xml)

                if not num_nfse:
                    num_nfse = extrair_numero_nfse_da_linha(linha)
                    if num_nfse:
                        print(f"AVISO - número não veio pelo XML; usando número da linha: {num_nfse}")

                if num_nfse:
                    print(f"Número da NFS-e identificado pelo XML: {num_nfse}")
                    if num_nfse in nfse_processadas:
                        print(f"  [SKIP] Nota {num_nfse} já processada (contingência/duplicata) — ignorando.")
                        continue
                    nfse_processadas.add(num_nfse)
                    nome_base = f"PADRAO_NACIONAL_NFSE_{num_nfse}_{COMPETENCIA}"
                else:
                    print(f"AVISO - não extraiu número da NFS-e; usando fallback p{pagina_atual}_{i + 1}")
                    nome_base = f"PADRAO_NACIONAL_NFSE_p{pagina_atual}_{i + 1}_{COMPETENCIA}"

                # Notas normais: salva XML + PDF normalmente
                # Notas especiais (cancelada/análise): salva XML + PDF em pasta temporária separada
                if status_nota in ("Cancelada", "Cancelamento em análise"):
                    _status_norm = (status_nota.lower()
                                    .replace(' ', '_')
                                    .replace('ã', 'a').replace('á', 'a').replace('â', 'a').replace('à', 'a')
                                    .replace('é', 'e').replace('ê', 'e')
                                    .replace('í', 'i').replace('î', 'i')
                                    .replace('õ', 'o').replace('ó', 'o').replace('ô', 'o')
                                    .replace('ú', 'u').replace('û', 'u')
                                    .replace('ç', 'c'))
                    pasta_especial = os.path.join(pasta_download, f"_especiais_{_status_norm}")
                    os.makedirs(pasta_especial, exist_ok=True)
                    salvar_conteudo_arquivo(
                        conteudo=conteudo_xml,
                        pasta=pasta_especial,
                        extensao=".xml",
                        nome_padrao=f"{nome_base}.xml"
                    )
                    if url_pdf:
                        _pdf_ok = False
                        for _tent_pdf in range(1, 4):
                            try:
                                baixar_arquivo_direto(
                                    driver=driver,
                                    url=url_pdf,
                                    pasta=pasta_especial,
                                    extensao=".pdf",
                                    nome_padrao=f"{nome_base}.pdf",
                                    timeout=15
                                )
                                _pdf_ok = True
                                break
                            except Exception as e_pdf_esp:
                                if _tent_pdf < 3:
                                    print(f"  [AVISO] PDF nota especial tentativa {_tent_pdf}/3 falhou — retentando em 3s... ({e_pdf_esp})")
                                    time.sleep(3)
                                else:
                                    _msg_pdf_err = f"PDF nota especial {nome_base} nao baixado apos 3 tentativas: {e_pdf_esp}"
                                    print(f"  [ERRO] {_msg_pdf_err}")
                                    erros_empresa.append(_msg_pdf_err)
                                    _log_diag_pn.append({
                                        "fase": "PDF_ESPECIAL",
                                        "empresa": f"{codigo} - {nome}",
                                        "resultado": "FALHA_PDF_ESPECIAL",
                                        "observacao": _msg_pdf_err,
                                    })
                    dados_nota = extrair_dados_xml_bytes(conteudo_xml, numero_fallback=num_nfse or "", status=status_nota)
                    xmls_execucao.append(dados_nota)
                    try:
                        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                    except Exception:
                        pass
                    continue

                caminho_xml_salvo = salvar_conteudo_arquivo(
                    conteudo=conteudo_xml,
                    pasta=pasta_download,
                    extensao=".xml",
                    nome_padrao=f"{nome_base}.xml"
                )
                qtd_xml_execucao += 1
                dados_nota = extrair_dados_xml_bytes(conteudo_xml, numero_fallback=num_nfse or "", status=status_nota)
                xmls_execucao.append(dados_nota)

                baixar_arquivo_direto(
                    driver=driver,
                    url=url_pdf,
                    pasta=pasta_download,
                    extensao=".pdf",
                    nome_padrao=f"{nome_base}.pdf"
                )
                qtd_pdf_execucao += 1

                try:
                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                except Exception:
                    pass

            except Exception as e:
                erro = f"Página {pagina_atual}, nota {i + 1}: {e}"
                print(f"ERRO - {erro}")
                erros_empresa.append(erro)

            time.sleep(1)

        if not ir_para_proxima_pagina(driver, pagina_atual):
            print("Última página ou próxima página não encontrada.")
            break

        pagina_atual += 1

    return total_registros, qtd_canceladas, qtd_analise, erros_empresa, qtd_pdf_execucao, qtd_xml_execucao, xmls_execucao




def gerar_detalhamento_nfse(pasta_download, nome_empresa, competencia,
                              xmls_emitidas=None, xmls_recebidas=None):
    """
    Gera arquivo com duas abas: PRESTADO e TOMADO.
    PRESTADO: notas emitidas — coluna CNPJ/CLIENTE = tomador (<toma>)
    TOMADO  : notas recebidas — coluna CNPJ/PRESTADOR = emitente (<emit>)
    """
    xmls_emitidas  = xmls_emitidas  or []
    xmls_recebidas = xmls_recebidas or []

    if not xmls_emitidas and not xmls_recebidas:
        return

    def _sort_key(d):
        try:
            from datetime import datetime as _dt
            return _dt.strptime(d.get("emissao", "01.01.1900"), "%d.%m.%Y")
        except Exception:
            return d.get("emissao", "")

    notas_prest = sorted(xmls_emitidas,  key=_sort_key)
    notas_toma  = sorted(xmls_recebidas, key=_sort_key)

    mm, aaaa   = competencia[:2], competencia[2:]
    nome_limpo = re.sub(r"\s+", "_", re.sub(r"[^\w\s\-]", "", nome_empresa).strip())
    caminho_saida = os.path.join(pasta_download,
                                 f"Detalhamento_de_NFSe_{nome_limpo}_{mm}{aaaa}.xlsx")

    # Regime do cliente — lido do primeiro XML prestado
    _REGIME_MAP = {
        "1": "NÃO OPTANTE",
        "2": "MEI — Microempreendedor Individual",
        "3": "Simples Nacional EPP",
        "4": "Simples Nacional Microempresa",
    }
    _op_sn  = notas_prest[0].get("op_simples", "") if notas_prest else ""
    _regime = _REGIME_MAP.get(_op_sn, "NÃO IDENTIFICADO")

    def _fill(h):  return PatternFill("solid", fgColor="FF" + h)
    def _font(bold=False, cor="000000", size=9):
        return Font(bold=bold, color="FF" + cor, name="Calibri", size=size)
    def _borda():
        s = Side(style="thin", color="FFB0B0B0")
        return Border(left=s, right=s, top=s, bottom=s)
    def _alin(h="center"):
        return Alignment(horizontal=h, vertical="center")

    # ── Função interna: monta uma aba ────────────────────────────────────────
    def _render_sheet(ws, notas, tipo):
        eh_toma = (tipo == "tomado")

        # Linha 1 — título
        ws.merge_cells("A1:W1")
        ws["A1"] = (f"DETALHAMENTO DE NFS-e  —  {nome_empresa.upper()}  —  "
                    f"{mm}/{aaaa} - REGIME: {_regime}")
        ws["A1"].fill      = _fill("1F4E78")
        ws["A1"].font      = _font(bold=True, cor="FFFFFF", size=11)
        ws["A1"].alignment = _alin("left")
        ws.row_dimensions[1].height = 22

        # Linha 2 — grupos
        label_a = "PRESTADOR" if eh_toma else "TOMADOR"
        grupos = [
            (label_a,              "A2", "B2", "2E75B6"),
            ("DOCUMENTO",          "C2", "I2", "1F6B75"),
            ("ISS",                "J2", "M2", "6C3483"),
            ("RETENÇÕES FEDERAIS", "N2", "V2", "1A5276"),
            ("NOTA",               "W2", "W2", "7D6608"),
        ]
        for label, ini, fim, cor in grupos:
            ws.merge_cells(f"{ini}:{fim}")
            c = ws[ini]
            c.value = label; c.fill = _fill(cor)
            c.font = _font(bold=True, cor="FFFFFF", size=10)
            c.alignment = _alin("center"); c.border = _borda()
        ws.row_dimensions[2].height = 18

        # Linha 3 — cabeçalhos
        label_b = "PRESTADOR" if eh_toma else "CLIENTE"
        cabecalhos = [
            "CNPJ", label_b,
            "EMISSÃO", "NÚMERO", "STATUS", "CNAE", "VALOR TOTAL NFSe", "DEDUÇÃO", "VALOR SERVIÇO",
            "ISS", "BASE DE CALCULO", "ALÍQUOTA", "VALOR ISS",
            "INSS", "IRRF", "CSLL",
            "BASE PIS", "ALIQUOTA PIS", "PIS",
            "BASE COFINS", "ALIQUOTA COFINS", "COFINS",
            "CHAVE",
        ]
        larguras = [18, 35, 13, 10, 20, 10, 18, 12, 16, 12, 16, 12, 14,
                    14, 14, 14, 14, 13, 14, 14, 15, 14, 40]
        for col, (cab, larg) in enumerate(zip(cabecalhos, larguras), 1):
            c = ws.cell(row=3, column=col, value=cab)
            c.fill = _fill("2E86C1"); c.font = _font(bold=True, cor="FFFFFF", size=9)
            c.alignment = _alin("center"); c.border = _borda()
            ws.column_dimensions[c.column_letter].width = larg
        ws.row_dimensions[3].height = 18
        ws.auto_filter.ref = "A3:W3"

        # Estilos de fundo
        fill_normal  = _fill("EBF5FB"); fill_alt     = _fill("F4F6F7")
        fill_cancel  = _fill("FADBD8"); fill_analise = _fill("FEF9E7")
        fill_substit = _fill("E8DAEF")

        total_g = total_i = 0.0
        n_regular = n_cancelada = n_analise = n_substit = 0

        monetarios    = {7, 8, 9, 11, 13, 14, 15, 16, 17, 19, 20, 22}
        aliquota_col  = 12
        aliquota_cols = {18, 21}

        for i, d in enumerate(notas, start=4):
            status     = d.get("status", "")
            _vnfse     = d.get("valor_nfse", 0.0)
            _vliq      = d.get("valor_liq",  0.0)
            _cancelada = (status == "Cancelada")

            if status == "Cancelada":
                base_fill = fill_cancel;  n_cancelada += 1
            elif status == "Cancelamento em análise":
                base_fill = fill_analise; n_analise += 1
                total_g += round(_vnfse, 2); total_i += round(_vliq, 2)
            elif status == "Substituída":
                base_fill = fill_substit; n_substit += 1
                total_g += round(_vnfse, 2); total_i += round(_vliq, 2)
            else:
                base_fill = fill_alt if i % 2 == 0 else fill_normal
                n_regular += 1
                total_g += round(_vnfse, 2); total_i += round(_vliq, 2)

            # TOMADO: CNPJ/nome vêm do emitente (<emit>); PRESTADO: do tomador (<toma>)
            _cnpj    = d.get("emit_cnpj", d.get("cnpj",    "")) if eh_toma else d.get("cnpj",    "")
            _cliente = d.get("emit_nome", d.get("cliente", "")) if eh_toma else d.get("cliente", "")

            valores_linha = [
                _cnpj, _cliente,
                d.get("emissao", ""), d.get("numero", ""),
                "",  # STATUS col E — abaixo
                d.get("cnae", ""),
                0.0  if _cancelada else d.get("valor_nfse",      0.0),  # G
                0.00,                                                      # H
                0.0  if _cancelada else d.get("valor_liq",        0.0),  # I
                ""   if _cancelada else d.get("tp_ret_issqn",     ""),   # J
                0.0  if _cancelada else d.get("base_calculo",     0.0),  # K
                0.0  if _cancelada else d.get("aliquota",         0.0),  # L
                0.0  if _cancelada else d.get("valor_iss",        0.0),  # M
                0.0  if _cancelada else d.get("v_ret_cp",         0.0),  # N
                0.0  if _cancelada else d.get("v_ret_irrf",       0.0),  # O
                0.0  if _cancelada else d.get("v_ret_csll",       0.0),  # P
                0.0  if _cancelada else d.get("v_bc_pis_cofins",  0.0),  # Q
                0.0  if _cancelada else d.get("p_aliq_pis",       0.0),  # R
                0.0  if _cancelada else d.get("v_pis",            0.0),  # S
                0.0  if _cancelada else d.get("v_bc_pis_cofins",  0.0),  # T
                0.0  if _cancelada else d.get("p_aliq_cofins",    0.0),  # U
                0.0  if _cancelada else d.get("v_cofins",         0.0),  # V
                d.get("chave", ""),                                        # W
            ]

            for col, val in enumerate(valores_linha, 1):
                c = ws.cell(row=i, column=col, value=val)
                c.fill = base_fill; c.font = _font(size=9); c.border = _borda()
                if col in monetarios:
                    c.number_format = 'R$ #,##0.00'; c.alignment = _alin("right")
                elif col == aliquota_col or col in aliquota_cols:
                    c.number_format = '0.00"%"'; c.alignment = _alin("center")
                elif col in (3, 4, 6, 10):
                    c.alignment = _alin("center")
                elif col == 23:
                    c.alignment = _alin("left"); c.font = _font(size=8)
                else:
                    c.alignment = _alin("left")

            # STATUS col E
            c_st = ws.cell(row=i, column=5)
            if status == "Cancelada":
                c_st.value = "✗  Cancelada"
                c_st.fill  = _fill("FADBD8")
                c_st.font  = Font(bold=True, color="FFC0392B", name="Calibri", size=9)
            elif status == "Cancelamento em análise":
                c_st.value = "⚠  Cancelamento em análise"
                c_st.fill  = _fill("FEF9E7")
                c_st.font  = Font(bold=True, color="FF9C6500", name="Calibri", size=9)
            elif status == "Substituída":
                c_st.value = "↺  Substituída"
                c_st.fill  = _fill("E8DAEF")
                c_st.font  = Font(bold=True, color="FF6C3483", name="Calibri", size=9)
            else:
                c_st.value = "✓"
                c_st.fill  = _fill("D5F5E3")
                c_st.font  = Font(bold=True, color="FF1E8449", name="Calibri", size=10)
            c_st.alignment = _alin("center"); c_st.border = _borda()
            ws.row_dimensions[i].height = 15

        # Rodapé
        lt = len(notas) + 4
        ws.merge_cells(f"A{lt}:F{lt}")
        _txt_tot = (f"TOTAL  |  Regular: {n_regular}  |  Substituídas: {n_substit}  |  "
                    f"Canceladas: {n_cancelada}  |  Cancelamento em análise: {n_analise}"
                    if n_substit else
                    f"TOTAL  |  Regular: {n_regular}  |  Canceladas: {n_cancelada}  |  "
                    f"Cancelamento em análise: {n_analise}")
        c_l = ws.cell(row=lt, column=1, value=_txt_tot)
        c_l.fill = _fill("F9E79F"); c_l.font = _font(bold=True, size=10)
        c_l.alignment = _alin("center"); c_l.border = _borda()

        p, u = 4, lt - 1

        # Quando não há notas (lista vazia), p > u → SUMIFS causaria referência circular.
        # Nesse caso escreve 0 direto em vez de fórmula.
        _tem_dados = len(notas) > 0

        def _tot_mon(col_num, formula):
            val = formula if _tem_dados else 0.0
            c = ws.cell(row=lt, column=col_num, value=val)
            c.fill = _fill("D5F5E3"); c.font = _font(bold=True, size=10)
            c.number_format = 'R$ #,##0.00'; c.alignment = _alin("right"); c.border = _borda()

        def _tot_vaz(col_num):
            ws.cell(row=lt, column=col_num, value="").fill  = _fill("F9E79F")
            ws.cell(row=lt, column=col_num).border = _borda()

        _tot_mon(7,  round(total_g, 2))
        _tot_vaz(8)
        _tot_mon(9,  round(total_i, 2))
        _tot_vaz(10)
        _tot_mon(11, f"=SUMIFS(K{p}:K{u},E{p}:E{u},\"<>*Cancelada*\")")
        _tot_vaz(12)
        _tot_mon(13, f"=SUMIFS(M{p}:M{u},E{p}:E{u},\"<>*Cancelada*\")")

        for cn, cl in {14:"N",15:"O",16:"P",17:"Q",18:None,19:"S",20:"T",21:None,22:"V"}.items():
            if cl:
                _tot_mon(cn, f"=SUMIFS({cl}{p}:{cl}{u},E{p}:E{u},\"<>*Cancelada*\")")
            else:
                _tot_vaz(cn)
        _tot_vaz(23)
        ws.row_dimensions[lt].height = 18

        # ISS Não Retido / Retido
        li1, li2 = lt + 1, lt + 2
        for li, txt, cor_fill, cor_font in [
            (li1, "Não Retido", "D5F5E3", "1E8449"),
            (li2, "Retido",     "FADBD8", "C0392B"),
        ]:
            c12 = ws.cell(row=li, column=12, value=txt)
            c12.fill = _fill(cor_fill); c12.font = _font(bold=True, size=10, cor=cor_font)
            c12.alignment = _alin("right"); c12.border = _borda()
            crit = f"\"Não Retido\"" if txt == "Não Retido" else f"\"Retido\""
            c13 = ws.cell(row=li, column=13,
                          value=f"=SUMIFS(M{p}:M{u},J{p}:J{u},{crit},E{p}:E{u},\"<>*Cancelada*\")" if _tem_dados else 0.0)
            c13.fill = _fill(cor_fill); c13.font = _font(bold=True, size=10, cor=cor_font)
            c13.number_format = 'R$ #,##0.00'; c13.alignment = _alin("right"); c13.border = _borda()

        ws.freeze_panes = "A4"
    # ── fim _render_sheet ────────────────────────────────────────────────────

    wb = openpyxl.Workbook()
    ws_prest = wb.active
    ws_prest.title = "PRESTADO"
    _render_sheet(ws_prest, notas_prest, "prestado")

    ws_toma = wb.create_sheet("TOMADO")
    _render_sheet(ws_toma, notas_toma, "tomado")

    if os.path.exists(caminho_saida):
        try:
            os.remove(caminho_saida)
        except PermissionError:
            print(f"\n  *** ERRO: Feche o arquivo '{os.path.basename(caminho_saida)}' no Excel e execute novamente. ***\n")
            return
    try:
        wb.save(caminho_saida)
        print(f"  [DETALHAMENTO] {os.path.basename(caminho_saida)}")
    except Exception as e_d:
        print(f"  [AVISO] Detalhamento não salvo: {e_d}")




# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def _definir_pasta_base() -> str:
    """Lê SENHAS SAO LUIS.xlsx varrendo todas as linhas até encontrar
    a coluna PASTA e o primeiro valor preenchido.
    Funciona independente de quantas linhas de cabeçalho existam."""
    try:
        wb = openpyxl.load_workbook(SENHAS_XLSX, read_only=True, data_only=True)
        ws = wb.active
        idx_pasta   = None
        header_row  = None

        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if row is None:
                continue
            # Procura linha de cabeçalho que contenha "PASTA"
            if idx_pasta is None:
                for j, cel in enumerate(row):
                    if cel and str(cel).strip().upper() == "PASTA":
                        idx_pasta  = j
                        header_row = i
                        break
            else:
                # Já achou o índice — pega o primeiro valor não-vazio
                if len(row) > idx_pasta and row[idx_pasta]:
                    val = str(row[idx_pasta]).strip()
                    if val:
                        wb.close()
                        return val
        wb.close()
    except Exception:
        pass
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# JANELA DE PROGRESSO — RWA Soluções
# ─────────────────────────────────────────────────────────────────────────────
import queue   as _q_mod
import tkinter as _tk_mod

class RWAProgressWindow:
    """Janela de progresso profissional — RWA Soluções."""

    def __init__(self):
        self._queue      = _q_mod.Queue()
        self._pronto     = _sys.modules["threading"].Event()
        self._fechado    = _sys.modules["threading"].Event()
        self._inicio     = None
        self._repro_fila = 0
        self._total       = 0   # total de empresas na lista
        self._repro_idx  = 0
        self._thread     = _sys.modules["threading"].Thread(target=self._run, daemon=True)
        self._thread.start()
        self._pronto.wait(timeout=4)

    # ── stdout ───────────────────────────────────────────────────────────────
    def write(self, texto):
        if texto:
            self._queue.put(("LOG", texto))
            import re as _re
            m = _re.search(r'\[(\d{1,3})/(\d{1,3})\]\s+(.+)', texto)
            if m:
                self._queue.put(("PROG", int(m.group(1)), int(m.group(2)), m.group(3).strip()))
                return
            if _re.search(r'\[REPROCESSAR\]', texto):
                # Só incrementa o contador se for tentativa 1 — tentativa 2 não conta
                if "1 tentativa" in texto or "Tentativa 1" in texto:
                    self._queue.put(("REPRO_ADD",))
                return
            m2 = _re.search(r'\[REPROCESSAMENTO\]\s+(.+)', texto)
            if m2:
                self._queue.put(("REPRO_START", m2.group(1).strip()))

    def flush(self):
        pass

    def finalizar(self):
        self._queue.put(("FIM",))
        self._fechado.wait()

    # ── UI ───────────────────────────────────────────────────────────────────
    def _run(self):
        import time as _time

        root = _tk_mod.Tk()
        root.title("RWA Soluções — STM São Luís")
        root.resizable(False, False)
        root.configure(bg="#13131f")

        W, H = 660, 600
        root.geometry(f"{W}x{H}+{(root.winfo_screenwidth()-W)//2}+{(root.winfo_screenheight()-H)//2}")

        # ── Barra superior roxa ───────────────────────────────────────────────
        _tk_mod.Frame(root, bg="#4f46e5", height=5).pack(fill="x")

        # ── Header ───────────────────────────────────────────────────────────
        _hdr = _tk_mod.Frame(root, bg="#13131f")
        _hdr.pack(fill="x", padx=30, pady=(18, 0))

        _tk_mod.Label(_hdr, text="🏛", font=("Arial", 22),
                      bg="#13131f", fg="#4f46e5").pack(side="left", padx=(0, 12))

        _col = _tk_mod.Frame(_hdr, bg="#13131f")
        _col.pack(side="left")
        _tk_mod.Label(_col, text="STM São Luís", font=("Arial", 15, "bold"),
                      bg="#13131f", fg="#ffffff", anchor="w").pack(anchor="w")
        _tk_mod.Label(_col, text="Automação do Sistema Tributário Municipal",
                      font=("Arial", 9), bg="#13131f", fg="#555577", anchor="w").pack(anchor="w")

        # Timer (canto direito)
        self._lbl_timer = _tk_mod.Label(_hdr, text="00:00", font=("Consolas", 13, "bold"),
                                         bg="#13131f", fg="#333355")
        self._lbl_timer.pack(side="right")

        # ── Separador ────────────────────────────────────────────────────────
        _tk_mod.Frame(root, bg="#1e1e30", height=1).pack(fill="x", padx=30, pady=(14, 0))

        # ── Card empresa atual ────────────────────────────────────────────────
        _card = _tk_mod.Frame(root, bg="#1a1a2e", pady=0)
        _card.pack(fill="x", padx=30, pady=(14, 0))

        _tk_mod.Label(_card, text="EMPRESA ATUAL", font=("Arial", 7, "bold"),
                      bg="#1a1a2e", fg="#333366", anchor="w").pack(anchor="w", padx=14, pady=(10, 2))
        self._lbl_emp = _tk_mod.Label(_card, text="Iniciando...",
                                       font=("Arial", 11, "bold"),
                                       bg="#1a1a2e", fg="#aaaadd", anchor="w")
        self._lbl_emp.pack(anchor="w", padx=14, pady=(0, 10))

        # ── Barra de progresso ────────────────────────────────────────────────
        _fp = _tk_mod.Frame(root, bg="#13131f")
        _fp.pack(fill="x", padx=30, pady=(14, 0))

        _topo = _tk_mod.Frame(_fp, bg="#13131f")
        _topo.pack(fill="x")
        _tk_mod.Label(_topo, text="PROGRESSO", font=("Arial", 7, "bold"),
                      bg="#13131f", fg="#333366").pack(side="left")
        self._lbl_pct = _tk_mod.Label(_topo, text="0 / 0",
                                       font=("Arial", 8, "bold"),
                                       bg="#13131f", fg="#4f46e5")
        self._lbl_pct.pack(side="right")

        _tk_mod.Frame(_fp, bg="#1e1e30", height=1).pack(fill="x", pady=(4, 6))

        self._canvas = _tk_mod.Canvas(_fp, height=10, bg="#1a1a2e",
                                       highlightthickness=0, bd=0)
        self._canvas.pack(fill="x")
        self._barra = self._canvas.create_rectangle(0, 0, 0, 10,
                                                      fill="#4f46e5", outline="")
        self._barra_shine = self._canvas.create_rectangle(0, 0, 0, 4,
                                                           fill="#7c73f5", outline="")

        # Label de reprocessamento — aparece abaixo da barra
        self._lbl_repro = _tk_mod.Label(
            root, text="",
            font=("Arial", 9, "bold"),
            bg="#13131f", fg="#f59e0b", anchor="w"
        )
        self._lbl_repro.pack(fill="x", padx=30, pady=(6, 0))

        # ── Log ───────────────────────────────────────────────────────────────
        _fl = _tk_mod.Frame(root, bg="#13131f")
        _fl.pack(fill="both", expand=True, padx=30, pady=(8, 0))

        _tk_mod.Label(_fl, text="LOG DE EXECUÇÃO", font=("Arial", 7, "bold"),
                      bg="#13131f", fg="#333366", anchor="w").pack(anchor="w")
        _tk_mod.Frame(_fl, bg="#1e1e30", height=1).pack(fill="x", pady=(4, 6))

        _fr_txt = _tk_mod.Frame(_fl, bg="#0d0d1a")
        _fr_txt.pack(fill="both", expand=True)

        self._txt = _tk_mod.Text(_fr_txt, bg="#0d0d1a", fg="#7777aa",
                                  font=("Consolas", 8), relief="flat",
                                  state="disabled", wrap="none",
                                  height=12,
                                  insertbackground="#4f46e5",
                                  selectbackground="#2a2a4a")
        self._txt.tag_config("ok",    foreground="#22c55e")
        self._txt.tag_config("erro",  foreground="#ef4444")
        self._txt.tag_config("aviso", foreground="#f59e0b")
        self._txt.tag_config("info",  foreground="#6366f1")

        _sb = _tk_mod.Scrollbar(_fr_txt, bg="#1a1a2e", troughcolor="#13131f",
                                 command=self._txt.yview)
        self._txt.configure(yscrollcommand=_sb.set)
        _sb.pack(side="right", fill="y")
        self._txt.pack(fill="both", expand=True, padx=6, pady=6)

        # ── Rodapé ────────────────────────────────────────────────────────────
        _rod = _tk_mod.Frame(root, bg="#13131f")
        _rod.pack(fill="x", padx=30, pady=(10, 14))

        _tk_mod.Label(_rod, text="Desenvolvido por RWA Soluções",
                      font=("Arial", 7), bg="#13131f", fg="#252540").pack(side="left")

        self._btn = _tk_mod.Button(_rod, text="⏳  Processando...",
                                    font=("Arial", 9, "bold"),
                                    bg="#252538", fg="#444466",
                                    relief="flat", padx=18, pady=5,
                                    state="disabled", cursor="arrow")
        self._btn.pack(side="right")

        self._root  = root
        self._inicio = _time.time()
        self._pronto.set()
        root.after(50,   self._poll)
        root.after(1000, self._tick)
        root.mainloop()
        self._fechado.set()

    # ── Timer ─────────────────────────────────────────────────────────────────
    def _tick(self):
        import time as _time
        if self._inicio and not self._fechado.is_set():
            s = int(_time.time() - self._inicio)
            self._lbl_timer.configure(
                text=f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}",
                fg="#4f46e5")
            self._root.after(1000, self._tick)

    # ── Poll fila ─────────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                item = self._queue.get_nowait()

                if item[0] == "LOG":
                    linha = item[1]
                    self._txt.configure(state="normal")
                    tag = ("ok"    if any(x in linha for x in ["[OK]", "✓", "sucesso"]) else
                           "erro"  if any(x in linha for x in ["[ERRO]", "✗", "BLOQUEIO"]) else
                           "aviso" if any(x in linha for x in ["[AVISO]", "REPROCESSAMENTO", "instável"]) else
                           "info"  if any(x in linha for x in ["[INFO]", "[LOGIN]", "[PDF]"]) else
                           None)
                    if tag:
                        self._txt.insert("end", linha, tag)
                    else:
                        self._txt.insert("end", linha)
                    self._txt.see("end")
                    self._txt.configure(state="disabled")

                elif item[0] == "PROG":
                    _, atual, total, nome = item
                    if total: self._total = total
                    pct    = int((atual / total) * 100) if total else 0
                    fill_w = int(598 * pct / 100)
                    self._canvas.coords(self._barra,       0, 0, fill_w, 10)
                    self._canvas.coords(self._barra_shine, 0, 0, fill_w,  3)
                    self._canvas.itemconfig(self._barra,       fill="#4f46e5")
                    self._canvas.itemconfig(self._barra_shine, fill="#7c73f5")
                    if self._repro_fila > 0:
                        self._lbl_pct.configure(
                            text=f"{atual} / {total}  ({pct}%)  +  {self._repro_fila} repro",
                            fg="#4f46e5"
                        )
                    else:
                        self._lbl_pct.configure(text=f"{atual} / {total}  ({pct}%)", fg="#4f46e5")
                    self._lbl_emp.configure(text=nome, fg="#ffffff")

                elif item[0] == "REPRO_ADD":
                    self._repro_fila += 1
                    self._lbl_repro.configure(
                        text=f"⟳  {self._repro_fila} empresa(s) na fila de reprocessamento",
                        fg="#f59e0b"
                    )

                elif item[0] == "REPRO_START":
                    _, nome = item
                    self._repro_idx += 1
                    self._lbl_emp.configure(
                        text=f"⟳  REPROCESSANDO ({self._repro_idx}/{self._repro_fila})  —  {nome}",
                        fg="#f59e0b"
                    )
                    self._lbl_repro.configure(
                        text=f"⟳  Reprocessando {self._repro_idx} de {self._repro_fila}  —  {nome}",
                        fg="#f59e0b"
                    )
                    self._canvas.itemconfig(self._barra,       fill="#f59e0b")
                    self._canvas.itemconfig(self._barra_shine, fill="#fbbf24")

                elif item[0] == "FIM":
                    self._canvas.coords(self._barra,       0, 0, 598, 10)
                    self._canvas.coords(self._barra_shine, 0, 0, 598,  3)
                    self._canvas.itemconfig(self._barra,       fill="#22c55e")
                    self._canvas.itemconfig(self._barra_shine, fill="#4ade80")
                    self._lbl_emp.configure(text="✅  Concluído com sucesso!", fg="#22c55e")
                    self._lbl_pct.configure(text="100%", fg="#22c55e")
                    self._lbl_timer.configure(fg="#22c55e")
                    if self._repro_fila > 0:
                        self._lbl_repro.configure(
                            text=f"✅  {self._total} empresa(s) processada(s) — concluído",
                            fg="#22c55e"
                        )
                    self._inicio = None
                    self._btn.configure(text="Concluído", state="normal",
                                        bg="#22c55e", fg="#0a0a0a",
                                        cursor="hand2", command=self._root.destroy)

        except _q_mod.Empty:
            pass

        if not self._fechado.is_set():
            self._root.after(50, self._poll)


def main():
    global PASTA_BASE
    _pw = RWAProgressWindow()
    _stdout_original = _sys.stdout
    _sys.stdout = _pw
    empresas = ler_empresas()

    # Define PASTA_BASE a partir da coluna PASTA da planilha,
    # varrendo todas as linhas independente da estrutura de cabeçalho.
    _pb = _definir_pasta_base()
    if _pb:
        PASTA_BASE = _pb
    else:
        # Fallback: primeiro empresa com salvamento preenchido (comportamento anterior)
        for emp in empresas:
            pasta_emp = emp.get("salvamento", "").strip()
            if pasta_emp:
                PASTA_BASE = pasta_emp
                break

    print(f"\n{'=' * 72}")
    print(f"  STM SEMFAZ — SERVICOS SAO LUIS  |  VERSAO 2.1")
    print(f"  Competencia  : {PERIODO}")
    print(f"  Empresas     : {len(empresas)}")
    print(f"{'=' * 72}\n")

    # ── VALIDAÇÃO PRÉVIA: todas as empresas devem ter PASTA informada ─────────
    _sem_pasta = [emp for emp in empresas if not emp.get("salvamento", "").strip()]
    if _sem_pasta:
        print("=" * 72)
        print("  ERRO — IMPORTAÇÃO BLOQUEADA")
        print("  As empresas abaixo estão sem PASTA informada na planilha.")
        print("  Preencha a coluna PASTA para todas as empresas e execute novamente.")
        print("=" * 72)
        for _emp in _sem_pasta:
            print(f"  ✗  Código {_emp.get('codigo','')} — {_emp.get('empresa','')}")
        print("=" * 72 + "\n")

        # Grava no DIAGNOSTICO de erros
        _msg_bloqueio = (
            f"IMPORTAÇÃO BLOQUEADA — {len(_sem_pasta)} empresa(s) sem PASTA informada:\n"
            + "\n".join(
                f"  Código {e.get('codigo','')} — {e.get('empresa','')}"
                for e in _sem_pasta
            )
        )
        try:
            with open(caminho_diagnostico_erros(), "a", encoding="utf-8") as _fd:
                _fd.write("\n" + "=" * 90 + "\n")
                _fd.write(f"Data/Hora  : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                _fd.write(_msg_bloqueio + "\n")
                _fd.write("=" * 90 + "\n")
        except Exception:
            pass

        sys.exit(1)
    # ─────────────────────────────────────────────────────────────────────────

    resultados            = []
    fila                  = []
    total_original        = len(empresas)
    tempo_inicio_total    = time.time()
    tempos_empresa        = []
    _email_50_enviado_stm = False  # garante disparo único do email de 50%

    # ── LOG DE DIAGNÓSTICO ───────────────────────────────────────────────────
    _log_empresas = []          # Um registro por tentativa
    _log_instabilidades = []    # Instabilidades capturadas com timestamp

    # ── Helpers de email HTML (compartilhados pelos 3 emails: início, 50%, conclusão) ──
    def _render_html_email_rwa(tag_texto, tag_bg, tag_fg, titulo, intro, secoes, callout, intro_destaque=False):
        """Monta HTML inline-friendly do email RWA Soluções.

        secoes: lista de tuplas (titulo_secao_ou_None, [(label, valor, destacar_vermelho), ...])
        callout: tuple (bg, borda, fg, texto)
        intro_destaque: se True, aplica negrito + verde sobrio + icone na frase do intro.
        """
        partes_secoes = []
        for titulo_secao, linhas in secoes:
            if titulo_secao:
                partes_secoes.append(
                    f'<tr><td style="padding:14px 0 4px 0;font-family:Arial,Helvetica,sans-serif;'
                    f'font-size:10px;color:#6B6B6B;letter-spacing:1px;text-transform:uppercase;'
                    f'font-weight:bold;">{titulo_secao}</td></tr>'
                )
            partes_secoes.append('<tr><td style="padding:0;"><table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">')
            for idx, item in enumerate(linhas):
                label, valor, destacar = item
                eh_ultima = (idx == len(linhas) - 1)
                borda = "" if eh_ultima else "border-bottom:1px solid #EFEFEF;"
                cor_valor = "#A32D2D" if destacar else "#1a1a1a"
                partes_secoes.append(
                    f'<tr>'
                    f'<td style="padding:8px 0;{borda}font-family:Arial,Helvetica,sans-serif;'
                    f'font-size:13px;color:#6B6B6B;width:55%;">{label}</td>'
                    f'<td style="padding:8px 0;{borda}font-family:Arial,Helvetica,sans-serif;'
                    f'font-size:13px;color:{cor_valor};text-align:right;font-weight:bold;">{valor}</td>'
                    f'</tr>'
                )
            partes_secoes.append('</table></td></tr>')

        html_secoes = "".join(partes_secoes)
        callout_bg, callout_borda, callout_fg, callout_texto = callout

        _intro_cor   = "#1F6B43" if intro_destaque else "#1a1a1a"
        _intro_peso  = "bold"    if intro_destaque else "normal"
        _intro_icone = "✓ "      if intro_destaque else ""

        return (
            f'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" '
            f'"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
            f'<html><head><meta charset="utf-8"/></head>'
            f'<body style="margin:0;padding:0;background-color:#f4f4f4;">'
            f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f4;padding:20px 0;">'
            f'<tr><td align="center">'
            f'<table width="480" cellpadding="0" cellspacing="0" border="0" '
            f'style="max-width:480px;background-color:#ffffff;border:1px solid #e0e0e0;border-radius:6px;overflow:hidden;">'
            # ── header escuro ──
            f'<tr><td style="background-color:#0F1B2D;padding:20px 18px;">'
            f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#7FB3E0;'
            f'letter-spacing:1.5px;font-weight:bold;margin-bottom:8px;">RWA SOLUÇÕES</div>'
            f'<div style="display:inline-block;background-color:{tag_bg};color:{tag_fg};'
            f'font-family:Arial,Helvetica,sans-serif;font-size:10px;padding:3px 8px;'
            f'border-radius:3px;font-weight:bold;letter-spacing:0.5px;margin-bottom:10px;">{tag_texto}</div>'
            f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:17px;color:#FFFFFF;'
            f'font-weight:bold;line-height:1.3;margin-top:4px;">{titulo}</div>'
            f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#B5D4F4;'
            f'margin-top:2px;">STM São Luís</div>'
            f'</td></tr>'
            # ── corpo ──
            f'<tr><td style="padding:18px;">'
            f'<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:{_intro_cor};'
            f'line-height:1.5;margin-bottom:12px;font-weight:{_intro_peso};">{_intro_icone}{intro}</div>'
            f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">'
            f'{html_secoes}'
            f'</table>'
            # ── callout ──
            f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;border-collapse:collapse;">'
            f'<tr><td style="background-color:{callout_bg};border-left:3px solid {callout_borda};'
            f'padding:10px 12px;font-family:Arial,Helvetica,sans-serif;font-size:12px;'
            f'color:{callout_fg};line-height:1.5;">{callout_texto}</td></tr>'
            f'</table>'
            # ── rodapé ──
            f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;border-top:1px solid #EFEFEF;">'
            f'<tr><td style="padding-top:10px;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#888;">'
            f'<span style="color:#0F1B2D;font-weight:bold;">RWA Soluções</span><br/>'
            f'<span style="font-size:11px;">Automação fiscal para escritórios contábeis</span>'
            f'</td></tr></table>'
            f'</td></tr>'
            f'</table>'
            f'</td></tr></table>'
            f'</body></html>'
        )

    def _enviar_email_rwa(conta, senha, destino, assunto, corpo_html, corpo_plain):
        """Envia email multipart/alternative (HTML + plain) via Gmail SMTP."""
        import smtplib as _smtp_h
        from email.mime.multipart import MIMEMultipart as _MMulti_h
        from email.mime.text import MIMEText as _MText_h
        msg = _MMulti_h("alternative")
        msg["From"]    = conta
        msg["To"]      = destino
        msg["Subject"] = assunto
        msg.attach(_MText_h(corpo_plain, "plain", "utf-8"))
        msg.attach(_MText_h(corpo_html,  "html",  "utf-8"))
        srv = _smtp_h.SMTP("smtp.gmail.com", 587)
        srv.starttls()
        srv.login(conta, senha)
        srv.sendmail(conta, destino, msg.as_string())
        srv.quit()

    # ── Email de início — disparado aqui, quando o processamento realmente começa ──
    try:
        _rwa_email_destino  = os.environ.get("RWA_EMAIL_DESTINO",   "").strip()
        _rwa_email_agendado = os.environ.get("RWA_EMAIL_AGENDADO",  "0") == "1"
        _rwa_email_conta    = os.environ.get("RWA_EMAIL_CONTA",     "").strip()
        _rwa_email_senha    = os.environ.get("RWA_EMAIL_SENHA_APP", "").strip()

        if _rwa_email_destino and _rwa_email_agendado and _rwa_email_conta and _rwa_email_senha:
            _ini_dt_stm = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            # ── plain text (fallback) ──
            _corpo_ini_stm = (
                f"Prezado(a),\n\n"
                f"A automação STM São Luís foi iniciada conforme agendado.\n\n"
                f"  Competência   : {PERIODO}\n"
                f"  Início        : {_ini_dt_stm}\n"
                f"  Total empresas: {total_original}\n\n"
                f"Acompanhe o andamento no computador onde a automação está sendo executada.\n\n"
                f"RWA Soluções"
            )
            # ── HTML (visual) ──
            _corpo_ini_html_stm = _render_html_email_rwa(
                tag_texto="INICIADA",
                tag_bg="#185FA5", tag_fg="#E6F1FB",
                titulo="Automação iniciada",
                intro="Prezado(a), a automação foi iniciada conforme agendado.",
                secoes=[
                    (None, [
                        ("Competência",       PERIODO,              False),
                        ("Início",            _ini_dt_stm,          False),
                        ("Total de empresas", str(total_original),  False),
                    ]),
                ],
                callout=(
                    "#F4F8FB", "#185FA5", "#0C447C",
                    "Acompanhe o andamento no computador onde a automação está sendo executada.",
                ),
            )
            _enviar_email_rwa(
                _rwa_email_conta, _rwa_email_senha, _rwa_email_destino,
                f"RWA Soluções — STM São Luís iniciada — {PERIODO}",
                _corpo_ini_html_stm, _corpo_ini_stm,
            )
            print("[EMAIL] Email de inicio enviado.")
    except Exception as _e_mail_ini_stm:
        print(f"[EMAIL] Falha ao enviar email de inicio: {_e_mail_ini_stm}")

    for idx, emp in enumerate(empresas, 1):
        fila.append({"emp": emp, "tentativa": 1, "ordem_original": idx})

    while fila:
        item           = fila.pop(0)
        emp            = item["emp"]
        tentativa      = item.get("tentativa", 1)
        ordem_original = item.get("ordem_original", 0)

        nome         = emp["nome"]
        codigo       = emp.get("codigo", "")
        empresa_nome = emp.get("empresa", "")
        regime       = emp.get("regime", "")
        salvamento   = emp.get("salvamento", "").strip()

        print("-" * 72)
        if tentativa == 1:
            print(f"  [{ordem_original:>03}/{total_original}] {nome}")
            print(f"  Regime: {regime or 'NAO INFORMADO'}")
        else:
            print(f"  [REPROCESSAMENTO] {nome}")
        print("-" * 72)

        # ── VALIDACAO: PASTA ───────────────────────────────────
        if not salvamento:
            print(f"  [ERRO] Pasta de salvamento nao informada na planilha. Empresa ignorada.")
            cnpj_emp = emp.get("login", "")
            resultados.append({
                "codigo":              codigo,
                "empresa":             empresa_nome,
                "cnpj":                cnpj_emp,
                "regime":              regime,
                "qtd_notas":           "",
                "valor_total":         "",
                "iss_proprio":         "",
                "iss_retido":          "",
                "qtd_tomado":          "",
                "valor_total_tomado":  "",
                "iss_proprio_tomado":  "",
                "iss_retido_tomado":   "",
                "pdf_prestado":        "Erro",
                "xml_prestado":        "Erro",
                "pdf_tomado":          "Erro",
                "xml_tomado":          "Erro",
                "talao_nf":            None,
                "talao_guia":          None,
                "login_senha":         "",
                "observacao":          "PASTA NAO INFORMADA",
            })
            salvar_relatorio_parcial(resultados)
            continue

        pasta         = pasta_empresa(nome, salvamento)
        pasta_int     = pasta_interna(pasta, codigo)
        tomados       = emp.get("tomados", "SIM")
        xml_prestados = emp.get("xml_prestados", "SIM")
        driver = None
        tempo_inicio_empresa = time.time()
        _hora_inicio_empresa = datetime.now().strftime("%H:%M:%S")

        # Timers por fase
        _t_login_ini = _t_login_fim = None
        _t_rel_ini   = _t_rel_fim   = None
        _t_down_ini  = _t_down_fim  = None
        _fase_atual  = "INICIO"

        status_pdf_prestado = "Erro"
        status_xml_prestado = "Erro"
        status_pdf_tomado   = "Erro"
        status_xml_tomado   = "Erro"
        status_talao_nf     = None
        status_talao_guia   = None
        status_login        = "✓"
        observacao          = ""
        reprocessar_no_final = False
        resumo_pdf = {"qtd_notas": "", "valor_total": "", "iss_proprio": "", "iss_retido": ""}
        resumo_pdf_tomado = {"qtd_notas": "", "valor_total": "", "iss_proprio": "", "iss_retido": ""}

        try:
            driver   = configurar_driver(pasta)

            # ── FASE: LOGIN ──────────────────────────────────────────────────
            _fase_atual  = "LOGIN"
            _t_login_ini = time.time()
            login_ok = fazer_login(driver, emp["login"], emp["senha"])
            _t_login_fim = time.time()

            if login_ok is not True:
                # login_ok é a mensagem de erro exata do portal (string)
                msg_erro_login      = login_ok if isinstance(login_ok, str) else "Erro de login não identificado."
                status_login        = "INCORRETO"
                status_pdf_prestado = "Erro"
                status_xml_prestado = "Erro"
                status_pdf_tomado   = "Erro"
                status_xml_tomado   = "Erro"
                status_talao_nf     = "Erro"
                status_talao_guia   = "Erro"
                observacao          = msg_erro_login
                print(f"  [INFO] Login falhou: {msg_erro_login}. Seguindo para a proxima empresa.")

            else:
                # ── FASE: RELATORIO (PDF + XML) ──────────────────────────────
                _fase_atual = "RELATORIO"
                _t_rel_ini  = time.time()

                _baixar_tomado = (tomados == "SIM")

                # ── 1+2. PDF PRESTADO + TOMADO ───────────────────────────────
                status_pdf_prestado, status_pdf_tomado = baixar_pdfs_prestado_e_tomado(
                    driver, pasta, pasta_int, codigo,
                    cnpj=emp.get("login", ""),
                    baixar_tomado=_baixar_tomado)

                if status_pdf_prestado == "PDF":
                    resumo_pdf = analisar_pdf_gerencial(pasta, nome_pdf_prestado(codigo))

                if status_pdf_tomado == "PDF":
                    resumo_pdf_tomado = analisar_pdf_gerencial(pasta_int, nome_pdf_tomado(codigo))

                # ── 3+4. XML PRESTADO + TOMADO ───────────────────────────────
                # XML Prestado respeita a coluna "XML PRESTADOS" da planilha:
                #   SIM (default) → baixa pelo STM
                #   NAO           → pula (cliente baixa pelo Padrao Nacional)
                _pular_xml_prest_config = (xml_prestados == "NAO")
                if _pular_xml_prest_config:
                    print("  [INFO] XML Prestado dispensado (XML PRESTADOS=NAO na planilha).")

                status_xml_prestado, status_xml_tomado = baixar_xmls_prestado_e_tomado(
                    driver, pasta, pasta_int, codigo,
                    pular_prestado=(status_pdf_prestado == "Sem movimento" or _pular_xml_prest_config),
                    pular_tomado  =(not _baixar_tomado or status_pdf_tomado in ("Sem movimento", "DISPENSADO")),
                )

                # Se pulou por config (e nao por falta de movimento), marca status DISPENSADO
                # para nao confundir com "Sem movimento" no gerencial
                if _pular_xml_prest_config and status_pdf_prestado != "Sem movimento":
                    status_xml_prestado = "DISPENSADO"

                if status_xml_prestado == "Portal Nacional":
                    _obs_xml = "XML nao disponivel no portal STM — consultar pelo Portal Nacional"
                    observacao = (observacao + " | " + _obs_xml) if observacao else _obs_xml

                _t_rel_fim = time.time()

                # ── FASE: DOWNLOAD (TALAO) ───────────────────────────────────
                _fase_atual = "DOWNLOAD"
                _t_down_ini = time.time()

                # ── 5+6. TALAO FISCAL ────────────────────────────────────────
                # Talao NF  : executa se tem movimento (Prestado OU Tomado), qualquer regime
                # Guia ISS  : executa se tem movimento E regime != SIMPLES
                tem_movimento = (
                    status_pdf_prestado not in ("Sem movimento", "Erro")
                    or status_pdf_tomado not in ("Sem movimento", "Erro", "DISPENSADO")
                )

                if not tem_movimento:
                    print("  [TALAO] Dispensado: nenhum movimento em Prestado e Tomado.")
                    status_talao_nf   = None
                    status_talao_guia = None
                else:
                    resultado_talao = baixar_talao(
                        driver, pasta, codigo,
                        emp=emp,
                        executar_guia=regime_tem_talao(regime),
                        pasta_print=pasta
                    )
                    status_talao_nf   = resultado_talao.get("nf",   "Erro")
                    status_talao_guia = resultado_talao.get("guia", None)
                    if resultado_talao.get("obs_talao"):
                        if observacao:
                            observacao += " | " + resultado_talao.get("obs_talao")
                        else:
                            observacao = resultado_talao.get("obs_talao")

                _t_down_fim = time.time()

        except PortalBloqueioAceoException:
            ESPERA_BLOQUEIO = 240
            print(f"  [BLOQUEIO] ACEO007-016 detectado — fechando browser.")
            try:
                forcar_fechar_driver(driver)  # timeout 8s — seguro mesmo com browser semi-morto
            except Exception:
                pass
            driver = None
            print(f"  [BLOQUEIO] Aguardando {ESPERA_BLOQUEIO}s com browser fechado...")
            for restante in range(ESPERA_BLOQUEIO, 0, -30):
                print(f"  [BLOQUEIO] Retomando em {restante}s...")
                time.sleep(30)
            print("  [BLOQUEIO] Verificando login apos desbloqueio...")
            driver_teste = None
            try:
                driver_teste = configurar_driver(pasta)
                login_ok = fazer_login(driver_teste, emp["login"], emp["senha"])
                try:
                    driver_teste.quit()
                except Exception:
                    pass
                driver_teste = None

                if login_ok is not True:
                    msg_erro_login      = login_ok if isinstance(login_ok, str) else "Login falhou apos desbloqueio ACEO007-016."
                    status_login        = "INCORRETO"
                    status_pdf_prestado = "Erro"
                    status_xml_prestado = "Erro"
                    status_pdf_tomado   = "Erro"
                    status_xml_tomado   = "Erro"
                    status_talao_nf     = "Erro"
                    status_talao_guia   = "Erro"
                    observacao          = msg_erro_login
                    print(f"  [BLOQUEIO] Login falhou apos espera: {msg_erro_login}")
                else:
                    # Login OK — reenfileira para processamento completo pelo fluxo normal
                    if tentativa == 1:
                        print("  [BLOQUEIO] Login OK. Reenfileirando para reprocessamento completo.")
                        fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                        reprocessar_no_final = True
                    else:
                        # Tentativa 2+ com ACEO — esgotou tentativas automaticas
                        observacao = "ACEO007-016 — bloqueio persistiu. Reprocessamento manual necessario."
                        print(f"  [BLOQUEIO] {observacao}")

            except Exception as e_aceo:
                if driver_teste is not None:
                    try:
                        driver_teste.quit()
                    except Exception:
                        pass
                observacao = f"Falha ao verificar login apos ACEO007-016: {type(e_aceo).__name__}"
                print(f"  [BLOQUEIO] Erro inesperado: {e_aceo}")

        except PortalInstabilidadeException as e:
            _e_str = str(e)
            _e_talao = "Talao" in _e_str and "nao gerado apos 2 tentativas" in _e_str

            # Renderer morto disfarçado de PortalInstabilidade (ex: login com renderer travado)
            # A causa raiz é TimeoutException com "receiving message from renderer"
            _causa_str = str(e.__cause__) if e.__cause__ else ""
            if "receiving message from renderer" in _causa_str:
                _msg_renderer = "Chrome renderer travado (detectado via causa) — empresa enviada ao final da fila sem espera"
                observacao = _msg_renderer
                print(f"  [RENDERER MORTO] {_msg_renderer}")
                _log_instabilidades.append({
                    "empresa":   nome,
                    "codigo":    codigo,
                    "tentativa": tentativa,
                    "fase":      _fase_atual,
                    "horario":   datetime.now().strftime("%H:%M:%S"),
                    "duracao_s": round(time.time() - tempo_inicio_empresa, 1),
                    "mensagem":  _msg_renderer,
                })
                # NÃO tira print — renderer morto, qualquer comando no driver trava por segundos
                _decisao_renderer = "REPROCESSAMENTO AGENDADO — tentativa 2 sera executada automaticamente" if tentativa == 1 else "REPROCESSAMENTO MANUAL — renderer morto em 2 tentativas"
                registrar_diagnostico(emp, "RENDERER MORTO", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao=_decisao_renderer)
                if tentativa == 1:
                    print("  [REPROCESSAR] Empresa enviada para o final da fila (1 tentativa).")
                    fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                    reprocessar_no_final = True
                else:
                    observacao = "Renderer morto em duas tentativas — reprocessar manualmente"
                    print("  [MANUAL] Renderer morto repetido. Reprocessamento manual necessario.")

            elif _e_talao:
                # Instabilidade exclusiva do talão (volume grande) — não reprocessa
                observacao = "Portal instável em três tentativas — limite de infraestrutura da prefeitura atingido, volume de notas excede a capacidade do servidor municipal. Emitir talão manualmente."
                print(f"  [INSTABILIDADE] {observacao}")
                _log_instabilidades.append({
                    "empresa":   nome,
                    "codigo":    codigo,
                    "tentativa": tentativa,
                    "fase":      _fase_atual,
                    "horario":   datetime.now().strftime("%H:%M:%S"),
                    "duracao_s": round(time.time() - tempo_inicio_empresa, 1),
                    "mensagem":  _e_str,
                })
                if driver is not None:
                    pass
                registrar_diagnostico(emp, "POSSIVEL INSTABILIDADE DO PORTAL", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao="EMITIR TALAO MANUALMENTE — instabilidade por volume, sem reprocessamento automatico")
            else:
                observacao = f"Portal instável durante o processamento — empresa enviada para reprocessamento"
                print(f"  [INSTABILIDADE] {observacao}")
                _log_instabilidades.append({
                    "empresa":   nome,
                    "codigo":    codigo,
                    "tentativa": tentativa,
                    "fase":      _fase_atual,
                    "horario":   datetime.now().strftime("%H:%M:%S"),
                    "duracao_s": round(time.time() - tempo_inicio_empresa, 1),
                    "mensagem":  _e_str,
                })
                registrar_diagnostico(emp, "POSSIVEL INSTABILIDADE DO PORTAL", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao="REPROCESSAMENTO AGENDADO — tentativa 2 sera executada automaticamente" if tentativa == 1 else "REPROCESSAMENTO MANUAL — instabilidade em 2 tentativas")

                if tentativa == 1:
                    print("  [REPROCESSAR] Empresa enviada para o final da fila (1 tentativa).")
                    fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                    reprocessar_no_final = True
                else:
                    observacao = "Portal instável em duas tentativas — reprocessar manualmente"
                    print("  [MANUAL] Instabilidade repetida. Reprocessamento manual necessario.")

        except TimeoutException as e:
            if "receiving message from renderer" in str(e):
                _msg_renderer = "Chrome renderer travado — empresa enviada ao final da fila sem espera"
                observacao = _msg_renderer
                print(f"  [RENDERER MORTO] {_msg_renderer}")
                _log_instabilidades.append({
                    "empresa":   nome,
                    "codigo":    codigo,
                    "tentativa": tentativa,
                    "fase":      _fase_atual,
                    "horario":   datetime.now().strftime("%H:%M:%S"),
                    "duracao_s": round(time.time() - tempo_inicio_empresa, 1),
                    "mensagem":  _msg_renderer,
                })
                # NÃO tira print — renderer morto, qualquer comando no driver trava por segundos
                _decisao_timeout = "REPROCESSAMENTO AGENDADO — tentativa 2 sera executada automaticamente" if tentativa == 1 else "REPROCESSAMENTO MANUAL — renderer morto em 2 tentativas"
                registrar_diagnostico(emp, "RENDERER MORTO", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao=_decisao_timeout)
                if tentativa == 1:
                    print("  [REPROCESSAR] Empresa enviada para o final da fila (1 tentativa).")
                    fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                    reprocessar_no_final = True
                else:
                    observacao = "Renderer morto em duas tentativas — reprocessar manualmente"
                    print("  [MANUAL] Renderer morto repetido. Reprocessamento manual necessario.")
            else:
                if tentativa == 1:
                    observacao = "Timeout inesperado — empresa enviada para reprocessamento"
                    print(f"  [REPROCESSAR] TimeoutException nao-renderer. Tentativa 1 — reprocessando.")
                else:
                    observacao = "Timeout inesperado em duas tentativas — reprocessamento manual"
                    print(f"  [MANUAL] TimeoutException repetida. Reprocessamento manual necessario.")
                if driver is not None:
                    try:
                        tirar_print(driver, os.path.join(pasta_download, "erro_geral.png"))
                    except Exception:
                        pass
                _decisao_timeout_else = "REPROCESSAMENTO AGENDADO — tentativa 2 sera executada automaticamente" if tentativa == 1 else "REPROCESSAMENTO MANUAL — timeout em 2 tentativas"
                registrar_diagnostico(emp, "TIMEOUT INESPERADO", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao=_decisao_timeout_else)
                if tentativa == 1:
                    fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                    reprocessar_no_final = True

        except Exception as e:
            _e_msg = str(e)
            _e_conn_timed_out = "ERR_CONNECTION_TIMED_OUT" in _e_msg or "net::" in _e_msg

            if driver is not None and getattr(driver, "_selecionar_empresa_stm", False) and tentativa == 1:
                observacao = "Falha no fluxo especial Selecionar Empresa — empresa enviada para reprocessamento"
                print(f"  [REPROCESSAR] {observacao}: {type(e).__name__}: {e}")
                registrar_diagnostico(emp, "FALHA FLUXO ESPECIAL - REPROCESSAMENTO", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao="REPROCESSAMENTO AGENDADO — tentativa 2 sera executada automaticamente")
                fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                reprocessar_no_final = True

            elif _e_conn_timed_out and tentativa == 1:
                # Portal fora do ar (ERR_CONNECTION_TIMED_OUT) — reprocessa automaticamente
                observacao = "Portal fora do ar (ERR_CONNECTION_TIMED_OUT) — empresa enviada para reprocessamento"
                print(f"  [REPROCESSAR] Portal offline detectado. Reprocessando empresa {codigo}.")
                try:
                    tirar_print(driver, os.path.join(pasta_download, "erro_geral.png"))
                except Exception:
                    pass
                registrar_diagnostico(emp, "PORTAL FORA DO AR", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao="REPROCESSAMENTO AGENDADO — portal offline, tentativa 2 sera executada automaticamente")
                fila.append({"emp": emp, "tentativa": 2, "ordem_original": ordem_original})
                reprocessar_no_final = True

            elif _e_conn_timed_out and tentativa > 1:
                observacao = "Portal fora do ar em duas tentativas — reprocessamento manual necessario"
                print(f"  [MANUAL] Portal offline repetido. Reprocessamento manual necessario.")
                try:
                    tirar_print(driver, os.path.join(pasta_download, "erro_geral.png"))
                except Exception:
                    pass
                registrar_diagnostico(emp, "PORTAL FORA DO AR", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual,
                                       status_atual={
                                           "PDF Prestado": status_pdf_prestado,
                                           "XML Prestado": status_xml_prestado,
                                           "PDF Tomado":   status_pdf_tomado,
                                           "XML Tomado":   status_xml_tomado,
                                           "Talao NF":     status_talao_nf,
                                           "Guia ISS":     status_talao_guia,
                                       },
                                       tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                       decisao="REPROCESSAMENTO MANUAL — portal offline em 2 tentativas")

            else:
                observacao = f"Erro inesperado durante o processamento: {type(e).__name__}"
                print(f"  [ERRO] Falha geral: {e}")
                if driver is not None:
                    try:
                        tirar_print(driver, os.path.join(pasta_download, "erro_geral.png"))
                    except Exception:
                        pass
                registrar_diagnostico(emp, "FALHA GERAL NA EMPRESA", e, driver, pasta,
                                   tentativa=tentativa, fase=_fase_atual,
                                   status_atual={
                                       "PDF Prestado": status_pdf_prestado,
                                       "XML Prestado": status_xml_prestado,
                                       "PDF Tomado":   status_pdf_tomado,
                                       "XML Tomado":   status_xml_tomado,
                                       "Talao NF":     status_talao_nf,
                                       "Guia ISS":     status_talao_guia,
                                   },
                                   tempo_decorrido=round(time.time() - tempo_inicio_empresa, 1),
                                   decisao="FALHA GERAL — verificar manualmente")

        finally:
            # ── Cronometro empresa ───────────────────────────────────────────
            tempo_empresa = time.time() - tempo_inicio_empresa
            mins  = int(tempo_empresa // 60)
            segs  = int(tempo_empresa % 60)
            label = f"EMPRESA {ordem_original:>03}/{total_original}"
            print(f"  [TEMPO] {label} — {mins}min{segs:02d}s")
            if tentativa == 1:
                tempos_empresa.append(tempo_empresa)

            # ── Registro de diagnóstico por empresa ──────────────────────────
            def _fmt(t_ini, t_fim):
                if t_ini is not None and t_fim is not None:
                    return round(t_fim - t_ini, 1)
                return ""

            _log_empresas.append({
                "codigo":         codigo,
                "empresa":        empresa_nome,
                "tentativa":      tentativa,
                "inicio":         _hora_inicio_empresa,
                "fim":            datetime.now().strftime("%H:%M:%S"),
                "total_s":        round(tempo_empresa, 1),
                "login_s":        _fmt(_t_login_ini, _t_login_fim),
                "relatorio_s":    _fmt(_t_rel_ini,   _t_rel_fim),
                "download_s":     _fmt(_t_down_ini,  _t_down_fim),
                "status_login":   status_login,
                "pdf_prestado":   status_pdf_prestado,
                "xml_prestado":   status_xml_prestado,
                "pdf_tomado":     status_pdf_tomado,
                "xml_tomado":     status_xml_tomado,
                "talao_nf":       status_talao_nf,
                "talao_guia":     status_talao_guia,
                "fase_parou":     _fase_atual,
                "observacao":     observacao,
            })

            # ────────────────────────────────────────────────────────────────

            perfil_temp = None
            try:
                if driver is not None:
                    perfil_temp = getattr(driver, "_perfil_temp_stm", None)
                    forcar_fechar_driver(driver)  # timeout 8s — nunca trava mesmo com renderer morto
            except Exception as e:
                registrar_diagnostico(emp, "FALHA AO FECHAR NAVEGADOR", e, driver, pasta,
                                       tentativa=tentativa, fase=_fase_atual)
            try:
                if perfil_temp and os.path.exists(perfil_temp):
                    shutil.rmtree(perfil_temp, ignore_errors=True)
            except Exception:
                pass
            print(f"  [FIM] Empresa finalizada: {nome}\n")

        if reprocessar_no_final:
            salvar_relatorio_parcial(resultados)
            continue

        # ── DETALHAMENTO POR NOTA — REMOVIDO ─────────────────────────────────
        # Era redundante com o relatorio gerencial do STM Sao Luis,
        # que ja apresenta os dados das notas em formato consolidado.
        # (Funcoes extrair_dados_xml_bytes e gerar_detalhamento_nfse continuam
        # definidas no arquivo mas nao sao mais chamadas no fluxo principal.)
        # ─────────────────────────────────────────────────────────────────────

        # ── RETENÇÕES FEDERAIS — extrai do ZIP Prestado ──────────────────────
        _retencoes = {}
        try:
            _zip_prest = os.path.join(pasta, nome_xml_prestado(codigo).replace(".xml", ".zip"))
            if os.path.exists(_zip_prest):
                _retencoes = extrair_retencoes_zip(_zip_prest)
        except Exception:
            pass
        # ─────────────────────────────────────────────────────────────────────

        # ── DETALHAMENTO POR NOTA — STM São Luís ─────────────────────────────
        try:
            gerar_detalhamento_stm(pasta, codigo, empresa_nome, regime)
        except Exception as _e_det:
            print(f"  [DETALH] Falha no detalhamento (nao impacta fluxo): {_e_det}")
        # ─────────────────────────────────────────────────────────────────────

        resultados.append({
            "codigo":              codigo,
            "empresa":             empresa_nome,
            "cnpj":                emp.get("login", ""),
            "regime":              regime,
            "qtd_notas":           resumo_pdf.get("qtd_notas",        ""),
            "valor_total":         resumo_pdf.get("valor_total",       ""),
            "iss_proprio":         resumo_pdf.get("iss_proprio",       ""),
            "iss_retido":          resumo_pdf.get("iss_retido",        ""),
            "qtd_tomado":          resumo_pdf_tomado.get("qtd_notas",  ""),
            "valor_total_tomado":  resumo_pdf_tomado.get("valor_total",""),
            "iss_proprio_tomado":  resumo_pdf_tomado.get("iss_proprio",""),
            "iss_retido_tomado":   resumo_pdf_tomado.get("iss_retido", ""),
            "pdf_prestado":        status_pdf_prestado,
            "xml_prestado":        status_xml_prestado,
            "pdf_tomado":          status_pdf_tomado,
            "xml_tomado":          status_xml_tomado,
            "talao_nf":            status_talao_nf,
            "talao_guia":          status_talao_guia,
            "login_senha":         status_login,
            "observacao":          observacao,
            "ret_base_pis":        _retencoes.get("ret_base_pis"),
            "ret_aliq_pis":        _retencoes.get("ret_aliq_pis"),
            "ret_aliq_pis_div":    _retencoes.get("ret_aliq_pis_divergente", False),
            "ret_valor_pis":       _retencoes.get("ret_valor_pis"),
            "ret_base_cofins":     _retencoes.get("ret_base_cofins"),
            "ret_aliq_cofins":     _retencoes.get("ret_aliq_cofins"),
            "ret_aliq_cofins_div": _retencoes.get("ret_aliq_cofins_divergente", False),
            "ret_valor_cofins":    _retencoes.get("ret_valor_cofins"),
            "ret_irrf":            _retencoes.get("ret_irrf"),
            "ret_csll":            _retencoes.get("ret_csll"),
        })

        salvar_relatorio_parcial(resultados)

        # ── Email de 50% — disparado uma única vez quando atinge a metade ──
        if (not _email_50_enviado_stm
                and total_original >= 2
                and len(resultados) * 2 >= total_original):
            _email_50_enviado_stm = True
            try:
                _rwa_email_destino  = os.environ.get("RWA_EMAIL_DESTINO",   "").strip()
                _rwa_email_conta    = os.environ.get("RWA_EMAIL_CONTA",     "").strip()
                _rwa_email_senha    = os.environ.get("RWA_EMAIL_SENHA_APP", "").strip()
                if _rwa_email_destino and _rwa_email_conta and _rwa_email_senha:
                    _t50_stm = int(time.time() - tempo_inicio_total)
                    _m50_stm = _t50_stm // 60
                    _s50_stm = _t50_stm % 60
                    _tempo50_str_stm = f"{_m50_stm}min{_s50_stm:02d}s"
                    _progresso50_stm = f"{len(resultados)} de {total_original} empresas"
                    # ── plain text (fallback) ──
                    _corpo50_stm = (
                        f"RWA Soluções — STM São Luís\n"
                        f"{'=' * 38}\n\n"
                        f"Status      : EM ANDAMENTO — 50% concluída\n"
                        f"Competência : {PERIODO}\n"
                        f"Progresso   : {len(resultados)} de {total_original} empresas processadas\n"
                        f"Tempo decorrido: {_tempo50_str_stm}\n\n"
                        f"Automação rodando com estabilidade.\n\n"
                        f"Metade da fila concluída. Você receberá um email quando a automação for concluída.\n\n"
                        f"RWA Soluções"
                    )
                    # ── HTML (visual) ──
                    _corpo50_html_stm = _render_html_email_rwa(
                        tag_texto="EM ANDAMENTO",
                        tag_bg="#BA7517", tag_fg="#FAEEDA",
                        titulo="50% concluída",
                        intro="Automação rodando com estabilidade.",
                        intro_destaque=True,
                        secoes=[
                            (None, [
                                ("Competência",     PERIODO,           False),
                                ("Progresso",       _progresso50_stm,  False),
                                ("Tempo decorrido", _tempo50_str_stm,  False),
                            ]),
                        ],
                        callout=(
                            "#FAEEDA", "#BA7517", "#633806",
                            "Metade da fila concluída. Você receberá um email quando a automação for concluída.",
                        ),
                    )
                    _enviar_email_rwa(
                        _rwa_email_conta, _rwa_email_senha, _rwa_email_destino,
                        f"RWA Soluções — STM São Luís 50% concluída — {PERIODO}",
                        _corpo50_html_stm, _corpo50_stm,
                    )
                    print("[EMAIL] Email de 50% enviado.")
            except Exception as _e50_stm:
                print(f"[EMAIL] Falha ao enviar email de 50%: {_e50_stm}")
    tempo_total = time.time() - tempo_inicio_total
    total_mins  = int(tempo_total // 60)
    total_segs  = int(tempo_total % 60)
    media_segs  = int(sum(tempos_empresa) / len(tempos_empresa)) if tempos_empresa else 0
    media_mins  = media_segs // 60
    media_s     = media_segs % 60

    ok        = sum(1 for r in resultados if r.get("pdf_prestado") == "PDF")
    sem_mov   = sum(1 for r in resultados if r.get("pdf_prestado") == "Sem movimento")
    erro      = sum(1 for r in resultados if r.get("pdf_prestado") == "Erro")
    incorreto = sum(1 for r in resultados if r.get("login_senha")  == "INCORRETO")
    reprocess = sum(1 for r in resultados
                    if "REPROCESSAR" in str(r.get("observacao", "")).upper())

    print("=" * 72)
    print(f"  RESUMO FINAL — STM Sao Luis v2.1  |  {PERIODO}")
    print("=" * 72)
    print(f"  Total empresas  : {len(resultados)}")
    print(f"  OK              : {ok}")
    print(f"  Sem movimento   : {sem_mov}")
    print(f"  Erro            : {erro}")
    print(f"  Login incorreto : {incorreto}")
    print(f"  Reprocessar     : {reprocess}")
    print(f"  Tempo total     : {total_mins}min{total_segs:02d}s")
    print(f"  Media/empresa   : {media_mins}min{media_s:02d}s")
    print("=" * 72)

    # Montar dicionario de tempo por codigo de empresa
    _tempos_por_empresa = {}
    for _le in _log_empresas:
        _cod = _le.get("codigo", "")
        _ts  = _le.get("total_s", 0)
        if _cod and _ts:
            _tempos_por_empresa[_cod] = _tempos_por_empresa.get(_cod, 0) + float(_ts or 0)

    gerar_relatorio(resultados)
    gerar_relatorio_conclusao(resultados, tempo_total_s=tempo_total, tempos_por_empresa=_tempos_por_empresa, total_original=total_original)
    finalizar_diagnostico()
    _imprimir_diagnostico_cmd(_log_empresas, _log_instabilidades)

    # ── Email de conclusão — disparado após todos os relatórios, antes do botão Concluído ──
    try:
        _rwa_email_destino  = os.environ.get("RWA_EMAIL_DESTINO",   "").strip()
        _rwa_email_conta    = os.environ.get("RWA_EMAIL_CONTA",     "").strip()
        _rwa_email_senha    = os.environ.get("RWA_EMAIL_SENHA_APP", "").strip()

        if _rwa_email_destino and _rwa_email_conta and _rwa_email_senha:
            _tempo_total_str_stm = f"{total_mins}min{total_segs:02d}s"
            _media_str_stm       = f"{media_mins}min{media_s:02d}s"

            _sep = "  " + "-" * 32
            # ── plain text (fallback) ──
            _corpo_fim_stm = (
                f"Prezado(a),\n\n"
                f"A automação STM São Luís foi concluída.\n\n"
                f"  Competência   : {PERIODO}\n"
                f"  Tempo total   : {_tempo_total_str_stm}\n"
                f"  Média/empresa : {_media_str_stm}\n\n"
                f"{_sep}\n"
                f"  Total de empresas   : {len(resultados)}\n"
                f"  Com movimento       : {ok}\n"
                f"  Sem movimento       : {sem_mov}\n"
                f"  Erro operacional    : {erro}\n"
                f"  Login incorreto     : {incorreto}\n"
                f"  Reprocessadas       : {reprocess}\n"
                f"{_sep}\n\n"
                f"Verifique os relatórios na pasta da competência.\n\n"
                f"RWA Soluções"
            )
            # ── HTML (visual) ──
            _corpo_fim_html_stm = _render_html_email_rwa(
                tag_texto="CONCLUÍDA",
                tag_bg="#3B6D11", tag_fg="#EAF3DE",
                titulo="Automação concluída",
                intro="Prezado(a), a automação foi concluída.",
                secoes=[
                    ("Tempos", [
                        ("Competência",     PERIODO,              False),
                        ("Tempo total",     _tempo_total_str_stm, False),
                        ("Média / empresa", _media_str_stm,       False),
                    ]),
                    ("Resultado", [
                        ("Total de empresas", str(len(resultados)), False),
                        ("Com movimento",     str(ok),              False),
                        ("Sem movimento",     str(sem_mov),         False),
                        ("Erro operacional",  str(erro),            erro      > 0),
                        ("Login incorreto",   str(incorreto),       incorreto > 0),
                        ("Reprocessadas",     str(reprocess),       reprocess > 0),
                    ]),
                ],
                callout=(
                    "#EAF3DE", "#639922", "#173404",
                    "Verifique os relatórios na pasta da competência.",
                ),
            )
            _enviar_email_rwa(
                _rwa_email_conta, _rwa_email_senha, _rwa_email_destino,
                f"RWA Soluções — STM São Luís concluída — {PERIODO}",
                _corpo_fim_html_stm, _corpo_fim_stm,
            )
            print("[EMAIL] Email de conclusao enviado.")
    except Exception as _e_mail_fim_stm:
        print(f"[EMAIL] Falha ao enviar email de conclusao: {_e_mail_fim_stm}")

    print("\n[CONCLUIDO] Processo finalizado com sucesso.")
    _sys.stdout = _stdout_original
    _pw.finalizar()


def _imprimir_diagnostico_cmd(log_empresas, log_instabilidades):
    """Imprime diagnostico completo no CMD - por empresa e resumo de instabilidades."""

    def _fmt_s(val):
        if val == "" or val is None:
            return "   - "
        try:
            s = float(val)
            m = int(s // 60)
            s2 = int(s % 60)
            return f"{m}m{s2:02d}s" if m > 0 else f"  {s2}s"
        except Exception:
            return str(val)

    def _resultado(r):
        """Determina o resultado final legivel da empresa."""
        status_login = r.get("status_login", "")
        pdf_prest    = r.get("pdf_prestado", "")
        talao_nf     = r.get("talao_nf", "")
        talao_guia   = r.get("talao_guia", "")
        obs          = r.get("observacao", "")

        if status_login == "INCORRETO":
            return "LOGIN INCORRETO"
        if pdf_prest == "Sem movimento":
            return "SEM MOVIMENTO"
        if "instavel" in obs.lower() or "reprocessar" in obs.lower():
            return "INSTABILIDADE"

        nf   = "NF:OK"   if talao_nf   == "PDF" else ("NF:ERR" if talao_nf == "Erro" else "NF:-")
        guia = "ISS:OK"  if talao_guia == "PDF" else ("ISS:DISP" if talao_guia in ("DISPENSADO", "Recolhimento via PGDAS") else ("ISS:ERR" if talao_guia == "Erro" else "ISS:-"))

        if pdf_prest == "PDF":
            return f"{nf} {guia}"
        return "ERRO"

    L = 110
    print()
    print("=" * L)
    print(f"  DIAGNOSTICO DE EXECUCAO --- STM Sao Luis")
    print("=" * L)
    print(f"  {'EMPRESA':<35} {'T':>2}  {'TOTAL':>6}  {'LOGIN':>5}  {'RELAT':>6}  {'DOWN':>5}  {'RESULTADO':<18}  OBSERVACAO")
    print("-" * L)

    reprocessadas = {}

    for r in log_empresas:
        cod       = r.get("codigo", "")
        emp       = r.get("empresa", "")[:30]
        label     = f"{cod}-{emp}"
        tent      = r.get("tentativa", 1)
        total     = _fmt_s(r.get("total_s"))
        login     = _fmt_s(r.get("login_s"))
        relat     = _fmt_s(r.get("relatorio_s"))
        down      = _fmt_s(r.get("download_s"))
        resultado = _resultado(r)
        obs       = r.get("observacao", "")

        if tent > 1:
            prefixo = f"  {'  >> REPROCESS ':<35}"
            reprocessadas.setdefault(cod, []).append(r)
        else:
            prefixo = f"  {label:<35}"

        print(f"{prefixo} {tent:>2}  {total:>6}  {login:>5}  {relat:>6}  {down:>5}  {resultado:<18}  {obs}")

    print("=" * L)

    if reprocessadas:
        print(f"\n  REPROCESSAMENTOS ({len(reprocessadas)} empresa(s)):")
        print("-" * L)
        for cod, tentativas in reprocessadas.items():
            emp_nome = tentativas[0].get("empresa", "")
            print(f"  {cod} - {emp_nome}")
            for t in tentativas:
                res = _resultado(t)
                obs = t.get("observacao", "")
                print(f"    Tentativa {t.get('tentativa')}: {_fmt_s(t.get('total_s'))}  |  {res}  |  {obs}")
        print("-" * L)

    if log_instabilidades:
        print(f"\n  INSTABILIDADES DETECTADAS ({len(log_instabilidades)}):")
        print("-" * L)
        for inst in log_instabilidades:
            print(f"  {inst['horario']}  |  {inst['empresa'][:35]:<35}  |  Tent {inst['tentativa']}  |  Fase: {inst['fase']:<12}  |  {inst['duracao_s']}s decorridos")
        print("-" * L)
    else:
        print(f"\n  INSTABILIDADES: nenhuma detectada.")

    total_emp  = len([r for r in log_empresas if r.get("tentativa") == 1])
    ok_count   = sum(1 for r in log_empresas if r.get("pdf_prestado") == "PDF"            and r.get("tentativa") == 1)
    sem_mov    = sum(1 for r in log_empresas if r.get("pdf_prestado") == "Sem movimento"   and r.get("tentativa") == 1)
    erro_count = sum(1 for r in log_empresas if r.get("pdf_prestado") == "Erro"           and r.get("tentativa") == 1)
    incorreto  = sum(1 for r in log_empresas if r.get("status_login") == "INCORRETO"      and r.get("tentativa") == 1)

    print()
    print(f"  Total processadas : {total_emp}")
    print(f"  OK (PDF baixado)  : {ok_count}")
    print(f"  Sem movimento     : {sem_mov}")
    print(f"  Erro              : {erro_count}")
    print(f"  Login incorreto   : {incorreto}")
    print(f"  Reprocessamentos  : {len(reprocessadas)}")
    print(f"  Instabilidades    : {len(log_instabilidades)}")
    print()
    print("=" * L)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[INTERROMPIDO] Execucao cancelada pelo usuario.")
        print("[DIAGNOSTICO] O log de execucao foi registrado durante o processo.")
    except Exception as _e_critico_inicial:
        # ❌ ERRO CRÍTICO NO INÍCIO (antes do loop de empresas)
        _msg_erro_critico = f"ERRO CRÍTICO INICIAL: {str(_e_critico_inicial)}\n\n{traceback.format_exc()}"
        print(f"\n❌ {_msg_erro_critico}")
        
        # Registra no DIAGNOSTICO
        try:
            with open(caminho_diagnostico_erros(), "a", encoding="utf-8") as _fd:
                _fd.write("\n" + "=" * 90 + "\n")
                _fd.write(f"Data/Hora  : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                _fd.write(_msg_erro_critico + "\n")
                _fd.write("=" * 90 + "\n")
        except Exception:
            pass
        
        # Manda EMAIL DE ERRO IMEDIATAMENTE
        try:
            _rwa_email_destino  = os.environ.get("RWA_EMAIL_DESTINO",   "").strip()
            _rwa_email_agendado = os.environ.get("RWA_EMAIL_AGENDADO",  "0") == "1"
            _rwa_email_conta    = os.environ.get("RWA_EMAIL_CONTA",     "").strip()
            _rwa_email_senha    = os.environ.get("RWA_EMAIL_SENHA_APP", "").strip()
            
            if _rwa_email_destino and _rwa_email_agendado and _rwa_email_conta and _rwa_email_senha:
                _erro_dt = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                
                # Plain text
                _corpo_erro = (
                    f"Prezado(a),\n\n"
                    f"❌ ERRO CRÍTICO — Automação não pôde ser iniciada.\n\n"
                    f"  Competência   : {PERIODO}\n"
                    f"  Horário       : {_erro_dt}\n"
                    f"  Erro          : {str(_e_critico_inicial)}\n\n"
                    f"Verifique os logs de diagnóstico para mais detalhes.\n\n"
                    f"RWA Soluções"
                )
                
                # HTML
                _corpo_erro_html = _render_html_email_rwa(
                    tag_texto="❌ ERRO",
                    tag_bg="#DC2626", tag_fg="#FEE2E2",
                    titulo="Erro crítico na automação",
                    intro="Prezado(a), a automação não pôde ser iniciada.",
                    secoes=[
                        (None, [
                            ("Competência",  PERIODO,                  False),
                            ("Horário",      _erro_dt,                 False),
                            ("Erro",         str(_e_critico_inicial),  True),
                        ]),
                    ],
                    callout=(
                        "#FEE2E2", "#DC2626", "#991B1B",
                        "Verifique os logs de diagnóstico para mais detalhes.",
                    ),
                )
                
                import smtplib as _smtp_h
                from email.mime.multipart import MIMEMultipart as _MMulti_h
                from email.mime.text import MIMEText as _MText_h
                
                _msg = _MMulti_h("alternative")
                _msg["From"]    = _rwa_email_conta
                _msg["To"]      = _rwa_email_destino
                _msg["Subject"] = f"❌ RWA Soluções — ERRO CRÍTICO — {PERIODO}"
                _msg.attach(_MText_h(_corpo_erro, "plain", "utf-8"))
                _msg.attach(_MText_h(_corpo_erro_html, "html", "utf-8"))
                
                _srv = _smtp_h.SMTP("smtp.gmail.com", 587)
                _srv.starttls()
                _srv.login(_rwa_email_conta, _rwa_email_senha)
                _srv.sendmail(_rwa_email_conta, _rwa_email_destino, _msg.as_string())
                _srv.quit()
                
                print("[EMAIL] ❌ Email de ERRO CRÍTICO enviado.")
        except Exception as _e_email_erro:
            print(f"[EMAIL] Falha ao enviar email de erro: {_e_email_erro}")
        
        import sys as _sys_exit
        _sys_exit.exit(1)