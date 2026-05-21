import time
import pyautogui
import os
from datetime import datetime

PASTA_RWA = r"C:\\RWA_AUTOMACOES"
os.makedirs(PASTA_RWA, exist_ok=True)

# Coordenadas de referencia baseadas no script 0505_15 (AJUSTE V7: 04/05/2026)
# LISTA_X = 768 | LINHA_Y = {1: 256, 2: 305, 3: 350} | OK_X = 1108 | OK_Y = 443
REF_OCR_LEFT   = 480
REF_OCR_TOP    = 240
REF_OCR_RIGHT  = 965
REF_OCR_BOTTOM = 377

TOLERANCIA = 15


def avaliar(capturado, referencia, eixo):
    diff = capturado - referencia
    if abs(diff) <= TOLERANCIA:
        return f"OK (diferenca {diff:+d}px)", True
    if eixo == "X":
        if diff < 0:
            return f"MAIS PARA A DIREITA ({abs(diff)}px faltando)", False
        else:
            return f"MAIS PARA A ESQUERDA ({abs(diff)}px sobrando)", False
    if eixo == "Y":
        if diff < 0:
            return f"MAIS PARA BAIXO ({abs(diff)}px faltando)", False
        else:
            return f"MAIS PARA CIMA ({abs(diff)}px sobrando)", False


def capturar(mensagem, segundos=2):
    print(f"\n>>> {mensagem}")
    for i in range(segundos, 0, -1):
        print(f"    {i}...", end="\r")
        time.sleep(1)
    x, y = pyautogui.position()
    print(f"    CAPTURADO! X={x}, Y={y}          ")
    return x, y


def capturar_ate_ok(titulo, instrucao, ref_x, ref_y, eixo_x, eixo_y):
    """Fica em loop ate o usuario atingir tolerancia aceitavel nos dois eixos."""
    tentativa = 1
    while True:
        print(f"""
============================================================
  {titulo}  [tentativa {tentativa}]
============================================================
{instrucao}
  >> Posicione o mouse e pressione ENTER
""")
        input()
        x, y = capturar(f"Aguardando 2s", segundos=2)

        msg_x, ok_x = avaliar(x, ref_x, eixo_x)
        msg_y, ok_y = avaliar(y, ref_y, eixo_y)

        print(f"""
  RESULTADO:
    X : capturado={x}  |  referencia={ref_x}  =>  {msg_x}
    Y : capturado={y}  |  referencia={ref_y}  =>  {msg_y}
""")

        if ok_x and ok_y:
            print("  >> APROVADO! Avancando para o proximo passo.\n")
            return x, y
        else:
            print("  >> Fora da tolerancia. Tente novamente.\n")
            tentativa += 1


print("=" * 60)
print("  CAPTURA DE COORDENADAS — CERTIFICADO DIGITAL  v3")
print("=" * 60)
print("""
INSTRUCOES:
- Abra o Chrome e clique em 'Acesso com Certificado'
- O popup vai abrir com a lista de certificados
- Para cada passo: posicione o mouse e pressione ENTER
- Se errar a tolerancia, o script pede de novo automaticamente
- So avanca quando estiver OK

Pressione ENTER para iniciar...
""")
input()

# PASSOS 1-4: linhas e botao OK (sem loop, apenas captura)
print("""
============================================================
  PASSO 1: LINHA 1 — primeiro certificado da lista
============================================================
  >> Posicione o mouse no centro do texto da LINHA 1 e pressione ENTER
""")
input()
x1, y1 = capturar("Aguardando 2s — LINHA 1")

print("""
============================================================
  PASSO 2: LINHA 2 — segundo certificado da lista
============================================================
  >> Posicione o mouse no centro do texto da LINHA 2 e pressione ENTER
""")
input()
x2, y2 = capturar("Aguardando 2s — LINHA 2")

print("""
============================================================
  PASSO 3: LINHA 3 — terceiro certificado da lista
============================================================
  >> Posicione o mouse no centro do texto da LINHA 3 e pressione ENTER
""")
input()
x3, y3 = capturar("Aguardando 2s — LINHA 3")

print("""
============================================================
  PASSO 4: BOTAO OK
============================================================
  >> Posicione o mouse sobre o botao OK e pressione ENTER
""")
input()
xok, yok = capturar("Aguardando 2s — BOTAO OK")

# PASSOS 5-8: 4 cantos OCR — captura simples, sem tolerancia
print("""
============================================================
  PASSO 5: OCR — CANTO SUPERIOR ESQUERDO
============================================================
  Inicio da coluna 'Tema', um pouco abaixo do cabecalho.
  Nao pegue o titulo nem a borda externa.
  >> Posicione o mouse e pressione ENTER
""")
input()
ocr_left, ocr_top = capturar("Aguardando 2s — CANTO SUPERIOR ESQUERDO")

print("""
============================================================
  PASSO 6: OCR — CANTO SUPERIOR DIREITO
============================================================
  Final da coluna 'Serial', antes da barra de rolagem.
  >> Posicione o mouse e pressione ENTER
""")
input()
ocr_right_top, ocr_top_2 = capturar("Aguardando 2s — CANTO SUPERIOR DIREITO")

print("""
============================================================
  PASSO 7: OCR — CANTO INFERIOR ESQUERDO
============================================================
  Abaixo da ultima linha de certificado visivel.
  Nao pegue os botoes OK ou Cancelar.
  >> Posicione o mouse e pressione ENTER
""")
input()
ocr_left_2, ocr_bottom_left = capturar("Aguardando 2s — CANTO INFERIOR ESQUERDO")

print("""
============================================================
  PASSO 8: OCR — CANTO INFERIOR DIREITO
============================================================
  Abaixo da ultima linha, antes da barra de rolagem.
  >> Posicione o mouse e pressione ENTER
""")
input()
ocr_right, ocr_bottom = capturar("Aguardando 2s — CANTO INFERIOR DIREITO")

# Normaliza retangulo OCR
OCR_LEFT   = min(ocr_left, ocr_left_2)
OCR_TOP    = min(ocr_top, ocr_top_2)
OCR_RIGHT  = max(ocr_right_top, ocr_right)
OCR_BOTTOM = max(ocr_bottom_left, ocr_bottom)

# Salva resultado no formato exato do COORDENADAS OCR.txt
caminho = os.path.join(PASTA_RWA, "COORDENADAS OCR.txt")
linhas = []
linhas.append("========================================================")
linhas.append("CONFIGURACAO DE COORDENADAS - CERTIFICADO DIGITAL")
linhas.append("========================================================")
linhas.append("")
linhas.append("IMPORTANTE:")
linhas.append("Nao altere os nomes.")
linhas.append("Altere apenas os numeros.")
linhas.append("")
linhas.append("========================================================")
linhas.append("LISTA DE CERTIFICADOS")
linhas.append("========================================================")
linhas.append("")
linhas.append(f"LINHA_1_X={x1}")
linhas.append(f"LINHA_1_Y={y1}")
linhas.append("")
linhas.append(f"LINHA_2_X={x2}")
linhas.append(f"LINHA_2_Y={y2}")
linhas.append("")
linhas.append(f"LINHA_3_X={x3}")
linhas.append(f"LINHA_3_Y={y3}")
linhas.append("")
linhas.append("========================================================")
linhas.append("BOTAO OK")
linhas.append("========================================================")
linhas.append("")
linhas.append(f"BOTAO_OK_X={xok}")
linhas.append(f"BOTAO_OK_Y={yok}")
linhas.append("")
linhas.append("========================================================")
linhas.append("AREA DE LEITURA OCR")
linhas.append("========================================================")
linhas.append("")
linhas.append(f"SUPERIOR_ESQUERDO_X={ocr_left}")
linhas.append(f"SUPERIOR_ESQUERDO_Y={ocr_top}")
linhas.append("")
linhas.append(f"SUPERIOR_DIREITO_X={ocr_right_top}")
linhas.append(f"SUPERIOR_DIREITO_Y={ocr_top_2}")
linhas.append("")
linhas.append(f"INFERIOR_ESQUERDO_X={ocr_left_2}")
linhas.append(f"INFERIOR_ESQUERDO_Y={ocr_bottom_left}")
linhas.append("")
linhas.append(f"INFERIOR_DIREITO_X={ocr_right}")
linhas.append(f"INFERIOR_DIREITO_Y={ocr_bottom}")
linhas.append("")
linhas.append("========================================================")
linhas.append("CONFIGURACOES AVANCADAS")
linhas.append("========================================================")
linhas.append("")
linhas.append("MAXIMO_ROLAGENS=150")

with open(caminho, "w", encoding="utf-8") as f:
    f.write("\n".join(linhas))

print("\n" + "=" * 60)
print("  CONCLUIDO!")
print("=" * 60)
for linha in linhas:
    print(linha)

print(f"\nArquivo salvo em: {caminho}")
input("\nPressione ENTER para fechar...")
