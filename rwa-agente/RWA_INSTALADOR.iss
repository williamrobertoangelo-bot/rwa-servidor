; =====================================================================
; RWA TECNOLOGIA OPERACIONAL — INSTALADOR PROFISSIONAL
; Inno Setup 6.7.2
; =====================================================================

[Setup]
AppName=RWA Tecnologia Operacional
AppVersion=2.0
AppPublisher=RWA Tecnologia Operacional
AppPublisherURL=https://rwasolucoes.com.br
DefaultDirName={commonpf64}\RWA_AUTOMACOES
DefaultGroupName=RWA Tecnologia Operacional
OutputDir=C:\rwa-servidor\rwa-agente\COMPILAR_RWA
OutputBaseFilename=RWA_Setup_v2.0
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
DisableProgramGroupPage=yes
UninstallDisplayName=RWA Tecnologia Operacional
ShowLanguageDialog=no
LanguageDetectionMethod=none

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Messages]
WelcomeLabel1=Bem-vindo ao instalador do RWA Tecnologia Operacional
WelcomeLabel2=Este instalador vai preparar sua máquina para executar as automações RWA.%n%nClique em Avançar para continuar.
FinishedLabel=A instalação foi concluída com sucesso.%n%nO agente RWA foi iniciado e está rodando em segundo plano.%n%nAcesse o portal para começar: https://rwasolucoes.com.br

[Files]
; ── Agente e Launcher compilados ───────────────────────────────────
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\exe\RWA_AGENTE.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\exe\RWA_LAUNCHER.exe"; DestDir: "{app}"; Flags: ignoreversion

; ── Módulos de automação ofuscados ────────────────────────────────
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\ofuscado\Sao_Luis.PY"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\ofuscado\Padrao_nacional.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\ofuscado\conferencias_sao_luis.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\ofuscado\conferencias_pn.py"; DestDir: "{app}"; Flags: ignoreversion

; ── PyArmor Runtime ────────────────────────────────────────────────
Source: "C:\rwa-servidor\rwa-agente\COMPILAR_RWA\ofuscado\pyarmor_runtime_012235\*"; DestDir: "{app}\pyarmor_runtime_012235"; Flags: ignoreversion recursesubdirs createallsubdirs

; ── Arquivos de configuração ───────────────────────────────────────
Source: "C:\rwa-servidor\rwa-agente\COORDENADAS OCR.txt"; DestDir: "{app}"; Flags: ignoreversion

; ── Python embutido ────────────────────────────────────────────────
Source: "C:\Users\willi\AppData\Local\Python\pythoncore-3.14-64\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs

; ── Tesseract OCR ──────────────────────────────────────────────────
Source: "C:\Program Files\Tesseract-OCR\*"; DestDir: "{app}\Tesseract-OCR"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{app}\config"
Name: "{app}\logs"

[Registry]
; ── Protocolo rwa:// ───────────────────────────────────────────────
Root: HKCR; Subkey: "rwa"; ValueType: string; ValueName: ""; ValueData: "URL:RWA Protocol"; Flags: uninsdeletekey
Root: HKCR; Subkey: "rwa"; ValueType: string; ValueName: "URL Protocol"; ValueData: ""
Root: HKCR; Subkey: "rwa\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\RWA_LAUNCHER.exe"" ""%1"""

; ── Agente no startup do Windows ──────────────────────────────────
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "RWA_AGENTE"; ValueData: """{app}\RWA_AGENTE.exe"""; Flags: uninsdeletevalue

; ── Caminho do Tesseract para o pytesseract ───────────────────────
Root: HKLM; Subkey: "SOFTWARE\RWA_AUTOMACOES"; ValueType: string; ValueName: "TesseractPath"; ValueData: "{app}\Tesseract-OCR\tesseract.exe"; Flags: uninsdeletekey

[Run]
; ── Abre o launcher ao final da instalação ─────────────────
Filename: "{app}\RWA_LAUNCHER.exe"; Parameters: "--primeiro-acesso"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent; Description: "Configurar acesso RWA"

[UninstallRun]
; ── Para o agente antes de desinstalar ────────────────────────────
Filename: "taskkill"; Parameters: "/F /IM RWA_AGENTE.exe"; Flags: runhidden

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Agente iniciado automaticamente pela seção [Run]
  end;
end;
