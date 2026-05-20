# -*- mode: python ; coding: utf-8 -*-
# RWA Tecnologia Operacional — RWA_LAUNCHER.spec
# Gerado para PyInstaller 6.x
# Coloque este arquivo na mesma pasta que RWA_LAUNCHER.py e execute:
#   pyinstaller RWA_LAUNCHER.spec

a = Analysis(
    ['RWA_LAUNCHER.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # stdlib — detectados via imports do script
        'uuid',
        'socket',
        'hashlib',
        'getpass',
        'platform',
        'threading',
        'subprocess',
        'json',
        'time',
        'pathlib',
        'datetime',
        'urllib',
        'urllib.request',
        'urllib.error',
        'urllib.parse',
        # tkinter e customtkinter
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'customtkinter',
        # pystray e dependências
        'pystray',
        'pystray._win32',
        # PIL / Pillow
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        # encodings necessários no Windows
        'encodings',
        'encodings.utf_8',
        'encodings.cp1252',
        'encodings.latin_1',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RWA_LAUNCHER',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # sem janela de console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icone_rwa.ico',  # descomente e ajuste o caminho se tiver ícone
)
