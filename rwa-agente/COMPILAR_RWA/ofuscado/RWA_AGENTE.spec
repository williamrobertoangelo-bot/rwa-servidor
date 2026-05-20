# -*- mode: python ; coding: utf-8 -*-
# RWA Tecnologia Operacional — RWA_AGENTE.spec
# PyInstaller 6.x
# Executar de dentro de: C:\rwa-servidor\rwa-agente\COMPILAR_RWA\ofuscado\
#   pyinstaller RWA_AGENTE.spec
# O .exe será gerado em: COMPILAR_RWA\exe\

a = Analysis(
    ['RWA_AGENTE.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('pyarmor_runtime_012235', 'pyarmor_runtime_012235'),
    ],
    hiddenimports=[
        'uuid',
        'socket',
        'hashlib',
        'getpass',
        'platform',
        'threading',
        'subprocess',
        'logging',
        'logging.handlers',
        'json',
        'time',
        'pathlib',
        'datetime',
        'urllib',
        'urllib.request',
        'urllib.error',
        'urllib.parse',
        'pystray',
        'pystray._win32',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
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
    name='RWA_AGENTE',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    distpath='../exe',
    # icon='../../icone_rwa.ico',
)
