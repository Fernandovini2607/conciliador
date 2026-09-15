# PyInstaller spec — build do Conciliador para distribuição
# Rodar: pyinstaller Conciliador.spec --noconfirm --clean
#
# Gera dist/Conciliador/ com Conciliador.exe + libs + pasta data/ com
# db_config.json (aponta pro servidor).

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Módulos que às vezes o PyInstaller não detecta sozinho (imports dinâmicos)
hidden_imports = [
    "pymysql",
    "pymysql.cursors",
    "pymysql.constants",
    "openpyxl",
    "ofxparse",
    "pyodbc",
    "reportlab",
    "reportlab.pdfbase",
    "reportlab.pdfbase.ttfonts",
    "auth",
    "config",
    "db",
    "dialogos_dominio",
    "dialogos_login",
    "dialogos_taxas",
    "exportar_xlsx",
    "gerar_pdf",
    "lancamentos",
    "matcher",
    "parser_dominio",
    "parser_ofx",
    "parser_pdf",
    "parser_xlsx",
    "pdfplumber",
    "pdfminer",
    "pdfminer.high_level",
    "pdfminer.pdfparser",
]

# Arquivos de dados a incluir no build
# Formato: (origem_no_projeto, destino_relativo_no_build)
datas = [
    ("data/db_config.json", "data"),
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "test",
        "unittest",
        "pytest",
        "pip",
        "setuptools",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Conciliador",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # False = sem janela preta atrás do app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Conciliador",
)
