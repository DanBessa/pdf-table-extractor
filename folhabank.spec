# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# ==========================================
# 📦 COLETA DE DADOS E DEPENDÊNCIAS
# ==========================================
datas = []
binaries = []
hiddenimports = [
    # Bibliotecas essenciais
    'customtkinter',
    'cryptography',
    'pdfplumber',
    'pypdf',
    'openpyxl',
    'pandas',
    'PIL',
    'PIL.Image',
    'requests',
    'packaging',
    
    # Módulos de conversores bancários dinâmicos
    'conversor_inter',
    'conversor_bradesco',
    'conversor_pagbank',
    'conversor_cefmod1',
    'conversor_cefmod2',
    'conversor_c6',
    'conversor_banestes',
    'conversor_paycash',
    'conversor_sicredi',
    'conversor_stone',
    'conversor_mp',
    'conversor_bbmod1',
    'conversor_bbmod2',
    'conversor_sicoobmod1',
    'conversor_sicoobmod2',
    'conversor_sicoobmod3',
    'conversor_santandermod1',
    'conversor_santandermod2',
    'conversor_saframod1',
    'conversor_saframod2',
    'conversor_itaumod1',
    'conversor_itaumod2',
    'conversor_itaumod3',
]

# 1. Coleta completa de temas e fontes do CustomTkinter
tmp_ctk = collect_all('customtkinter')
datas += tmp_ctk[0]
binaries += tmp_ctk[1]
hiddenimports += tmp_ctk[2]

# 2. Coleta FORÇADA de toda a biblioteca charset_normalizer (Resolve o erro MD)
tmp_charset = collect_all('charset_normalizer')
datas += tmp_charset[0]
binaries += tmp_charset[1]
hiddenimports += tmp_charset[2]

# 3. Empacotamento das pastas de assets (se existirem)
if os.path.exists('icons'):
    datas.append(('icons', 'icons'))

if os.path.exists('conversores'):
    datas.append(('conversores', 'conversores'))

# ==========================================
# ⚙️ ANÁLISE DO PROJETO
# ==========================================
a = Analysis(
    ['folhabank.py'],  # Nome do seu script principal
    pathex=['.', 'conversores'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ==========================================
# 🚀 GERAÇÃO DO EXECUTÁVEL (ONEFILE & WINDOWED)
# ==========================================
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='FolhaBank Integrador',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False = Oculta a janela preta do terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icons/convert.ico'
)