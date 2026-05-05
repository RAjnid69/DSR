# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['dsr_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/*.png', 'assets'),
        ('assets/*.xlsx', 'assets')
    ],
    hiddenimports=['xlsxwriter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PySide6', 'tensorflow', 'torch', 'scipy', 'matplotlib', 'notebook', 'jupyter', 'sympy'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Agro_DSR_Pro',
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
    icon='assets/xtreme.png'
)
