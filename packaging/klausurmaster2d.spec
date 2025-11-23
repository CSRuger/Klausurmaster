# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


spec_dir = Path(__file__).resolve().parent
project_root = spec_dir.parent
assets_dir = project_root / 'assets'
resource_files = [
    (str(assets_dir / 'card_app_data.json'), 'assets'),
    (str(assets_dir / 'config.json'), 'assets'),
    (str(assets_dir / 'favicon.ico'), 'assets'),
]


a = Analysis(
    [str(project_root / 'main.py')],
    pathex=[str(project_root)],
    binaries=[],
    datas=resource_files,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='klausurmaster2d',
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
    icon=[str(assets_dir / 'favicon.ico')],
)
