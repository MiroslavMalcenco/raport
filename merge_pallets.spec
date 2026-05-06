# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path


# PyInstaller defines SPECPATH (dir containing this .spec).
project_dir = Path(SPECPATH).resolve()

a = Analysis(
    ['merge_pallets.py'],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# На macOS оставляем .app (onedir), а на Windows собираем onefile exe
# без дополнительной папки рядом.
if sys.platform == 'darwin':
    exe = EXE(
        pyz,
        a.scripts,
        [],
        [],
        [],
        name='merge_pallets',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        exclude_binaries=True,
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
        upx=True,
        upx_exclude=[],
        name='merge_pallets',
    )

    app = BUNDLE(
        coll,
        name='merge_pallets.app',
        icon=None,
        bundle_identifier=None,
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name='merge_pallets',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        exclude_binaries=False,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
