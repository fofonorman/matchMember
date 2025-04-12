# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['match.py'],
    pathex=[],
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Match',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=True,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['MyIcon.icns'],
)
app = BUNDLE(
    exe,
    name='Match.app',
    icon='MyIcon.icns',
    bundle_identifier='com.tingpan.matchmember',
    info_plist={
        'NSHighResolutionCapable': 'True',
        'LSBackgroundOnly': 'False',
        'LSUIElement': '0',
        'CFBundleDisplayName': 'Match',
        'CFBundleShortVersionString': '1.0.0',
        'NSRequiresAquaSystemAppearance': 'False',
        'NSAppleScriptEnabled': 'False',
    },
)
