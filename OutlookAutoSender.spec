# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

qt_datas, qt_binaries, qt_hiddenimports = collect_all('PyQt5')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=qt_binaries,
    datas=[
        ('config.yaml', '.'),
        ('templates.yaml', '.'),
        ('resources', 'resources'),
        *qt_datas,
    ],
    hiddenimports=[
        *qt_hiddenimports,
        'PyQt5.sip',
        'win32com.client',
        'openpyxl',
        'yaml',
        'jinja2',
        'apscheduler',
    ],
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
    name='OutlookAutoSender',
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
)
