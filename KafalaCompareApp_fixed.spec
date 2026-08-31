# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['kafala_compare_app.py'],
    pathex=[],
    datas=[
        ('auto_update_gui.py', '.'),
        ('auto_update_from_diff.js', '.'),
        ('package.json', '.'),
        ('package-lock.json', '.'),
        ('KafalaCompareApp_build/node.exe', '.'),
        ('KafalaCompareApp_build/node_modules', 'node_modules')
    ],
    hiddenimports=['openpyxl', 'xlrd', 'tkinter.scrolledtext'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['pyi_cleanup_child_processes.py'],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='KafalaCompareApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
