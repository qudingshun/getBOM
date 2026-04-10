# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\qudin\\Desktop\\getBOM\\OA_Plan_BOM2_portable.py'],
    pathex=[],
    binaries=[('C:\\Users\\qudin\\Desktop\\getBOM\\chromedriver.exe', '.')],
    datas=[('C:\\Users\\qudin\\Desktop\\getBOM\\portable_chrome', 'portable_chrome')],
    hiddenimports=['selenium.webdriver.chrome.webdriver', 'selenium.webdriver.chromium.webdriver'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pyautogui', 'pyperclip'],
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
    name='OA_Plan_BOM2_portable_GUI',
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
