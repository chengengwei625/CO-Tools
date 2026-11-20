# -*- mode: python ; coding: utf-8 -*-
# 装箱单生成工具 — 独立测试版 exe（确认后并入主程序）

_EXCLUDES = [
    'scipy', 'pyarrow', 'cryptography', 'lxml',
    # PIL is a hard dependency of reportlab.lib.utils — do NOT exclude
    'pandas.tests',
    'pandas.io.formats.style',
    'jinja2', 'markupsafe',
    'xlrd', 'xlwt', 'xlsxwriter', 'pyxlsb',
    'bs4', 'html5lib',
    'numexpr', 'bottleneck',
    'tables', 'sqlalchemy',
    'numpy.testing', 'numpy.distutils', 'numpy.f2py',
    'unittest', 'pytest', '_pytest',
    'pydoc', 'doctest', 'lib2to3',
    'distutils', 'setuptools', 'pkg_resources',
    'tkinter.test',
    'xmlrpc',
    'ftplib', 'imaplib', 'poplib', 'smtplib', 'telnetlib',
    'reportlab.graphics.renderSVG',
    'reportlab.graphics.renderPS',
]

a = Analysis(
    ['packing_list_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl.cell._writer'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=_EXCLUDES,
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
    name='装箱单生成工具',
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
