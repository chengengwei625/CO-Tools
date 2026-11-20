# -*- mode: python ; coding: utf-8 -*-
# 出货表导出工具 — 独立测试版 exe（第四功能）

_EXCLUDES = [
    'scipy', 'pyarrow', 'cryptography', 'lxml',
    'pandas', 'pandas.tests', 'pandas.io.formats.style',
    'jinja2', 'markupsafe',
    'xlrd', 'xlwt', 'xlsxwriter', 'pyxlsb',
    'bs4', 'html5lib',
    'numexpr', 'bottleneck',
    'tables', 'sqlalchemy',
    'numpy', 'numpy.testing', 'numpy.distutils', 'numpy.f2py',
    'unittest', 'pytest', '_pytest',
    'pydoc', 'doctest', 'lib2to3',
    'distutils', 'setuptools', 'pkg_resources',
    'tkinter.test',
    'xmlrpc',
    'ftplib', 'imaplib', 'poplib', 'smtplib', 'telnetlib',
    'reportlab',
    'PIL',
]

a = Analysis(
    ['shipping_table_app.py'],
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
    name='出货表导出工具',
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
