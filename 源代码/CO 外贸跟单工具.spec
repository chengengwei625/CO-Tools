# -*- mode: python ; coding: utf-8 -*-
# 集杰跟单工具 — 合并多功能 exe

from PyInstaller.utils.hooks import collect_data_files, collect_submodules
from PyInstaller.utils.win32.versioninfo import (
    VSVersionInfo, FixedFileInfo, StringFileInfo, StringTable, StringStruct,
    VarFileInfo, VarStruct,
)
from pathlib import Path
import runpy

release = runpy.run_path(str(Path(SPECPATH) / 'app_version.py'))
version_tuple = tuple(map(int, release['APP_VERSION'].split('.'))) + (0,)
version_info = VSVersionInfo(
    ffi=FixedFileInfo(filevers=version_tuple, prodvers=version_tuple,
                     mask=0x3f, flags=0, OS=0x40004, fileType=1,
                     subtype=0, date=(0, 0)),
    kids=[StringFileInfo([StringTable('080404B0', [
        StringStruct('FileDescription', release['APP_NAME']),
        StringStruct('FileVersion', release['APP_VERSION']),
        StringStruct('ProductName', release['APP_NAME']),
        StringStruct('ProductVersion', release['APP_VERSION']),
        StringStruct('OriginalFilename', release['APP_TITLE'] + '.exe'),
    ])]), VarFileInfo([VarStruct('Translation', [2052, 1200])])],
)

_EXCLUDES = [
    # ── Large unused packages ─────────────────────────────────────
    'scipy',
    'pyarrow',
    'cryptography',
    # lxml is required by python-docx (③ 断针记录) — do NOT exclude
    # PIL is a hard dependency of reportlab — do NOT exclude
    # fitz (PyMuPDF) is required for ⑥ 装箱贴纸 — do NOT exclude

    # ── Unused pandas optional back-ends ──────────────────────────
    'pandas.tests',
    'pandas.io.formats.style',
    'jinja2', 'markupsafe',
    'xlwt', 'xlsxwriter', 'pyxlsb',
    'bs4', 'html5lib',
    'numexpr', 'bottleneck',
    'tables', 'sqlalchemy',
    # All inputs are local files; pandas cloud-storage backends are unused.
    'boto3', 'botocore', 'fsspec', 's3fs',

    # ── Unused numpy extras ────────────────────────────────────────
    'numpy.testing', 'numpy.distutils', 'numpy.f2py',

    # ── Test / development bloat ───────────────────────────────────
    'unittest', 'pytest', '_pytest',
    'pydoc', 'doctest', 'lib2to3',
    'distutils', 'setuptools', 'pkg_resources',
    'tkinter.test',

    # ── Rarely-needed stdlib ───────────────────────────────────────
    'xmlrpc',
    'ftplib', 'imaplib', 'poplib', 'smtplib', 'telnetlib',

    # ── Unused reportlab sub-packages ─────────────────────────────
    'reportlab.graphics.renderSVG',
    'reportlab.graphics.renderPS',

    # ── ML / vision libs (not used anywhere) ──────────────────────
    'torch', 'torchvision', 'torchaudio', 'tensorboard',
    'sklearn', 'IPython', 'jupyter', 'notebook',
    'cv2', 'numba', 'llvmlite',
]

a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=[],
    # ASOS labels draw barcodes with the template's Free 3 of 9 Extended font (MIT-style licence kept alongside).
    datas=collect_data_files('docx') + [('fre3of9x.ttf', '.'), ('fre3of9x-license.txt', '.')],
    # ReportLab's barcode registry imports its encoders dynamically during startup.
    hiddenimports=['openpyxl.cell._writer', 'docx', 'lxml._elementpath', 'xlrd']
                  + collect_submodules('reportlab.graphics.barcode', filter=lambda name: not name.endswith('.test')),
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
    name=release['APP_TITLE'],
    version=version_info,
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
