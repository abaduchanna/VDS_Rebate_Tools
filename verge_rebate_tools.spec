# -*- mode: python ; coding: utf-8 -*-
import datetime as _dt
_year = _dt.date.today().year

SPEC_DOC = f"""PyInstaller spec
Developed by Abad Umair Channa \u00a9 {_year}
Build command: pyinstaller verge_verge_rebate_tools.spec
"""


block_cipher = None

a = Analysis(
    ['verge_verge_rebate_tools.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('assets', 'assets'),
        ('verge_icon.ico', '.'),
        ('Verge_Logo.png', '.'),
        ('stores.json', '.'),
        ('theme_manager.py', '.'),
        ('logo_handler.py', '.'),
        ('header_manager.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        '_tkinter',
        'openpyxl',
        'xlrd',
        'xlwt',
        'xlutils',
        'PIL',
        'theme_manager',
        'logo_handler',
        'header_manager',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={
        'matplotlib': {'backends': ''},
    },
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
        'pandas',
        'numpy',
        'numba',
        'llvmlite',
        'matplotlib',
        'matplotlib.pyplot',
        'matplotlib.backends',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'shiboken2',
        'shiboken6',
        'scipy',
        'sympy',
        'torch',
        'torchvision',
        'torchaudio',
        'tensorflow',
        'sklearn',
        'scikit-learn',
        'speech_recognition',
        'SpeechRecognition',
        'imageio',
        'imageio_ffmpeg',
        'soundfile',
        'gi',
        'pygments',
        'fsspec',
        'tensorboard',
        'IPython',
        'ipython',
        'jupyter',
        'notebook',
        'selenium',
        'requests',
        'pyautogui',
        'pyperclip',
    ],
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
    name='rebate_tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='verge_icon.ico',
)
