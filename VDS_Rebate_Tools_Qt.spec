# -*- mode: python ; coding: utf-8 -*-
import datetime as _dt
_year = _dt.date.today().year

SPEC_DOC = f"""PyInstaller spec (PySide6/Qt build)
Developed by www.3SVerse.com © {_year}
Build command: pyinstaller VDS_Rebate_Tools_Qt.spec
"""


block_cipher = None

a = Analysis(
    ['VDS_Rebate_Tools_Qt.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('runtime_guard.py', '.'),
        ('assets', 'assets'),
        ('VDS_icon.ico', '.'),
        ('VDS_Logo.png', '.'),
        ('stores.json', '.'),
        ('theme_manager_qt.py', '.'),
        ('header_manager_qt.py', '.'),
        ('native_titlebar_qt.py', '.'),
    ],
    hiddenimports=[
        'PySide6',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'shiboken6',
        'openpyxl',
        'xlrd',
        'xlwt',
        'xlutils',
        'PIL',
        'theme_manager_qt',
        'header_manager_qt',
        'native_titlebar_qt',
        'runtime_guard',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
        'tkinter',
        '_tkinter',
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


splash = Splash(
    'splash.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(60, 316),
    text_size=11,
    text_color='#96a8cd',
    text_default='Starting...',
    minify_script=True,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='VDS_Rebate_Tools_Qt',
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
    icon='VDS_icon.ico',
)
