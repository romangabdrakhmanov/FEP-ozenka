# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules
ROOT = Path(SPECPATH)
a = Analysis(
    [str(ROOT / 'app' / 'exe_launcher.py')],
    pathex=[str(ROOT / 'app')],
    binaries=[],
    datas=[
        (str(ROOT / 'references' / 'cascade-metrics.csv'), 'references'),
        (str(ROOT / 'references' / 'metric-families.json'), 'references'),
        (str(ROOT / 'references' / 'methodology.md'), 'references'),
    ],
    hiddenimports=collect_submodules('openpyxl'),
    hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False, optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='GoalAudit', debug=False,
          bootloader_ignore_signals=False, strip=False, upx=False, console=True,
          disable_windowed_traceback=False)
