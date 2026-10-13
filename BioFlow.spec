# -*- mode: python ; coding: utf-8 -*-
# 빌드: pyinstaller BioFlow.spec
# 결과물: dist/BioFlow-<VERSION 파일 내용>/ (exe + _internal 폴더) 전체를 그대로 복사해서 배포.
# 새 버전을 낼 땐 빌드 전에 VERSION 파일 내용만 올리면 됨 -> 이전 버전 폴더를 덮어쓰지 않고
# dist/ 밑에 버전별로 따로 남음.
# data/ 폴더는 exe에 번들되어 최초 실행 시 내 문서\BioFlow\data로 시드 복사됨
# (config/paths.py의 ensure_data_seeded 참고). 이후 실행부터는 그쪽만 읽고 씀.

import json
import os
import shutil
import sys
from pathlib import Path

# Tcl cannot reliably resolve the Python install path when it contains Korean text.
# Stage its scripts under the build directory; no files in the Python install are changed.
_tcl_stage = Path('build/tcl-runtime').resolve()
for _library in ('tcl8.6', 'tk8.6'):
    shutil.copytree(Path(sys.base_prefix) / 'tcl' / _library, _tcl_stage / _library, dirs_exist_ok=True)
os.environ['TCL_LIBRARY'] = str(_tcl_stage / 'tcl8.6')
os.environ['TK_LIBRARY'] = str(_tcl_stage / 'tk8.6')
# Splash imports this helper before the spec runs; refresh its earlier failed probe.
from PyInstaller.utils.hooks import tcl_tk as _tcl_hook
(_tcl_hook.tcl_dir, _tcl_hook.tcl_version,
 _tcl_hook.tk_version, _tcl_hook.tcl_threaded) = _tcl_hook._get_tcl_tk_info()

with open("VERSION", encoding="utf-8") as _f:
    APP_VERSION = _f.read().strip()

# 빌드 = 릴리즈 기준이라, 앱 자체 업데이트 알림용 app_version.json도 VERSION에 맞춰 같이
# 갱신함 (services/data_update_service.py의 check_for_app_update가 이 파일을 원격에서
# 읽어 새 버전 알림을 띄움 -> 두 파일을 따로 손으로 맞출 필요 없게).
_app_version_path = Path("app_version.json")
try:
    _app_version_data = json.loads(_app_version_path.read_text(encoding="utf-8"))
except Exception:
    _app_version_data = {}
_app_version_data["version"] = APP_VERSION
_app_version_data.setdefault("download_url", "https://github.com/jin123346/BioFlow/releases/latest")
_app_version_path.write_text(json.dumps(_app_version_data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

# Exclude local backups and in-progress drafts from distribution.
seed_data = [
    (str(path), str(path.parent))
    for path in Path('data').rglob('*')
    if path.is_file() and 'history' not in path.parts
    and path.name not in {'new_data_draft.json', 'new_specimen_draft.json'}
]

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=seed_data + [('assets', 'assets'), ('VERSION', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # PyQt5: requirements.txt에 있지만 코드에서 실제로 쓰지 않음
    # scipy/pyarrow: 코드에서 안 쓰는데 pandas 설치 환경에 있으면 훅이 같이 끌고 옴
    # (둘이 unpacked 기준 150MB+ 차지) -> 제외해서 빌드 용량 줄임
    excludes=['PyQt5', 'scipy', 'pyarrow', 'IPython', 'pytest', 'jedi', 'matplotlib'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# 더블클릭 후 창이 뜨기까지 몇 초 걸리는데 그동안 아무 반응이 없어 여러 번 누르게 되는
# 문제 때문에 추가 - 파이썬 인터프리터가 뜨기도 전에 부트로더가 바로 이 이미지를 띄움
# (app.py에서 메인창이 뜨면 pyi_splash.close()로 닫음)
splash = Splash(
    'assets/splash/splash_light.png',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    always_on_top=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    exclude_binaries=True,
    name='BioFlow',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/favicon/favicon.ico',
)
coll = COLLECT(
    exe,
    splash.binaries,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=f'BioFlow-{APP_VERSION}',
)
