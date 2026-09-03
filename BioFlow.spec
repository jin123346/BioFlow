# -*- mode: python ; coding: utf-8 -*-
# 빌드: pyinstaller BioFlow.spec
# 결과물: dist/BioFlow-<VERSION 파일 내용>/ (exe + _internal 폴더) 전체를 그대로 복사해서 배포.
# 새 버전을 낼 땐 빌드 전에 VERSION 파일 내용만 올리면 됨 -> 이전 버전 폴더를 덮어쓰지 않고
# dist/ 밑에 버전별로 따로 남음.
# data/ 폴더는 exe에 번들되어 최초 실행 시 내 문서\BioFlow\data로 시드 복사됨
# (config/paths.py의 ensure_data_seeded 참고). 이후 실행부터는 그쪽만 읽고 씀.

with open("VERSION", encoding="utf-8") as _f:
    APP_VERSION = _f.read().strip()

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('data', 'data')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5'],  # requirements.txt에 있지만 코드에서 실제로 쓰지 않음
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=f'BioFlow-{APP_VERSION}',
)
