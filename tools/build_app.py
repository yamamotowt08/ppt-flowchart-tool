#!/usr/bin/env python3
"""Windows 用のアプリ（EXE 1 つ）を作る。

    python3 tools/build_app.py

1. tools/build_web.py でツール本体（app/tool.html）を作り直す
2. app/ の起動アプリ（.NET Framework 4.8 + WebView2）をビルドする。
   ツール本体と WebView2 の DLL は EXE の中に入る
3. フローチャート作成ツール.exe をフォルダの一番上に置き、
   GitHub のリリース用に英数字の名前の ZIP（dist/FlowchartTool.zip）を作る
"""
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EXE = ROOT / 'app' / 'bin' / 'Release' / 'net48' / 'FlowchartTool.exe'
APP_NAME = 'フローチャート作成ツール.exe'
DIST = ROOT / 'dist'
ZIP_NAME = 'FlowchartTool.zip'


class ShiftJisName(zipfile.ZipInfo):
    """ZIP の中の日本語のファイル名を Shift_JIS で書く（Windows のエクスプローラーで文字化けしないように）。"""
    def _encodeFilenameFlags(self):
        return self.filename.encode('cp932'), self.flag_bits & ~0x800


def run(cmd, cwd):
    print('>', ' '.join(str(c) for c in cmd))
    subprocess.run(cmd, cwd=cwd, check=True)


run([sys.executable, 'tools/build_web.py'] + sys.argv[1:], ROOT)
run(['dotnet', 'build', '-c', 'Release', '--nologo'], ROOT / 'app')

shutil.copy2(EXE, ROOT / APP_NAME)
if DIST.exists():
    shutil.rmtree(DIST)
DIST.mkdir()
shutil.copy2(EXE, DIST / APP_NAME)
with zipfile.ZipFile(DIST / ZIP_NAME, 'w') as z:
    info = ShiftJisName.from_file(EXE, APP_NAME)
    info.compress_type = zipfile.ZIP_DEFLATED
    z.writestr(info, EXE.read_bytes())
print(f'{APP_NAME} を作りました（{EXE.stat().st_size // 1024} KB）')
print(f'dist/{ZIP_NAME} を作りました（{(DIST / ZIP_NAME).stat().st_size // 1024} KB）')
