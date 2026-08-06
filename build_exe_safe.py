# -*- coding: utf-8 -*-
"""勤怠チェッカーの配布用exeを「一時フォルダへ」安全にビルドする。

稼働中の共有 exe を使ったまま `dist\\` へ直接ビルドしてはいけない。PyInstaller は
出力先を削除してから書き込むため、誰かが exe を起動中だと `dist` が半削除されて壊れる。
このスクリプトは必ず一時の distpath（既定 `dist_new`）へビルドする。入れ替え(swap)は別工程。

**ビルドオプションは build_exe.bat から読み取る。** ここに書き写さないのは、二重管理にすると
必ず片方が古くなるため。実際 2026-08-06 時点で kintai-exe-builder エージェントの手順書に
書かれていたコマンドは `--collect-all pdfplumber` と `--hidden-import services.*` を
すべて落としており、そのまま使うと壊れた exe ができる状態だった。

使い方:
    python build_exe_safe.py                     # dist_new へビルド
    python build_exe_safe.py --distpath dist_x   # 出力先を変える
    python build_exe_safe.py --show              # 実行するコマンドを表示するだけ
"""

from __future__ import annotations

import argparse
import shlex
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BAT = ROOT / "build_exe.bat"


def _split(piece: str) -> list[str]:
    """bat の1行分をトークンへ分解する。

    posix=False にしないと Windows のパス区切り `\\` がエスケープとして食われる。
    ただし posix=False はクォートを**残す**ため、`"templates;templates"` がクォート込みで
    PyInstaller に渡り `Unable to find '..."templates'` で失敗する（2026-08-06 に実際に踏んだ）。
    subprocess にリスト形式で渡す以上クォートは不要なので、ここで外す。
    """
    out = []
    for tok in shlex.split(piece, posix=False):
        if len(tok) >= 2 and tok[0] == tok[-1] and tok[0] in "\"'":
            tok = tok[1:-1]
        out.append(tok)
    return out


def parse_bat_args(bat_path: Path) -> list[str]:
    """build_exe.bat の PyInstaller 呼び出しから引数だけを取り出す。

    `python -m PyInstaller ^` から `launcher.py` までを連結し、行末の `^`、
    REM コメント行、空行を落とす。
    """
    text = bat_path.read_text(encoding="utf-8")
    args: list[str] = []
    collecting = False
    for raw in text.splitlines():
        line = raw.strip()
        if not collecting:
            if line.lower().startswith("python -m pyinstaller"):
                collecting = True
                rest = line[len("python -m PyInstaller"):].strip()
                if rest.rstrip("^").strip():
                    args += _split(rest.rstrip("^").strip())
                if not line.endswith("^"):
                    break
            continue
        if not line or line.upper().startswith("REM"):
            continue
        piece = line.rstrip("^").strip()
        if piece:
            args += _split(piece)
        if not line.endswith("^"):
            break
    if not args:
        raise ValueError(f"build_exe.bat から PyInstaller の引数を読み取れませんでした: {bat_path}")
    if args[-1] != "launcher.py":
        raise ValueError(f"想定と違う末尾です（launcher.py のはず）: {args[-1]}")
    return args


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="勤怠チェッカーexeを一時フォルダへ安全にビルドする")
    ap.add_argument("--distpath", default="dist_new", help="出力先（既定: dist_new）")
    ap.add_argument("--workpath", default="build_new", help="作業フォルダ（既定: build_new）")
    ap.add_argument("--show", action="store_true", help="実行するコマンドを表示するだけ")
    args = ap.parse_args(argv)

    if args.distpath.rstrip("\\/").lower() == "dist":
        print("[ERROR] 出力先に dist を指定してはいけません（稼働中の共有exeが壊れます）。",
              file=sys.stderr)
        return 1

    pyi_args = parse_bat_args(BAT)
    # 末尾の launcher.py の前に一時パスを差し込む
    cmd = ([sys.executable, "-m", "PyInstaller"] + pyi_args[:-1]
           + ["--distpath", args.distpath, "--workpath", args.workpath, pyi_args[-1]])

    print("build_exe.bat から読み取ったオプション:")
    for a in pyi_args[:-1]:
        print(f"    {a}")
    print()
    print("実行するコマンド:")
    print("   ", " ".join(cmd))
    print()
    if args.show:
        return 0

    proc = subprocess.run(cmd, cwd=ROOT)
    if proc.returncode != 0:
        print(f"[ERROR] ビルドに失敗しました（終了コード {proc.returncode}）", file=sys.stderr)
        return proc.returncode

    exe = ROOT / args.distpath / "KintaiChecker" / "KintaiChecker.exe"
    if not exe.exists():
        print(f"[ERROR] exe が生成されていません: {exe}", file=sys.stderr)
        return 1
    print()
    print(f"[OK] ビルド成功: {exe}")
    print("     次は同梱チェック → 稼働ゼロ確認 → swap の順に進めること（dist へは直接書かない）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
