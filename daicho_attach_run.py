# -*- coding: utf-8 -*-
"""派遣台帳のjinjer添付ジョブを python から起動するCLI口（デタッチ子プロセス用）。

exe では launcher.py の --daicho-attach 分岐が同じ処理をする。.env の読み込みは
main_cli 側で行う。使い方:
    python -X utf8 daicho_attach_run.py --execute [--limit 1] [--start-at 22:00]
"""
import sys

from services.daicho.attach_job import main_cli

if __name__ == "__main__":
    sys.exit(main_cli(sys.argv))
