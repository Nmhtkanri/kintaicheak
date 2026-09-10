# -*- coding: utf-8 -*-
"""賞与エンジンを経理担当の最終 CSV（実データ）と突合する。共有フォルダが無ければ skip。

2026-09-10 の結果:
  - 2026-09 FE部賞与（30 人・発生日 9/15）: 支給・厚生年金は行順・金額・日付まで一致。健康保険は
    支援金の合算行の品目名だけ違う（経理担当は 9 月だけ「（預り分）」無し。6 月は付き。マスタの決まりどおり付きに統一）。
  - 2026-06 本社賞与（10 人・発生日 6/15）: 支給は完全一致。社保は 5 か所で 1 円差
    （経理担当が 6 月は jinjer の事業主列、9 月は 標準賞与額×総率÷2 の四捨五入 を使っており月で方法が違う。
      エンジンは 9 月方式。差は要確認 md の「会社負担が jinjer の事業主列と 1 円ずれた人」に出る）。

実行: python -X utf8 -m pytest tests/test_keiri_bonus_realdata.py -q
"""
from __future__ import annotations

import csv
import io
import os
import shutil
import sys
from collections import Counter
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

CASES = [
    ("2026-09", "FE部賞与", "2026-09-15",
     r"Y:\給与明細\R8年\9月\jinjer_賞与支給控除項目一覧表_9637_20260904.csv", r"Y:\給与明細\R8年\9月\freee",
     {"支給": 0, "健康保険": 2, "厚生年金": 0}),        # 健保の 2 = 支援金の品目名（（預り分）の有無）
    ("2026-06", "本社賞与", "2026-06-15",
     r"Y:\給与明細\R8年\6月\jinjer_賞与支給控除項目一覧表_9637_20260610.csv", r"Y:\給与明細\R8年\6月\freee",
     {"支給": 0, "健康保険": 4, "厚生年金": 1}),        # 1 円差（経理担当の 6 月の方法が違う）
]
RAW_DIR = ROOT / "outputs" / "keiri" / "raw"


def _load(path):
    return list(csv.DictReader(io.open(path, encoding="cp932", errors="replace")))


def _sums(rows):
    c = Counter()
    for r in rows:
        c[(r["従業員"].replace("\u3000", " "), r["品目"], r["勘定科目"], r["部門"])] += int(float(str(r["金額"] or 0).replace(",", "")))
    return c


def _heads(rows):
    """取引の先頭行（管理番号・発生日・支払期日・取引先）の並び。"""
    return [(r["管理番号"], r["発生日"], r["支払期日"], r["取引先"]) for r in rows if r.get("収支区分")]


@pytest.mark.parametrize("month,label,hassei,src_csv,final_dir,expected_diffs", CASES,
                         ids=[c[1] for c in CASES])
def test_matches_accountant_files(tmp_path, month, label, hassei, src_csv, final_dir, expected_diffs):
    if not (os.path.exists(src_csv) and os.path.isdir(final_dir)
            and (RAW_DIR / "roster.json").exists() and (RAW_DIR / "custom_items.json").exists()):
        pytest.skip("共有フォルダ（Y:）または名簿・部門キャッシュ（outputs/keiri/raw）が無い")
    from services.keiri_bonus import generate_bonus

    out_base = tmp_path / "keiri"
    (out_base / "raw").mkdir(parents=True)
    for f in ("roster.json", "custom_items.json"):
        shutil.copy(RAW_DIR / f, out_base / "raw" / f)
    res = generate_bonus(month, src_csv, label, hassei, out_base=str(out_base), client=None)
    for kind, info in res["files"].items():
        final = os.path.join(final_dir, info["name"])
        assert os.path.exists(final), final
        g, f = _load(info["path"]), _load(final)
        assert len(g) == len(f), (kind, len(g), len(f))
        gs, fs = _sums(g), _sums(f)
        diffs = [(k, gs.get(k), fs.get(k)) for k in sorted(set(gs) | set(fs)) if gs.get(k) != fs.get(k)]
        assert len(diffs) == expected_diffs[kind], (kind, diffs)
        assert _heads(g) == _heads(f), (kind, _heads(g), _heads(f))     # 取引の切り方・日付・管理番号まで一致
        if kind == "支給":
            assert [(r["従業員"], r["品目"], r["金額"]) for r in g] == \
                   [(r["従業員"].replace("\u3000", " "), r["品目"], str(r["金額"]).replace(",", "")) for r in f]
