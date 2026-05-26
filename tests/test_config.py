import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config


def test_template_csv_discovery_falls_back_to_unc_dir(monkeypatch, tmp_path):
    mapped_dir = tmp_path / "mapped"
    unc_dir = tmp_path / "unc"
    unc_dir.mkdir()
    latest = unc_dir / "スケジュール雛形一覧_2026-05-22.csv"
    latest.write_text("header\n", encoding="utf-8")

    monkeypatch.delenv("JINJER_TEMPLATE_CSV_PATH", raising=False)
    monkeypatch.setattr(
        config,
        "DEFAULT_JINJER_TEMPLATE_DIRS",
        (str(mapped_dir), str(unc_dir)),
    )

    assert config._resolve_jinjer_template_csv_path() == str(latest)
