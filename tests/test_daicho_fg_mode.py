# -*- coding: utf-8 -*-
"""派遣台帳の fg_mode（auto / legacy / report）判定のテスト（合成ファイル。実データは使わない）。

2026-09-10 追加。build_quarter から切り出した resolve_fg_source / uses_ericsson_overlay を
3 通り＋ no_fg・--fg 明示の組み合わせで固定する。台帳本体の出力はここでは見ない。

実行: python -X utf8 -m pytest tests/test_daicho_fg_mode.py -q
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.daicho.build import FG_MODES, resolve_fg_source, uses_ericsson_overlay  # noqa: E402


def _touch(folder: Path, name: str, *, mtime_offset: float = 0.0) -> Path:
    p = folder / name
    p.write_bytes(b"x")
    if mtime_offset:
        t = time.time() + mtime_offset
        import os
        os.utime(p, (t, t))
    return p


@pytest.fixture
def input_dir(tmp_path: Path) -> Path:
    d = tmp_path / "input"
    d.mkdir()
    return d


def test_modes_constant_matches_cli_choices():
    assert FG_MODES == ("auto", "legacy", "report")


@pytest.mark.parametrize("mode", FG_MODES)
def test_cli_accepts_every_mode_and_passes_it_through(monkeypatch, mode: str):
    # CLI の --fg-mode の選択肢は FG_MODES と同じ 1 か所で管理する（二重管理の検知）
    from services.daicho import __main__ as cli
    seen = {}

    def fake_build(args):
        seen["fg_mode"] = args.fg_mode
        return 0

    monkeypatch.setattr(cli, "cmd_build", fake_build)
    assert cli.main(["build", "--quarter", "2026Q2", "--fg-mode", mode]) == 0
    assert seen["fg_mode"] == mode


def test_cli_rejects_unknown_mode(monkeypatch, capsys):
    from services.daicho import __main__ as cli
    monkeypatch.setattr(cli, "cmd_build", lambda args: 0)
    with pytest.raises(SystemExit):
        cli.main(["build", "--quarter", "2026Q2", "--fg-mode", "bogus"])
    assert "bogus" in capsys.readouterr().err


@pytest.mark.parametrize("no_fg", [False, True])
def test_invalid_mode_raises_value_error(input_dir: Path, no_fg: bool):
    # no_fg=True でも不正値は黙って通さない（検証が早期 return より先）
    with pytest.raises(ValueError):
        resolve_fg_source("bogus", no_fg=no_fg, input_dir=input_dir)


@pytest.mark.parametrize("mode", FG_MODES)
def test_no_fg_reads_nothing_in_every_mode(input_dir: Path, mode: str):
    # レポートも旧CSVも置いてあるが、no_fg なら一切拾わない
    _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    _touch(input_dir, "WorkOrder_2026Q2.csv")
    assert resolve_fg_source(mode, no_fg=True, input_dir=input_dir) == (False, None, None)


def test_legacy_ignores_report_and_uses_workorder_csv(input_dir: Path):
    _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    wo = _touch(input_dir, "WorkOrder_2026Q2.csv")
    use_report, report_path, fg_path = resolve_fg_source("legacy", input_dir=input_dir)
    assert use_report is False
    assert report_path is None
    assert fg_path == wo


def test_legacy_without_workorder_means_no_fieldglass(input_dir: Path):
    _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    assert resolve_fg_source("legacy", input_dir=input_dir) == (False, None, None)


def test_auto_prefers_report_when_present(input_dir: Path):
    rep = _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    _touch(input_dir, "WorkOrder_2026Q2.csv")
    assert resolve_fg_source("auto", input_dir=input_dir) == (True, rep, None)


def test_auto_falls_back_to_legacy_without_report(input_dir: Path):
    wo = _touch(input_dir, "WorkOrder_2026Q2.csv")
    assert resolve_fg_source("auto", input_dir=input_dir) == (False, None, wo)


def test_auto_with_no_inputs_means_no_fieldglass(input_dir: Path):
    # 新レポートも旧CSVも無い四半期（レポート未着など）は Fieldglass 分ゼロで進む
    assert resolve_fg_source("auto", input_dir=input_dir) == (False, None, None)


def test_auto_picks_newest_report(input_dir: Path):
    old = _touch(input_dir, "ユニアデックス_業務内容_2026Q1.xlsx", mtime_offset=-3600)
    new = _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    use_report, report_path, _ = resolve_fg_source("auto", input_dir=input_dir)
    assert use_report and report_path == new and report_path != old


def test_report_requires_report_file(input_dir: Path):
    _touch(input_dir, "WorkOrder_2026Q2.csv")
    with pytest.raises(FileNotFoundError):
        resolve_fg_source("report", input_dir=input_dir)


def test_report_uses_report_even_if_workorder_exists(input_dir: Path):
    rep = _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    _touch(input_dir, "WorkOrder_2026Q2.csv")
    assert resolve_fg_source("report", input_dir=input_dir) == (True, rep, None)


@pytest.mark.parametrize("mode", ["auto", "legacy"])
def test_explicit_fg_skips_report_lookup(input_dir: Path, tmp_path: Path, mode: str):
    # --fg を明示したら、input に新レポートがあっても探さず、その CSV を使う
    _touch(input_dir, "ユニアデックス_業務内容_2026Q2.xlsx")
    given = tmp_path / "given_WorkOrder.csv"
    given.write_bytes(b"x")
    assert resolve_fg_source(mode, fg=str(given), input_dir=input_dir) == (False, None, given)


def test_explicit_fg_with_report_mode_still_requires_report(input_dir: Path, tmp_path: Path):
    # report は新レポート必須なので、--fg を渡しても新レポートが無ければ止まる（移設前と同じ挙動）
    given = tmp_path / "given_WorkOrder.csv"
    given.write_bytes(b"x")
    with pytest.raises(FileNotFoundError):
        resolve_fg_source("report", fg=str(given), input_dir=input_dir)


def test_ericsson_overlay_only_outside_legacy():
    assert uses_ericsson_overlay("auto") is True
    assert uses_ericsson_overlay("report") is True
    assert uses_ericsson_overlay("legacy") is False


def test_ericsson_overlay_rejects_unknown_mode():
    with pytest.raises(ValueError):
        uses_ericsson_overlay("bogus")
