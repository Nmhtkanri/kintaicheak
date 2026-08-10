# -*- coding: utf-8 -*-
"""uploads の自動掃除のテスト

uploads に入るのは処理のために保存した入力のコピー。成果物は outputs 側にある。
掃除で成果物まで消えると取り返しがつかないので、そこを一番きつく見張る。
"""

from __future__ import annotations

import os
import sys
import time

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module  # noqa: E402
from config import Config  # noqa: E402


def make(path, *, days_old=0):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"x")
    if days_old:
        old = time.time() - days_old * 24 * 3600
        os.utime(path, (old, old))
    return path


@pytest.fixture
def folders(tmp_path, monkeypatch):
    uploads = tmp_path / "uploads"
    outputs = tmp_path / "outputs"
    (uploads / "sessions").mkdir(parents=True)
    outputs.mkdir()
    monkeypatch.setattr(Config, "UPLOAD_FOLDER", str(uploads))
    monkeypatch.setattr(Config, "OUTPUT_FOLDER", str(outputs))
    monkeypatch.setattr(Config, "UPLOAD_RETENTION_DAYS", 7)
    return {"uploads": uploads, "outputs": outputs}


class TestCleanupUploads:
    def test_removes_old_and_keeps_recent(self, folders):
        old = make(folders["uploads"] / "ts_old.xlsx", days_old=30)
        recent = make(folders["uploads"] / "ts_recent.xlsx", days_old=1)

        removed = app_module.cleanup_uploads_on_start()

        assert removed == 1
        assert not old.exists()
        assert recent.exists(), "最近のものは残す（作業中かもしれない）"

    def test_never_touches_outputs(self, folders):
        """成果物は消さない。ここが壊れると作ったものが失われる。"""
        made = make(folders["outputs"] / "差異一覧_202603.xlsx", days_old=365)
        nested = make(folders["outputs"] / "keiri" / "202603" / "kyuyo.csv", days_old=365)

        app_module.cleanup_uploads_on_start()

        assert made.exists()
        assert nested.exists()

    def test_cleans_sessions_subfolder(self, folders):
        old = make(folders["uploads"] / "sessions" / "abc.pkl", days_old=30)
        recent = make(folders["uploads"] / "sessions" / "def.pkl", days_old=2)

        app_module.cleanup_uploads_on_start()

        assert not old.exists()
        assert recent.exists()

    def test_boundary(self, folders):
        just_inside = make(folders["uploads"] / "a.xlsx", days_old=6)
        just_outside = make(folders["uploads"] / "b.xlsx", days_old=8)

        app_module.cleanup_uploads_on_start()

        assert just_inside.exists()
        assert not just_outside.exists()

    def test_disabled_with_zero(self, folders, monkeypatch):
        monkeypatch.setattr(Config, "UPLOAD_RETENTION_DAYS", 0)
        old = make(folders["uploads"] / "ts_old.xlsx", days_old=365)

        assert app_module.cleanup_uploads_on_start() == 0
        assert old.exists(), "0 なら掃除しない（止めたいときの逃げ道）"

    def test_explicit_days_wins(self, folders):
        old = make(folders["uploads"] / "ts.xlsx", days_old=3)

        app_module.cleanup_uploads_on_start(max_age_days=1)

        assert not old.exists()

    def test_empty_folder_is_fine(self, folders):
        assert app_module.cleanup_uploads_on_start() == 0

    def test_missing_folder_is_fine(self, tmp_path, monkeypatch):
        monkeypatch.setattr(Config, "UPLOAD_FOLDER", str(tmp_path / "ない"))
        assert app_module.cleanup_uploads_on_start() == 0

    def test_folders_are_kept(self, folders):
        """フォルダ自体は消さない（次の実行で作り直す手間を増やさない）。"""
        make(folders["uploads"] / "sessions" / "old.pkl", days_old=30)

        app_module.cleanup_uploads_on_start()

        assert (folders["uploads"] / "sessions").is_dir()


class TestNotCalledOnImport:
    def test_import_does_not_clean(self, folders):
        """テストで app を import しただけで実ファイルが消えては困る。

        呼ぶのは launcher.py と __main__ からだけ、という約束を固定する。
        """
        old = make(folders["uploads"] / "ts_old.xlsx", days_old=365)

        import importlib
        importlib.reload(app_module)

        assert old.exists()
