import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.jinjer_parser import parse_jinjer_csv

SAMPLE_CSV = os.path.join(os.path.dirname(__file__), "sample_data", "sample_jinjer.csv")


def test_parse_jinjer_basic():
    df = parse_jinjer_csv(SAMPLE_CSV)
    assert len(df) > 0, "レコードが0件"
    assert "氏名" in df.columns
    assert "日付" in df.columns
    assert "出勤時刻" in df.columns
    assert "退勤時刻" in df.columns
    assert "総労働時間(分)" in df.columns
    assert "データソース" in df.columns
    assert (df["データソース"] == "jinjer").all()
    print(f"解析件数: {len(df)}")
    print(df.head())


def test_parse_jinjer_names():
    # パーサは氏名の内部スペースを保持する（突合時の正規化は services/matcher.py が担当）。
    # サンプルCSVは jinjer 標準どおり半角スペース区切り。
    df = parse_jinjer_csv(SAMPLE_CSV)
    names = df["氏名"].unique().tolist()
    assert "山田 太郎" in names
    assert "佐藤 花子" in names


def test_parse_jinjer_times():
    df = parse_jinjer_csv(SAMPLE_CSV)
    from datetime import time
    first = df[df["氏名"] == "山田 太郎"].iloc[0]
    assert first["出勤時刻"] == time(9, 0)
    assert first["退勤時刻"] == time(18, 0)


def test_parse_jinjer_uses_total_work_time_not_actual_work_time(tmp_path):
    path = tmp_path / "jinjer_total.csv"
    path.write_text(
        "名前,*年月日,出勤1,退勤1,総労働時間,実労働時間\n"
        "中澤 寿代,2026/06/01,09:45,19:26,08:41,09:17\n",
        encoding="utf-8",
    )

    df = parse_jinjer_csv(path)

    assert len(df) == 1
    assert df.iloc[0]["総労働時間(分)"] == 521


def test_parse_jinjer_missing_total_column_keeps_none():
    df = parse_jinjer_csv(SAMPLE_CSV)
    assert df["総労働時間(分)"].isna().all()


if __name__ == "__main__":
    test_parse_jinjer_basic()
    test_parse_jinjer_names()
    test_parse_jinjer_times()
    print("全テスト通過")
