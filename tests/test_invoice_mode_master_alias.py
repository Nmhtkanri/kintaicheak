from services import invoice_mode as im


def test_override_master_registers_alias_and_partner_override(tmp_path):
    master_csv = tmp_path / "master.csv"
    master_csv.write_text(
        "氏名,氏名別名,社員番号,freee取引先,部門\n"
        "奈良隆宏,Nara Takahiro|NaraTakahiro,2022013,エリクソン・ジャパン株式会社,SI：その他\n",
        encoding="utf-8-sig",
    )
    master = im.load_employee_master(None, master_csv)
    got = master[im.normalize_name("Nara, Takahiro")]
    assert got["employee_no"] == "2022013"
    assert got["employee_name"] == "奈良隆宏"
    assert got["partner"] == "エリクソン・ジャパン株式会社"
    assert got["partner_override"] is True
