# -*- coding: utf-8 -*-
"""Contract（e-staffing）＋ Person（jinjer）→ 台帳1枚分の項目辞書と警告。

項目名はそのまま CSV のヘッダーと、template.CELL_MAP のキーになる。
法定記載事項（労働者派遣法37条・施行規則31条）に沿って、旧フォームに無かった項目も持つ。
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .config import DISPATCH_LICENSE_NO, DISPATCH_SOURCE_NAME
from .estaffing import SHIFT_NOTE, Contract, parse_date, split_time
from .roster import Person

MANUAL = "※別途記録（該当があれば記入）"
CMD = "事業所の名称及び所在地その他派遣就業場所"   # CPI の長い接頭辞


@dataclass
class LedgerRecord:
    fields: dict[str, str]
    warnings: list[str] = field(default_factory=list)
    contract: Contract | None = None
    person: Person | None = None

    @property
    def emp_id(self) -> str:
        return self.fields.get("社員番号", "")

    @property
    def name(self) -> str:
        return self.fields.get("氏名", "")


def _fmt_date(d: dt.date | None) -> str:
    return d.strftime("%Y/%m/%d") if d else ""


def _person_block(c: Contract, prefix: str) -> tuple[str, str, str]:
    """CPI の『{prefix} 部署/役職/氏名/TEL』を (部署, 役職 氏名, TEL:…) の3行にする。"""
    dept = c.c(f"{prefix} 部署")
    post = c.c(f"{prefix} 役職")
    name = c.c(f"{prefix} 氏名")
    tel = c.c(f"{prefix} TEL")
    return dept, f"{post} {name}".strip(), (f"TEL:{tel}" if tel else "")


def _working_hours(c: Contract) -> tuple[str, str, str]:
    """(就業時間, 休憩1, 休憩2以降) の表示文字列。"""
    s, sh1 = split_time(c.c("就業時間 開始時間"))
    e, sh2 = split_time(c.c("就業時間 終了時間"))
    work, _ = split_time(c.c("就業時間 就業時間"))
    # 片側だけのとき（シフト制の説明文など）は～を付けず原文のまま
    hours = f"{s}～{e}" if (s and e) else (s or e)
    if work:
        hours += f"（実働 {work}）"
    if sh1 or sh2:
        hours += f"　{SHIFT_NOTE}"
    breaks = []
    for n in (1, 2, 3):
        bs, _ = split_time(c.c(f"休憩時間{n} 開始時間"))
        be, _ = split_time(c.c(f"休憩時間{n} 終了時間"))
        bl, _ = split_time(c.c(f"休憩時間{n} 時間"))
        if bs or be:
            rng = f"{bs}～{be}" if (bs and be) else (bs or be)
            breaks.append(f"休憩{n}: {rng}" + (f"（{bl}）" if bl else ""))
    b1 = breaks[0] if breaks else ""
    b2 = "　".join(breaks[1:]) if len(breaks) > 1 else ""
    return hours, b1, b2


def _insurance(c: Contract, kind: str) -> tuple[str, str, str]:
    """(有/無の表示, 理由, 資格取得届の行の文言)。"""
    status = c.c(kind)                    # 有 / 無(加入対象外) / 無(手続中)
    note = c.c(f"{kind} 補足")
    if status.startswith("有"):
        return "有", "", f"{kind}：提出済（加入）"
    if "手続" in status:
        return "手続中", note, f"{kind}：未提出（手続中）" + (f"／{note}" if note else "")
    if status.startswith("無"):
        reason = f"／理由: {note}" if note else "／理由: 未記載 ※要確認"
        return "無", note, f"{kind}：未提出（加入対象外）{reason}"
    return status or "", note, f"{kind}：{status or '不明 ※要確認'}"


def _conveniences(c: Contract) -> str:
    items = [k for k in ("診療施設", "給食施設", "休憩室", "更衣室") if c.t(f"便宜供与：{k}") == "1"]
    items += [c.t(f"便宜供与：{k}") for k in ("その他1", "その他2", "その他3") if c.t(f"便宜供与：{k}")]
    return "、".join(items) if items else "―"


def build_record(c: Contract, person: Person | None, match_state: str,
                 q_start: dt.date, q_end: dt.date,
                 generated_at: dt.datetime | None = None) -> LedgerRecord:
    w: list[str] = []
    f: dict[str, str] = {}
    start, end = c.start, c.end

    # --- 派遣労働者 ---
    f["氏名"] = c.worker_name
    f["社員番号"] = person.emp_id if person else ""
    if person is None:
        w.append("jinjer の従業員一覧に氏名が見つからない（退職者か表記ゆれ）→ 社員番号・生年月日・性別が空"
                 if match_state == "none" else "同姓同名が複数いて特定できない → 社員番号・生年月日・性別が空")
    # 契約データの性別はコード（0=男, 1=女。jinjer と突き合わせて確認済み）
    sex = person.sex if (person and person.sex) else {"0": "男", "1": "女"}.get(c.t("スタッフ性別"), c.t("スタッフ性別"))
    f["性別"] = sex
    f["生年月日"] = _fmt_date(person.birth) if person else ""
    age = person.age_at(start) if person else None
    f["年齢"] = f"{age}歳" if age is not None else ""
    f["氏名・年齢・性別"] = "　".join(x for x in ((f"契約開始時 {f['年齢']}" if age is not None else ""), sex) if x)

    # 60歳以上か否か: 契約書の宣言（期間制限の対象外理由）を正とし、生年月日で検算
    reason = c.c("期間制限の対象外理由")
    declared_60 = "60歳以上" in reason or c.t("期間制限の対象外理由：60歳以上派遣労働者") == "1"
    f["60歳以上の者であるか否かの別"] = "60歳以上" if declared_60 else "60歳未満"
    if age is not None and (age >= 60) != declared_60:
        w.append(f"60歳判定が不一致: 契約書={'60歳以上' if declared_60 else '60歳未満'} / jinjer生年月日から {age}歳")

    kyotei = c.c("協定対象派遣労働者に該当するか否かの別")
    f["協定対象派遣労働者であるか否かの別"] = {
        "該当する": "協定対象派遣労働者", "該当しない": "協定対象派遣労働者ではない"}.get(kyotei, kyotei or "※要確認")
    if not kyotei:
        w.append("協定対象か否かが契約書・通知書データに無い")

    emp_form = c.c("派遣元での雇用形態") or {"0": "無期雇用契約", "1": "有期雇用契約"}.get(c.t("派遣元での雇用形態"), "")
    f["有期か無期かの別"] = {"無期雇用契約": "無期雇用派遣労働者", "有期雇用契約": "有期雇用派遣労働者"}.get(emp_form, emp_form or "※要確認")
    if person and person.employment_type and emp_form:
        jinjer_form = "無期雇用契約" if person.employment_type in ("正社員", "役員") else "有期雇用契約"
        if jinjer_form != emp_form:
            w.append(f"雇用形態が不一致: 契約書={emp_form} / jinjer雇用区分={person.employment_type}")

    # --- 派遣先 ---
    f["派遣先名称"] = c.client_name
    addr = c.c(f"{CMD} 事業所の所在地及び就業場所") or c.t("就業先住所")
    tel = c.c(f"{CMD} 部署TEL")
    f["派遣先の事業所所在地"] = addr
    f["派遣先TEL"] = f"TEL:{tel}" if tel else ""
    office = c.t("就業先正式事業所名称") or c.c(f"{CMD} 事業所の名称")
    dept = c.c(f"{CMD} 部署名称") or c.t("就業先正式部署")
    f["就業場所"] = "　".join(x for x in (office, dept) if x)
    f["就業場所住所"] = addr
    f["就業場所TEL"] = f["派遣先TEL"]
    unit = c.t("組織単位") or dept
    chief = c.t("組織の長の職名")
    f["組織単位"] = unit + (f"（組織の長の職名：{chief}）" if chief else "")
    f["業務内容"] = c.c("業務内容") or c.t("業務内容")
    f["責任の程度"] = c.c("責任の程度")
    if not f["責任の程度"]:
        w.append("責任の程度が空（契約書・通知書データ未結合）")
    f["職種"] = c.c("職種") or c.t("職種")

    # --- 期間・就業条件 ---
    f["契約期間"] = f"{_fmt_date(start)}～{_fmt_date(end)}"
    if start is None or end is None:
        w.append("派遣期間が読めない")
    hours, b1, b2 = _working_hours(c)
    f["就業時間"] = hours
    f["休憩時間1"] = b1
    f["休憩時間2"] = b2
    f["就業曜日"] = c.c("勤務日") or c.t("勤務日")
    shift_flag = c.t("シフト") == "1" or SHIFT_NOTE in f["就業曜日"]
    f["就業曜日備考"] = SHIFT_NOTE if shift_flag else ""
    f["休日"] = c.c("休日")
    f["休日備考"] = c.t("法定休日")
    other_holiday = c.t("休日（その他）")
    if other_holiday and other_holiday not in f["休日"]:
        f["休日備考"] = "／".join(x for x in (other_holiday, f["休日備考"]) if x)
    f["就業時間外の労働"] = c.c("36協定1 時間外労働、休日労働") or c.t("36協定")
    f["休日勤務"] = "有（36協定の範囲内）" if c.c("休日労働") == "有" else c.c("休日労働")
    f["就業状況"] = "日々の始業・終業・休憩はjinjer勤怠記録（派遣先承認済み勤怠）による"

    # --- 責任者・苦情・指揮命令 ---
    f["派遣先責任者_部署"], f["派遣先責任者_役職氏名"], f["派遣先責任者_TEL"] = _person_block(c, "派遣先責任者")
    f["苦情申出先_部署"], f["苦情申出先_役職氏名"], f["苦情申出先_TEL"] = _person_block(c, "派遣先苦情申出先")
    f["指揮命令者_部署"] = c.c(f"{CMD} 指揮命令者部署")
    f["指揮命令者_役職氏名"] = f"{c.c(f'{CMD} 指揮命令者役職')} {c.c(f'{CMD} 指揮命令者氏名')}".strip()
    f["指揮命令者_TEL"] = f["派遣先TEL"]
    f["派遣元責任者_部署"], f["派遣元責任者_役職氏名"], f["派遣元責任者_TEL"] = _person_block(c, "派遣元責任者")
    (f["派遣元苦情処理担当者_部署"], f["派遣元苦情処理担当者_役職氏名"],
     f["派遣元苦情処理担当者_TEL"]) = _person_block(c, "派遣元苦情申出先")
    f["製造業務専門派遣先責任者"] = "―（製造業務への派遣なし）"
    f["製造業務専門派遣元責任者"] = "―（製造業務への派遣なし）"
    for k in ("派遣先責任者_役職氏名", "苦情申出先_役職氏名", "指揮命令者_役職氏名", "派遣元責任者_役職氏名"):
        if not f[k]:
            w.append(f"{k.split('_')[0]}が空")

    # --- 期間制限・抵触日 ---
    f["期間制限の対象外理由"] = reason or (
        "無期雇用派遣労働者" if c.t("期間制限の対象外理由：無期雇用派遣労働者") == "1" else "")
    f["有期プロジェクト業務"] = "該当" if c.t("期間制限の対象外理由：有期プロジェクト業務") == "1" else "―"
    f["日数限定業務・産休代替等"] = "該当" if (
        c.t("期間制限の対象外理由：日数限定業務") == "1"
        or c.t("期間制限の対象外理由：産前産後、育児休業、介護休業等の代替要員") == "1") else "―"
    f["事業所単位の抵触日"] = _fmt_date(parse_date(c.t("事業所抵触日"))) or c.t("事業所抵触日")
    personal = (c.c("個人抵触日") or c.t("個人抵触日") or "").replace("(期間制限の対象外)", "").strip()
    if f["期間制限の対象外理由"]:
        personal = f"{personal}（期間制限の対象外：{f['期間制限の対象外理由']}）".strip()
    f["個人単位の抵触日"] = personal

    # --- 保険 ---
    for kind in ("健康保険", "厚生年金", "雇用保険"):
        disp, note, line = _insurance(c, kind)
        f[f"{kind}_加入"] = disp
        f[f"{kind}_未加入理由"] = note
        f[f"資格取得届_{kind}"] = line
        if disp == "無" and not note:
            w.append(f"{kind}が未加入なのに理由が空")

    # --- その他 ---
    f["便宜供与"] = _conveniences(c)
    f["安全及び衛生"] = c.t("安全及び衛生")
    # FGの新レポートは教育訓練の記載を持つ（e-staffing・旧details JSONでは常に空＝従来どおり手書き）
    f["教育訓練の日時及び内容"] = c.c("教育訓練") or MANUAL
    f["キャリアコンサルティングの日時及び内容"] = MANUAL
    f["雇用安定措置の内容"] = MANUAL
    f["苦情の処理状況"] = c.t("苦情処理結果") or "申出なし"
    f["派遣先へ通知した事項"] = ("氏名・協定対象派遣労働者であるか否かの別・無期雇用か有期雇用かの別・"
                              "60歳以上であるか否かの別・社会保険及び雇用保険の被保険者資格取得届の提出の有無"
                              "（e-staffing 派遣先通知書による）")
    f["契約No"] = c.contract_no
    f["契約書確定日"] = _fmt_date(parse_date(c.t("契約確定日"))) or c.t("契約確定日")
    f["派遣元名称"] = c.c("派遣元企業 名称") or DISPATCH_SOURCE_NAME
    f["派遣許可番号"] = c.c("派遣元企業 派遣許可番号") or DISPATCH_LICENSE_NO
    f["契約書備考"] = c.t("契約書備考")
    f["所属グループ"] = person.group if person else ""
    gen = (generated_at or dt.datetime.now()).strftime("%Y/%m/%d %H:%M")
    f["備考"] = (f"契約No {c.contract_no}／契約書確定日 {f['契約書確定日']}／派遣許可番号 {f['派遣許可番号']}"
                f"／出所: e-staffing 契約データ・契約書通知書データ＋jinjer 従業員情報／作成 {gen}")
    return LedgerRecord(fields=f, warnings=w, contract=c, person=person)


# CSV に出す列順（台帳フォームの順に近い）
CSV_COLUMNS = [
    "契約No", "社員番号", "氏名", "性別", "生年月日", "年齢", "所属グループ",
    "協定対象派遣労働者であるか否かの別", "有期か無期かの別", "60歳以上の者であるか否かの別",
    "派遣先名称", "派遣先の事業所所在地", "派遣先TEL", "就業場所", "組織単位", "業務内容", "責任の程度", "職種",
    "契約期間", "就業時間", "休憩時間1", "休憩時間2", "就業曜日", "就業曜日備考", "休日", "休日備考",
    "就業時間外の労働", "休日勤務",
    "派遣先責任者_部署", "派遣先責任者_役職氏名", "派遣先責任者_TEL",
    "苦情申出先_部署", "苦情申出先_役職氏名", "苦情申出先_TEL",
    "指揮命令者_部署", "指揮命令者_役職氏名", "指揮命令者_TEL",
    "派遣元責任者_部署", "派遣元責任者_役職氏名", "派遣元責任者_TEL",
    "派遣元苦情処理担当者_部署", "派遣元苦情処理担当者_役職氏名", "派遣元苦情処理担当者_TEL",
    "期間制限の対象外理由", "事業所単位の抵触日", "個人単位の抵触日",
    "健康保険_加入", "健康保険_未加入理由", "厚生年金_加入", "厚生年金_未加入理由", "雇用保険_加入", "雇用保険_未加入理由",
    "便宜供与", "苦情の処理状況", "契約書確定日", "派遣許可番号", "契約書備考",
]
