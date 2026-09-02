# -*- coding: utf-8 -*-
"""健康診断申込: 「選択肢」シート（健診機関・健診種別・追加検査・続柄）の解釈。

正本は Google の「選択肢」シート。HPM の場所コード／健診種別コードは前ゼロ・英字X を
含むので、コードは一切数値化せず文字列のまま扱う。

「別名」列は jinjer 側のテキスト表記（例: 前年度の健診機関名、健診内容の選択肢名）を
コードへ寄せるための変換表を兼ねる。コードを触らずシートで直せるようにしている。
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass

from services.health_apply.schema import OPTION_HEADERS, SchemaError

KIND_INSTITUTION = "機関"
KIND_EXAM_TYPE = "種別"
KIND_EXTRA = "追加検査"
KIND_RELATIONSHIP = "続柄"
KINDS = (KIND_INSTITUTION, KIND_EXAM_TYPE, KIND_EXTRA, KIND_RELATIONSHIP)

# 「その他」の機関コード（HPM の場所コードと衝突しない予約語）
OTHER_INSTITUTION_CODE = "OTHER"
# 婦人科検診（HPM の健診種別には無い追加検査）
GYN_CODE = "GYN"

# jinjer の旧表記 → 画面表記
LEGACY_EXTRA_NAMES = {
    "婦人病検査": "婦人科検診",
    "婦人科健診": "婦人科検診",
}

_ACTIVE_FALSE = {"0", "false", "no", "無効", "×", "x"}
_ALIAS_SEPARATORS = (";", "；", "\n")


def normalize_key(text: str | None) -> str:
    """照合キー。NFKC → 空白（全角含む）を全部除去 → 小文字。表示には使わない。"""
    s = unicodedata.normalize("NFKC", str(text or ""))
    return "".join(ch for ch in s if not ch.isspace()).casefold()


def normalize_extra_name(name: str | None) -> str:
    """追加検査の旧表記（婦人病検査 など）を画面表記に寄せる。"""
    s = str(name or "").strip()
    return LEGACY_EXTRA_NAMES.get(s, s)


def _parse_active(value: str) -> bool:
    s = str(value or "").strip().casefold()
    if not s:
        return True  # 空欄は有効扱い（貼り付け直後に列を埋め忘れても選べるように）
    return s not in _ACTIVE_FALSE


def _parse_order(value: str, row_no: int) -> int:
    s = unicodedata.normalize("NFKC", str(value or "")).strip()
    if not s:
        return 9999
    try:
        return int(s)
    except ValueError as e:
        raise SchemaError(f"選択肢シート{row_no}行目の並び順が数値ではありません: 「{value}」") from e


def _split_aliases(value: str) -> tuple[str, ...]:
    s = str(value or "")
    for sep in _ALIAS_SEPARATORS[1:]:
        s = s.replace(sep, _ALIAS_SEPARATORS[0])
    return tuple(a.strip() for a in s.split(_ALIAS_SEPARATORS[0]) if a.strip())


@dataclass(frozen=True)
class Option:
    kind: str
    code: str
    name: str
    active: bool = True
    order: int = 9999
    aliases: tuple[str, ...] = ()
    note: str = ""

    def as_dict(self) -> dict:
        return {"kind": self.kind, "code": self.code, "name": self.name,
                "active": self.active, "order": self.order}


class OptionCatalog:
    """区分ごとの選択肢。lookup はコードの文字列一致、resolve_name は表示名・別名のゆれ吸収。"""

    def __init__(self, options: list[Option]):
        self._options = list(options)
        self._by_code: dict[tuple[str, str], Option] = {}
        for opt in self._options:
            key = (opt.kind, opt.code)
            if key in self._by_code:
                raise SchemaError(f"選択肢シートで区分「{opt.kind}」のコード「{opt.code}」が重複しています")
            self._by_code[key] = opt
        self._by_name: dict[tuple[str, str], list[Option]] = {}
        for opt in self._options:
            names = [opt.name, *opt.aliases]
            if opt.kind == KIND_EXTRA:
                names.append(normalize_extra_name(opt.name))
            for n in names:
                k = normalize_key(n)
                if k:
                    self._by_name.setdefault((opt.kind, k), []).append(opt)

    @classmethod
    def from_rows(cls, rows: list[dict[str, str]]) -> "OptionCatalog":
        """rows_to_dicts(OPTION_HEADERS, ...) の結果から作る。"""
        options: list[Option] = []
        for i, row in enumerate(rows, start=2):  # 1行目はヘッダー
            kind = str(row.get(OPTION_HEADERS[0], "")).strip()
            code = str(row.get(OPTION_HEADERS[1], "")).strip()
            name = str(row.get(OPTION_HEADERS[2], "")).strip()
            if not kind and not code and not name:
                continue
            if kind not in KINDS:
                raise SchemaError(
                    f"選択肢シート{i}行目の区分が想定外です: 「{kind}」（使えるのは "
                    + "・".join(KINDS) + "）")
            if not code:
                raise SchemaError(f"選択肢シート{i}行目のコードが空です（表示名: {name}）")
            if not name:
                raise SchemaError(f"選択肢シート{i}行目の表示名が空です（コード: {code}）")
            options.append(Option(
                kind=kind, code=code, name=name,
                active=_parse_active(row.get(OPTION_HEADERS[3], "")),
                order=_parse_order(row.get(OPTION_HEADERS[4], ""), i),
                aliases=_split_aliases(row.get(OPTION_HEADERS[5], "")),
                note=str(row.get(OPTION_HEADERS[6], "")).strip(),
            ))
        return cls(options)

    def all(self) -> list[Option]:
        return list(self._options)

    def of_kind(self, kind: str, active_only: bool = True) -> list[Option]:
        items = [o for o in self._options if o.kind == kind and (o.active or not active_only)]
        return sorted(items, key=lambda o: (o.order, o.name))

    def lookup(self, kind: str, code: str | None) -> Option | None:
        return self._by_code.get((kind, str(code or "").strip()))

    def resolve_name(self, kind: str, text: str | None) -> Option | None:
        """表示名・別名で引く。有効なものを優先し、複数なら並び順が先のもの。"""
        if kind == KIND_EXTRA:
            text = normalize_extra_name(text)
        hits = self._by_name.get((kind, normalize_key(text)), [])
        if not hits:
            return None
        return sorted(hits, key=lambda o: (not o.active, o.order, o.name))[0]

    def display(self, kind: str, code: str | None) -> str:
        opt = self.lookup(kind, code)
        if opt is None:
            return f"（不明: {str(code or '').strip() or '空'}）"
        return opt.name

    def counts(self) -> dict[str, int]:
        return {
            "institutions": len(self.of_kind(KIND_INSTITUTION)),
            "exam_types": len(self.of_kind(KIND_EXAM_TYPE)),
            "extras": len(self.of_kind(KIND_EXTRA)),
            "relationships": len(self.of_kind(KIND_RELATIONSHIP)),
        }
