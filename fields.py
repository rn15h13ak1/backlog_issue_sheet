"""
Excel の列と Backlog の項目の対応
================================
見出しの名前で列を特定し、セルの値を「正規化した値」に変換する。

正規化した値は、Excel から読んだ値と Backlog の現在値を同じ形で比べるために使う。

| 種類   | 正規化した値            |
|--------|-------------------------|
| text   | str                     |
| name   | int（ID）               |
| names  | tuple[int, ...]（昇順） |
| date   | "YYYY-MM-DD"            |
| hours / number | float           |

値が無いことは None で表す。`(削除)` と書かれたセルは CLEAR になる。
"""

from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass

from master import (
    CF_DATE,
    CF_NUMBER,
    CF_SENTENCE,
    CF_TEXT,
    CustomField,
    Master,
    ResolveError,
)

REGISTER_SHEET = "登録"
UPDATE_SHEET = "更新"

KEY_HEADER = "課題キー"
PARENT_HEADER = "親課題キー"
CHILD_HEADER = "子課題"

CLEAR_TOKEN = "(削除)"
DETACH_TOKEN = "(解除)"


def is_token(raw, token: str) -> bool:
    """
    セルの値が (削除) / (解除) の印か。

    人手で書くと、日本語入力のまま全角の「（削除）」になったり、空白が混ざったりしやすい。
    全角・半角の違い（NFKC で半角にそろえる）と空白は無視して比べる。
    括弧の無い「削除」は受け付けない（担当者などの名前として読む）。
    """
    if not isinstance(raw, str):
        return False
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", raw)) == token

# 子課題列で「子にする」と読む値。これ以外の値はエラーにする。
# 「×」「-」などを「子ではない」の意味で書く人がいるため、何でも受け付けると誤読する。
CHILD_MARKS = {"○", "〇", "◯", "o", "O", "1", "true", "TRUE", "True", "yes", "y", "Y"}
CHILD_MARK = "○"

# 見出しがこの文字で始まる列は読み飛ばす（メモ用）
COMMENT_PREFIX = "#"

# Backlog の課題キー（例: PROJ-123, 10_AAA-20）
ISSUE_KEY_RE = re.compile(r"^[A-Za-z0-9_]+-\d+$")

# 複数の値を書くときの区切り
MULTI_SEPARATORS = re.compile(r"[\n,、]")


class _Clear:
    """`(削除)` を表す目印。"""

    def __repr__(self) -> str:
        return "CLEAR"


CLEAR = _Clear()


class ValueProblem(ValueError):
    """セルの値を変換できなかった。"""


@dataclass(frozen=True)
class Column:
    header: str
    key: str                       # 差分・送信の識別子（API パラメータ名）
    kind: str                      # text / name / names / date / hours / number
    table: str | None = None       # Master.tables のキー
    required: bool = False         # 作成時に必須
    clearable: bool = True         # (削除) で消せる
    on_create: bool = True         # 作成時に送れる（False なら作成直後に更新で送る）
    custom: CustomField | None = None

    @property
    def is_custom(self) -> bool:
        return self.custom is not None


STANDARD_COLUMNS: list[Column] = [
    Column("件名", "summary", "text", required=True, clearable=False),
    Column("詳細", "description", "text"),
    Column("種別", "issueTypeId", "name", "issue_types", required=True, clearable=False),
    Column("優先度", "priorityId", "name", "priorities", required=True, clearable=False),
    # 状態と完了理由は、Backlog の作成 API では指定できない
    Column("状態", "statusId", "name", "statuses", clearable=False, on_create=False),
    Column("完了理由", "resolutionId", "name", "resolutions", on_create=False),
    Column("担当者", "assigneeId", "name", "users"),
    Column("カテゴリー", "categoryId", "names", "categories"),
    Column("発生バージョン", "versionId", "names", "versions"),
    Column("マイルストーン", "milestoneId", "names", "versions"),
    Column("開始日", "startDate", "date"),
    Column("期限日", "dueDate", "date"),
    Column("予定時間", "estimatedHours", "hours"),
    Column("実績時間", "actualHours", "hours"),
]

STANDARD_BY_HEADER = {c.header: c for c in STANDARD_COLUMNS}
STRUCTURE_HEADERS = (KEY_HEADER, PARENT_HEADER, CHILD_HEADER)


def custom_column(cf: CustomField) -> Column:
    if cf.type_id in (CF_TEXT, CF_SENTENCE):
        kind = "text"
    elif cf.type_id == CF_NUMBER:
        kind = "number"
    elif cf.type_id == CF_DATE:
        kind = "date"
    elif cf.is_multi:
        kind = "names"
    else:
        kind = "name"
    return Column(cf.name, f"customField_{cf.id}", kind, custom=cf)


def all_columns(master: Master) -> list[Column]:
    """標準の項目とカスタム属性をすべて並べる（ひな形・書き出しの列順）。"""
    return STANDARD_COLUMNS + [custom_column(cf) for cf in master.custom_fields]


def column_for_header(header: str, master: Master) -> Column | None:
    """見出しから列を引く。標準の項目を優先し、無ければカスタム属性の名前として探す。"""
    if header in STANDARD_BY_HEADER:
        return STANDARD_BY_HEADER[header]
    cf = master.custom_field_by_name(header)
    return custom_column(cf) if cf else None


# ---------------------------------------------------------------------------
# セル → 正規化した値
# ---------------------------------------------------------------------------

def is_blank(raw) -> bool:
    return raw is None or (isinstance(raw, str) and raw.strip() == "")


def cell_text(raw) -> str:
    """セルの値を文字列にする。整数の float は小数点を付けない。"""
    if raw is None:
        return ""
    if isinstance(raw, bool):
        return "TRUE" if raw else "FALSE"
    if isinstance(raw, float) and raw.is_integer():
        return str(int(raw))
    if isinstance(raw, dt.datetime):
        if raw.time() == dt.time(0, 0):
            return raw.date().isoformat()
        return raw.isoformat(sep=" ")
    if isinstance(raw, dt.date):
        return raw.isoformat()
    return str(raw)


def normalize_text(text: str) -> str:
    """改行コードを揃え、前後の空白を除く。Backlog の現在値と比べるときも同じ処理をする。"""
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


_DATE_RE = re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$")


def parse_date(raw) -> str:
    if isinstance(raw, dt.datetime):
        return raw.date().isoformat()
    if isinstance(raw, dt.date):
        return raw.isoformat()
    if isinstance(raw, (int, float)):
        # 書式の付いていない日付セルはシリアル値（数値）で読まれる
        raise ValueProblem(
            f"日付として読めません（{cell_text(raw)}）。セルの書式を日付にするか、"
            "2026-10-31 の形で書いてください"
        )
    text = str(raw).strip()
    m = _DATE_RE.match(text)
    if not m:
        raise ValueProblem(f"日付として読めません（{text}）。2026-10-31 の形で書いてください")
    try:
        return dt.date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    except ValueError:
        raise ValueProblem(f"存在しない日付です（{text}）") from None


def parse_number(raw, *, non_negative: bool) -> float:
    if isinstance(raw, bool):
        raise ValueProblem(f"数値として読めません（{cell_text(raw)}）")
    if isinstance(raw, (int, float)):
        value = float(raw)
    else:
        text = str(raw).strip().replace(",", "")
        try:
            value = float(text)
        except ValueError:
            raise ValueProblem(f"数値として読めません（{text}）") from None
    if non_negative and value < 0:
        raise ValueProblem(f"負の値は書けません（{cell_text(raw)}）")
    return round(value, 2)


def _resolve(col: Column, name: str, master: Master) -> int:
    table = col.custom.items if col.is_custom else master.table(col.table)
    return table.resolve(name)


def _has_name(col: Column, name: str, master: Master) -> bool:
    table = col.custom.items if col.is_custom else master.table(col.table)
    return table.has(name.strip())


def split_names(text: str) -> list[str]:
    """
    複数の名前に分ける。改行があれば改行だけで分ける。

    書き出しは改行で区切るため、名前に「,」を含んでいても取り込みに戻せる。
    改行が無いときだけ「,」「、」でも分ける（手で書くときに Alt+Enter は面倒なため）。
    """
    pattern = r"\n" if "\n" in text else MULTI_SEPARATORS
    return [p.strip() for p in re.split(pattern, text) if p.strip()]


def parse_value(col: Column, raw, master: Master):
    """
    セルの値を正規化した値にする。空なら None、`(削除)` なら CLEAR。

    変換できなければ ValueProblem を送出する（名前が無い場合も含む）。
    """
    if is_blank(raw):
        return None
    if is_token(raw, CLEAR_TOKEN):
        return CLEAR

    try:
        if col.kind == "text":
            return normalize_text(cell_text(raw))
        if col.kind == "date":
            return parse_date(raw)
        if col.kind == "hours":
            return parse_number(raw, non_negative=True)
        if col.kind == "number":
            return parse_number(raw, non_negative=False)
        text = cell_text(raw).strip()
        if col.kind == "name":
            return _resolve(col, text, master)
        if col.kind == "names":
            # 名前そのものに区切り文字（「,」など）を含む場合に備え、全体で一致すれば 1 件とみなす
            if _has_name(col, text, master):
                return (_resolve(col, text, master),)
            return tuple(sorted({_resolve(col, n, master) for n in split_names(text)}))
    except ResolveError as e:
        raise ValueProblem(str(e)) from None
    raise AssertionError(f"未知の列の種類: {col.kind}")


# ---------------------------------------------------------------------------
# Backlog の課題 → 正規化した値
# ---------------------------------------------------------------------------

def _id_of(obj) -> int | None:
    return obj.get("id") if isinstance(obj, dict) else None


def _ids_of(objs) -> tuple:
    return tuple(sorted(o["id"] for o in (objs or []) if isinstance(o, dict) and "id" in o))


def _float_or_none(value) -> float | None:
    if value is None or value == "":
        return None
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def current_value(col: Column, issue: dict):
    """課題の現在値を正規化した値で返す。値が無ければ None。"""
    if col.is_custom:
        return _current_custom_value(col, issue)

    key = col.key
    if key == "summary":
        return normalize_text(issue.get("summary") or "") or None
    if key == "description":
        return normalize_text(issue.get("description") or "") or None
    if key == "issueTypeId":
        return _id_of(issue.get("issueType"))
    if key == "priorityId":
        return _id_of(issue.get("priority"))
    if key == "statusId":
        return _id_of(issue.get("status"))
    if key == "resolutionId":
        return _id_of(issue.get("resolution"))
    if key == "assigneeId":
        return _id_of(issue.get("assignee"))
    if key == "categoryId":
        return _ids_of(issue.get("category")) or None
    if key == "versionId":
        return _ids_of(issue.get("versions")) or None
    if key == "milestoneId":
        return _ids_of(issue.get("milestone")) or None
    if key in ("startDate", "dueDate"):
        value = issue.get(key)
        return value[:10] if value else None
    if key in ("estimatedHours", "actualHours"):
        return _float_or_none(issue.get(key))
    raise AssertionError(f"未知の項目: {key}")


def _current_custom_value(col: Column, issue: dict):
    cf = col.custom
    entry = next(
        (c for c in issue.get("customFields") or [] if c.get("id") == cf.id), None
    )
    value = entry.get("value") if entry else None
    if value is None or value == "" or value == []:
        return None
    if col.kind == "text":
        return normalize_text(str(value)) or None
    if col.kind == "number":
        return _float_or_none(value)
    if col.kind == "date":
        return str(value)[:10]
    if col.kind == "name":
        return _id_of(value)
    if col.kind == "names":
        return _ids_of(value if isinstance(value, list) else [value]) or None
    raise AssertionError(f"未知の列の種類: {col.kind}")


# ---------------------------------------------------------------------------
# 正規化した値 → API パラメータ・表示・セル
# ---------------------------------------------------------------------------

def format_number(value: float) -> str:
    return f"{value:g}" if abs(value) < 1e15 else str(value)


def to_param(col: Column, value):
    """
    API に送る値にする。CLEAR は空文字（複数値は [""]）で送る。

    複数値の空は `categoryId[]=` の形になり、Backlog はこれを「全部外す」と解釈する。
    空のリストを送ると何も送られず、値が消えない。
    """
    if value is CLEAR:
        return [""] if col.kind == "names" else ""
    if col.kind == "names":
        return list(value)
    if col.kind in ("hours", "number"):
        return format_number(value)
    return value


def display(col: Column, value, master: Master) -> str:
    """計画の表示用の文字列。"""
    if value is None:
        return "（なし）"
    if value is CLEAR:
        # セルに書く印と同じ形で示す（全角で示すと、それをまねて書かれる）
        return CLEAR_TOKEN
    if col.kind == "name":
        return _name_of(col, value, master)
    if col.kind == "names":
        return "、".join(_name_of(col, v, master) for v in value)
    if col.kind in ("hours", "number"):
        return format_number(value)
    if col.kind == "text":
        one_line = value.replace("\n", "⏎")
        return one_line if len(one_line) <= 40 else one_line[:39] + "…"
    return str(value)


def _name_of(col: Column, item_id: int, master: Master) -> str:
    table = col.custom.items if col.is_custom else master.table(col.table)
    return table.name_of(item_id)


def to_cell(col: Column, value, master: Master):
    """Excel に書く値。取り込みに戻したとき、同じ正規化した値になるようにする。"""
    if value is None:
        return None
    if col.kind == "name":
        return _name_of(col, value, master)
    if col.kind == "names":
        return "\n".join(_name_of(col, v, master) for v in value)
    if col.kind == "date":
        return dt.date.fromisoformat(value)
    if col.kind in ("hours", "number"):
        return int(value) if float(value).is_integer() else value
    return value
