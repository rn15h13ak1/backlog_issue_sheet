"""
計画（検証・親子の決定・差分）
==============================
Excel から読んだ行を検証し、「何を作り、何をどう変えるか」を決める。
ここでは Backlog に書き込まない。読むのは、渡された client の
get_issue / get_child_issues だけ（テストでは偽物に差し替える）。

エラーが 1 件でもあれば、実行側は何も送らない。全件のエラーを一度に
示すため、途中で止めずに最後まで検証する。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from fields import (
    CHILD_HEADER,
    CHILD_MARKS,
    CLEAR,
    DETACH_TOKEN,
    ISSUE_KEY_RE,
    KEY_HEADER,
    PARENT_HEADER,
    REGISTER_SHEET,
    STANDARD_BY_HEADER,
    STANDARD_COLUMNS,
    STRUCTURE_HEADERS,
    UPDATE_SHEET,
    Column,
    ValueProblem,
    cell_text,
    column_for_header,
    current_value,
    is_blank,
    parse_value,
)
from master import STATUS_OPEN_ID, Master, suggest
from sheet_io import SheetData


@dataclass
class Problem:
    sheet: str
    row: int | None
    column: str | None
    message: str

    def __str__(self) -> str:
        where = f"{self.sheet}"
        if self.row is not None:
            where += f" {self.row}行目"
        if self.column:
            where += f"「{self.column}」"
        return f"{where}: {self.message}"


@dataclass
class CreatePlan:
    row_no: int
    summary: str
    values: dict[str, object]            # 作成時に送る {列のキー: 正規化した値}
    follow_up: dict[str, object]         # 作成の直後に更新で送る（状態・完了理由）
    columns: dict[str, Column] = field(default_factory=dict)
    parent_row: int | None = None        # 同じシートの親の行番号
    parent_key: str | None = None        # 既存の親の課題キー
    parent_issue_id: int | None = None   # 既存の親の課題 ID


@dataclass
class Change:
    column: Column
    old: object
    new: object


@dataclass
class UpdatePlan:
    row_no: int
    issue_key: str
    issue: dict
    changes: list[Change] = field(default_factory=list)
    # None: 変えない / ("set", 親の課題, 今の親の課題キー) / ("detach", None, 今の親の課題キー)
    parent_change: tuple | None = None

    @property
    def summary(self) -> str:
        return self.issue.get("summary", "")

    @property
    def has_changes(self) -> bool:
        return bool(self.changes) or self.parent_change is not None

    @property
    def phase(self) -> int:
        """
        実行の順番。親を付ける行は後に回す。

        親から外す行を先に実行しないと、「子課題を親にする」形になって
        Backlog に拒否されることがある（例: A を親から外し、B を A の子にする）。
        """
        return 2 if self.parent_change and self.parent_change[0] == "set" else 1


@dataclass
class Plan:
    creates: list[CreatePlan] = field(default_factory=list)
    updates: list[UpdatePlan] = field(default_factory=list)
    problems: list[Problem] = field(default_factory=list)
    warnings: list[Problem] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.problems


class IssueCache:
    """同じ課題を何度も取りにいかないためのキャッシュ。"""

    def __init__(self, client, project_id: int):
        self.client = client
        self.project_id = project_id
        self._by_key: dict[str, dict | None] = {}
        self._children: dict[int, list] = {}

    def get(self, key: str) -> dict | None:
        if key not in self._by_key:
            self._by_key[key] = self.client.get_issue(key)
        return self._by_key[key]

    def children(self, issue_id: int) -> list:
        if issue_id not in self._children:
            self._children[issue_id] = self.client.get_child_issues(self.project_id, issue_id)
        return self._children[issue_id]

    def key_of_id(self, issue_id: int | None) -> str | None:
        if issue_id is None:
            return None
        for issue in self._by_key.values():
            if issue and issue.get("id") == issue_id:
                return issue.get("issueKey")
        issue = self.client.get_issue(str(issue_id))
        if issue:
            self._by_key[issue["issueKey"]] = issue
            return issue["issueKey"]
        return str(issue_id)


# ---------------------------------------------------------------------------
# 見出し
# ---------------------------------------------------------------------------

def map_headers(
    sheet: SheetData, master: Master, problems: list[Problem]
) -> dict[int, Column | str]:
    """
    列番号 → Column（標準・カスタム属性）または構造の見出し（課題キーなど）。

    使えない見出しは problems に積む。
    """
    is_register = sheet.name == REGISTER_SHEET
    mapping: dict[int, Column | str] = {}
    seen: dict[str, int] = {}
    custom_names = [cf.name for cf in master.custom_fields]

    for idx, header in sheet.headers:
        if header in seen:
            problems.append(Problem(sheet.name, 1, header, "同じ見出しの列が 2 つあります"))
            continue
        seen[header] = idx

        if header in STRUCTURE_HEADERS:
            if is_register and header == KEY_HEADER:
                problems.append(Problem(
                    sheet.name, 1, header,
                    f"{REGISTER_SHEET}シートには{KEY_HEADER}の列を置けません。"
                    f"既存の課題を変えるときは{UPDATE_SHEET}シートを使ってください",
                ))
                continue
            if not is_register and header == CHILD_HEADER:
                problems.append(Problem(
                    sheet.name, 1, header,
                    f"{UPDATE_SHEET}シートでは{CHILD_HEADER}の列は使えません。"
                    f"親は{PARENT_HEADER}に課題キーで書いてください",
                ))
                continue
            mapping[idx] = header
            continue

        col = column_for_header(header, master)
        if col is None:
            message = (
                "標準の項目にも、カスタム属性にも無い見出しです。"
                "メモ用の列なら見出しを # で始めてください"
            )
            hint = suggest(header, custom_names + [c.header for c in STANDARD_COLUMNS] + list(STRUCTURE_HEADERS))
            if hint:
                message += f"。{hint}"
            problems.append(Problem(sheet.name, 1, header, message))
            continue
        mapping[idx] = col

    present = {h for _, h in sheet.headers}
    if is_register:
        for required in ("件名", "種別", "優先度"):
            if required not in present:
                problems.append(Problem(sheet.name, 1, None, f"必須の列「{required}」がありません"))
    elif KEY_HEADER not in present:
        problems.append(Problem(sheet.name, 1, None, f"必須の列「{KEY_HEADER}」がありません"))
    return mapping


def _parse_row(
    sheet: SheetData, row, mapping, master: Master, problems: list[Problem]
) -> tuple[dict[str, object], dict[str, Column], dict[str, object]]:
    """
    1 行を読む。返り値は ({列のキー: 値}, {列のキー: Column}, {構造の見出し: 生の値})。

    値が None（空のセル）の列は含めない。
    """
    values: dict[str, object] = {}
    columns: dict[str, Column] = {}
    structure: dict[str, object] = {}
    for idx, target in mapping.items():
        raw = row.values.get(idx)
        if isinstance(target, str):
            structure[target] = raw
            continue
        try:
            value = parse_value(target, raw, master)
        except ValueProblem as e:
            problems.append(Problem(sheet.name, row.row_no, target.header, str(e)))
            continue
        if value is None:
            continue
        values[target.key] = value
        columns[target.key] = target
    return values, columns, structure


def _check_custom_applicability(
    sheet: SheetData, row_no: int, columns: dict[str, Column], values: dict,
    issue_type_id: int | None, master: Master, problems: list[Problem],
) -> None:
    if issue_type_id is None:
        return
    type_name = master.table("issue_types").name_of(issue_type_id)
    for key, col in columns.items():
        if col.is_custom and values.get(key) is not CLEAR and not col.custom.applies_to(issue_type_id):
            problems.append(Problem(
                sheet.name, row_no, col.header,
                f"このカスタム属性は種別「{type_name}」では使えません",
            ))


def _parse_issue_key(
    raw, sheet: SheetData, row_no: int, column: str, master: Master, problems: list[Problem]
) -> str | None:
    key = cell_text(raw).strip().upper()
    if not ISSUE_KEY_RE.match(key):
        problems.append(Problem(sheet.name, row_no, column, f"課題キーの形ではありません（{key}）"))
        return None
    prefix = master.project_key.upper() + "-"
    if master.project_key and not key.startswith(prefix):
        problems.append(Problem(
            sheet.name, row_no, column,
            f"別のプロジェクトの課題です（{key}）。扱えるのは {master.project_key} の課題だけです",
        ))
        return None
    return key


# ---------------------------------------------------------------------------
# 登録シート
# ---------------------------------------------------------------------------

def plan_register(sheet: SheetData, master: Master, issues: IssueCache, plan: Plan) -> None:
    problems = plan.problems
    mapping = map_headers(sheet, master, problems)
    # 行ごとの「子かどうか」。親を探すとき、エラーのある行も含めて数える
    # （エラーの行を飛ばすと、直すまで別の行が親として選ばれ、表示が食い違う）。
    is_child_by_index: list[bool] = []
    rows = sheet.rows
    uses_parent = False

    for i, row in enumerate(rows):
        values, columns, structure = _parse_row(sheet, row, mapping, master, problems)

        for key, value in values.items():
            if value is CLEAR:
                problems.append(Problem(
                    sheet.name, row.row_no, columns[key].header,
                    f"{REGISTER_SHEET}シートでは (削除) は使えません。値を入れないときは空にしてください",
                ))

        for required_key, header in (("summary", "件名"), ("issueTypeId", "種別"), ("priorityId", "優先度")):
            if required_key not in values and any(h == header for _, h in sheet.headers):
                problems.append(Problem(sheet.name, row.row_no, header, "必須の項目が空です"))

        issue_type_id = values.get("issueTypeId")
        _check_custom_applicability(sheet, row.row_no, columns, values, issue_type_id, master, problems)
        if isinstance(issue_type_id, int):
            for cf in master.custom_fields:
                if cf.required and cf.applies_to(issue_type_id) and f"customField_{cf.id}" not in values:
                    problems.append(Problem(
                        sheet.name, row.row_no, cf.name,
                        "必須のカスタム属性が空です"
                        + ("" if any(h == cf.name for _, h in sheet.headers) else "（列もありません）"),
                    ))

        # ---- 親子 ----
        child_raw = structure.get(CHILD_HEADER)
        is_child_flag = False
        if not is_blank(child_raw):
            mark = cell_text(child_raw).strip()
            if mark in CHILD_MARKS:
                is_child_flag = True
            else:
                problems.append(Problem(
                    sheet.name, row.row_no, CHILD_HEADER,
                    f"「{mark}」は読めません。子課題にするなら ○ を、しないなら空にしてください",
                ))

        parent_key = None
        parent_raw = structure.get(PARENT_HEADER)
        if not is_blank(parent_raw):
            if cell_text(parent_raw).strip() == DETACH_TOKEN:
                problems.append(Problem(
                    sheet.name, row.row_no, PARENT_HEADER,
                    f"{DETACH_TOKEN} は{UPDATE_SHEET}シートでだけ使えます",
                ))
            else:
                parent_key = _parse_issue_key(parent_raw, sheet, row.row_no, PARENT_HEADER, master, problems)
            # 親課題キーを書いた行は、それ自体が子課題になる
            is_child_by_index.append(True)
        else:
            is_child_by_index.append(is_child_flag)

        if is_child_flag and not is_blank(parent_raw):
            problems.append(Problem(
                sheet.name, row.row_no, None,
                f"{PARENT_HEADER}と{CHILD_HEADER}の両方が書かれています。どちらか一方にしてください",
            ))

        parent_row = None
        if is_child_flag and is_blank(parent_raw):
            uses_parent = True
            skipped_existing_child = None
            for j in range(i - 1, -1, -1):
                if not is_child_by_index[j]:
                    parent_row = rows[j].row_no
                    break
                if skipped_existing_child is None and _has_parent_key(rows[j], mapping):
                    skipped_existing_child = rows[j].row_no
            if parent_row is None:
                problems.append(Problem(
                    sheet.name, row.row_no, CHILD_HEADER,
                    f"上に親になれる行（{CHILD_HEADER}の印も{PARENT_HEADER}も無い行）がありません",
                ))
            elif skipped_existing_child is not None:
                plan.warnings.append(Problem(
                    sheet.name, row.row_no, CHILD_HEADER,
                    f"{skipped_existing_child}行目は既存課題の子課題のため親にできません。"
                    f"さらに上の {parent_row}行目を親にします",
                ))

        parent_issue_id = None
        if parent_key:
            uses_parent = True
            parent = issues.get(parent_key)
            if parent:
                parent_issue_id = parent.get("id")
            if parent is None:
                problems.append(Problem(sheet.name, row.row_no, PARENT_HEADER, f"{parent_key} が見つかりません"))
            elif parent.get("parentIssueId"):
                problems.append(Problem(
                    sheet.name, row.row_no, PARENT_HEADER,
                    f"{parent_key} は子課題のため親にできません（Backlog の親子は 1 階層まで）",
                ))

        create_values = {k: v for k, v in values.items() if columns[k].on_create}
        follow_up = {k: v for k, v in values.items() if not columns[k].on_create}
        if follow_up.get("statusId") == STATUS_OPEN_ID:
            del follow_up["statusId"]     # 作成直後の状態と同じなので更新しない

        plan.creates.append(CreatePlan(
            row_no=row.row_no,
            summary=str(values.get("summary", "")),
            values=create_values,
            follow_up=follow_up,
            columns=columns,
            parent_row=parent_row,
            parent_key=parent_key,
            parent_issue_id=parent_issue_id,
        ))

    if uses_parent and not master.subtasking_enabled:
        problems.append(Problem(
            sheet.name, None, None,
            "このプロジェクトでは親子課題が無効になっています（プロジェクト設定で有効にしてください）",
        ))


def _has_parent_key(row, mapping) -> bool:
    for idx, target in mapping.items():
        if target == PARENT_HEADER and not is_blank(row.values.get(idx)):
            return True
    return False


# ---------------------------------------------------------------------------
# 更新シート
# ---------------------------------------------------------------------------

def plan_update(sheet: SheetData, master: Master, issues: IssueCache, plan: Plan) -> None:
    problems = plan.problems
    mapping = map_headers(sheet, master, problems)
    seen_keys: dict[str, int] = {}
    planned: list[UpdatePlan] = []

    for row in sheet.rows:
        values, columns, structure = _parse_row(sheet, row, mapping, master, problems)

        key_raw = structure.get(KEY_HEADER)
        if is_blank(key_raw):
            problems.append(Problem(sheet.name, row.row_no, KEY_HEADER, "課題キーが空です"))
            continue
        key = _parse_issue_key(key_raw, sheet, row.row_no, KEY_HEADER, master, problems)
        if key is None:
            continue
        if key in seen_keys:
            problems.append(Problem(
                sheet.name, row.row_no, KEY_HEADER, f"{key} は {seen_keys[key]}行目にもあります",
            ))
            continue
        seen_keys[key] = row.row_no

        for k, value in values.items():
            if value is CLEAR and not columns[k].clearable:
                problems.append(Problem(
                    sheet.name, row.row_no, columns[k].header, "この項目は (削除) で消せません",
                ))

        issue = issues.get(key)
        if issue is None:
            problems.append(Problem(sheet.name, row.row_no, KEY_HEADER, f"{key} が見つかりません"))
            continue

        new_type = values.get("issueTypeId")
        final_type = new_type if isinstance(new_type, int) else current_value(
            STANDARD_BY_HEADER["種別"], issue
        )
        _check_custom_applicability(sheet, row.row_no, columns, values, final_type, master, problems)

        up = UpdatePlan(row_no=row.row_no, issue_key=key, issue=issue)
        for k, new in values.items():
            col = columns[k]
            old = current_value(col, issue)
            if new is CLEAR:
                if old is not None:
                    up.changes.append(Change(col, old, CLEAR))
            elif new != old:
                up.changes.append(Change(col, old, new))

        # ---- 親子 ----
        parent_raw = structure.get(PARENT_HEADER)
        current_parent_id = issue.get("parentIssueId")
        if not is_blank(parent_raw):
            text = cell_text(parent_raw).strip()
            if text == DETACH_TOKEN:
                if current_parent_id:
                    up.parent_change = ("detach", None, issues.key_of_id(current_parent_id))
            else:
                parent_key = _parse_issue_key(parent_raw, sheet, row.row_no, PARENT_HEADER, master, problems)
                if parent_key == key:
                    problems.append(Problem(sheet.name, row.row_no, PARENT_HEADER, "自分自身は親にできません"))
                elif parent_key:
                    parent = issues.get(parent_key)
                    if parent is None:
                        problems.append(Problem(
                            sheet.name, row.row_no, PARENT_HEADER, f"{parent_key} が見つかりません",
                        ))
                    elif parent.get("id") != current_parent_id:
                        up.parent_change = ("set", parent, issues.key_of_id(current_parent_id))
        planned.append(up)

    _check_final_hierarchy(sheet, planned, issues, problems)
    if any(u.parent_change and u.parent_change[0] == "set" for u in planned) and not master.subtasking_enabled:
        problems.append(Problem(
            sheet.name, None, None,
            "このプロジェクトでは親子課題が無効になっています（プロジェクト設定で有効にしてください）",
        ))
    plan.updates.extend(planned)


def _check_final_hierarchy(
    sheet: SheetData, planned: list[UpdatePlan], issues: IssueCache, problems: list[Problem]
) -> None:
    """
    実行後の親子関係が Backlog の制約（1 階層まで）を満たすか確かめる。

    行の順番によらず、シート全体を反映した後の状態で判定する。
    """
    final_parent: dict[int, int | None] = {}
    for up in planned:
        issue_id = up.issue["id"]
        if up.parent_change is None:
            final_parent[issue_id] = up.issue.get("parentIssueId")
        elif up.parent_change[0] == "detach":
            final_parent[issue_id] = None
        else:
            final_parent[issue_id] = up.parent_change[1]["id"]

    def parent_after(issue: dict) -> int | None:
        return final_parent.get(issue["id"], issue.get("parentIssueId"))

    for up in planned:
        if not up.parent_change or up.parent_change[0] != "set":
            continue
        parent = up.parent_change[1]
        if parent_after(parent):
            problems.append(Problem(
                sheet.name, up.row_no, PARENT_HEADER,
                f"{parent['issueKey']} は（この実行の後も）子課題のため親にできません"
                "（Backlog の親子は 1 階層まで）",
            ))

        me = up.issue["id"]
        remaining = [
            c["issueKey"] for c in issues.children(me) if parent_after(c) == me
        ]
        remaining += [
            other.issue_key for other in planned
            if other is not up and final_parent.get(other.issue["id"]) == me
            and other.issue_key not in remaining
        ]
        if remaining:
            problems.append(Problem(
                sheet.name, up.row_no, PARENT_HEADER,
                f"{up.issue_key} には子課題（{'、'.join(remaining[:5])}"
                f"{' ほか' if len(remaining) > 5 else ''}）があるため、子課題にできません",
            ))


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def build_plan(sheets: dict[str, SheetData], master: Master, client) -> Plan:
    plan = Plan()
    issues = IssueCache(client, master.project_id)
    if REGISTER_SHEET in sheets:
        plan_register(sheets[REGISTER_SHEET], master, issues, plan)
    if UPDATE_SHEET in sheets:
        plan_update(sheets[UPDATE_SHEET], master, issues, plan)
    return plan
