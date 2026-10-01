"""
書き出し（Backlog → Excel）
===========================
課題を「更新」シートの書式で書き出す。編集してそのまま取り込みに戻せる。

空のセルは「変更しない」と読まれるため、書き出した値を変えずに戻しても
差分は出ない。tests/test_roundtrip.py がこれを確かめている。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from fields import KEY_HEADER, PARENT_HEADER, all_columns, current_value, to_cell
from master import STATUS_CLOSED_ID, Master, ResolveError
from sheet_io import write_workbook

# Excel のセルに入る文字数の上限
EXCEL_CELL_LIMIT = 32767

URL_HEADER = "#URL"
UPDATED_HEADER = "#更新日時"
REFERENCE_HEADERS = [URL_HEADER, UPDATED_HEADER]


class ExportError(Exception):
    """書き出しの条件が不正。"""


@dataclass
class ExportFilter:
    statuses: list[str] = field(default_factory=list)   # 状態の名前
    types: list[str] = field(default_factory=list)      # 種別の名前
    keyword: str = ""
    parent_key: str = ""
    keys: list[str] = field(default_factory=list)
    include_closed: bool = False


def issue_url(base_url: str, issue_key: str) -> str:
    """API の URL（…/api/v2）から課題の画面の URL を作る。"""
    return re.sub(r"/api/v2$", "", base_url) + f"/view/{issue_key}"


def issue_to_row(
    issue: dict,
    master: Master,
    parent_key: str | None,
    *,
    base_url: str = "",
    warnings: list[str] | None = None,
) -> dict:
    """課題 1 件を {見出し: セルの値} にする。"""
    row = {KEY_HEADER: issue.get("issueKey"), PARENT_HEADER: parent_key}
    for col in all_columns(master):
        value = to_cell(col, current_value(col, issue), master)
        if isinstance(value, str) and len(value) > EXCEL_CELL_LIMIT:
            # 切り詰めて書くと、取り込みに戻したときに切り詰めた内容で上書きしてしまう。
            # 空のセルは「変更しない」なので、空にしておけば安全。
            if warnings is not None:
                warnings.append(
                    f"{issue.get('issueKey')}: {col.header}が Excel の上限"
                    f"（{EXCEL_CELL_LIMIT} 文字）を超えるため空にしました"
                )
            value = None
        row[col.header] = value
    if base_url:
        row[URL_HEADER] = issue_url(base_url, issue.get("issueKey", ""))
    updated = issue.get("updated")
    row[UPDATED_HEADER] = updated.replace("T", " ").replace("Z", "") if updated else None
    return row


def _key_number(issue: dict) -> int:
    try:
        return int(str(issue.get("issueKey", "")).rsplit("-", 1)[1])
    except (IndexError, ValueError):
        return 0


def order_issues(issues: list[dict]) -> list[dict]:
    """親の直後にその子が並ぶようにする。それ以外は課題番号の順。"""
    by_id = {i["id"]: i for i in issues}
    children: dict[int, list[dict]] = {}
    roots = []
    for issue in issues:
        parent_id = issue.get("parentIssueId")
        if parent_id and parent_id in by_id:
            children.setdefault(parent_id, []).append(issue)
        else:
            roots.append(issue)
    ordered = []
    for root in sorted(roots, key=_key_number):
        ordered.append(root)
        ordered.extend(sorted(children.get(root["id"], []), key=_key_number))
    return ordered


def _resolve_names(names: list[str], table) -> list[int]:
    try:
        return [table.resolve(n) for n in names]
    except ResolveError as e:
        raise ExportError(str(e)) from None


def fetch_issues(client, master: Master, flt: ExportFilter) -> list[dict]:
    """条件に合う課題を集める。子だけが当たった場合は、その親も加える。"""
    if flt.keys:
        issues = []
        for key in flt.keys:
            issue = client.get_issue(key)
            if issue is None:
                raise ExportError(f"{key} が見つかりません")
            issues.append(issue)
    else:
        params: dict = {}
        if flt.statuses:
            params["statusId"] = _resolve_names(flt.statuses, master.table("statuses"))
        elif not flt.include_closed:
            params["statusId"] = [
                item["id"] for item in master.table("statuses").items
                if item["id"] != STATUS_CLOSED_ID
            ]
        if flt.types:
            params["issueTypeId"] = _resolve_names(flt.types, master.table("issue_types"))
        if flt.keyword:
            params["keyword"] = flt.keyword
        issues = []
        if flt.parent_key:
            parent = client.get_issue(flt.parent_key)
            if parent is None:
                raise ExportError(f"{flt.parent_key} が見つかりません")
            params["parentIssueId"] = [parent["id"]]
            issues.append(parent)
        issues += client.get_issues(master.project_id, params)

    # 重複を除く（--parent の親が条件にも当たった場合など）
    unique = {i["id"]: i for i in issues}
    for issue in list(unique.values()):
        parent_id = issue.get("parentIssueId")
        if parent_id and parent_id not in unique:
            parent = client.get_issue(str(parent_id))
            if parent:
                unique[parent["id"]] = parent
    return order_issues(list(unique.values()))


def parent_key_map(issues: list[dict], client) -> dict[int, str]:
    """課題 ID → 課題キー。親のキーを書くために使う。"""
    keys = {i["id"]: i["issueKey"] for i in issues}
    for issue in issues:
        parent_id = issue.get("parentIssueId")
        if parent_id and parent_id not in keys:
            parent = client.get_issue(str(parent_id))
            keys[parent_id] = parent["issueKey"] if parent else str(parent_id)
    return keys


def write_issues(
    path: str | Path, issues: list[dict], master: Master, keys: dict[int, str], *, base_url: str = ""
) -> tuple[Path, list[str]]:
    warnings: list[str] = []
    rows = [
        issue_to_row(i, master, keys.get(i.get("parentIssueId")), base_url=base_url, warnings=warnings)
        for i in issues
    ]
    out = write_workbook(
        path, master, update_rows=rows, extra_update_headers=REFERENCE_HEADERS,
        include_register=False,
    )
    return out, warnings


def export(client, master: Master, path: str | Path, flt: ExportFilter) -> tuple[Path, int, list[str]]:
    issues = fetch_issues(client, master, flt)
    keys = parent_key_map(issues, client)
    out, warnings = write_issues(path, issues, master, keys, base_url=getattr(client, "base_url", ""))
    return out, len(issues), warnings
