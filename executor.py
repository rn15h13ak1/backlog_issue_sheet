"""
実行（Excel → Backlog）
=======================
計画どおりに課題を作成・更新する。順番は次のとおり。

1. 登録シート（上の行から。親は必ず子より上にあるので、この順で親が先に作られる）
2. 更新シートのうち、親を付けない行
3. 更新シートのうち、親を付ける行（planner.UpdatePlan.phase を参照）

1 件の失敗では止めない。ただし認証・権限のエラー（fatal）は全体を止める。
親の作成に失敗した子は、作っても親子にならないため作らない。
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from backlog_client import BacklogAPIError, BacklogNoChangeError
from fields import REGISTER_SHEET, UPDATE_SHEET, to_param
from planner import CreatePlan, Plan, UpdatePlan

CREATED = "作成"
UPDATED = "更新"
NO_CHANGE = "変更なし"
FAILED = "失敗"
SKIPPED = "スキップ"
NOT_RUN = "未実行"


@dataclass
class Result:
    sheet: str
    row_no: int
    outcome: str
    issue_key: str = ""
    summary: str = ""              # 更新できた行は更新後の件名、それ以外は Backlog の今の件名
    message: str = ""
    issue: dict | None = None      # 作成・更新後の課題（結果の Excel に使う）
    fatal: bool = False            # 認証・権限のエラーで失敗した（以降は送らない）


class RunLog:
    """
    実行ログ（CSV）。1 件ごとに書き込んで flush する。

    途中で落ちても、どこまで作ったかが残る。作成済みの行をもう一度
    実行すると二重に作られるため、再実行の前にこれを見る。
    """

    HEADERS = ["シート", "行", "結果", "課題キー", "件名", "内容"]

    def __init__(self, path: Path | None):
        self.path = path
        self._fh = None
        self._writer = None

    def __enter__(self) -> "RunLog":
        if self.path:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # Excel で開いたときに文字化けしないよう BOM を付ける
            self._fh = open(self.path, "w", encoding="utf-8-sig", newline="")
            self._writer = csv.writer(self._fh)
            self._writer.writerow(self.HEADERS)
            self._fh.flush()
        return self

    def __exit__(self, *exc) -> bool:
        if self._fh:
            self._fh.close()
        return False

    def record(self, r: Result) -> None:
        if self._writer:
            self._writer.writerow([r.sheet, r.row_no, r.outcome, r.issue_key, r.summary, r.message])
            self._fh.flush()


def create_params(cp: CreatePlan, project_id: int, parent_id: int | None) -> dict:
    params = {"projectId": project_id}
    for key, value in cp.values.items():
        params[key] = to_param(cp.columns[key], value)
    if parent_id:
        params["parentIssueId"] = parent_id
    return params


def follow_up_params(cp: CreatePlan) -> dict:
    return {key: to_param(cp.columns[key], value) for key, value in cp.follow_up.items()}


def update_params(up: UpdatePlan) -> dict:
    params = {c.column.key: to_param(c.column, c.new) for c in up.changes}
    if up.parent_change:
        kind, parent, _ = up.parent_change
        # 空文字で親子関係を外す
        params["parentIssueId"] = parent["id"] if kind == "set" else ""
    return params


def ordered_updates(plan: Plan) -> list[UpdatePlan]:
    # sorted は安定なので、同じ phase の中ではシートの順が保たれる
    return sorted(plan.updates, key=lambda u: u.phase)


def count_actions(plan: Plan) -> int:
    """送信が発生する行の数（--limit の対象）。"""
    return len(plan.creates) + sum(1 for u in plan.updates if u.has_changes)


def execute(plan: Plan, project_id: int, client, *, limit: int | None = None, log: RunLog | None = None) -> list[Result]:
    if not plan.ok:
        raise ValueError("エラーのある計画は実行できません")
    log = log or RunLog(None)
    results: list[Result] = []
    created: dict[int, dict] = {}      # 登録シートの行番号 → 作成した課題
    budget = limit if limit is not None else float("inf")
    aborted: str | None = None

    def emit(r: Result) -> None:
        results.append(r)
        log.record(r)
        mark = {CREATED: "✓", UPDATED: "✓", NO_CHANGE: "-", FAILED: "✗"}.get(r.outcome, "・")
        where = f"{r.sheet} {r.row_no}行目"
        key = f" {r.issue_key}" if r.issue_key else ""
        tail = f" … {r.message}" if r.message else ""
        print(f"  {mark} {r.outcome:<4} {where}{key} {r.summary}{tail}")

    for cp in plan.creates:
        if aborted:
            emit(Result(REGISTER_SHEET, cp.row_no, NOT_RUN, summary=cp.summary, message=aborted))
            continue
        if budget <= 0:
            emit(Result(REGISTER_SHEET, cp.row_no, NOT_RUN, summary=cp.summary, message="--limit に達したため"))
            continue

        parent_id = cp.parent_issue_id
        if cp.parent_row is not None:
            parent = created.get(cp.parent_row)
            if parent is None:
                emit(Result(
                    REGISTER_SHEET, cp.row_no, SKIPPED, summary=cp.summary,
                    message=f"親（{cp.parent_row}行目）が作成されていないため",
                ))
                continue
            parent_id = parent["id"]

        budget -= 1
        try:
            issue = client.create_issue(create_params(cp, project_id, parent_id))
        except BacklogAPIError as e:
            emit(Result(
                REGISTER_SHEET, cp.row_no, FAILED, summary=cp.summary, message=_one_line(e), fatal=e.fatal,
            ))
            if e.fatal:
                aborted = "認証・権限のエラーで中止したため"
            continue

        created[cp.row_no] = issue
        result = Result(REGISTER_SHEET, cp.row_no, CREATED, issue.get("issueKey", ""), cp.summary, issue=issue)
        if cp.follow_up:
            try:
                result.issue = client.update_issue(issue["issueKey"], follow_up_params(cp))
            except BacklogNoChangeError:
                pass
            except BacklogAPIError as e:
                # 課題は作成済み。失敗として数えるが、キーは残す（重複して作らないため）
                result.outcome = FAILED
                result.message = f"作成はできたが、状態・完了理由の更新に失敗: {_one_line(e)}"
                result.fatal = e.fatal
                if e.fatal:
                    aborted = "認証・権限のエラーで中止したため"
        emit(result)

    for up in ordered_updates(plan):
        if not up.has_changes:
            emit(Result(UPDATE_SHEET, up.row_no, NO_CHANGE, up.issue_key, up.summary))
            continue
        if aborted:
            emit(Result(UPDATE_SHEET, up.row_no, NOT_RUN, up.issue_key, up.summary, message=aborted))
            continue
        if budget <= 0:
            emit(Result(UPDATE_SHEET, up.row_no, NOT_RUN, up.issue_key, up.summary, message="--limit に達したため"))
            continue
        budget -= 1
        try:
            issue = client.update_issue(up.issue_key, update_params(up))
        except BacklogNoChangeError:
            emit(Result(UPDATE_SHEET, up.row_no, NO_CHANGE, up.issue_key, up.summary, message="Backlog 側で変更なしと判定"))
            continue
        except BacklogAPIError as e:
            emit(Result(
                UPDATE_SHEET, up.row_no, FAILED, up.issue_key, up.summary, message=_one_line(e), fatal=e.fatal,
            ))
            if e.fatal:
                aborted = "認証・権限のエラーで中止したため"
            continue
        # 件名を変えた行は、ログを見る人が変更後の課題を探せるよう、更新後の件名を残す
        summary = (issue or {}).get("summary") or up.summary
        emit(Result(UPDATE_SHEET, up.row_no, UPDATED, up.issue_key, summary, issue=issue))

    return results


def _one_line(e: Exception) -> str:
    return " ".join(str(e).split())
