"""実行（送信の順番・失敗時の扱い・実行ログ）"""

import csv

import pytest

import executor
from executor import CREATED, FAILED, NO_CHANGE, NOT_RUN, SKIPPED, UPDATED, RunLog, execute

REG = ["親課題キー", "子課題", "件名", "種別", "優先度"]


def run(plan, fake, **kw):
    assert plan.ok, [str(p) for p in plan.problems]
    return execute(plan, 42, fake, **kw)


def outcomes(results):
    return [(r.row_no, r.outcome) for r in results]


class TestCreate:
    def test_親を先に作り_子にその_ID_を渡す(self, plan_of, fake):
        plan = plan_of({"登録": (REG, [
            [None, None, "親", "タスク", "中"],
            [None, "○", "子1", "タスク", "中"],
            [None, "○", "子2", "タスク", "中"],
        ])})
        results = run(plan, fake)
        assert outcomes(results) == [(2, CREATED), (3, CREATED), (4, CREATED)]
        parent_id = results[0].issue["id"]
        assert [c[1].get("parentIssueId") for c in fake.calls] == [None, parent_id, parent_id]

    def test_既存課題の子として作る(self, plan_of, fake):
        p = fake.add("既存")
        plan = plan_of({"登録": (REG, [["DEMO-1", None, "子", "タスク", "中"]])})
        run(plan, fake)
        assert fake.calls[0][1]["parentIssueId"] == p["id"]

    def test_送る値(self, plan_of, fake):
        headers = REG + ["カテゴリー", "予定時間", "区分", "タグ"]
        plan = plan_of({"登録": (headers, [[None, None, "A", "タスク", "中", "設計,開発", 2, "A", "x"]])})
        run(plan, fake)
        params = fake.calls[0][1]
        assert params["projectId"] == 42
        assert params["categoryId"] == [100, 101]
        assert params["estimatedHours"] == "2"
        assert params["customField_302"] == 1          # 単一は値のまま
        assert params["customField_303"] == [11]       # 複数はリスト

    def test_状態は作成の直後に更新する(self, plan_of, fake):
        plan = plan_of({"登録": (REG + ["状態"], [[None, None, "A", "タスク", "中", "処理中"]])})
        results = run(plan, fake)
        assert [c[0] for c in fake.calls] == ["create", "update"]
        assert fake.calls[1][2] == {"statusId": 2}
        assert results[0].issue["status"]["id"] == 2

    def test_親の作成に失敗したら子は作らない(self, plan_of, fake):
        fake.fail_create_matching = "親"
        plan = plan_of({"登録": (REG, [
            [None, None, "親", "タスク", "中"],
            [None, "○", "子", "タスク", "中"],
            [None, None, "別", "タスク", "中"],
        ])})
        results = run(plan, fake)
        assert outcomes(results) == [(2, FAILED), (3, SKIPPED), (4, CREATED)]
        assert "2行目" in results[1].message

    def test_認証エラーで以降を止める(self, plan_of, fake):
        fake.fatal_on_create = True
        plan = plan_of({"登録": (REG, [[None, None, "a", "タスク", "中"], [None, None, "b", "タスク", "中"]])})
        results = run(plan, fake)
        assert outcomes(results) == [(2, FAILED), (3, NOT_RUN)]
        assert len(fake.calls) == 1


class TestUpdate:
    def test_差分だけを送る(self, plan_of, fake):
        fake.add("x", assigneeId=10)
        plan = plan_of({"更新": (["課題キー", "件名", "担当者"], [["DEMO-1", "y", "山田太郎"]])})
        results = run(plan, fake)
        assert outcomes(results) == [(2, UPDATED)]
        assert fake.calls == [("update", "DEMO-1", {"summary": "y"})]

    def test_削除は空で送る(self, plan_of, fake):
        fake.add("x", assigneeId=10, categoryId=[100], customField_303=[11])
        plan = plan_of({"更新": (["課題キー", "担当者", "カテゴリー", "タグ"], [["DEMO-1", "(削除)", "(削除)", "(削除)"]])})
        run(plan, fake)
        assert fake.calls[0][2] == {"assigneeId": "", "categoryId": [""], "customField_303": [""]}
        issue = fake.get_issue("DEMO-1")
        assert issue["assignee"] is None and issue["category"] == []

    def test_変更なしの行は送らない(self, plan_of, fake):
        fake.add("x")
        plan = plan_of({"更新": (["課題キー", "件名"], [["DEMO-1", "x"]])})
        assert outcomes(run(plan, fake)) == [(2, NO_CHANGE)]
        assert fake.calls == []

    def test_Backlog_が変更なしと返したら変更なし(self, plan_of, fake):
        fake.add("x")
        fake.no_change_keys.add("DEMO-1")
        plan = plan_of({"更新": (["課題キー", "件名"], [["DEMO-1", "y"]])})
        assert outcomes(run(plan, fake)) == [(2, NO_CHANGE)]

    def test_親から外す行を先に送る(self, plan_of, fake):
        """
        シートの順に送ると、DEMO-2 が子課題のまま DEMO-3 の親にされて拒否される。
        FakeBacklog は本物と同じくこれを拒否する。
        """
        p = fake.add("親")
        fake.add("子", parent=p)      # DEMO-2
        fake.add("x")                 # DEMO-3
        plan = plan_of({"更新": (["課題キー", "親課題キー"], [["DEMO-3", "DEMO-2"], ["DEMO-2", "(解除)"]])})
        results = run(plan, fake)
        assert [r.issue_key for r in results] == ["DEMO-2", "DEMO-3"]
        assert all(r.outcome == UPDATED for r in results)
        assert fake.calls[0][2] == {"parentIssueId": ""}
        assert fake.get_issue("DEMO-3")["parentIssueId"] == fake.get_issue("DEMO-2")["id"]

    def test_順番を変えなければ拒否されることを確かめておく(self, plan_of, fake, monkeypatch):
        """上のテストが順番の効果を見ていることの確認（偽物が制約を守っているか）。"""
        p = fake.add("親")
        fake.add("子", parent=p)
        fake.add("x")
        plan = plan_of({"更新": (["課題キー", "親課題キー"], [["DEMO-3", "DEMO-2"], ["DEMO-2", "(解除)"]])})
        monkeypatch.setattr(executor, "ordered_updates", lambda plan: plan.updates)
        results = run(plan, fake)
        assert results[0].outcome == FAILED


class TestRunLog:
    def test_1件ごとに書く(self, plan_of, fake, tmp_path):
        plan = plan_of({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        path = tmp_path / "out" / "run.csv"
        with RunLog(path) as log:
            run(plan, fake, log=log)
        rows = list(csv.reader(path.open(encoding="utf-8-sig")))
        assert rows[0] == RunLog.HEADERS
        assert rows[1][:4] == ["登録", "2", CREATED, "DEMO-1"]

    def test_件名を変えた行は更新後の件名を残す(self, plan_of, fake, tmp_path):
        """変更前の件名だと、ログだけを見る人には変更後の課題と区別がつかない。"""
        fake.add("旧い件名")
        fake.add("そのまま")
        plan = plan_of({"更新": (["課題キー", "件名", "担当者"], [
            ["DEMO-1", "新しい件名", None],
            ["DEMO-2", None, "山田太郎"],
        ])})
        path = tmp_path / "run.csv"
        with RunLog(path) as log:
            results = run(plan, fake, log=log)
        assert [r.summary for r in results] == ["新しい件名", "そのまま"]
        rows = list(csv.reader(path.open(encoding="utf-8-sig")))
        assert [r[4] for r in rows[1:]] == ["新しい件名", "そのまま"]

    def test_更新に失敗した行は今の件名を残す(self, plan_of, fake):
        """送っていない件名を書くと、反映されたように読めてしまう。"""
        from backlog_client import BacklogAPIError

        fake.add("旧い件名")
        plan = plan_of({"更新": (["課題キー", "件名"], [["DEMO-1", "新しい件名"]])})

        def failing(key, params):
            raise BacklogAPIError("更新できません", status=400)

        fake.update_issue = failing
        results = run(plan, fake)
        assert (results[0].outcome, results[0].summary) == (FAILED, "旧い件名")

    def test_エラーのある計画は実行しない(self, plan_of, fake):
        plan = plan_of({"登録": (REG, [[None, "○", "a", "タスク", "中"]])})
        with pytest.raises(ValueError):
            execute(plan, 42, fake)


class TestFollowUpFailure:
    def test_作成後の更新に失敗しても課題キーは残す(self, plan_of, fake):
        """作成は済んでいるので、キーを失うと再実行で二重に作ってしまう。"""
        plan = plan_of({"登録": (REG + ["状態"], [[None, None, "A", "タスク", "中", "処理中"]])})
        original = fake.update_issue

        def failing(key, params):
            from backlog_client import BacklogAPIError
            raise BacklogAPIError("状態を変えられません", status=400)

        fake.update_issue = failing
        results = run(plan, fake)
        fake.update_issue = original
        assert results[0].outcome == FAILED
        assert results[0].issue_key == "DEMO-1"
        assert "作成はできたが" in results[0].message
