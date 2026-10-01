"""
二重登録の防止（同じ件名の未完了の課題）
========================================
元の Excel には課題キーを書き戻さないため、同じ登録シートをもう一度実行すると
二重に作られる。Backlog に同じ件名の未完了の課題があれば止める。
"""

import pytest

from planner import build_plan
from sheet_io import read_workbook

H = ["親課題キー", "子課題", "件名", "種別", "優先度"]


def reg(*rows):
    return {"登録": (H, [list(r) for r in rows])}


def row(summary, *, parent_key=None, child=None):
    return [parent_key, child, summary, "タスク", "中"]


@pytest.fixture
def plan_with(make_book, master, fake):
    def _plan(sheets, **kw):
        return build_plan(read_workbook(make_book(sheets)), master, fake, **kw)
    return _plan


def messages(plan):
    return [str(p) for p in plan.problems]


class TestTopLevel:
    def test_同じ件名の未完了の課題があれば止める(self, plan_with, fake):
        fake.add("画面設計")                      # DEMO-1
        plan = plan_with(reg(row("画面設計"), row("帳票設計")))
        assert messages(plan) == [
            "登録 2行目「件名」: 同じ件名の未完了の課題があります（DEMO-1）。二重に登録しないよう、送りません"
        ]
        assert plan.duplicate_rows == [2]

    def test_前後の空白は同じとみなす(self, plan_with, fake):
        fake.add("画面設計")
        assert not plan_with(reg(row(" 画面設計 "))).ok

    def test_同じ件名が複数あればすべて示す(self, plan_with, fake):
        fake.add("画面設計")
        fake.add("画面設計")
        assert "（DEMO-1 / DEMO-2）" in messages(plan_with(reg(row("画面設計"))))[0]

    def test_完了した課題は比べない(self, plan_with, fake):
        """同じ件名の課題を定期的に作り直す使い方がある。"""
        fake.add("週次の定例", statusId=4)
        assert plan_with(reg(row("週次の定例"))).ok

    def test_子課題とは比べない(self, plan_with, fake):
        parent = fake.add("親")
        fake.add("レビュー", parent=parent)
        assert plan_with(reg(row("レビュー"))).ok

    def test_指定すれば止めない(self, plan_with, fake):
        fake.add("画面設計")
        plan = plan_with(reg(row("画面設計")), allow_duplicates=True)
        assert plan.ok and plan.duplicate_rows == []


class TestChild:
    def test_同じ親の下に同じ件名があれば止める(self, plan_with, fake):
        parent = fake.add("親")                   # DEMO-1
        fake.add("レビュー", parent=parent)       # DEMO-2
        plan = plan_with(reg(row("レビュー", parent_key="DEMO-1")))
        assert messages(plan) == [
            "登録 2行目「件名」: DEMO-1 の子課題に同じ件名の未完了の課題があります（DEMO-2）。"
            "二重に登録しないよう、送りません"
        ]

    def test_別の親の下の同じ件名は止めない(self, plan_with, fake):
        """「レビュー」のように、親ごとに同じ件名の子を作るのはよくある。"""
        a = fake.add("A")                         # DEMO-1
        fake.add("B")                             # DEMO-2
        fake.add("レビュー", parent=a)
        assert plan_with(reg(row("レビュー", parent_key="DEMO-2"))).ok

    def test_同じシートの行を親にする子は比べない(self, plan_with, fake):
        """親が新しく作られるので、その下に既存の課題は無い。"""
        fake.add("レビュー")
        assert plan_with(reg(row("新しい親"), row("レビュー", child="○"))).ok

    def test_親が見つからない行は_親の無い課題と比べない(self, plan_with, fake):
        fake.add("レビュー")
        plan = plan_with(reg(row("レビュー", parent_key="DEMO-99")))
        assert messages(plan) == ["登録 2行目「親課題キー」: DEMO-99 が見つかりません"]


class TestCalls:
    @pytest.fixture
    def counted(self, fake, monkeypatch):
        calls = []
        original = fake.get_issues

        def counting(project_id, params=None):
            calls.append(params)
            return original(project_id, params)

        monkeypatch.setattr(fake, "get_issues", counting)
        return calls

    def test_未完了の課題を1回だけまとめて取る(self, plan_with, counted):
        plan_with(reg(row("a"), row("b"), row("c")))
        assert counted == [{"statusId": [1, 2, 3]}]

    def test_登録する行が無ければ取りにいかない(self, plan_with, fake, counted):
        fake.add("x")
        plan_with({"更新": (["課題キー", "件名"], [["DEMO-1", "y"]])})
        assert counted == []


class TestInSheet:
    """シートの中の重複は、書き間違いのことが多いが意図することもあるため、注意にとどめる。"""

    def warnings(self, plan):
        return [str(w) for w in plan.warnings if w.column == "件名"]

    def test_同じ親の下に同じ件名の子が2行あれば注意(self, plan_with):
        plan = plan_with(reg(row("帳票設計"), row("レビュー", child="○"), row("レビュー", child="○")))
        assert plan.ok and len(plan.creates) == 3
        assert self.warnings(plan) == [
            "登録 4行目「件名」: 同じ親（2行目）の下に、3行目と同じ件名の行があります。どちらも作ります"
        ]

    def test_親の無い行どうしでも注意(self, plan_with):
        plan = plan_with(reg(row("画面設計"), row(" 画面設計 ")))
        assert self.warnings(plan) == [
            "登録 3行目「件名」: 親の無い行に、2行目と同じ件名の行があります。どちらも作ります"
        ]

    def test_既存の親の下でも注意(self, plan_with, fake):
        fake.add("親")
        plan = plan_with(reg(row("レビュー", parent_key="DEMO-1"), row("レビュー", parent_key="DEMO-1")))
        assert self.warnings(plan) == [
            "登録 3行目「件名」: 同じ親（DEMO-1）の下に、2行目と同じ件名の行があります。どちらも作ります"
        ]

    def test_親が違えば注意しない(self, plan_with):
        """「レビュー」のように、親ごとに同じ件名の子を作るのはよくある。"""
        plan = plan_with(reg(
            row("画面設計"), row("レビュー", child="○"),
            row("帳票設計"), row("レビュー", child="○"),
        ))
        assert self.warnings(plan) == []

    def test_3行あれば2行目以降のそれぞれに出す(self, plan_with):
        plan = plan_with(reg(row("a"), row("a"), row("a")))
        assert [w.row for w in plan.warnings if w.column == "件名"] == [3, 4]

    def test_同じ件名の課題を許しても注意は出す(self, plan_with):
        plan = plan_with(reg(row("a"), row("a")), allow_duplicates=True)
        assert len(self.warnings(plan)) == 1
