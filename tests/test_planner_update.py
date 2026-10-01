"""更新シートの計画（差分・親課題キー・親子の制約）"""

import pytest

import executor
from fields import CLEAR


def upd(rows, headers=("課題キー", "親課題キー", "件名", "担当者", "期限日", "状態")):
    return {"更新": (list(headers), rows)}


def messages(plan):
    return [str(p) for p in plan.problems]


@pytest.fixture
def issue(fake):
    return fake.add("元の件名", assigneeId=10, dueDate="2026-10-01", statusId=2)   # DEMO-1


class TestDiff:
    def test_空のセルは変更しない(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", None, None, None, None, None]]))
        assert plan.ok
        assert not plan.updates[0].has_changes

    def test_同じ値を書いても変更なし(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", None, "元の件名", "山田太郎", "2026-10-01", "処理中"]]))
        assert plan.ok
        assert plan.updates[0].changes == []

    def test_変わる項目だけを差分にする(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", None, "新しい件名", "山田太郎", "2026-10-31", None]]))
        changes = {c.column.header: (c.old, c.new) for c in plan.updates[0].changes}
        assert changes == {"件名": ("元の件名", "新しい件名"), "期限日": ("2026-10-01", "2026-10-31")}

    def test_削除で値を消す(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", None, None, "(削除)", None, None]]))
        assert [(c.column.header, c.new) for c in plan.updates[0].changes] == [("担当者", CLEAR)]

    def test_もともと空の項目を削除しても変更なし(self, plan_of, fake):
        fake.add("x")
        plan = plan_of(upd([["DEMO-1", None, None, "(削除)", None, None]]))
        assert plan.updates[0].changes == []

    def test_削除できない項目(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", None, "(削除)", None, None, None]]))
        assert any("(削除) で消せません" in m for m in messages(plan))

    def test_カスタム属性の差分(self, plan_of, fake):
        fake.add("x", customField_302=1, customField_303=[11])
        plan = plan_of(upd([["DEMO-1", "B", "x\ny"]], ["課題キー", "区分", "タグ"]))
        changes = {c.column.header: (c.old, c.new) for c in plan.updates[0].changes}
        assert changes == {"区分": (1, 2), "タグ": ((11,), (11, 12))}

    def test_種別を変えると使えなくなる属性はエラー(self, plan_of, fake):
        fake.add("x", issueTypeId=2)
        plan = plan_of(upd([["DEMO-1", "タスク", "値"]], ["課題キー", "種別", "バグ専用"]))
        assert any("種別「タスク」では使えません" in m for m in messages(plan))


class TestDescription:
    """詳細（本文）。複数行で、Excel と Backlog で改行コードが違いうる。"""

    def sent(self, plan, fake, master):
        executor.execute(plan, master.project_id, fake)
        return [c for c in fake.calls if c[0] == "update"]

    def test_書き換えると詳細だけを送る(self, plan_of, fake, master):
        fake.add("x", description="元の本文")
        plan = plan_of(upd([["DEMO-1", "新しい本文\n2行目"]], ["課題キー", "詳細"]))
        assert plan.ok
        assert self.sent(plan, fake, master) == [("update", "DEMO-1", {"description": "新しい本文\n2行目"})]
        assert fake.get_issue("DEMO-1")["description"] == "新しい本文\n2行目"

    def test_削除で本文を空にする(self, plan_of, fake, master):
        fake.add("x", description="消す本文")
        plan = plan_of(upd([["DEMO-1", "(削除)"]], ["課題キー", "詳細"]))
        assert plan.ok
        assert self.sent(plan, fake, master) == [("update", "DEMO-1", {"description": ""})]
        assert fake.get_issue("DEMO-1")["description"] == ""

    def test_空のセルは本文を変えない(self, plan_of, fake):
        fake.add("x", description="元の本文")
        plan = plan_of(upd([["DEMO-1", None, "新しい件名"]], ["課題キー", "詳細", "件名"]))
        assert [c.column.header for c in plan.updates[0].changes] == ["件名"]

    def test_改行コードだけの違いは変更なし(self, plan_of, fake):
        """
        Backlog の本文は \r\n を含みうるが、Excel から読むと \n になる（openpyxl が
        XML の改行をそろえる）。違いとみなすと、書き出して戻すたびに本文を送ってしまう。
        """
        fake.add("x", description="1行目\r\n2行目")
        plan = plan_of(upd([["DEMO-1", "1行目\n2行目"]], ["課題キー", "詳細"]))
        assert plan.ok
        assert plan.updates[0].changes == []


class TestKeys:
    def test_課題キーが空(self, plan_of):
        plan = plan_of(upd([[None, None, "件名", None, None, None]]))
        assert any("課題キーが空" in m for m in messages(plan))

    def test_課題キーの列が無い(self, plan_of):
        plan = plan_of(upd([["件名"]], ["件名"]))
        assert any("必須の列「課題キー」" in m for m in messages(plan))

    def test_重複(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", None, "a", None, None, None], ["DEMO-1", None, "b", None, None, None]]))
        assert any("2行目にもあります" in m for m in messages(plan))

    def test_存在しない(self, plan_of):
        plan = plan_of(upd([["DEMO-9", None, "a", None, None, None]]))
        assert any("DEMO-9 が見つかりません" in m for m in messages(plan))

    def test_別のプロジェクト(self, plan_of):
        plan = plan_of(upd([["OTHER-1", None, "a", None, None, None]]))
        assert any("別のプロジェクト" in m for m in messages(plan))

    def test_子課題列は使えない(self, plan_of, issue):
        plan = plan_of(upd([["DEMO-1", "○"]], ["課題キー", "子課題"]))
        assert any("子課題の列は使えません" in m for m in messages(plan))


class TestParent:
    def test_親を付ける(self, plan_of, fake):
        p = fake.add("親")
        fake.add("子にする")
        plan = plan_of(upd([["DEMO-2", "DEMO-1"]], ["課題キー", "親課題キー"]))
        assert plan.ok, messages(plan)
        kind, parent, old = plan.updates[0].parent_change
        assert (kind, parent["id"], old) == ("set", p["id"], None)
        assert plan.updates[0].phase == 2

    def test_今と同じ親なら変更なし(self, plan_of, fake):
        p = fake.add("親")
        fake.add("子", parent=p)
        plan = plan_of(upd([["DEMO-2", "DEMO-1"]], ["課題キー", "親課題キー"]))
        assert plan.ok
        assert not plan.updates[0].has_changes

    def test_解除で親から外す(self, plan_of, fake):
        p = fake.add("親")
        fake.add("子", parent=p)
        plan = plan_of(upd([["DEMO-2", "(解除)"]], ["課題キー", "親課題キー"]))
        assert plan.updates[0].parent_change == ("detach", None, "DEMO-1")
        assert plan.updates[0].phase == 1

    def test_親のない課題を解除しても変更なし(self, plan_of, fake):
        fake.add("x")
        plan = plan_of(upd([["DEMO-1", "(解除)"]], ["課題キー", "親課題キー"]))
        assert not plan.updates[0].has_changes

    def test_親を付け替える(self, plan_of, fake):
        p1 = fake.add("親1")
        fake.add("親2")
        fake.add("子", parent=p1)
        plan = plan_of(upd([["DEMO-3", "DEMO-2"]], ["課題キー", "親課題キー"]))
        kind, parent, old = plan.updates[0].parent_change
        assert (kind, parent["issueKey"], old) == ("set", "DEMO-2", "DEMO-1")

    def test_自分自身は親にできない(self, plan_of, fake):
        fake.add("x")
        plan = plan_of(upd([["DEMO-1", "DEMO-1"]], ["課題キー", "親課題キー"]))
        assert any("自分自身" in m for m in messages(plan))

    def test_子課題を親にはできない(self, plan_of, fake):
        p = fake.add("親")
        fake.add("子", parent=p)
        fake.add("x")
        plan = plan_of(upd([["DEMO-3", "DEMO-2"]], ["課題キー", "親課題キー"]))
        assert any("DEMO-2 は（この実行の後も）子課題" in m for m in messages(plan))

    def test_同じシートで親から外す課題なら親にできる(self, plan_of, fake):
        p = fake.add("親")
        fake.add("子", parent=p)     # DEMO-2
        fake.add("x")                # DEMO-3
        plan = plan_of(upd([["DEMO-3", "DEMO-2"], ["DEMO-2", "(解除)"]], ["課題キー", "親課題キー"]))
        assert plan.ok, messages(plan)

    def test_子を持つ課題は子にできない(self, plan_of, fake):
        p = fake.add("親")          # DEMO-1
        fake.add("子", parent=p)    # DEMO-2
        fake.add("別の親")           # DEMO-3
        plan = plan_of(upd([["DEMO-1", "DEMO-3"]], ["課題キー", "親課題キー"]))
        assert any("DEMO-1 には子課題（DEMO-2）" in m for m in messages(plan))

    def test_子を全部よそへ移すなら子にできる(self, plan_of, fake):
        p = fake.add("親")          # DEMO-1
        fake.add("子", parent=p)    # DEMO-2
        fake.add("別の親")           # DEMO-3
        plan = plan_of(upd([["DEMO-1", "DEMO-3"], ["DEMO-2", "(解除)"]], ["課題キー", "親課題キー"]))
        assert plan.ok, messages(plan)

    def test_シート内で親子を連鎖させるとエラー(self, plan_of, fake):
        for s in ("a", "b", "c"):
            fake.add(s)
        # DEMO-2 を DEMO-1 の子にしつつ、DEMO-3 を DEMO-2 の子にする（2 階層）
        plan = plan_of(upd([["DEMO-2", "DEMO-1"], ["DEMO-3", "DEMO-2"]], ["課題キー", "親課題キー"]))
        assert any("DEMO-2 は（この実行の後も）子課題" in m for m in messages(plan))
        assert any("DEMO-2 には子課題（DEMO-3）" in m for m in messages(plan))



class TestProblemOrder:
    """
    各行を読んで見つかるエラーと、シート全体で判定する親子の制約のエラーは、
    別々の段で見つかる。見つかった順のままだと、同じ行のエラーが離れて並ぶ。
    """

    def test_エラーは行の順に並ぶ(self, plan_of, fake):
        a = fake.add("A")                         # DEMO-1
        fake.add("A の子", parent=a)              # DEMO-2
        fake.add("B")                             # DEMO-3
        fake.add("C")                             # DEMO-4
        plan = plan_of(upd([
            ["DEMO-4", "DEMO-4", None, None, None, None],   # 自分自身（行を読んで見つかる）
            ["DEMO-3", "DEMO-2", None, None, None, None],   # 子課題を親に（シート全体で判定）
            ["DEMO-1", "DEMO-99", None, None, None, None],  # 見つからない（行を読んで見つかる）
        ]))
        assert [p.row for p in plan.problems] == [2, 3, 4]

    def test_同じ行の中では見つかった順を保ち_シートの順は登録_更新(self, plan_of, fake):
        fake.add("A")
        plan = plan_of({
            "更新": (["課題キー", "件名"], [["DEMO-1", "(削除)"]]),
            "登録": (["件名", "種別", "優先度"], [["x", "無い種別", "無い優先度"]]),
        })
        sheets = [p.sheet for p in plan.problems]
        assert sheets == sorted(sheets, key=["登録", "更新"].index)
        # 同じ行の中は、列を読んだ順（種別 → 優先度）のまま
        assert [p.column for p in plan.problems if p.sheet == "登録"][:2] == ["種別", "優先度"]


class TestBothSheets:
    def test_登録と更新を同時に扱う(self, plan_of, issue):
        plan = plan_of({
            "登録": (["件名", "種別", "優先度"], [["新規", "タスク", "中"]]),
            "更新": (["課題キー", "件名"], [["DEMO-1", "変更"]]),
            "メモ": (["何でも"], [["無視される"]]),
        })
        assert plan.ok
        assert len(plan.creates) == 1 and len(plan.updates) == 1
