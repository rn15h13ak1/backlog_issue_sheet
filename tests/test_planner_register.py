"""登録シートの計画（親の決定・検証）"""

import datetime as dt

from fields import CLEAR

H = ["親課題キー", "子課題", "件名", "種別", "優先度"]


def reg(rows, headers=H):
    return {"登録": (headers, rows)}


def messages(plan):
    return [str(p) for p in plan.problems]


class TestParentByMark:
    def test_丸の行は直前の印の無い行の子になる(self, plan_of):
        plan = plan_of(reg([
            [None, None, "親A", "タスク", "中"],
            [None, "○", "子A-1", "タスク", "中"],
            [None, "○", "子A-2", "タスク", "中"],
            [None, None, "親B", "タスク", "中"],
            [None, "○", "子B-1", "タスク", "中"],
        ]))
        assert plan.ok, messages(plan)
        parents = [(c.summary, c.parent_row) for c in plan.creates]
        assert parents == [("親A", None), ("子A-1", 2), ("子A-2", 2), ("親B", None), ("子B-1", 5)]

    def test_空行を挟んでも行番号は_Excel_のまま(self, plan_of):
        plan = plan_of(reg([
            [None, None, "親A", "タスク", "中"],
            [None, None, None, None, None],
            [None, "○", "子A-1", "タスク", "中"],
        ]))
        assert plan.ok
        assert [(c.row_no, c.parent_row) for c in plan.creates] == [(2, None), (4, 2)]

    def test_1行目から子課題だとエラー(self, plan_of):
        plan = plan_of(reg([[None, "○", "子", "タスク", "中"]]))
        assert any("親になれる行" in m for m in messages(plan))

    def test_丸の別表記も受け付ける(self, plan_of):
        plan = plan_of(reg([
            [None, None, "親", "タスク", "中"],
            [None, "〇", "子1", "タスク", "中"],
            [None, 1, "子2", "タスク", "中"],
        ]))
        assert plan.ok
        assert [c.parent_row for c in plan.creates] == [None, 2, 2]

    def test_読めない印はエラー(self, plan_of):
        plan = plan_of(reg([
            [None, None, "親", "タスク", "中"],
            [None, "×", "子", "タスク", "中"],
        ]))
        assert any("「×」は読めません" in m for m in messages(plan))

    def test_既存課題の子の行は飛ばしてさらに上を親にし_注意を出す(self, plan_of, fake):
        fake.add("既存の親")   # DEMO-1
        plan = plan_of(reg([
            [None, None, "親A", "タスク", "中"],
            ["DEMO-1", None, "既存の子", "タスク", "中"],
            [None, "○", "子", "タスク", "中"],
        ]))
        assert plan.ok, messages(plan)
        assert plan.creates[2].parent_row == 2
        assert any("さらに上の 2行目" in str(w) for w in plan.warnings)


class TestParentByKey:
    def test_既存課題の子として作る(self, plan_of, fake):
        parent = fake.add("既存")
        plan = plan_of(reg([["DEMO-1", None, "子", "タスク", "中"]]))
        assert plan.ok
        assert plan.creates[0].parent_key == "DEMO-1"
        assert plan.creates[0].parent_issue_id == parent["id"]

    def test_小文字の課題キーも受け付ける(self, plan_of, fake):
        fake.add("既存")
        plan = plan_of(reg([["demo-1", None, "子", "タスク", "中"]]))
        assert plan.ok
        assert plan.creates[0].parent_key == "DEMO-1"

    def test_丸と親課題キーの両方はエラー(self, plan_of, fake):
        fake.add("既存")
        plan = plan_of(reg([
            [None, None, "親", "タスク", "中"],
            ["DEMO-1", "○", "子", "タスク", "中"],
        ]))
        assert any("両方" in m for m in messages(plan))

    def test_存在しない課題(self, plan_of):
        plan = plan_of(reg([["DEMO-99", None, "子", "タスク", "中"]]))
        assert any("DEMO-99 が見つかりません" in m for m in messages(plan))

    def test_子課題は親にできない(self, plan_of, fake):
        p = fake.add("親")
        fake.add("子", parent=p)   # DEMO-2
        plan = plan_of(reg([["DEMO-2", None, "孫", "タスク", "中"]]))
        assert any("1 階層まで" in m for m in messages(plan))

    def test_別のプロジェクトの課題はエラー(self, plan_of):
        plan = plan_of(reg([["OTHER-1", None, "子", "タスク", "中"]]))
        assert any("別のプロジェクト" in m for m in messages(plan))

    def test_解除は登録シートでは使えない(self, plan_of):
        plan = plan_of(reg([["(解除)", None, "子", "タスク", "中"]]))
        assert any("更新シートでだけ" in m for m in messages(plan))

    def test_親子課題が無効なプロジェクトではエラー(self, make_book, fake):
        from master import Master
        from planner import build_plan
        from sheet_io import read_workbook

        fake.project["subtaskingEnabled"] = False
        master = Master.load(fake, "DEMO")
        plan = build_plan(read_workbook(make_book(reg([
            [None, None, "親", "タスク", "中"],
            [None, "○", "子", "タスク", "中"],
        ]))), master, fake)
        assert any("親子課題が無効" in m for m in messages(plan))


class TestValues:
    def test_名前を_ID_に変換する(self, plan_of):
        headers = H + ["担当者", "カテゴリー", "期限日", "予定時間"]
        plan = plan_of(reg([[None, None, "A", "バグ", "高", "山田太郎", "設計\n開発", dt.datetime(2026, 10, 31), 1.5]], headers))
        assert plan.ok, messages(plan)
        v = plan.creates[0].values
        assert v["issueTypeId"] == 2 and v["priorityId"] == 2 and v["assigneeId"] == 10
        assert v["categoryId"] == (100, 101)
        assert v["dueDate"] == "2026-10-31"
        assert v["estimatedHours"] == 1.5

    def test_ログイン_ID_でも担当者を引ける(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "sato"]], H + ["担当者"]))
        assert plan.creates[0].values["assigneeId"] == 11

    def test_無い名前は候補を添えてエラー(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "山田"]], H + ["担当者"]))
        assert any("山田太郎" in m and "もしかして" in m for m in messages(plan))

    def test_必須の項目が空(self, plan_of):
        plan = plan_of(reg([[None, None, "A", None, "中"]]))
        assert any("種別" in m and "必須" in m for m in messages(plan))

    def test_必須の列が無い(self, plan_of):
        plan = plan_of(reg([["A", "タスク"]], ["件名", "種別"]))
        assert any("必須の列「優先度」" in m for m in messages(plan))

    def test_削除は登録シートでは使えない(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "(削除)"]], H + ["担当者"]))
        assert any("(削除) は使えません" in m for m in messages(plan))

    def test_状態は作成後に更新で送る(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "処理中", "対応済み"]], H + ["状態", "完了理由"]))
        cp = plan.creates[0]
        assert "statusId" not in cp.values
        assert cp.follow_up == {"statusId": 2, "resolutionId": 0}

    def test_状態が未対応なら作成後の更新はしない(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "未対応"]], H + ["状態"]))
        assert plan.creates[0].follow_up == {}

    def test_日付として読めない(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "来週"]], H + ["期限日"]))
        assert any("日付として読めません" in m for m in messages(plan))



class TestNoDoubleReport:
    """名前が見つからない列は、空ではない。「必須の項目が空です」を重ねて出さない。"""

    def test_必須の列の名前が見つからなければ_そのエラーだけ(self, plan_of):
        plan = plan_of(reg([[None, None, "a", "無い種別", "無い優先度"]]))
        assert messages(plan) == [
            "登録 2行目「種別」: 種別「無い種別」が見つかりません",
            "登録 2行目「優先度」: 優先度「無い優先度」が見つかりません",
        ]

    def test_空なら必須のエラーは出る(self, plan_of):
        plan = plan_of(reg([[None, None, "a", None, "中"]]))
        assert messages(plan) == ["登録 2行目「種別」: 必須の項目が空です"]

    def test_必須のカスタム属性も重ねて出さない(self, fake, make_book, monkeypatch):
        import conftest
        from master import Master
        from planner import build_plan
        from sheet_io import read_workbook

        required = {"id": 306, "typeId": 5, "name": "必須区分", "required": True,
                    "items": [{"id": 21, "name": "甲"}]}
        monkeypatch.setattr(conftest, "CUSTOM_FIELDS", conftest.CUSTOM_FIELDS + [required])
        master = Master.load(fake, "DEMO")

        def problems(value):
            book = make_book(reg([[None, None, "a", "タスク", "中", value]], H + ["必須区分"]))
            return messages(build_plan(read_workbook(book), master, fake))

        assert problems("乙") == ["登録 2行目「必須区分」: 「必須区分」の選択肢「乙」が見つかりません"]
        assert problems(None) == ["登録 2行目「必須区分」: 必須のカスタム属性が空です"]

class TestHeaders:
    def test_シャープで始まる列は読み飛ばす(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "メモ"]], H + ["#備考"]))
        assert plan.ok

    def test_知らない見出しはエラー(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "x"]], H + ["顧客"]))
        assert any("「顧客」" in m and "顧客名" in m for m in messages(plan))

    def test_課題キー列は置けない(self, plan_of):
        plan = plan_of(reg([["DEMO-1", "A", "タスク", "中"]], ["課題キー", "件名", "種別", "優先度"]))
        assert any("課題キーの列を置けません" in m for m in messages(plan))

    def test_同じ見出しが2つ(self, plan_of):
        plan = plan_of(reg([["A", "B", "タスク", "中"]], ["件名", "件名", "種別", "優先度"]))
        assert any("同じ見出し" in m for m in messages(plan))

    def test_列の並びは問わない(self, plan_of):
        plan = plan_of(reg([["中", "タスク", "A"]], ["優先度", "種別", "件名"]))
        assert plan.ok
        assert plan.creates[0].values["summary"] == "A"


class TestCustomFields:
    def test_列を足すだけで使える(self, plan_of):
        headers = H + ["顧客名", "見積", "区分", "タグ", "締切"]
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "ACME", "3", "B", "x, y", "2026/1/5"]], headers))
        assert plan.ok, messages(plan)
        v = plan.creates[0].values
        assert v["customField_300"] == "ACME"
        assert v["customField_301"] == 3.0
        assert v["customField_302"] == 2
        assert v["customField_303"] == (11, 12)
        assert v["customField_304"] == "2026-01-05"

    def test_種別で使えない属性はエラー(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "x"]], H + ["バグ専用"]))
        assert any("種別「タスク」では使えません" in m for m in messages(plan))

    def test_選択肢に無い値(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", "C"]], H + ["区分"]))
        assert any("選択肢" in m and "「C」" in m for m in messages(plan))

    def test_必須のカスタム属性(self, make_book, fake):
        import conftest
        from master import Master
        from planner import build_plan
        from sheet_io import read_workbook

        fields = [dict(cf) for cf in conftest.CUSTOM_FIELDS]
        fields[0]["required"] = True
        fake.get_custom_fields = lambda key: fields
        master = Master.load(fake, "DEMO")
        plan = build_plan(read_workbook(make_book(reg([[None, None, "A", "タスク", "中"]]))), master, fake)
        assert any("顧客名" in m and "必須のカスタム属性" in m for m in messages(plan))

    def test_空のセルは送らない(self, plan_of):
        plan = plan_of(reg([[None, None, "A", "タスク", "中", None]], H + ["顧客名"]))
        assert "customField_300" not in plan.creates[0].values
        assert CLEAR not in plan.creates[0].values.values()
