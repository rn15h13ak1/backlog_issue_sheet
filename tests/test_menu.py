"""対話メニュー（入力から組み立てる引数）"""

import builtins

import pytest

import backlog_issue_sheet as cli
import menu


def feed(monkeypatch, answers):
    it = iter(answers)
    monkeypatch.setattr(builtins, "input", lambda prompt="": next(it))


@pytest.fixture(autouse=True)
def no_history(monkeypatch, tmp_path):
    monkeypatch.setattr(menu, "HISTORY_PATH", tmp_path / "history.json")


def test_実行は_execute_を付け_件数も渡せる(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    feed(monkeypatch, [f'"{book}"', "3"])
    assert menu.build_args("register_execute", None, {}) == [
        "import", str(book), "--sheet", "登録", "--from-menu", "--execute", "--limit", "3",
    ]


def test_前回のファイルを既定にする(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    history = {}
    feed(monkeypatch, [str(book)])
    menu.build_args("register_dry", "c.yaml", history)
    feed(monkeypatch, [""])
    assert menu.build_args("update_dry", "c.yaml", history) == [
        "import", str(book), "--sheet", "更新", "--config", "c.yaml", "--from-menu",
    ]


def test_無いファイルは聞き直す(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    feed(monkeypatch, [str(tmp_path / "none.xlsx"), str(book)])
    assert menu.build_args("register_dry", None, {})[1] == str(book)


def test_書き出しで親を指定(monkeypatch):
    feed(monkeypatch, ["3", "PROJ-1"])
    assert menu.build_args("export", None, {}) == ["export", "--parent", "PROJ-1"]


def test_件数に数以外を入れると戻る(monkeypatch, tmp_path):
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    feed(monkeypatch, [str(book), "abc"])
    assert menu.build_args("register_execute", None, {}) is None


def test_件数の欄は空_Enter_が全件だと案内する(monkeypatch, tmp_path):
    """空 Enter は全件を送る。「戻る」と案内すると、取りやめるつもりで送ってしまう。"""
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    prompts = []
    answers = iter([str(book), ""])
    monkeypatch.setattr(builtins, "input", lambda prompt="": prompts.append(prompt) or next(answers))
    assert "--limit" not in menu.build_args("update_execute", None, {})
    assert prompts[-1] == "  先頭から何件だけ送るか（空 Enter で全件）: "


def test_ドライランの案内に使う表示名がメニューにある():
    labels = {key: label for key, label, _ in menu.ACTIONS}
    for action, (sheet, execute) in menu.IMPORTS.items():
        if execute:
            assert labels[action] == cli.menu_execute_label(sheet)


def test_登録と更新は最初のメニューで選び分ける(monkeypatch, tmp_path):
    """途中でシートを選ばせると、確認と実行で別のシートを選ぶ取り違えが起きうる。"""
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    expected = {
        "register_dry": ["--sheet", "登録", "--from-menu"],
        "register_execute": ["--sheet", "登録", "--from-menu", "--execute"],
        "update_dry": ["--sheet", "更新", "--from-menu"],
        "update_execute": ["--sheet", "更新", "--from-menu", "--execute"],
    }
    for action, tail in expected.items():
        feed(monkeypatch, [str(book), ""])
        assert menu.build_args(action, None, {}) == ["import", str(book), *tail]


def test_使い方シートが案内するメニューの項目はメニューにある():
    """メニューの項目名を変えたのに使い方シートを直し忘れると、無い項目を案内してしまう。"""
    import re

    from sheet_io import GUIDE_LINES

    referred = re.findall(r"メニューの「(.+?)」", "\n".join(GUIDE_LINES))
    assert referred, "使い方シートがメニューの項目を案内していない"
    labels = [label for _, label, _ in menu.ACTIONS]
    assert [r for r in referred if r not in labels] == []
