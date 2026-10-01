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


def test_実行は_execute_を付け_件数は聞かない(monkeypatch, tmp_path):
    """件数の指定はなくした。一部だけ送ると、続きの扱いが複雑になり混乱を招くため。"""
    book = tmp_path / "a.xlsx"
    book.write_bytes(b"")
    feed(monkeypatch, [f'"{book}"'])        # ファイルのほかに何か聞かれると、入力が尽きて落ちる
    assert menu.build_args("register_execute", None, {}) == [
        "import", str(book), "--sheet", "登録", "--from-menu", "--execute",
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
        feed(monkeypatch, [str(book)])
        assert menu.build_args(action, None, {}) == ["import", str(book), *tail]


def test_使い方シートが案内するメニューの項目はメニューにある():
    """メニューの項目名を変えたのに使い方シートを直し忘れると、無い項目を案内してしまう。"""
    import re

    from sheet_io import GUIDE_LINES

    referred = re.findall(r"メニューの「(.+?)」", "\n".join(GUIDE_LINES))
    assert referred, "使い方シートがメニューの項目を案内していない"
    labels = [label for _, label, _ in menu.ACTIONS]
    assert [r for r in referred if r not in labels] == []


# ------------------------------------------------------------------
# 本体の呼び出しとメニューの繰り返し
# ------------------------------------------------------------------

class TestRun:
    """
    メニューは本体（backlog_issue_sheet.py）を別のプロセスとして起動する。
    本物の Backlog に接続しないよう、接続より前に必ず終わる呼び出しだけで確かめる
    （リポジトリ直下に本物の API キーの .env があっても安全なように）。
    """

    def test_本体を起動して終了コードを返す(self, capfd):
        assert menu.run(["import", "--help"]) == 0
        assert "--sheet" in capfd.readouterr().out

    def test_本体の引数の誤りは終了コード2(self, capfd):
        assert menu.run(["import", "a.xlsx"]) == 2          # --sheet が無い
        assert "--sheet" in capfd.readouterr().err


def feed_then_eof(monkeypatch, answers):
    """答えを順に返し、尽きたら EOFError（入力の終わり）にする。"""
    it = iter(answers)

    def fake_input(prompt=""):
        try:
            return next(it)
        except StopIteration:
            raise EOFError from None
    monkeypatch.setattr(builtins, "input", fake_input)


class TestMain:
    @pytest.fixture
    def calls(self, monkeypatch):
        recorded = []
        monkeypatch.setattr(menu, "run", lambda args: recorded.append(args) or 0)
        monkeypatch.setattr("sys.argv", ["menu.py"])
        return recorded

    def test_選んだ項目を本体に渡し_0で終わる(self, monkeypatch, tmp_path, calls, capsys):
        book = tmp_path / "a.xlsx"
        book.write_bytes(b"")
        feed_then_eof(monkeypatch, ["4", str(book), "", "9", "7", "", "0"])
        with pytest.raises(SystemExit) as e:
            menu.main()
        assert e.value.code == menu.EXIT_OK
        assert calls == [
            ["import", str(book), "--sheet", "更新", "--from-menu"],
            ["master"],
        ]
        out = capsys.readouterr().out
        assert "無効な入力" in out and "終了コード: 0" in out

    def test_設定ファイルを指定すると本体にも渡す(self, monkeypatch, tmp_path, calls):
        monkeypatch.setattr("sys.argv", ["menu.py", "--config", "my.yaml"])
        monkeypatch.chdir(tmp_path)
        feed_then_eof(monkeypatch, ["7", "", "0"])
        with pytest.raises(SystemExit):
            menu.main()
        assert calls == [["master", "--config", str(tmp_path / "my.yaml")]]

    def test_入力が尽きたら終了コード1(self, monkeypatch, calls):
        feed_then_eof(monkeypatch, [])
        with pytest.raises(SystemExit) as e:
            menu.main()
        assert e.value.code == menu.EXIT_EOF and calls == []

    def test_途中で戻ると本体を呼ばない(self, monkeypatch, calls):
        feed_then_eof(monkeypatch, ["2", "", "0"])            # ファイルを聞かれて空 Enter で戻る
        with pytest.raises(SystemExit):
            menu.main()
        assert calls == []


class TestExportChoices:
    @pytest.mark.parametrize("answers, expected", [
        (["1"], ["export"]),
        (["2"], ["export", "--include-closed"]),
        (["4", "ログイン"], ["export", "--keyword", "ログイン"]),
        (["0"], None),
        (["3", ""], None),
    ])
    def test_書き出す範囲(self, monkeypatch, answers, expected):
        feed(monkeypatch, answers)
        assert menu.build_args("export", None, {}) == expected
