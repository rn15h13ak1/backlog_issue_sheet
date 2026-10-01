"""CLI（設定・API キー・サブコマンドの通し）"""

import pytest

import backlog_issue_sheet as cli
import sheet_io
from sheet_io import SheetError, read_workbook, unique_stamp

CONFIG = """
backlog:
  space_host: "example.backlog.com"
  project_key: "DEMO"
"""

REG = ["親課題キー", "子課題", "件名", "種別", "優先度"]


@pytest.fixture
def env(tmp_path, monkeypatch, fake):
    """設定ファイル・API キー・偽のクライアントを用意して、作業ディレクトリを移す。"""
    (tmp_path / "config.yaml").write_text(CONFIG, encoding="utf-8")
    monkeypatch.setenv(cli.API_KEY_ENV, "dummy")
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(cli, "make_client", lambda config, key, debug: fake)
    return tmp_path


class TestConfig:
    def test_API_キーを設定ファイルに書ける(self, tmp_path, monkeypatch):
        monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
        monkeypatch.chdir(tmp_path)
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG + '  api_key: "from-config"\n', encoding="utf-8")
        assert cli.find_api_key(path, cli.load_config(path)) == "from-config"

    def test_環境変数は設定ファイルより先(self, tmp_path, monkeypatch):
        """設定ファイルを書き換えずに差し替えられるように（共通規約 C）。"""
        monkeypatch.setenv(cli.API_KEY_ENV, "from-env")
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG + '  api_key: "from-config"\n', encoding="utf-8")
        assert cli.find_api_key(path, cli.load_config(path)) == "from-env"

    def test_設定ファイルは_env_より先(self, tmp_path, monkeypatch):
        monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
        monkeypatch.chdir(tmp_path)
        (tmp_path / ".env").write_text(f"{cli.API_KEY_ENV}=from-dotenv\n", encoding="utf-8")
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG + '  api_key: "from-config"\n', encoding="utf-8")
        assert cli.find_api_key(path, cli.load_config(path)) == "from-config"

    def test_例の値のままの_api_key_は書いていないものとみなす(self, tmp_path, monkeypatch):
        """テンプレートをコピーしただけの config.yaml でも、.env が使われるように。"""
        monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
        monkeypatch.chdir(tmp_path)
        (tmp_path / ".env").write_text(f"{cli.API_KEY_ENV}=from-dotenv\n", encoding="utf-8")
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG + '  api_key: "YOUR_API_KEY_HERE"\n', encoding="utf-8")
        assert cli.find_api_key(path, cli.load_config(path)) == "from-dotenv"

    def test_例の値のままで_ほかにも無ければエラー(self, tmp_path, monkeypatch):
        monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
        monkeypatch.chdir(tmp_path)
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG + '  api_key: "YOUR_API_KEY_HERE"\n', encoding="utf-8")
        with pytest.raises(cli.ConfigError, match="backlog.api_key"):
            cli.find_api_key(path, cli.load_config(path))

    def test_同梱の設定例の_api_key_の項目は受け付ける(self, tmp_path):
        """例の値のままのホスト名だけがエラーになる。"""
        path = tmp_path / "c.yaml"
        path.write_text((cli.TOOL_DIR / "config.example.yaml").read_text(encoding="utf-8"), encoding="utf-8")
        with pytest.raises(cli.ConfigError) as e:
            cli.load_config(path)
        # 1 行目はファイルのパス（テスト名を含む）なので、問題を並べた 2 行目以降を見る
        assert str(e.value).splitlines()[1:] == ["  backlog.space_host が例のままです"]

    def test_例のままのホストはエラー(self, tmp_path):
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG.replace("example.backlog.com", "yourcompany.backlog.com"), encoding="utf-8")
        with pytest.raises(cli.ConfigError, match="例のまま"):
            cli.load_config(path)

    def test_未知の項目はエラー(self, tmp_path):
        path = tmp_path / "c.yaml"
        path.write_text(CONFIG + "extra: 1\n", encoding="utf-8")
        with pytest.raises(cli.ConfigError, match="未知の項目: extra"):
            cli.load_config(path)

    def test_API_キーは_env_ファイルからも読む(self, tmp_path, monkeypatch):
        monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
        monkeypatch.chdir(tmp_path)
        (tmp_path / ".env").write_text(f'# memo\nexport {cli.API_KEY_ENV}="abc"\n', encoding="utf-8")
        assert cli.find_api_key(tmp_path / "config.yaml") == "abc"

    def test_環境変数を優先する(self, tmp_path, monkeypatch):
        monkeypatch.setenv(cli.API_KEY_ENV, "from-env")
        (tmp_path / ".env").write_text(f"{cli.API_KEY_ENV}=from-file\n", encoding="utf-8")
        assert cli.find_api_key(tmp_path / "config.yaml") == "from-env"

    def test_API_キーが無ければエラー(self, tmp_path, monkeypatch):
        monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
        monkeypatch.chdir(tmp_path)
        with pytest.raises(cli.ConfigError, match="API キーがありません"):
            cli.find_api_key(tmp_path / "config.yaml")

    def test_設定ファイルが無ければ終了コード2(self, tmp_path, monkeypatch, capsys):
        monkeypatch.chdir(tmp_path)
        monkeypatch.setattr(cli, "TOOL_DIR", tmp_path)
        assert cli.main(["master"]) == cli.EXIT_USAGE
        assert "config.yaml が見つかりません" in capsys.readouterr().err


class TestImport:
    def test_ドライランでは送らない(self, env, fake, make_book, capsys):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録"]) == cli.EXIT_OK
        assert fake.calls == []
        assert "ドライラン" in capsys.readouterr().out

    def test_エラーがあれば終了コード1で何も送らない(self, env, fake, make_book, capsys):
        book = make_book({"登録": (REG, [[None, "○", "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_FAILED
        assert fake.calls == []
        assert "エラー（1 件）" in capsys.readouterr().out

    def test_実行すると実行ログと結果のファイルを出す(self, env, fake, make_book, capsys):
        book = make_book({"登録": (REG, [
            [None, None, "親", "タスク", "中"],
            [None, "○", "子", "タスク", "中"],
        ])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_OK
        assert len(fake.calls) == 2
        out = env / "output"
        logs = list(out.glob("run_*.csv"))
        results = list(out.glob("結果_*.xlsx"))
        assert len(logs) == 1 and len(results) == 1

        # 結果のファイルは更新シートの書式で、そのまま取り込みに戻せる
        sheets = read_workbook(results[0])
        assert list(sheets) == ["更新"]
        rows = [[r.values[1], r.values[2]] for r in sheets["更新"].rows]
        assert rows == [["DEMO-1", None], ["DEMO-2", "DEMO-1"]]

    def test_対話できなければ確認なしには送らない(self, env, fake, make_book, monkeypatch):
        monkeypatch.setattr("sys.stdin.isatty", lambda: False)
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute"]) == cli.EXIT_OK
        assert fake.calls == []

    def test_失敗があれば終了コード1(self, env, fake, make_book):
        fake.fail_create_matching = "a"
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_FAILED

    def test_limit_は1以上(self, env, make_book):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--limit", "0"]) == cli.EXIT_USAGE

    def test_別のプロジェクト向けのブックは送らずに終了コード1(self, env, fake, make_book, capsys):
        book = make_book({
            "登録": (REG, [[None, None, "a", "タスク", "中"]]),
            "プロジェクト": (["プロジェクトキー", "OTHER"], []),
        })
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_FAILED
        assert fake.calls == []
        assert "OTHER 向け" in capsys.readouterr().out

    def test_ドライランの案内はメニューから呼ばれたらメニューの操作で示す(self, env, fake, make_book, capsys):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録"]) == cli.EXIT_OK
        assert "--execute を付けてください" in capsys.readouterr().out
        assert cli.main(["import", str(book), "--sheet", "登録", "--from-menu"]) == cli.EXIT_OK
        out = capsys.readouterr().out
        assert "メニューの「登録（実行）」を選んでください" in out and "--execute" not in out

        fake.add("x")
        book = make_book({"更新": (["課題キー", "件名"], [["DEMO-1", "y"]])})
        assert cli.main(["import", str(book), "--sheet", "更新", "--from-menu"]) == cli.EXIT_OK
        assert "メニューの「更新（実行）」を選んでください" in capsys.readouterr().out



class TestSelectSheet:
    """取り込むシートは 1 つだけ指定する。書いたまま残っていたもう一方のシートの行を送らないため。"""

    @pytest.fixture
    def both(self, fake, make_book):
        fake.add("既存")
        return make_book({
            "登録": (REG, [[None, None, "新規", "タスク", "中"]]),
            "更新": (["課題キー", "件名"], [["DEMO-1", "変更"]]),
        })

    def test_登録を指定すると更新シートは送らない(self, env, fake, both, capsys):
        assert cli.main(["import", str(both), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_OK
        assert [c[0] for c in fake.calls] == ["create"]
        assert fake.get_issue("DEMO-1")["summary"] == "既存"
        assert "「更新」シートにも 1 行ありますが、今回は読みません" in capsys.readouterr().out

    def test_更新を指定すると登録シートは送らない(self, env, fake, both, capsys):
        assert cli.main(["import", str(both), "--sheet", "更新", "--execute", "--yes"]) == cli.EXIT_OK
        assert fake.calls == [("update", "DEMO-1", {"summary": "変更"})]
        assert "「登録」シートにも 1 行ありますが、今回は読みません" in capsys.readouterr().out

    def test_もう一方のシートが空なら知らせない(self, env, make_book, capsys):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]]), "更新": (["課題キー"], [])})
        assert cli.main(["import", str(book), "--sheet", "登録"]) == cli.EXIT_OK
        assert "今回は読みません" not in capsys.readouterr().out

    def test_指定したシートが無ければ終了コード2(self, env, make_book, capsys):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_USAGE
        assert "「更新」シートがありません" in capsys.readouterr().err

    def test_シートの指定は必須(self, env, make_book, capsys):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        with pytest.raises(SystemExit) as e:
            cli.main(["import", str(book)])
        assert e.value.code == 2
        assert "--sheet" in capsys.readouterr().err


class TestDuplicates:
    def test_同じ登録シートを2回実行しても二重には作らない(self, env, fake, make_book, capsys):
        book = make_book({"登録": (REG, [
            [None, None, "親", "タスク", "中"],
            [None, "○", "子", "タスク", "中"],
        ])})
        args = ["import", str(book), "--sheet", "登録", "--execute", "--yes"]
        assert cli.main(args) == cli.EXIT_OK
        capsys.readouterr()
        assert cli.main(args) == cli.EXIT_FAILED
        assert len(fake.calls) == 2              # 1 回目の 2 件だけ
        out = capsys.readouterr().out
        assert "同じ件名の未完了の課題があります（DEMO-1）" in out
        assert "--allow-duplicates を付けて実行してください" in out

    def test_メニューからはコマンドで付けるよう案内する(self, env, fake, make_book, capsys):
        fake.add("a")
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--from-menu"]) == cli.EXIT_FAILED
        assert "コマンドで --allow-duplicates を付けて実行してください" in capsys.readouterr().out

class TestOtherCommands:
    def test_template(self, env):
        assert cli.main(["template", "-o", "t.xlsx"]) == cli.EXIT_OK
        assert (env / "t.xlsx").is_file()

    def test_export(self, env, fake):
        fake.add("x")
        assert cli.main(["export", "-o", "e.xlsx"]) == cli.EXIT_OK
        assert list(read_workbook(env / "e.xlsx")) == ["更新"]

    def test_export_の名前の誤りは終了コード2(self, env, capsys):
        assert cli.main(["export", "--status", "完了済み"]) == cli.EXIT_USAGE
        assert "もしかして" in capsys.readouterr().err

    def test_master(self, env, capsys):
        assert cli.main(["master"]) == cli.EXIT_OK
        out = capsys.readouterr().out
        assert "区分  [単一リスト]" in out
        assert "山田太郎（yamada） / 佐藤花子（sato）" in out

    def test_master_は同姓同名をログイン_ID_で見分けられるよう両方出す(self, env, fake, capsys, monkeypatch):
        """同姓同名は名前では書けず、ログイン ID で書くしかない。その ID を確かめる場所が要る。"""
        import conftest

        monkeypatch.setattr(conftest, "USERS", conftest.USERS + [{"id": 12, "name": "山田太郎", "userId": "yamada2"}])
        assert cli.main(["master"]) == cli.EXIT_OK
        assert "山田太郎（yamada） / 佐藤花子（sato） / 山田太郎（yamada2）" in capsys.readouterr().out



class TestReadWorkbook:
    def test_ファイルが無い(self, tmp_path):
        with pytest.raises(SheetError, match="見つかりません"):
            read_workbook(tmp_path / "none.xlsx")

    def test_xls_は読めない(self, tmp_path):
        path = tmp_path / "a.xls"
        path.write_bytes(b"")
        with pytest.raises(SheetError, match="xlsx / xlsm 以外"):
            read_workbook(path)

    def test_どちらのシートも無い(self, make_book):
        with pytest.raises(SheetError, match="Sheet"):
            read_workbook(make_book({"Sheet": (["件名"], [["a"]])}))


class TestPlanDisplay:
    def test_更新の差分を旧と新で示す(self, env, fake, make_book, capsys):
        p = fake.add("親")
        fake.add("x", assigneeId=10)
        book = make_book({"更新": (["課題キー", "親課題キー", "担当者"], [["DEMO-2", "DEMO-1", "佐藤花子"]])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_OK
        out = capsys.readouterr().out
        assert "親課題キー: （なし） → DEMO-1" in out
        assert "担当者: 山田太郎 → 佐藤花子" in out
        assert "更新 1 件 / 変更なし 0 件" in out

    def test_作成は親子をツリーで示し_作成後の更新も示す(self, env, make_book, capsys):
        book = make_book({"登録": (REG + ["状態"], [
            [None, None, "親", "タスク", "中", None],
            [None, "○", "子", "タスク", "中", "処理中"],
        ])})
        cli.main(["import", str(book), "--sheet", "登録"])
        out = capsys.readouterr().out
        assert "└ 子（親: 2行目）" in out
        assert "状態=処理中（作成後に更新）" in out


    def test_エラーのある行に印を付け_見出しにエラーの行数を出す(self, env, make_book, capsys):
        """エラーのある行も予定に並ぶ。印が無いと、件名の空の行まで作られるように読める。"""
        book = make_book({"登録": (REG, [
            [None, None, "a", "タスク", "中"],
            [None, None, None, "タスク", "中"],
            [None, None, "c", "タスケ", "中"],
        ])})
        assert cli.main(["import", str(book), "--sheet", "登録"]) == cli.EXIT_FAILED
        lines = capsys.readouterr().out.splitlines()
        assert "■ 登録シート（作成 3 件（うちエラー 2 行））" in lines
        assert any(line.startswith("  2行目   a") for line in lines)
        assert any(line.startswith("✗ 3行目") for line in lines)
        assert any(line.startswith("✗ 4行目   c") for line in lines)

    def test_エラーが無ければ印も行数も出さない(self, env, make_book, capsys):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録"]) == cli.EXIT_OK
        out = capsys.readouterr().out
        assert "■ 登録シート（作成 1 件）" in out and "✗" not in out

    def test_更新シートにも印と行数を出す(self, env, fake, make_book, capsys):
        fake.add("x")
        fake.add("y")
        book = make_book({"更新": (["課題キー", "親課題キー"], [["DEMO-1", "DEMO-1"], ["DEMO-2", None]])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_FAILED
        out = capsys.readouterr().out
        assert "■ 更新シート（更新 0 件 / 変更なし 2 件（うちエラー 1 行））" in out
        assert "✗ 2行目   DEMO-1 x  変更なし" in out and "  3行目   DEMO-2 y  変更なし" in out

    def test_課題キーの誤りで一覧に載らない行を示す(self, env, fake, make_book, capsys):
        """どの課題か決まらない行は一覧に載らず、見出しの件数にも入らない。"""
        fake.add("x")
        book = make_book({"更新": (["課題キー", "件名"], [
            ["DEMO-1", "y"], [None, "a"], ["DEMO-99", "b"], ["DEMO-1", "c"],
        ])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_FAILED
        out = capsys.readouterr().out
        assert "✗ ほかに、課題キーの誤りで一覧に載らない行が 3 行あります（3、4、5 行目）" in out

    def test_すべての行が課題キーの誤りでも見出しと行を出す(self, env, make_book, capsys):
        book = make_book({"更新": (["課題キー", "件名"], [["DEMO-99", "b"]])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_FAILED
        out = capsys.readouterr().out
        assert "■ 更新シート（更新 0 件 / 変更なし 0 件）" in out
        assert "一覧に載らない行が 1 行あります（2 行目）" in out

    def test_課題キーの誤りが無ければ案内しない(self, env, fake, make_book, capsys):
        fake.add("x")
        book = make_book({"更新": (["課題キー", "件名"], [["DEMO-1", "(削除)"]])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_FAILED
        # 読み込みの行に出るファイルのパスにはテスト名が入る。テスト名に無い語句で確かめる
        assert "誤りで一覧に載らない" not in capsys.readouterr().out

    def test_見出しの誤りは一覧に載らない行に数えない(self, env, fake, make_book, capsys):
        """見出しの誤りは 1 行目のエラー。データの行ではない。"""
        fake.add("x")
        book = make_book({"更新": (["課題キー", "担当"], [["DEMO-1", "a"]])})
        assert cli.main(["import", str(book), "--sheet", "更新"]) == cli.EXIT_FAILED
        assert "誤りで一覧に載らない" not in capsys.readouterr().out

class TestOutputNames:
    """日時は秒までなので、同じ秒に続けて実行すると名前が重なる。前のファイルを消さない。"""

    @pytest.fixture
    def same_second(self, monkeypatch):
        monkeypatch.setattr(sheet_io, "timestamp", lambda now=None: "20261001_120000")

    def test_重ならなければ日時のまま(self, tmp_path, same_second):
        assert unique_stamp(tmp_path, ["a_{}.xlsx"]) == "20261001_120000"

    def test_既にあれば番号を付ける(self, tmp_path, same_second):
        (tmp_path / "a_20261001_120000.xlsx").touch()
        (tmp_path / "a_20261001_120000_2.xlsx").touch()
        assert unique_stamp(tmp_path, ["a_{}.xlsx"]) == "20261001_120000_3"

    def test_対になるファイルは同じ番号にそろえる(self, tmp_path, same_second):
        """結果のファイルだけが残っていても、実行ログと番号がずれないようにする。"""
        (tmp_path / "結果_20261001_120000.xlsx").touch()
        assert unique_stamp(tmp_path, ["run_{}.csv", "結果_{}.xlsx"]) == "20261001_120000_2"

    def test_同じ秒に2回書き出しても両方残る(self, env, fake, same_second):
        fake.add("x")
        assert cli.main(["export"]) == cli.EXIT_OK
        assert cli.main(["export"]) == cli.EXIT_OK
        names = sorted(f.name for f in (env / "output").iterdir())
        assert names == ["書き出し_20261001_120000.xlsx", "書き出し_20261001_120000_2.xlsx"]

    def test_同じ秒に2回取り込んでも実行ログが両方残る(self, env, fake, make_book, same_second):
        book = make_book({"登録": (REG, [[None, None, "a", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_OK
        # 同じ件名の課題ができているので、2 回目は止められる。ここではファイル名だけを見る
        args = ["import", str(book), "--sheet", "登録", "--execute", "--yes", "--allow-duplicates"]
        assert cli.main(args) == cli.EXIT_OK
        names = sorted(f.name for f in (env / "output").iterdir())
        assert names == [
            "run_20261001_120000.csv", "run_20261001_120000_2.csv",
            "結果_20261001_120000.xlsx", "結果_20261001_120000_2.xlsx",
        ]
