"""
認証・権限のエラー（401・403）の案内
====================================
API キーは環境変数・設定ファイル・.env のどれからも読む。エラーの案内は項目名で
名指しせず、今回どこから読んだかを添えて、直す場所が分かるようにする。
"""

import io
import json
import urllib.error
import urllib.request

import pytest

import backlog_issue_sheet as cli
from backlog_client import BacklogAPIError, BacklogClient

CONFIG = """
backlog:
  space_host: "example.backlog.com"
  project_key: "DEMO"
"""


def http_error(status):
    body = json.dumps({"errors": [{"message": "Authentication failure.", "code": 11}]}).encode()

    def opener(*args, **kwargs):
        raise urllib.error.HTTPError("url", status, "reason", {}, io.BytesIO(body))
    return opener


class TestHint:
    @pytest.mark.parametrize("status, hint", [
        (401, "→ API キーを確認してください。"),
        (403, "→ API キーの持ち主に、この操作の権限があるか確認してください"),
    ])
    def test_項目名で名指ししない(self, monkeypatch, status, hint):
        monkeypatch.setattr(urllib.request, "urlopen", http_error(status))
        with pytest.raises(BacklogAPIError) as e:
            BacklogClient("example.backlog.com", "k").get_project("DEMO")
        assert hint in str(e.value) and "api_key" not in str(e.value)
        assert e.value.fatal


@pytest.fixture
def workdir(tmp_path, monkeypatch, fake):
    """設定ファイルを置いて作業ディレクトリを移す。API キーの置き場は各テストで決める。"""
    monkeypatch.delenv(cli.API_KEY_ENV, raising=False)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(cli, "make_client", lambda config, key, debug: fake)
    (tmp_path / "config.yaml").write_text(CONFIG, encoding="utf-8")
    return tmp_path


def unauthorized(fake, monkeypatch):
    def fail(key):
        raise BacklogAPIError("API 呼び出しに失敗しました（HTTP 401）", status=401, fatal=True)
    monkeypatch.setattr(fake, "get_project", fail)


class TestSource:
    def test_環境変数から読んだとき(self, workdir, fake, monkeypatch, capsys):
        monkeypatch.setenv(cli.API_KEY_ENV, "k")
        unauthorized(fake, monkeypatch)
        assert cli.main(["master"]) == cli.EXIT_FAILED
        assert "  使った API キー: 環境変数 BACKLOG_API_KEY" in capsys.readouterr().err

    def test_設定ファイルから読んだとき(self, workdir, fake, monkeypatch, capsys):
        (workdir / "config.yaml").write_text(CONFIG + '  api_key: "k"\n', encoding="utf-8")
        unauthorized(fake, monkeypatch)
        assert cli.main(["master"]) == cli.EXIT_FAILED
        assert f"  使った API キー: 設定ファイル {workdir / 'config.yaml'} の backlog.api_key" in capsys.readouterr().err

    def test_env_から読んだとき(self, workdir, fake, monkeypatch, capsys):
        (workdir / ".env").write_text(f"{cli.API_KEY_ENV}=k\n", encoding="utf-8")
        unauthorized(fake, monkeypatch)
        assert cli.main(["master"]) == cli.EXIT_FAILED
        assert f"  使った API キー: {workdir / '.env'} の BACKLOG_API_KEY" in capsys.readouterr().err

    def test_認証・権限以外のエラーでは示さない(self, workdir, fake, monkeypatch, capsys):
        monkeypatch.setenv(cli.API_KEY_ENV, "k")

        def not_found(key):
            raise BacklogAPIError("API 呼び出しに失敗しました（HTTP 404）", status=404)
        monkeypatch.setattr(fake, "get_project", not_found)
        assert cli.main(["master"]) == cli.EXIT_FAILED
        assert "使った API キー" not in capsys.readouterr().err

    def test_取り込みの途中で権限のエラーになったとき(self, workdir, fake, make_book, monkeypatch, capsys):
        """閲覧はできても、課題の追加の権限が無いと、送る段階で初めてエラーになる。"""
        monkeypatch.setenv(cli.API_KEY_ENV, "k")
        fake.fatal_on_create = True
        book = make_book({"登録": (["件名", "種別", "優先度"], [["a", "タスク", "中"], ["b", "タスク", "中"]])})
        assert cli.main(["import", str(book), "--sheet", "登録", "--execute", "--yes"]) == cli.EXIT_FAILED
        out, err = capsys.readouterr()
        assert "認証・権限のエラーで中止しました。" in out
        assert "  使った API キー: 環境変数 BACKLOG_API_KEY" in err


    def test_更新の途中で権限のエラーになったとき(self, workdir, fake, make_book, monkeypatch, capsys):
        monkeypatch.setenv(cli.API_KEY_ENV, "k")
        fake.add("x")

        def forbidden(key, params):
            raise BacklogAPIError("API 呼び出しに失敗しました（HTTP 403）", status=403, fatal=True)
        monkeypatch.setattr(fake, "update_issue", forbidden)
        book = make_book({"更新": (["課題キー", "件名"], [["DEMO-1", "y"]])})
        assert cli.main(["import", str(book), "--sheet", "更新", "--execute", "--yes"]) == cli.EXIT_FAILED
        out, err = capsys.readouterr()
        assert "認証・権限のエラーで中止しました。" in out
        assert "  使った API キー: 環境変数 BACKLOG_API_KEY" in err
