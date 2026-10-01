"""
backlog-issue-sheet
===================
Excel の「登録」「更新」シートから Backlog の課題を作成・更新する。
Backlog の課題を、同じ書式の Excel に書き出すこともできる。

  backlog-issue-sheet template                 ひな形を作る
  backlog-issue-sheet import 課題.xlsx           ドライラン（既定）
  backlog-issue-sheet import 課題.xlsx --execute
  backlog-issue-sheet export                   Backlog → Excel
  backlog-issue-sheet master                   使える名前の一覧
"""

from __future__ import annotations

import argparse
import os
import sys
import unicodedata
from pathlib import Path

try:
    import openpyxl  # noqa: F401
    import yaml
except ModuleNotFoundError as e:  # pragma: no cover - 依存が無い環境でだけ通る
    print(f"必要なライブラリ {e.name} が入っていません。", file=sys.stderr)
    print(f"実行中の Python: {sys.executable}", file=sys.stderr)
    print("README の「インストール」の手順で入れてください。", file=sys.stderr)
    sys.exit(1)

import executor
from backlog_client import BacklogAPIError, BacklogClient
from exporter import ExportError, ExportFilter, export, parent_key_map, write_issues
from fields import (
    PARENT_HEADER,
    REGISTER_SHEET,
    UPDATE_SHEET,
    all_columns,
    display,
)
from master import CF_TYPE_NAMES, Master
from planner import Plan, build_plan
from sheet_io import SheetError, read_workbook, unique_stamp, write_workbook

TOOL_DIR = Path(__file__).resolve().parent
# メニュー（menu.py）の「取り込み（実行）」の表示名。ドライランの案内に使う
MENU_EXECUTE_LABEL = "取り込み（実行）"
API_KEY_ENV = "BACKLOG_API_KEY"

EXIT_OK = 0
EXIT_FAILED = 1        # 検証エラー・送信の失敗
EXIT_USAGE = 2         # 設定・引数の誤り
EXIT_INTERRUPTED = 130


class ConfigError(Exception):
    pass


# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------

ALLOWED_KEYS = {"backlog", "output_dir"}
ALLOWED_BACKLOG_KEYS = {"space_host", "project_key", "ssl_verify", "base_path"}


def find_config(explicit: str | None) -> Path:
    if explicit:
        path = Path(explicit)
        if not path.is_file():
            raise ConfigError(f"設定ファイルが見つかりません: {path}")
        return path
    for candidate in (Path.cwd() / "config.yaml", TOOL_DIR / "config.yaml"):
        if candidate.is_file():
            return candidate
    raise ConfigError(
        "設定ファイル config.yaml が見つかりません。"
        "config.example.yaml をコピーして作ってください（--config で場所を指定することもできます）"
    )


def load_config(path: Path) -> dict:
    try:
        config = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as e:
        raise ConfigError(f"設定ファイルを読めません: {path}\n  {e}") from e
    if not isinstance(config, dict):
        raise ConfigError(f"設定ファイルの形が不正です: {path}")

    problems = [f"未知の項目: {k}" for k in config if k not in ALLOWED_KEYS]
    backlog = config.get("backlog")
    if not isinstance(backlog, dict):
        problems.append("backlog の節がありません")
        backlog = {}
    problems += [f"未知の項目: backlog.{k}" for k in backlog if k not in ALLOWED_BACKLOG_KEYS]
    if "api_key" in backlog:
        problems.append(
            f"API キーは設定ファイルに書かず、環境変数 {API_KEY_ENV} か .env に書いてください"
        )
    for required in ("space_host", "project_key"):
        if not backlog.get(required):
            problems.append(f"backlog.{required} が空です")
    if backlog.get("space_host") == "yourcompany.backlog.com":
        problems.append("backlog.space_host が例のままです")
    if problems:
        raise ConfigError(f"設定ファイルに誤りがあります: {path}\n  " + "\n  ".join(problems))
    return config


def read_dotenv(path: Path) -> dict[str, str]:
    """KEY=VALUE の行だけを読む簡易版。# で始まる行と空行は飛ばす。"""
    values = {}
    if not path.is_file():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip().removeprefix("export ").strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key] = value
    return values


def find_api_key(config_path: Path) -> str:
    """環境変数、設定ファイルと同じ場所の .env、作業ディレクトリの .env の順に探す。"""
    if os.environ.get(API_KEY_ENV):
        return os.environ[API_KEY_ENV]
    for env_path in (config_path.parent / ".env", Path.cwd() / ".env"):
        value = read_dotenv(env_path).get(API_KEY_ENV)
        if value:
            return value
    raise ConfigError(
        f"API キーがありません。環境変数 {API_KEY_ENV} を設定するか、"
        f"設定ファイルと同じ場所の .env に {API_KEY_ENV}=... と書いてください"
    )


def output_dir(config: dict, config_path: Path, override: str | None) -> Path:
    if override:
        return Path(override)
    out = Path(config.get("output_dir") or "output")
    return out if out.is_absolute() else config_path.parent / out


def stamped_path(out_dir: Path, name: str) -> Path:
    """name の {} に日時を入れたパス。既にあるファイルとは重ならない。"""
    return out_dir / name.format(unique_stamp(out_dir, [name]))


def make_client(config: dict, api_key: str, debug: bool) -> BacklogClient:
    b = config["backlog"]
    return BacklogClient(
        space_host=b["space_host"],
        api_key=api_key,
        ssl_verify=b.get("ssl_verify", True),
        base_path=b.get("base_path", ""),
        debug=debug,
    )


# ---------------------------------------------------------------------------
# 表示
# ---------------------------------------------------------------------------

def display_width(text: str) -> int:
    return sum(2 if unicodedata.east_asian_width(c) in ("W", "F") else 1 for c in text)


def pad(text: str, width: int) -> str:
    return text + " " * max(0, width - display_width(text))


def print_problems(plan: Plan) -> None:
    if plan.warnings:
        print(f"\n注意（{len(plan.warnings)} 件）:")
        for w in plan.warnings:
            print(f"  ! {w}")
    if plan.problems:
        print(f"\nエラー（{len(plan.problems)} 件）: 直すまで Backlog には何も送りません")
        for p in plan.problems:
            print(f"  ✗ {p}")


def print_plan(plan: Plan, master: Master) -> None:
    if plan.creates:
        print(f"\n■ {REGISTER_SHEET}シート（作成 {len(plan.creates)} 件）")
        for cp in plan.creates:
            is_child = cp.parent_row is not None or cp.parent_key
            branch = "  └ " if is_child else ""
            parent = ""
            if cp.parent_row is not None:
                parent = f"（親: {cp.parent_row}行目）"
            elif cp.parent_key:
                parent = f"（親: {cp.parent_key}）"
            print(f"  {pad(f'{cp.row_no}行目', 8)}{branch}{cp.summary}{parent}")
            attrs = [
                f"{cp.columns[k].header}={display(cp.columns[k], v, master)}"
                for k, v in cp.values.items() if k != "summary"
            ]
            attrs += [
                f"{cp.columns[k].header}={display(cp.columns[k], v, master)}（作成後に更新）"
                for k, v in cp.follow_up.items()
            ]
            if attrs:
                print(" " * 12 + " / ".join(attrs))

    if plan.updates:
        changed = [u for u in plan.updates if u.has_changes]
        print(
            f"\n■ {UPDATE_SHEET}シート（更新 {len(changed)} 件 / "
            f"変更なし {len(plan.updates) - len(changed)} 件）"
        )
        for up in plan.updates:
            mark = "" if up.has_changes else "  変更なし"
            print(f"  {pad(f'{up.row_no}行目', 8)}{up.issue_key} {up.summary}{mark}")
            if up.parent_change:
                kind, parent, old_key = up.parent_change
                new = parent["issueKey"] if kind == "set" else "（なし）"
                print(" " * 12 + f"{PARENT_HEADER}: {old_key or '（なし）'} → {new}")
            for c in up.changes:
                print(
                    " " * 12 + f"{c.column.header}: "
                    f"{display(c.column, c.old, master)} → {display(c.column, c.new, master)}"
                )


def ask(prompt: str, assume_yes: bool) -> bool:
    if assume_yes:
        print(f"{prompt} y（--yes 指定）")
        return True
    if not sys.stdin.isatty():
        print("  確認できないため中止します（自動で実行するなら --yes を付けてください）", file=sys.stderr)
        return False
    try:
        return input(prompt).strip().lower() in ("y", "yes")
    except EOFError:
        return False


# ---------------------------------------------------------------------------
# サブコマンド
# ---------------------------------------------------------------------------

def connect(args) -> tuple[dict, Path, BacklogClient, Master]:
    config_path = find_config(args.config)
    config = load_config(config_path)
    client = make_client(config, find_api_key(config_path), args.debug)
    print(f"接続先: {config['backlog']['space_host']} / {config['backlog']['project_key']}")
    master = Master.load(client, config["backlog"]["project_key"])
    return config, config_path, client, master


def cmd_template(args) -> int:
    config, config_path, _, master = connect(args)
    out_dir = output_dir(config, config_path, args.output_dir)
    path = Path(args.output) if args.output else stamped_path(out_dir, "課題シート_{}.xlsx")
    write_workbook(path, master, include_guide=True)
    print(f"ひな形を作りました: {path}")
    print(f"  列: {len(all_columns(master))} 項目（うちカスタム属性 {len(master.custom_fields)}）")
    return EXIT_OK


def cmd_master(args) -> int:
    _, _, _, master = connect(args)
    labels = [
        ("種別", "issue_types"), ("優先度", "priorities"), ("状態", "statuses"),
        ("完了理由", "resolutions"), ("担当者（括弧内はログイン ID）", "users"), ("カテゴリー", "categories"),
        ("発生バージョン / マイルストーン", "versions"),
    ]
    for label, key in labels:
        names = master.table(key).labels()
        print(f"\n{label}")
        print("  " + (" / ".join(names) if names else "（なし）"))
    print("\nカスタム属性（列の見出しにこの名前を書く）")
    if not master.custom_fields:
        print("  （なし）")
    for cf in master.custom_fields:
        flags = []
        if cf.required:
            flags.append("必須")
        if cf.applicable_issue_types:
            types = master.table("issue_types")
            flags.append("種別: " + "、".join(types.name_of(t) for t in cf.applicable_issue_types))
        extra = f"（{' / '.join(flags)}）" if flags else ""
        print(f"  {cf.name}  [{CF_TYPE_NAMES.get(cf.type_id, cf.type_id)}]{extra}")
        if cf.items:
            print(f"      選択肢: {' / '.join(cf.items.names())}")
    if not master.subtasking_enabled:
        print("\n※ このプロジェクトでは親子課題が無効です")
    return EXIT_OK


def cmd_import(args) -> int:
    config, config_path, client, master = connect(args)
    sheets = read_workbook(args.excel)
    print(f"読み込み: {args.excel}（{' / '.join(f'{n} {len(s.rows)} 行' for n, s in sheets.items())}）")
    print("Backlog の現在の状態と照らし合わせています…")
    plan = build_plan(sheets, master, client)

    print_plan(plan, master)
    print_problems(plan)
    if not plan.ok:
        return EXIT_FAILED

    n_actions = executor.count_actions(plan)
    if n_actions == 0:
        print("\n送信するものはありません。")
        return EXIT_OK
    if not args.execute:
        how = f"メニューの「{MENU_EXECUTE_LABEL}」を選んでください" if args.from_menu else "--execute を付けてください"
        print(f"\nドライランです（送信予定 {n_actions} 件）。Backlog に反映するには{how}。")
        return EXIT_OK

    limit_note = f"（--limit {args.limit} のため先頭 {min(args.limit, n_actions)} 件）" if args.limit else ""
    if not ask(f"\nBacklog に反映しますか？ 送信 {n_actions} 件{limit_note} [y/N]: ", args.yes):
        print("中止しました。")
        return EXIT_OK

    out_dir = output_dir(config, config_path, args.output_dir)
    ts = unique_stamp(out_dir, ["run_{}.csv", "結果_{}.xlsx"])
    log_path = out_dir / f"run_{ts}.csv"
    print()
    with executor.RunLog(log_path) as log:
        results = executor.execute(plan, master.project_id, client, limit=args.limit, log=log)

    counts: dict[str, int] = {}
    for r in results:
        counts[r.outcome] = counts.get(r.outcome, 0) + 1
    print("\n結果: " + " / ".join(f"{k} {v} 件" for k, v in counts.items()))
    print(f"実行ログ: {log_path}")

    touched = [r.issue for r in results if r.issue]
    if touched:
        result_path = out_dir / f"結果_{ts}.xlsx"
        _, warnings = write_issues(
            result_path, touched, master, parent_key_map(touched, client), base_url=client.base_url,
        )
        print(f"作成・更新した課題（{UPDATE_SHEET}シートの書式）: {result_path}")
        for w in warnings:
            print(f"  ! {w}")
    return EXIT_FAILED if counts.get(executor.FAILED) or counts.get(executor.NOT_RUN) or counts.get(executor.SKIPPED) else EXIT_OK


def cmd_export(args) -> int:
    config, config_path, client, master = connect(args)
    out_dir = output_dir(config, config_path, args.output_dir)
    path = Path(args.output) if args.output else stamped_path(out_dir, "書き出し_{}.xlsx")
    flt = ExportFilter(
        statuses=args.status or [],
        types=args.type or [],
        keyword=args.keyword or "",
        parent_key=(args.parent or "").upper(),
        keys=[k.upper() for k in (args.key or [])],
        include_closed=args.include_closed,
    )
    out, count, warnings = export(client, master, path, flt)
    for w in warnings:
        print(f"  ! {w}")
    print(f"{count} 件を書き出しました: {out}")
    return EXIT_OK


# ---------------------------------------------------------------------------
# 引数
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--config", help="設定ファイル（既定: 作業ディレクトリかツールの場所の config.yaml）")
    common.add_argument("--output-dir", help="出力先フォルダ（既定: 設定の output_dir、なければ output）")
    common.add_argument("--debug", action="store_true", help="送信内容を表示する（API キーは出さない）")

    parser = argparse.ArgumentParser(
        prog="backlog-issue-sheet",
        description="Excel の「登録」「更新」シートから Backlog の課題を作成・更新する",
    )
    sub = parser.add_subparsers(dest="command", required=True, metavar="コマンド")

    p = sub.add_parser("template", parents=[common], help="ひな形の Excel を作る")
    p.add_argument("-o", "--output", help="出力するファイル")
    p.set_defaults(func=cmd_template)

    p = sub.add_parser("import", parents=[common], help="Excel の内容を Backlog に反映する（既定はドライラン）")
    p.add_argument("excel", help="取り込む Excel（xlsx / xlsm）")
    p.add_argument("--execute", action="store_true", help="実際に送信する")
    p.add_argument("--yes", action="store_true", help="確認を省く（--execute と一緒に使う）")
    p.add_argument("--limit", type=int, help="先頭から N 件だけ送信する（試しに少しだけ反映するとき）")
    # メニューから呼ばれたとき、案内をメニューの操作で示す。利用者が打つものではないため隠す
    p.add_argument("--from-menu", action="store_true", help=argparse.SUPPRESS)
    p.set_defaults(func=cmd_import)

    p = sub.add_parser("export", parents=[common], help="Backlog の課題を Excel に書き出す")
    p.add_argument("-o", "--output", help="出力するファイル")
    p.add_argument("--status", nargs="+", metavar="名前", help="状態で絞る（既定: 完了以外）")
    p.add_argument("--type", nargs="+", metavar="名前", help="種別で絞る")
    p.add_argument("--keyword", help="キーワードで絞る")
    p.add_argument("--parent", metavar="課題キー", help="この課題とその子課題を書き出す")
    p.add_argument("--key", nargs="+", metavar="課題キー", help="課題キーを指定して書き出す")
    p.add_argument("--include-closed", action="store_true", help="完了した課題も含める")
    p.set_defaults(func=cmd_export)

    p = sub.add_parser("master", parents=[common], help="Excel に書ける名前（種別・担当者・カスタム属性など）を一覧する")
    p.set_defaults(func=cmd_master)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if getattr(args, "limit", None) is not None and args.limit < 1:
        print("--limit には 1 以上を指定してください", file=sys.stderr)
        return EXIT_USAGE
    try:
        return args.func(args)
    except (ConfigError, SheetError, ExportError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return EXIT_USAGE
    except BacklogAPIError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return EXIT_FAILED
    except KeyboardInterrupt:
        print("\n中断しました。", file=sys.stderr)
        return EXIT_INTERRUPTED


if __name__ == "__main__":
    sys.exit(main())
