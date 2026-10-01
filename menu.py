"""
backlog-issue-sheet 対話メニュー
================================
コマンドのオプションを覚えなくても、番号を選んで実行できるようにする。
作りは backlog_issue_cloner の menu.py にならった。

  python menu.py
  python menu.py --config my.yaml

無人実行にはメニューではなく本体を使うこと:
  backlog-issue-sheet import 課題.xlsx --execute --yes
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

WIDTH = 60
TOOL_DIR = Path(__file__).resolve().parent
SCRIPT = TOOL_DIR / "backlog_issue_sheet.py"
HISTORY_PATH = Path.home() / ".backlog_issue_sheet_menu.json"

EXIT_OK = 0
EXIT_EOF = 1
EXIT_INTERRUPTED = 130

ACTIONS = [
    ("template", "ひな形を作る", "登録・更新シートの空の Excel を作る（プルダウン付き）"),
    ("dry", "取り込み（確認だけ）", "Excel の内容を検証し、作成・変更の予定を表示する"),
    ("execute", "取り込み（実行）", "確認のうえで Backlog に反映する"),
    ("export", "書き出し", "Backlog の課題を更新シートの書式で Excel に出す"),
    ("master", "使える名前の一覧", "種別・担当者・カスタム属性などの名前を表示する"),
]


def hr(char="="):
    print(char * WIDTH)


def print_menu(title: str, items: list[str], back_label: str = "戻る") -> int:
    while True:
        print()
        hr()
        print(f"  {title}")
        hr()
        for i, item in enumerate(items, 1):
            print(f"  {i}. {item}")
        hr("-")
        print(f"  0. {back_label}")
        hr()
        choice = input("番号を入力してください: ").strip()
        if choice == "0":
            return 0
        if choice.isdigit() and 1 <= int(choice) <= len(items):
            return int(choice)
        print("  ※ 無効な入力です。もう一度入力してください。")


def clean_path(text: str) -> str:
    """ドラッグ＆ドロップで付く引用符や前後の空白を除く。"""
    text = text.strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in "\"'":
        text = text[1:-1]
    return text.replace("\\ ", " ") if sys.platform != "win32" else text


def input_text(prompt: str, default: str = "", empty: str = "戻る") -> str | None:
    """empty は、既定値が無いときに空 Enter が何を意味するかの案内。"""
    suffix = f" [Enter={default}]" if default else f"（空 Enter で{empty}）"
    answer = input(f"  {prompt}{suffix}: ").strip()
    if not answer:
        return default or None
    return answer


def load_history() -> dict:
    try:
        return json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_history(data: dict) -> None:
    try:
        HISTORY_PATH.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass   # 記憶できなくても動作には影響しない


def ask_excel(history: dict) -> str | None:
    while True:
        answer = input_text("Excel ファイル（ドラッグ＆ドロップ可）", history.get("excel", ""))
        if answer is None:
            return None
        # 本体はツールのフォルダで動くため、相対パスは今の場所を基準に絶対パスへ直す
        path = str(Path(clean_path(answer)).resolve())
        if Path(path).is_file():
            history["excel"] = path
            save_history(history)
            return path
        print(f"  ※ ファイルが見つかりません: {path}")


# 取り込むシート。もう一方のシートは読まないので、書いたまま残った行を送らない
SHEETS = [("登録", "課題を新しく作る"), ("更新", "既存の課題を変更する")]


def ask_sheet() -> str | None:
    choice = print_menu("取り込むシート", [f"{name} ― {desc}" for name, desc in SHEETS])
    return None if choice == 0 else SHEETS[choice - 1][0]


def build_args(action: str, config: str | None, history: dict) -> list[str] | None:
    common = ["--config", config] if config else []
    if action == "template":
        return ["template", *common]
    if action == "master":
        return ["master", *common]
    if action in ("dry", "execute"):
        excel = ask_excel(history)
        if excel is None:
            return None
        sheet = ask_sheet()
        if sheet is None:
            return None
        args = ["import", excel, "--sheet", sheet, *common]
        if action == "dry":
            # 「--execute を付けて」の案内を、メニューの操作に置き換えてもらう
            args.append("--from-menu")
        if action == "execute":
            args.append("--execute")
            limit = input_text("先頭から何件だけ送るか", empty="全件")
            if limit:
                if not limit.isdigit() or int(limit) < 1:
                    print("  ※ 1 以上の数を入力してください。")
                    return None
                args += ["--limit", limit]
        return args
    if action == "export":
        choice = print_menu("書き出す範囲", [
            "完了以外のすべて",
            "完了も含めてすべて",
            "ある課題とその子課題",
            "キーワードで絞る",
        ])
        if choice == 0:
            return None
        args = ["export", *common]
        if choice == 2:
            args.append("--include-closed")
        elif choice == 3:
            key = input_text("親の課題キー（例: PROJ-123）")
            if not key:
                return None
            args += ["--parent", key]
        elif choice == 4:
            keyword = input_text("キーワード")
            if not keyword:
                return None
            args += ["--keyword", keyword]
        return args
    raise AssertionError(action)


def run(args: list[str]) -> int:
    print()
    hr("-")
    try:
        completed = subprocess.run([sys.executable, str(SCRIPT), *args], cwd=TOOL_DIR)
    except KeyboardInterrupt:
        print("\n  中断しました。")
        return EXIT_INTERRUPTED
    return completed.returncode


def main() -> None:
    parser = argparse.ArgumentParser(description="backlog-issue-sheet 対話メニュー")
    parser.add_argument("--config", help="設定ファイル（既定: ツールの場所の config.yaml）")
    config = parser.parse_args().config
    if config:
        config = str(Path(config).resolve())
    history = load_history()
    items = [f"{label} ― {desc}" for _, label, desc in ACTIONS]

    try:
        while True:
            choice = print_menu("Backlog 課題シート", items, back_label="終了")
            if choice == 0:
                print("  終了します。")
                sys.exit(EXIT_OK)
            args = build_args(ACTIONS[choice - 1][0], config, history)
            if args is None:
                continue
            rc = run(args)
            print()
            hr("-")
            print(f"  終了コード: {rc}")
            input("\n  Enter キーでメニューに戻ります...")
    except EOFError:
        print("\n  入力が終了しました。")
        sys.exit(EXIT_EOF)
    except KeyboardInterrupt:
        print("\n  中断しました。")
        sys.exit(EXIT_INTERRUPTED)


if __name__ == "__main__":
    main()
