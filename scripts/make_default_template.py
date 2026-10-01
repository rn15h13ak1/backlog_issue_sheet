#!/usr/bin/env python3
"""
同梱のひな形（templates/課題シート.xlsx）を作る
==============================================
Backlog に接続せずに使い始められるよう、標準の項目だけの列を持つひな形を作る。
プルダウンには、Backlog が新しいプロジェクトに用意する既定の値を入れる。
プロジェクトに合わせたひな形（カスタム属性・担当者入り）は
`backlog-issue-sheet template` で作る。

生成物を追跡するため、何度作っても同じバイト列になるようにする
（共通規約 A「生成物の扱い」）。openpyxl は保存のたびに作成・更新日時と
zip の各項目の時刻を書き込むので、保存後にそれらを固定値へ置き換える。

  python scripts/make_default_template.py           # 作り直す
  python scripts/make_default_template.py --check   # 最新か確かめる（差があれば終了コード 1）
"""

from __future__ import annotations

import argparse
import datetime as dt
import io
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from fields import REGISTER_SHEET, UPDATE_SHEET  # noqa: E402
from master import Master  # noqa: E402
from sheet_io import EXAMPLE_NOTE_HEADER, build_workbook  # noqa: E402

OUTPUT = ROOT / "templates" / "課題シート.xlsx"

# 生成物に埋め込む固定の時刻（zip の時刻は 1980 年より前を表せない）
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
FIXED_W3CDTF = "2026-01-01T00:00:00Z"


def default_master() -> Master:
    """Backlog の既定値だけで作ったマスタ。ID はプルダウンには使わないので仮の値。"""
    named = lambda names, start=1: [{"id": i, "name": n} for i, n in enumerate(names, start)]  # noqa: E731
    return Master.build(
        project={"id": 0, "projectKey": ""},
        issue_types=named(["タスク", "バグ", "要望", "その他"]),
        priorities=named(["高", "中", "低"], 2),
        statuses=named(["未対応", "処理中", "処理済み", "完了"]),
        resolutions=named(["対応済み", "対応しない", "無効", "重複", "再現しない"], 0),
        users=[], categories=[], versions=[], custom_fields=[],
    )


D = dt.date

REGISTER_EXAMPLE = [
    {"件名": "画面設計", "種別": "タスク", "優先度": "中", "担当者": "山田太郎", "期限日": D(2026, 10, 31),
     EXAMPLE_NOTE_HEADER: "親のない課題。次の ○ の行の親になる"},
    {"子課題": "○", "件名": "ログイン画面の設計", "種別": "タスク", "優先度": "中", "予定時間": 4,
     EXAMPLE_NOTE_HEADER: "上の「画面設計」の子課題になる"},
    {"子課題": "○", "件名": "一覧画面の設計", "種別": "タスク", "優先度": "低", "状態": "処理中",
     EXAMPLE_NOTE_HEADER: "これも「画面設計」の子課題。状態は作成の直後に更新で反映する"},
    {"親課題キー": "PROJ-50", "件名": "既存課題の下に付ける子", "種別": "タスク", "優先度": "中",
     EXAMPLE_NOTE_HEADER: "既にある PROJ-50 の子課題として作る"},
    {"件名": "ログインできない", "詳細": "再現手順:\n1. …\n2. …", "種別": "バグ", "優先度": "高",
     "カテゴリー": "設計\n開発", "開始日": D(2026, 10, 1), "期限日": D(2026, 10, 3),
     EXAMPLE_NOTE_HEADER: "複数の値は改行か「,」で区切る。カテゴリー名はプロジェクトのものを書く"},
]

UPDATE_EXAMPLE = [
    {"課題キー": "PROJ-101", "親課題キー": "PROJ-50",
     EXAMPLE_NOTE_HEADER: "PROJ-50 の子課題にする（ほかの項目は空なので変えない）"},
    {"課題キー": "PROJ-102", "親課題キー": "(解除)",
     EXAMPLE_NOTE_HEADER: "親から外し、子課題ではない課題にする"},
    {"課題キー": "PROJ-103", "状態": "処理中", "担当者": "(削除)", "期限日": D(2026, 10, 31),
     EXAMPLE_NOTE_HEADER: "状態と期限日を変え、担当者を外す"},
]


def render() -> bytes:
    wb = build_workbook(
        default_master(),
        include_guide=True,
        examples={
            f"記入例（{REGISTER_SHEET}）": (REGISTER_SHEET, REGISTER_EXAMPLE),
            f"記入例（{UPDATE_SHEET}）": (UPDATE_SHEET, UPDATE_EXAMPLE),
        },
    )
    buf = io.BytesIO()
    wb.save(buf)
    return normalize(buf.getvalue())


def normalize(data: bytes) -> bytes:
    """zip の各項目の時刻と、文書の作成・更新日時を固定値にする。"""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == "docProps/core.xml":
                text = content.decode("utf-8")
                for tag in ("dcterms:created", "dcterms:modified"):
                    pattern = rf"(<{tag}[^>]*>)[^<]*(</{tag}>)"
                    if not re.search(pattern, text):
                        raise RuntimeError(f"置換対象が見つかりません: {tag}")
                    text = re.sub(pattern, rf"\g<1>{FIXED_W3CDTF}\g<2>", text)
                content = text.encode("utf-8")
            info = zipfile.ZipInfo(item.filename, date_time=FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            dst.writestr(info, content)
    return out.getvalue()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="同梱のひな形を作る")
    parser.add_argument("--check", action="store_true", help="作り直さずに、最新かどうかだけ確かめる")
    args = parser.parse_args(argv)

    data = render()
    rel = OUTPUT.relative_to(ROOT)
    if args.check:
        if OUTPUT.is_file() and OUTPUT.read_bytes() == data:
            print(f"{rel}: 最新です")
            return 0
        print(f"{rel}: 古いか、ありません。python scripts/make_default_template.py で作り直してください")
        return 1
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_bytes(data)
    print(f"{rel} を作りました（{len(data):,} バイト）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
