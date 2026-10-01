"""セルの値の変換"""

import datetime as dt

import pytest

from fields import (
    CLEAR,
    STANDARD_BY_HEADER,
    ValueProblem,
    parse_date,
    parse_number,
    parse_value,
    split_names,
    to_cell,
    to_param,
)
from master import Master, NameTable

COL = STANDARD_BY_HEADER


class TestSplitNames:
    def test_改行があれば改行だけで分ける(self):
        # 書き出しは改行区切り。名前の「,」で分けてしまうと取り込みに戻せない
        assert split_names("設計, 実装\n開発") == ["設計, 実装", "開発"]

    def test_改行が無ければカンマと読点でも分ける(self):
        assert split_names("設計,開発、試験") == ["設計", "開発", "試験"]

    def test_空の要素は捨てる(self):
        assert split_names("a,,、") == ["a"]
        assert split_names("a\n\n") == ["a"]


class TestNamesWithComma:
    @pytest.fixture
    def master_with_comma(self, fake):
        fake.get_categories = lambda key: [{"id": 1, "name": "設計, 実装"}, {"id": 2, "name": "開発"}]
        return Master.load(fake, "DEMO")

    def test_カンマを含む名前1つ(self, master_with_comma):
        assert parse_value(COL["カテゴリー"], "設計, 実装", master_with_comma) == (1,)

    def test_カンマを含む名前を含む複数(self, master_with_comma):
        col = COL["カテゴリー"]
        cell = to_cell(col, (1, 2), master_with_comma)
        assert parse_value(col, cell, master_with_comma) == (1, 2)


class TestParse:
    def test_空は_None_削除は_CLEAR(self, master):
        assert parse_value(COL["件名"], "  ", master) is None
        assert parse_value(COL["担当者"], " (削除) ", master) is CLEAR

    def test_文字列は改行を揃えて前後の空白を除く(self, master):
        assert parse_value(COL["詳細"], " a\r\nb ", master) == "a\nb"

    def test_整数の_float_は小数点を付けない(self, master):
        assert parse_value(COL["件名"], 12.0, master) == "12"

    @pytest.mark.parametrize("raw", [dt.datetime(2026, 1, 5), dt.date(2026, 1, 5), "2026/1/5", "2026-01-05", "2026.1.5"])
    def test_日付(self, raw):
        assert parse_date(raw) == "2026-01-05"

    def test_存在しない日付(self):
        with pytest.raises(ValueProblem, match="存在しない"):
            parse_date("2026-02-30")

    def test_書式の無い日付セル(self):
        with pytest.raises(ValueProblem, match="セルの書式"):
            parse_date(46000)

    def test_時間は負にできない(self):
        with pytest.raises(ValueProblem, match="負"):
            parse_number(-1, non_negative=True)

    def test_数値は文字列でもよい(self):
        assert parse_number("1,234.5", non_negative=False) == 1234.5


class TestToParam:
    def test_削除は空文字_複数値は空の要素1つ(self):
        assert to_param(COL["担当者"], CLEAR) == ""
        assert to_param(COL["カテゴリー"], CLEAR) == [""]

    def test_時間は余分な_0_を付けない(self):
        assert to_param(COL["予定時間"], 2.0) == "2"
        assert to_param(COL["予定時間"], 1.25) == "1.25"


class TestNameTableLabels:
    def test_ログイン_ID_が無いユーザーは名前だけ(self):
        table = NameTable("ユーザー", [
            {"id": 1, "name": "山田太郎", "userId": "yamada"},
            {"id": 2, "name": "外部の人", "userId": None},
        ], alt_key="userId")
        assert table.labels() == ["山田太郎（yamada）", "外部の人"]

    def test_ログイン_ID_を持たない表は名前のまま(self):
        assert NameTable("種別", [{"id": 1, "name": "タスク"}]).labels() == ["タスク"]
