"""
Backlog のマスタ（名前 ⇄ ID）
=============================
Excel には名前で書き、API には ID で送る。その変換をここに集める。

名前が一致しないときは、似た名前を候補として添えて ResolveError を送出する。
打ち間違いを、送信してから Backlog のエラーで知るのでは遅いため。
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass, field

# カスタム属性の型（Backlog の typeId）
CF_TEXT = 1          # 文字列
CF_SENTENCE = 2      # 文章
CF_NUMBER = 3        # 数値
CF_DATE = 4          # 日付
CF_SINGLE_LIST = 5   # 単一リスト
CF_MULTI_LIST = 6    # 複数リスト
CF_CHECKBOX = 7      # チェックボックス
CF_RADIO = 8         # ラジオ

CF_TYPE_NAMES = {
    CF_TEXT: "文字列",
    CF_SENTENCE: "文章",
    CF_NUMBER: "数値",
    CF_DATE: "日付",
    CF_SINGLE_LIST: "単一リスト",
    CF_MULTI_LIST: "複数リスト",
    CF_CHECKBOX: "チェックボックス",
    CF_RADIO: "ラジオ",
}

# 完了の状態。Backlog では全プロジェクト共通で ID 4 が「完了」。
STATUS_CLOSED_ID = 4
# 作成直後の状態。作成時は状態を指定できず、必ずこれになる。
STATUS_OPEN_ID = 1


class ResolveError(ValueError):
    """名前を ID に変換できなかった。"""


def suggest(name: str, candidates: list[str]) -> str:
    """似た名前があれば「もしかして」の一文を返す。無ければ空文字。"""
    close = difflib.get_close_matches(name, candidates, n=3, cutoff=0.5)
    # 前方一致・部分一致も拾う（「山田」→「山田太郎」は difflib では遠い）
    for c in candidates:
        if c not in close and name and (name in c or c in name):
            close.append(c)
    if not close:
        return ""
    return "もしかして: " + " / ".join(f"「{c}」" for c in close[:3])


class NameTable:
    """
    名前 → ID の対応表。

    同じ名前が 2 つ以上ある場合（同姓同名のユーザーなど）は、名前では
    特定できないとしてエラーにする。ユーザーはログイン ID でも引ける。
    """

    def __init__(self, label: str, items: list[dict], *, alt_key: str | None = None):
        self.label = label
        self.items = list(items)
        self._by_name: dict[str, list[int]] = {}
        self._by_alt: dict[str, int] = {}
        self._name_of: dict[int, str] = {}
        for item in self.items:
            name = str(item.get("name", "")).strip()
            self._by_name.setdefault(name, []).append(item["id"])
            self._name_of[item["id"]] = name
            if alt_key and item.get(alt_key):
                self._by_alt[str(item[alt_key])] = item["id"]
        self._alt_of = {v: k for k, v in self._by_alt.items()}

    def names(self) -> list[str]:
        """表示順の名前一覧（重複は 1 回だけ）。"""
        seen = []
        for item in self.items:
            name = str(item.get("name", "")).strip()
            if name not in seen:
                seen.append(name)
        return seen

    def has(self, name: str) -> bool:
        return name in self._by_name or name in self._by_alt

    def resolve(self, name: str) -> int:
        name = name.strip()
        ids = self._by_name.get(name)
        if ids and len(ids) == 1:
            return ids[0]
        if name in self._by_alt:
            return self._by_alt[name]
        if ids:
            raise ResolveError(
                f"{self.label}「{name}」が {len(ids)} 件あり、名前では特定できません"
                + ("。ログイン ID で書いてください" if self._by_alt else "")
            )
        message = f"{self.label}「{name}」が見つかりません"
        hint = suggest(name, self.names() + list(self._by_alt))
        if hint:
            message += f"。{hint}"
        raise ResolveError(message)

    def name_of(self, item_id: int | None) -> str:
        """ID を Excel に書く名前にする。同名があればログイン ID を返す。"""
        if item_id is None:
            return ""
        name = self._name_of.get(item_id)
        if name is None:
            return str(item_id)
        if len(self._by_name.get(name, [])) > 1 and item_id in self._alt_of:
            return self._alt_of[item_id]
        return name


@dataclass
class CustomField:
    id: int
    name: str
    type_id: int
    required: bool = False
    applicable_issue_types: list[int] = field(default_factory=list)
    items: NameTable | None = None

    @property
    def type_name(self) -> str:
        return CF_TYPE_NAMES.get(self.type_id, f"typeId={self.type_id}")

    @property
    def is_list(self) -> bool:
        return self.type_id in (CF_SINGLE_LIST, CF_MULTI_LIST, CF_CHECKBOX, CF_RADIO)

    @property
    def is_multi(self) -> bool:
        return self.type_id in (CF_MULTI_LIST, CF_CHECKBOX)

    def applies_to(self, issue_type_id: int | None) -> bool:
        """その種別で使える属性か。空の一覧は「全種別」を意味する。"""
        if not self.applicable_issue_types or issue_type_id is None:
            return True
        return issue_type_id in self.applicable_issue_types

    @classmethod
    def from_api(cls, cf: dict) -> "CustomField":
        type_id = cf.get("typeId")
        items = None
        if type_id in (CF_SINGLE_LIST, CF_MULTI_LIST, CF_CHECKBOX, CF_RADIO):
            items = NameTable(f"「{cf.get('name')}」の選択肢", cf.get("items") or [])
        return cls(
            id=cf["id"],
            name=str(cf.get("name", "")).strip(),
            type_id=type_id,
            required=bool(cf.get("required")),
            applicable_issue_types=list(cf.get("applicableIssueTypes") or []),
            items=items,
        )


@dataclass
class Master:
    project_id: int
    project_key: str
    subtasking_enabled: bool
    tables: dict[str, NameTable]
    custom_fields: list[CustomField]

    def table(self, key: str) -> NameTable:
        return self.tables[key]

    def custom_field_by_name(self, name: str) -> CustomField | None:
        for cf in self.custom_fields:
            if cf.name == name:
                return cf
        return None

    def custom_field_by_id(self, cf_id: int) -> CustomField | None:
        for cf in self.custom_fields:
            if cf.id == cf_id:
                return cf
        return None

    @classmethod
    def load(cls, client, project_key: str) -> "Master":
        """Backlog からマスタをまとめて取得する。"""
        project = client.get_project(project_key)
        key = project.get("projectKey", project_key)
        return cls.build(
            project=project,
            issue_types=client.get_issue_types(key),
            priorities=client.get_priorities(),
            statuses=client.get_statuses(key),
            resolutions=client.get_resolutions(),
            users=client.get_project_users(key),
            categories=client.get_categories(key),
            versions=client.get_versions(key),
            custom_fields=client.get_custom_fields(key),
        )

    @classmethod
    def build(
        cls,
        *,
        project: dict,
        issue_types: list,
        priorities: list,
        statuses: list,
        resolutions: list,
        users: list,
        categories: list,
        versions: list,
        custom_fields: list,
    ) -> "Master":
        # アーカイブ済みのバージョンも残す。除くと、それが付いた既存課題を
        # 書き出したときに名前にできず、取り込みに戻すとエラーになる。
        return cls(
            project_id=project["id"],
            project_key=project.get("projectKey", ""),
            # 項目が無い（古い版・テスト）ときは、使えるものとして扱う
            subtasking_enabled=bool(project.get("subtaskingEnabled", True)),
            tables={
                "issue_types": NameTable("種別", issue_types),
                "priorities": NameTable("優先度", priorities),
                "statuses": NameTable("状態", statuses),
                "resolutions": NameTable("完了理由", resolutions),
                "users": NameTable("ユーザー", users, alt_key="userId"),
                "categories": NameTable("カテゴリー", categories),
                "versions": NameTable("バージョン", versions),
            },
            custom_fields=[CustomField.from_api(cf) for cf in custom_fields],
        )
