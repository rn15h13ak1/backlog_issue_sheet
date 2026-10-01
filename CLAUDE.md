# backlog_issue_sheet 固有の取り決め

共通規約は [`../ws-conventions/README.md`](../ws-conventions/README.md) に従う。
本ファイルには、本リポジトリだけの事情を書く。

commit / push は、修正ごとに行う。検査とテストを通してから commit し、そのまま
push する。共通規約は「利用者が指示したときだけ」とするが、本リポジトリでは
利用者が修正ごとに行うよう常設で指示したため。検査やテストが落ちたら commit しない。
GitHub のリポジトリは公開している。ライセンスは MIT（`LICENSE`）。

## 作業のたびに実行する

Markdown を編集したら、commit 前に検査する。

```bash
../ws-conventions/bin/check-markdown.sh .
../ws-conventions/bin/check-privacy.sh .
```

手順そのものも検査する。

```bash
../ws-conventions/bin/check-commands.sh .
```

ADR は使っていないため、語の検査と索引の生成は対象外。`docs/term-rules.md` を
置いて決定を記録するようになったら、次も実行する。

```bash
../ws-conventions/bin/check-terms.sh .
../ws-conventions/bin/gen-decision-index.py .
```

テストは、依存（openpyxl / PyYAML / pytest）を入れた仮想環境の Python で実行する。

```bash
.venv/bin/python -m pytest -q -p no:warnings
```

`sheet_io.py` の書式や記入例を変えたら、同梱のひな形を作り直す。作り直し忘れは
`tests/test_default_template.py` が検出する。

```bash
.venv/bin/python scripts/make_default_template.py
```

## 同梱のひな形は生成物として追跡する

`templates/課題シート.xlsx` は `.gitignore` の `*.xlsx` から外して追跡する。
openpyxl は保存のたびに日時を書き込むため、生成スクリプトでそれらを固定値に
置き換え、何度作っても同じバイト列になるようにしている（共通規約 A「生成物の扱い」）。

## 元の Excel は変更しない

取り込んだ Excel に課題キーを書き戻さない。作成した課題は `output/結果_*.xlsx` に
「更新」シートの書式で出す。openpyxl で開き直して保存すると、数式やグラフなどが
失われるため（`../excel_to_backlog` と同じ方針）。

## backlog_client.py はコピー

`../excel_to_backlog/backlog_client.py` をコピーし、取得系を足したもの。
import で依存させないのは、他のリポジトリを変更しない（共通規約 B）うえ、
向こうの変更でこちらが壊れるため。元のファイルで不具合が直ったら、こちらにも反映する。
`tests/test_backlog_client.py` / `test_http_layer.py` / `test_retry.py` も同じくコピー。

## テストの偽物

`tests/conftest.py` の `FakeBacklog` は、Backlog の親子の制約（1 階層まで）をまねる。
送信の順番を検証するテストはこれに頼っているので、制約を外さないこと。
`test_executor.py` の「順番を変えなければ拒否される」テストが、偽物が制約を守っていることを確かめている。
