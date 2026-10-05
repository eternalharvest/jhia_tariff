# jhia_tariff

協会けんぽ「[都道府県毎の保険料額表](https://www.kyoukaikenpo.or.jp/about/business/insurance_rate/premium_prefectures/)」の
Excel を取得して JSON に整形し、GitHub Pages で公開します。Excel の Power Query から読み込めます。

## 仕組み

1. `scripts/build.py` が一覧ページから年度ページ(`r08` など)を探す。デフォルトは最新年度のみ、`--year all` なら変換可能な全年度
2. 年度ページにある `.xlsx` をダウンロードし、シート内容で「一般(健保・厚年)」「任意継続」を判別して 47 都道府県分を 1 つのフラットな JSON に変換
3. 件数・等級の連続性などを検証(失敗すると Actions が落ち、**前回のデータが公開されたまま**になる)。`--year all` の場合、Excel のレイアウトが異なる古い年度はスキップして警告を出し、残りの年度の処理は継続する
4. 最新年度の出力を `latest/` にもコピー
5. GitHub Actions が毎日 06:00 JST に `--year all` で実行し、GitHub Pages にデプロイ

## 公開 URL

`https://<user>.github.io/<repo>/` 配下、年度ディレクトリ単位(例 `r08/`)と、最新年度のコピーである `latest/` に同じ構成のファイルが置かれます。

| ファイル | 内容 | 行数 |
|---|---|---|
| `index.json` (ルート) | 変換できた年度一覧・最新年度 | - |
| `{year}/index.json`, `latest/index.json` | その年度の生成時刻・取得元 Excel URL | - |
| `{year}/general.json`, `latest/general.json` | 一般: 健康保険/介護/子ども子育て支援金/厚生年金 | 47 × 50 等級 |
| `{year}/optional.json`, `latest/optional.json` | 任意継続被保険者 | 47 × 23 等級 |
| `{year}/general.ja.json` / `optional.ja.json`(各 `latest/` も同様) | 上記と同じ内容で**キーを日本語にしたもの** | 同上 |

英語キー版は他ツールからの利用向け、`*.ja.json` は Excel の Power Query でそのまま日本語列名にしたい場合向けです。年度をまたいだ比較をしたい場合は `{year}/` を、常に最新だけでよい場合は `latest/` を参照してください。

## Power Query (Excel)

データ → データの取得 → その他のデータソースから → 空のクエリ → 詳細エディター:

```m
let
    Source = Json.Document(Web.Contents("https://<user>.github.io/<repo>/latest/general.json")),
    Table  = Table.FromRecords(Source),
    Typed  = Table.TransformColumnTypes(Table, {
        {"prefecture_code", Int64.Type}, {"grade", Int64.Type},
        {"remuneration_std", Int64.Type},
        {"health_full", type number}, {"health_half", type number}})
in
    Typed
```

`latest/general.ja.json` を指定すれば列名は最初から日本語になります(型変換の列名も日本語に読み替えてください)。`optional.json` も同様です。特定年度を固定して参照したい場合は `latest/` の部分を `r08/` などに変えてください。Web.Contents の URL は文字列リテラルで書くと、
Power BI/Excel のスケジュール更新でも動的ソースエラーになりません。

## ローカル実行

[uv](https://docs.astral.sh/uv/) を使います。依存関係は `pyproject.toml`、固定バージョンは `uv.lock` で管理しています(追加は `uv add <pkg>`)。

```sh
uv sync --no-install-project                  # 依存関係をインストール (.venv 作成)
uv run scripts/build.py --out site            # 最新年度のみ (site/{最新年度}/, site/latest/)
uv run scripts/build.py --year all --out site # 変換可能な全年度
uv run scripts/build.py --year r07 --out site # 特定の過去年度を明示指定(レイアウトが異なると失敗する)
```

## 注意

- 現在 `r03`〜`r08` に対応しています。`r01`/`r02` は一覧ページに今の形式(都道府県別47シートの1つの `.xlsx`)の表が見つからないため未対応です(`--year all` では自動でスキップされます)。
- 過去年度は支部ラベルの表記(「北海道」/「北海道支部」、「青森」/「青森県」)や列構成(「子ども・子育て支援金」列の有無、r08〜)が異なり、`build.py` はこれらを吸収していますが、さらに古い形式・未知のレイアウト変更には対応していません。`--year all` では未対応のレイアウトは自動でスキップしますが、`--year <年度>` で明示指定した場合は検証失敗で停止します。
- 数値は協会けんぽの Excel 値そのまま(浮動小数の誤差のみ小数 2 桁に丸め)。正式な金額は協会けんぽの公表資料を参照してください。
- `general.json` の `pension_*` は厚生年金の対象等級のみ値があり、それ以外は `null`。
