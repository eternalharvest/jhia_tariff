# jhia_tariff

協会けんぽ「[都道府県毎の保険料額表](https://www.kyoukaikenpo.or.jp/about/business/insurance_rate/premium_prefectures/)」の
Excel を取得して JSON に整形し、GitHub Pages で公開します。Excel の Power Query から読み込めます。

> [!IMPORTANT]
> このリポジトリは個人が作成した非公式のものであり、全国健康保険協会(協会けんぽ)・厚生労働省その他いかなる公的機関とも関係ありません。公式な情報・正式な金額については、必ず[協会けんぽの公式サイト](https://www.kyoukaikenpo.or.jp/)の公表資料をご確認ください。
>
> 本リポジトリで公開しているデータ・スクリプトは現状有姿で提供されるものであり、その正確性・完全性・最新性を保証しません。本リポジトリのデータ・スクリプトの利用(給与計算等の実務利用を含む)により生じたいかなる損害についても、作者は一切の責任を負いません。利用は自己責任でお願いします。

## 仕組み

1. `scripts/build.py` が一覧ページから年度ページ(`r08` など)を探す。デフォルトは最新年度のみ、`--year all` なら変換可能な全年度
2. 年度ページ(およびその下に改定ごとのサブページがあればそれも)にある `.xlsx` をダウンロードし、シート内容で「一般(健保・厚年)」「任意継続」を判別して 47 都道府県分を 1 つのフラットな JSON に変換。年度内に保険料率の改定が複数回あった年度(例: r01, r02)は、改定ごとに分けて全て変換する
3. 件数・等級の連続性などを検証(失敗すると Actions が落ち、**前回のデータが公開されたまま**になる)。`--year all` の場合、レイアウトが異なる・必要な情報が揃わない年度・改定はスキップして警告を出し、残りの処理は継続する
4. 年度ごとに最新の改定を `{year}/latest/` に、さらに全体で最新の年度・改定を `latest/` にもコピー
5. GitHub Actions が毎日 06:00 JST に `--year all` で実行し、GitHub Pages にデプロイ

## 公開 URL

`https://<user>.github.io/<repo>/` 配下、年度・改定ごとのディレクトリ(例 `r08/001/`)と、最新のコピーである `latest/`・`{year}/latest/` に同じ構成のファイルが置かれます。

| ファイル | 内容 | 行数 |
|---|---|---|
| `index.json` (ルート) | 変換できた年度一覧・最新年度・全年度全改定のファイル一覧(`revisions`、下記参照) | - |
| `{year}/index.json` | その年度の改定一覧(`revisions`)・最新の改定(`latest`) | - |
| `{year}/{revision}/index.json`, `{year}/latest/index.json`, `latest/index.json` | その改定の生成時刻・実施時期(`effective_from`)。`files.general`/`files.optional` にはそれぞれの取得元Excel URLと、**その種類自身が実際に更新された実施時期**(`files.<kind>.effective_from`)も入る | - |
| `{year}/{revision}/general.json`, `{year}/latest/general.json`, `latest/general.json` | 一般: 健康保険/介護/子ども子育て支援金/厚生年金 | 47 × 50 等級 |
| `{year}/{revision}/optional.json`, 同上の `latest/` | 任意継続被保険者 | 47 × 23 等級 |
| `{year}/{revision}/general.ja.json` / `optional.ja.json`(各 `latest/` も同様) | 上記と同じ内容で**キーを日本語にしたもの** | 同上 |

英語キー版は他ツールからの利用向け、`*.ja.json` は Excel の Power Query でそのまま日本語列名にしたい場合向けです。年度・改定をまたいだ比較をしたい場合は `{year}/{revision}/` を、常に最新だけでよい場合は `latest/` を参照してください。

改定(`{revision}`、`001`から始まる連番)は、ページのサブページ番号ではなく、表に書かれた実施時期(「○年○月分」)の古い順に振り直しています。協会けんぽのサブページ番号は年度によって新しい順/古い順が不統一で、そのままでは信頼できないためです。また、改定によっては「一般」だけ新しい Excel が公開され「任意継続」は変わらず前の改定のものが使われる(そのページには存在しない)ことがあるため、実施時期の古い順に「直前までに見つかったファイル」を引き継いで出力します。その年度でまだ一度も「任意継続」が公開されていない期間は `general[.ja].json` だけが生成され、`optional[.ja].json` は存在しません。

「一般」は改定のたびに(内容が変わっていなくても)再公開されることが多いのに対し、「任意継続」はそうではないため、ある改定の `index.json` が示す `effective_from`(その改定全体の実施時期)と、`files.optional.effective_from`(任意継続のデータ自体が最後に更新された実施時期)は一致しないことがあります。一致していなければ、その改定の任意継続データは今回新しく更新されたものではなく、より古い改定から引き継がれたものだと分かります。`files.general.effective_from` は通常 `effective_from` と一致します。

ルートの `index.json` の `revisions` には、全年度・全改定について `{year, revision, path, effective_from, files}` が実施時期の古い順に列挙されています。`files` にはその改定で実際に生成されたファイル名(`general.json` のみ、または `general.json`+`optional.json` の両方)が入るので、例えば「`optional.json` を含む最新の改定」を知りたい場合は、この一覧を `files` でフィルタしたあと `effective_from` が最大のものを選び、その `path` 配下のファイルを参照してください。

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

`latest/general.ja.json` を指定すれば列名は最初から日本語になります(型変換の列名も日本語に読み替えてください)。`optional.json` も同様です。特定年度・改定を固定して参照したい場合は `latest/` の部分を `r08/001/` などに変えてください。Web.Contents の URL は文字列リテラルで書くと、
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

- 現在 `r01`〜`r08` に対応しています。
- 過去年度は支部ラベルの表記(「北海道」/「北海道支部」、「青森」/「青森県」)や列構成(「子ども・子育て支援金」列の有無、r08〜)が異なり、`build.py` はこれらを吸収していますが、さらに古い形式・未知のレイアウト変更には対応していません。`--year all` では未対応のレイアウトは自動でスキップしますが、`--year <年度>` で明示指定した場合は検証失敗で停止します。
- 2026年10月分から、「保険料調整制度」という新しい仕組みの改定に対応しています。従来「健康保険料・介護保険料」として合算されていた列が「健康保険料」(`health`)と「介護保険料」(単独、`health_ltc` の算出に合算)に分かれ、さらに健康保険料・厚生年金保険料それぞれの被保険者負担分が「被保険者負担」と「保険料調整額」(`health_adjustment`/`pension_adjustment`、対象外の等級・年度では `null`)に分かれています。`health_half`/`pension_half` は、常に「被保険者が実際に支払う金額」を表すように定義しています(調整額が無い場合は従来通り全額の半分)。
- 数値は協会けんぽの Excel 値そのまま(浮動小数の誤差のみ小数 2 桁に丸め)。正式な金額は協会けんぽの公表資料を参照してください。
- `general.json` の `pension_*`(`pension_full`/`pension_half`/`pension_adjustment`/`pension_grade`)は厚生年金の対象等級のみ値があり、それ以外は `null`。

## ライセンス

[MIT License](LICENSE) は、本リポジトリのソースコード(`scripts/build.py` 等)に適用されます。

生成される数値データ(`general.json` など)は、協会けんぽが公表している数値をそのまま機械的に変換したものであり、独自の著作物として著作権を主張できるものではないと考えています。そのためデータ自体にライセンスを設定する意図はありません。データの正確性等を保証しない点は冒頭の免責事項の通りです。
