#!/usr/bin/env python3
"""協会けんぽ 都道府県別保険料額表 (Excel) を取得し、JSON に整形する。

使い方:
    python scripts/build.py [--year r08|all] [--out site] [--cache .cache]
    --year 省略時は最新年度のみ、'all' で変換可能な全年度を処理する。

年度内に保険料率の改定が複数回あった場合(例: r01, r02)は、改定ごとに
サブディレクトリを分けて出力する。改定のサブページ番号は年度によって新しい順/
古い順が不統一なため、「最新」は各改定の実施時期(表に書かれた「○年○月分」)で
判定している。

出力 (--out):
    {year}/{revision}/general.json   一般の被保険者 (健康保険・介護・子ども子育て支援金・厚生年金)
    {year}/{revision}/optional.json  任意継続被保険者
    {year}/{revision}/index.json     メタ情報 (取得元 URL, 生成時刻, 実施時期, 件数)
    {year}/latest/                   その年度で最新の改定 ({year}/{revision}/ と同一) のコピー
    {year}/index.json                その年度の改定一覧・最新の改定
    latest/                          全体で最新の年度・改定のコピー
    index.json                       変換できた年度一覧・最新年度
"""
from __future__ import annotations

import argparse
import io
import json
import re
import shutil
import sys
import urllib.request
import warnings
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urljoin

import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE = "https://www.kyoukaikenpo.or.jp"
INDEX_URL = f"{BASE}/about/business/insurance_rate/premium_prefectures/"
UA = "Mozilla/5.0 (compatible; jhia-tariff-bot/1.0; +https://github.com)"

PREFECTURES = [
    "北海道", "青森", "岩手", "宮城", "秋田", "山形", "福島", "茨城", "栃木", "群馬",
    "埼玉", "千葉", "東京", "神奈川", "新潟", "富山", "石川", "福井", "山梨", "長野",
    "岐阜", "静岡", "愛知", "三重", "滋賀", "京都", "大阪", "兵庫", "奈良", "和歌山",
    "鳥取", "島根", "岡山", "広島", "山口", "徳島", "香川", "愛媛", "高知", "福岡",
    "佐賀", "長崎", "熊本", "大分", "宮崎", "鹿児島", "沖縄",
]
GRADE_RE = re.compile(r"^(\d+)(?:[(（](\d+)[)）])?$")
JST = timezone(timedelta(hours=9))


def fetch(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def num(v):
    # 注記テキスト(結合セル)が数値列に混入することがあるので、数値以外は欠損扱い
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    v = round(float(v), 4)
    return int(v) if v == int(v) else v


def money(v):
    return num(round(v, 2)) if isinstance(v, (int, float)) else None


# ---------------------------------------------------------------- discovery
def list_years() -> list[str]:
    html = fetch(INDEX_URL).decode("utf-8", "replace")
    years = sorted(set(re.findall(r"premium_prefectures/(r\d+)", html)))
    if not years:
        raise SystemExit("年度ページのリンクが見つかりません。ページ構造が変わった可能性があります。")
    return years


def find_xlsx_links(page_url: str) -> list[str]:
    html = fetch(page_url).decode("utf-8", "replace")
    links = re.findall(r'href="([^"]+\.xlsx)"', html, flags=re.I)
    return sorted({urljoin(page_url, l) for l in links})


def classify(wb: openpyxl.Workbook) -> str | None:
    ws = wb[wb.sheetnames[0]]
    head = " ".join(str(c) for row in ws.iter_rows(min_row=1, max_row=3, values_only=True)
                    for c in row if c)
    if "任意継続" in head:
        return "optional"
    if "健康保険・厚生年金保険の保険料額表" in head:
        return "general"
    return None


# ------------------------------------------------------------------ parsing
def strip_ws(s):
    """全角・半角の空白を除いた文字列(文字列以外はそのまま返す)。"""
    return re.sub(r"\s+", "", s) if isinstance(s, str) else s


def row_values(ws, row_idx: int) -> tuple:
    return next(ws.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True))


def find_col(row: tuple, *texts: str) -> int | None:
    """row 内で空白を除いた文字列が texts のいずれかと完全一致する最初の列 index。"""
    for i, c in enumerate(row):
        if strip_ws(c) in texts:
            return i
    return None


def find_col_contains(row: tuple, *substrs: str) -> int | None:
    """row 内で substrs のいずれかを含む文字列セルの最初の列 index。"""
    for i, c in enumerate(row):
        if isinstance(c, str) and any(s in c for s in substrs):
            return i
    return None


def header_info(ws) -> tuple[int, int, str]:
    """(「等級」ヘッダ行 index(1-based), 等級列 index(0-based), 支部ラベル) を返す。

    列位置を決め打ちにせず、ヘッダーのテキスト(「等級」や支部名)を実際に探して
    対応関係を把握する。将来フォーマットが変わって列が増減・並び替わっても、
    該当するヘッダーが見つからなければ例外で分かるようにする。
    """
    pref = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
        for c in row:
            if isinstance(c, str):
                # 支部表記は年度によって「北海道支部」「北海道」の両方があり、
                # 都道府県名も古い年度では「青森県」のように正式名で書かれることが
                # ある(「北海道」はそれ自体が正式名なので対象外)。「支部」は任意に
                # しつつ、都道府県名そのものであることで他の括弧書き
                # (「（４月納付分）」等)との誤マッチを防ぐ
                m = re.search(r"[（(](.+?)(?:支部)?[）)]", c)
                if m:
                    cand = m.group(1)
                    if cand not in PREFECTURES and cand[-1:] in "都道府県" and cand[:-1] in PREFECTURES:
                        cand = cand[:-1]
                    if pref is None and cand in PREFECTURES:
                        pref = cand
        grade_col = find_col(row, "等級")
        if grade_col is not None:
            return i, grade_col, pref
    raise ValueError("「等級」ヘッダが見つかりません")


def find_amount_cols(ws, hdr: int, grade_col: int) -> tuple[int, int, int]:
    """(標準報酬月額, 報酬月額下限, 報酬月額上限) の列 index を、ヘッダーのテキスト
    (「月額」「円以上」「円未満」)を探して返す。"""
    std_col = grade_col + 1
    hdr_row = row_values(ws, hdr)
    if std_col >= len(hdr_row) or strip_ws(hdr_row[std_col]) != "月額":
        raise ValueError(f"'等級' の次の列が '月額' ヘッダではありません: {hdr_row[std_col:std_col + 1]!r}")

    min_col = max_col = None
    for offset in range(1, 4):
        row = row_values(ws, hdr + offset)
        if min_col is None:
            min_col = find_col(row, "円以上")
        if max_col is None:
            max_col = find_col(row, "円未満")
        if min_col is not None and max_col is not None:
            break
    if min_col is None or max_col is None:
        raise ValueError("報酬月額の下限/上限ヘッダ(「円以上」/「円未満」)が見つかりません")
    return std_col, min_col, max_col


def find_pair_row(ws, hdr: int) -> tuple:
    """「全額」と、それに対応する「折半額」または「被保険者負担」の見出しが並ぶ行を、
    hdr の直後数行の窓から探して返す。固定の相対位置(hdr + 1)に決め打ちせず、
    間に行が増えてもテキストで追従する。"""
    for offset in range(1, 4):
        row = row_values(ws, hdr + offset)
        if find_col(row, "全額") is not None and (
            find_col(row, "折半額") is not None or find_col(row, "被保険者負担") is not None
        ):
            return row
    raise ValueError("「全額」の見出し行が見つかりません")


def scan_amount_groups(row: tuple) -> list[dict]:
    """「全額」から始まる金額の列組を、左から順に検出する。

    旧レイアウト: 全額・折半額 の2列組(health_col=None)。
    新レイアウト(令和8年10月分〜の「保険料調整制度」): 全額・被保険者負担
    [・保険料調整額] の2〜3列組(health_col=None, employee_col, adjustment_col)。
    ブロック数も列幅もレイアウトによって異なるため、都度その場で判定する。"""
    groups = []
    i, n = 0, len(row)
    while i < n:
        if strip_ws(row[i]) != "全額":
            i += 1
            continue
        nxt = strip_ws(row[i + 1]) if i + 1 < n else None
        if nxt == "折半額":
            groups.append({"total_col": i, "half_col": i + 1, "employee_col": None, "adjustment_col": None})
            i += 2
        elif nxt == "被保険者負担":
            adj = i + 2 if i + 2 < n and strip_ws(row[i + 2]) == "保険料調整額" else None
            groups.append({"total_col": i, "half_col": None, "employee_col": i + 1, "adjustment_col": adj})
            i += 3 if adj is not None else 2
        else:
            raise ValueError(f"「全額」(列{i})の次に「折半額」も「被保険者負担」もありません")
    return groups


def amount_pair(row: tuple, group: dict) -> tuple:
    """1つの金額グループから (全額, 被保険者が実際に払う額, 調整額) を取り出す。

    旧レイアウト(全額/折半額)では調整額は常に None。新レイアウト(全額/被保険者負担
    [/保険料調整額])でも、調整制度の対象外の等級では被保険者負担欄が空になるため、
    その場合は従来通り全額の半分を実際の負担額として補う(実額を優先する設計)。"""
    total = row[group["total_col"]]
    if group["half_col"] is not None:
        return total, row[group["half_col"]], None
    employee = row[group["employee_col"]]
    if employee is None:
        half = total / 2 if isinstance(total, (int, float)) else None
        return total, half, None
    adjustment = row[group["adjustment_col"]] if group["adjustment_col"] is not None else None
    return total, employee, adjustment


def add_or_none(a, b):
    return None if a is None or b is None else a + b


def classify_amount_block(ws, hdr: int, col: int) -> str | None:
    """金額グループが、ヘッダー上段のどのブロック(health/health_ltc/care/
    childcare_levy/pension)に属するかを、その列の見出しテキストから判別する。
    health_ltc は旧レイアウトの「健康保険料・介護保険料」合算ブロック、care は
    新レイアウトで分離された「介護保険料」単独ブロック。"""
    sub = row_values(ws, hdr - 1)[col] if hdr - 1 >= 1 else None
    title = row_values(ws, hdr - 2)[col] if hdr - 2 >= 1 else None
    sub_text = sub or ""
    # 旧レイアウトでは、sub行(hdr-1)だけが「該当しない/該当する」を持ち、
    # title行(hdr-2)は health/health_ltc の両列にまたがる共通の見出し
    # (「...健康保険料・介護保険料」)なので当てにできない。また新レイアウトの
    # 「介護保険料」ブロックの title には「...に該当する場合のみ」という説明文が
    # 含まれるため、sub行単体で先に判定しないと誤判定する。
    if "該当しない" in sub_text:
        return "health"
    if "該当する" in sub_text:
        return "health_ltc"
    text = f"{sub_text}{title or ''}"
    if "介護保険料" in text:
        return "care"
    if "子ども" in text:
        return "childcare_levy"
    if "厚生年金" in text or "坑内員" in text or "一般" in text:
        return "pension"
    if "健康保険料" in text:
        return "health"
    return None


ERA_BASE = {"令和": 2019, "平成": 1989, "昭和": 1926}


def effective_key(wb) -> tuple[int, int]:
    """表の先頭にある「令和〇年〇月分」のような実施時期を西暦の (年, 月) に変換して返す。
    同じ年度でも改定のサブページ番号は昇順/降順が年度によって不統一(実際に r01 は
    古い順、r02 は新しい順だった)なので、どれが最新の改定かはこの実施時期で判定する。"""
    ws = wb[wb.sheetnames[0]]
    for row in ws.iter_rows(min_row=1, max_row=4, values_only=True):
        for c in row:
            if isinstance(c, str):
                m = re.search(r"(令和|平成|昭和)(元|\d+)年\s*(\d{1,2})月分", c)
                if m:
                    era, yr, month = m.group(1), m.group(2), int(m.group(3))
                    year = ERA_BASE[era] + (0 if yr == "元" else int(yr) - 1)
                    return year, month
    raise ValueError("実施時期(「○年○月分」)が見つかりません")


def now_month() -> tuple[int, int]:
    """現在の (年, 月) を JST で返す。実施時期が未来の改定(協会けんぽが先行して
    公開しているもの)を「最新」として latest/ に反映しないよう、この値との
    比較に使う。"""
    now = datetime.now(JST)
    return now.year, now.month


def parse_general(wb) -> list[dict]:
    out = []
    for code, (name, sn) in enumerate(zip(PREFECTURES, wb.sheetnames), 1):
        ws = wb[sn]
        hdr, grade_col, label = header_info(ws)
        if label != name:
            raise ValueError(f"シート {sn!r}: 支部ラベル {label!r} が想定 {name!r} と不一致")
        std_col, min_col, max_col = find_amount_cols(ws, hdr, grade_col)

        # 「全額」から始まる列組を左から順に拾う(年度によってブロック数や、
        # 全額/折半額 2列組か 全額/被保険者負担[/調整額] 2〜3列組かが違うため、
        # 個数・幅を決め打ちしない)。
        pair_row = find_pair_row(ws, hdr)
        groups = scan_amount_groups(pair_row)

        blocks: dict[str, dict] = {}
        for group in groups:
            block = classify_amount_block(ws, hdr, group["total_col"])
            if block is None:
                raise ValueError(f"シート {sn!r}: {group['total_col']} 列目の金額ブロック"
                                  f"(健康保険/介護/支援金/年金)を判別できません")
            if block in blocks:
                raise ValueError(f"シート {sn!r}: 金額ブロック {block!r} が複数見つかりました")
            blocks[block] = group

        # 介護保険料は、旧レイアウトでは「健康保険料・介護保険料」として合算された
        # 1つの列組(health_ltc)だが、令和8年10月分〜の新レイアウト(保険料調整制度)
        # では「健康保険料」(health)と「介護保険料」(care)に分離されている。
        # どちらか一方が揃っていれば十分とする。
        has_combined_ltc = "health_ltc" in blocks
        has_split_care = "health" in blocks and "care" in blocks
        missing = {"health", "pension"} - set(blocks)
        if missing or not (has_combined_ltc or has_split_care):
            if not (has_combined_ltc or has_split_care):
                missing = missing | {"health_ltc(またはhealth+care)"}
            raise ValueError(f"シート {sn!r}: 必須の金額ブロックが見つかりません: {missing}")

        hdr_row = row_values(ws, hdr)
        rates = {block: num(hdr_row[group["total_col"]]) for block, group in blocks.items()}
        rate_health_ltc = rates["health_ltc"] if has_combined_ltc else \
            num((rates.get("health") or 0) + (rates.get("care") or 0))

        for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
            g = row[grade_col]
            m = GRADE_RE.match(str(g).strip()) if g is not None else None
            if not m or row[std_col] is None:
                continue

            def amount(block):
                if block not in blocks:
                    return None, None, None
                return amount_pair(row, blocks[block])

            health_full, health_half, health_adjustment = amount("health")
            childcare_levy_full, childcare_levy_half, _ = amount("childcare_levy")
            pension_full, pension_half, pension_adjustment = amount("pension")

            if has_combined_ltc:
                health_ltc_full, health_ltc_half, _ = amount("health_ltc")
            else:
                care_full, care_half, _ = amount("care")
                health_ltc_full = add_or_none(health_full, care_full)
                health_ltc_half = add_or_none(health_half, care_half)

            out.append({
                "prefecture_code": code,
                "prefecture_name": name,
                "grade": int(m.group(1)),
                "pension_grade": int(m.group(2)) if m.group(2) else None,
                "remuneration_std": num(row[std_col]),
                "remuneration_min": num(row[min_col]),
                "remuneration_max": num(row[max_col]),
                "rate_health": rates.get("health"),
                "rate_health_ltc": rate_health_ltc,
                "rate_childcare_levy": rates.get("childcare_levy"),
                "rate_pension": rates.get("pension"),
                "health_full": money(health_full),
                "health_half": money(health_half),
                "health_adjustment": money(health_adjustment),
                "health_ltc_full": money(health_ltc_full),
                "health_ltc_half": money(health_ltc_half),
                "childcare_levy_full": money(childcare_levy_full),
                "childcare_levy_half": money(childcare_levy_half),
                "pension_full": money(pension_full),
                "pension_half": money(pension_half),
                "pension_adjustment": money(pension_adjustment),
            })
    return out


def parse_optional(wb) -> list[dict]:
    out = []
    for code, (name, sn) in enumerate(zip(PREFECTURES, wb.sheetnames), 1):
        ws = wb[sn]
        if sn != f"{code:02d}{name}":
            raise ValueError(f"シート名 {sn!r} が想定 {code:02d}{name!r} と不一致")
        hdr, grade_col, label = header_info(ws)
        if label != name:
            raise ValueError(f"シート {sn!r}: 支部ラベル {label!r} が想定 {name!r} と不一致")
        std_col, min_col, max_col = find_amount_cols(ws, hdr, grade_col)

        sub_row = row_values(ws, hdr - 1)
        plain_col = find_col_contains(sub_row, "該当しない")
        care_col = find_col_contains(sub_row, "該当する")
        if plain_col is None or care_col is None:
            raise ValueError(f"シート {sn!r}: 介護保険の該当有無ヘッダが見つかりません")

        hdr_row = row_values(ws, hdr)
        r_plain, r_care = hdr_row[plain_col], hdr_row[care_col]
        for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
            g = row[grade_col]
            if not isinstance(g, int) or row[std_col] is None:
                continue
            out.append({
                "prefecture_code": code,
                "prefecture_name": name,
                "grade": g,
                "remuneration_std": num(row[std_col]),
                "remuneration_min": num(row[min_col]),
                "remuneration_max": num(row[max_col]),
                "rate_health": num(r_plain),
                "rate_health_ltc": num(r_care),
                "health_full": money(row[plain_col]),
                "health_ltc_full": money(row[care_col]),
            })
    return out


PARSERS = {"general": parse_general, "optional": parse_optional}

# *.ja.json 用の日本語キー (英語キー → 日本語キー。出力順も兼ねる)
JA_KEYS = {
    "general": {
        "prefecture_code": "都道府県コード",
        "prefecture_name": "都道府県",
        "grade": "健保等級",
        "pension_grade": "厚年等級",
        "remuneration_std": "標準報酬月額",
        "remuneration_min": "報酬月額下限",
        "remuneration_max": "報酬月額上限",
        "rate_health": "健康保険料率",
        "rate_health_ltc": "介護保険料率込み健康保険料率",
        "rate_childcare_levy": "子ども・子育て支援金率",
        "rate_pension": "厚生年金保険料率",
        "health_full": "健康保険料(全額)",
        "health_half": "健康保険料(被保険者負担額)",
        "health_adjustment": "健康保険料(調整額)",
        "health_ltc_full": "健康保険料・介護あり(全額)",
        "health_ltc_half": "健康保険料・介護あり(被保険者負担額)",
        "childcare_levy_full": "子ども・子育て支援金(全額)",
        "childcare_levy_half": "子ども・子育て支援金(被保険者負担額)",
        "pension_full": "厚生年金保険料(全額)",
        "pension_half": "厚生年金保険料(被保険者負担額)",
        "pension_adjustment": "厚生年金保険料(調整額)",
    },
    "optional": {
        "prefecture_code": "都道府県コード",
        "prefecture_name": "都道府県",
        "grade": "等級",
        "remuneration_std": "標準報酬月額",
        "remuneration_min": "報酬月額下限",
        "remuneration_max": "報酬月額上限",
        "rate_health": "保険料率(介護なし)",
        "rate_health_ltc": "保険料率(介護あり)",
        "health_full": "保険料(介護なし)",
        "health_ltc_full": "保険料(介護あり)",
    },
}


def to_ja(kind: str, rows: list[dict]) -> list[dict]:
    keys = JA_KEYS[kind]
    if set(rows[0]) != set(keys):
        raise ValueError(f"{kind}: 日本語キー対応表が不一致: {set(rows[0]) ^ set(keys)}")
    return [{ja: r[en] for en, ja in keys.items()} for r in rows]


def validate(kind: str, rows: list[dict]) -> None:
    if len({r["prefecture_code"] for r in rows}) != 47:
        raise ValueError(f"{kind}: 47 都道府県そろっていません")
    per = {}
    for r in rows:
        per.setdefault(r["prefecture_code"], []).append(r["grade"])
    counts = {len(v) for v in per.values()}
    if len(counts) != 1:
        raise ValueError(f"{kind}: 都道府県ごとの等級数が不一致: {counts}")
    for code, grades in per.items():
        if grades != sorted(grades) or len(set(grades)) != len(grades):
            raise ValueError(f"{kind}: 等級が単調増加でない (pref={code})")
    prem = "health_full"
    if any(r[prem] is None or r[prem] <= 0 for r in rows):
        raise ValueError(f"{kind}: 保険料が欠損/非正の行があります")


def find_revision_candidates(page_url: str, year: str) -> list[str]:
    """改定ごとの候補ページURL一覧を返す。年度ページ自身に直接 .xlsx がある場合は
    それも1つの候補として含め、さらにその下に改定ごとのサブページ(例: r01/001,
    r01/002)があれば合わせて返す。"""
    candidates = []
    if find_xlsx_links(page_url):
        candidates.append(page_url)
    html = fetch(page_url).decode("utf-8", "replace")
    hrefs = re.findall(r'href="([^"]+)"', html)
    subpages = sorted({urljoin(page_url, h) for h in hrefs if re.search(rf"/{re.escape(year)}/\d+/?$", h)})
    for sp in subpages:
        if sp not in candidates:
            candidates.append(sp)
    return candidates


def write_revision(out: Path, year: str, rid: str, page: str, found: dict, effective: tuple[int, int]) -> dict:
    """1 つの改定分(found: kind -> {wb, url})を out 直下に書き出す。found に
    存在する種類(general/optional)だけを書き出す(片方しか無ければ片方だけ)。"""
    out.mkdir(parents=True, exist_ok=True)
    files = {}
    for kind, parse in PARSERS.items():
        if kind not in found:
            continue
        wb = found[kind]["wb"]
        rows = parse(wb)
        validate(kind, rows)
        ws = wb[wb.sheetnames[0]]
        title = " ".join(str(c) for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True)) if c)
        (out / f"{kind}.json").write_text(
            json.dumps(rows, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        (out / f"{kind}.ja.json").write_text(
            json.dumps(to_ja(kind, rows), ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        kind_eff = found[kind]["effective"]
        files[kind] = {"file": f"{kind}.json", "file_ja": f"{kind}.ja.json",
                       "rows": len(rows), "title": title, "source": found[kind]["url"],
                       "effective_from": f"{kind_eff[0]}-{kind_eff[1]:02d}"}
        print(f"  [{rid}] {kind}.json, {kind}.ja.json: {len(rows)} rows")

    index = {
        "generated_at": datetime.now(JST).isoformat(timespec="seconds"),
        "year": year,
        "revision": rid,
        "revision_page": page,
        "effective_from": f"{effective[0]}-{effective[1]:02d}",
        "files": files,
    }
    (out / "index.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return index


def copy_dir(src: Path, dst: Path) -> None:
    """src 直下のファイルを dst にコピーする。optional[.ja].json は、その改定で
    任意継続の Excel がまだ存在しない場合は無いことがあるので、無ければスキップする。"""
    dst.mkdir(parents=True, exist_ok=True)
    names = ["general.json", "general.ja.json", "optional.json", "optional.ja.json", "index.json"]
    for name in names:
        src_file = src / name
        if src_file.exists():
            (dst / name).write_bytes(src_file.read_bytes())


def build_year(year: str, out: Path, cache: Path) -> dict:
    """1 年度分(年度内に複数の改定があれば全て)を取得・変換し、年度ディレクトリに書き出す。

    改定のサブページ番号は年度によって新しい順/古い順と不統一(実際に r01 は古い順、
    r02 は新しい順だった)なので、順序は各改定の実施時期(effective_key)で判定する。
    また、改定によっては一般(health)だけ新しい Excel が公開され、任意継続
    (optional)は前の改定のものが変わらず使われる(=そのページには存在しない)
    ことがあるため、実施時期の古い順に処理しながら「直前までに見つかった
    ファイル」を引き継ぐ(carry)。任意継続がまだ一度も公開されていない期間は
    general[.ja].json だけを出力し、optional[.ja].json は無い(片方だけ揃って
    いれば片方だけ書き出す)。
    """
    page = f"{INDEX_URL}{year}"
    print(f"年度ページ: {page}")
    out.mkdir(parents=True, exist_ok=True)

    candidates = find_revision_candidates(page, year)
    if not candidates:
        candidates = [page]

    loaded = []
    for cpage in candidates:
        found: dict[str, dict] = {}
        for url in find_xlsx_links(cpage):
            data = fetch(url)
            cache_dir = cache / Path(url).stem
            cache_dir.mkdir(parents=True, exist_ok=True)
            (cache_dir / Path(url).name).write_bytes(data)
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
            kind = classify(wb)
            print(f"  {cpage} : {url} -> {kind} ({len(wb.sheetnames)} sheets)")
            if kind and kind not in found:
                # 種類ごとに自分自身の実施時期を持たせる。carry-forward で引き継がれた
                # 場合に、latest/ の optional が実際にいつ更新されたものかを
                # (この改定自体の effective_from とは別に)追跡できるようにするため。
                try:
                    kind_effective = effective_key(wb)
                except ValueError as e:
                    print(f"  [警告] {year}: {url} の実施時期が分からずスキップ ({e})", file=sys.stderr)
                    continue
                found[kind] = {"wb": wb, "url": url, "effective": kind_effective}
        if not found:
            continue
        effective = (found.get("general") or found.get("optional"))["effective"]
        loaded.append({"page": cpage, "found": found, "effective": effective})

    if not loaded:
        raise SystemExit(f"[{year}] Excel が見つかりません。")

    loaded.sort(key=lambda r: r["effective"])  # 実施時期の古い順

    revisions = []
    carry: dict[str, dict] = {}
    for rev in loaded:
        carry = {**carry, **rev["found"]}
        revisions.append({"page": rev["page"], "effective": rev["effective"], "found": dict(carry)})

    built = []
    built_effective: dict[str, tuple[int, int]] = {}
    for i, rev in enumerate(revisions, 1):
        rid = f"{i:03d}"
        try:
            write_revision(out / rid, year, rid, rev["page"], rev["found"], rev["effective"])
            built.append(rid)
            built_effective[rid] = rev["effective"]
        except (Exception, SystemExit) as e:
            # 1つの改定だけが異常な場合、同じ年度の他の改定は生かしたままにする。
            # ただし改定が1つしかない年度では、想定通りの失敗としてそのまま止める。
            if len(revisions) == 1:
                raise
            print(f"  [警告] {year}/{rid}: スキップ ({e})", file=sys.stderr)
            shutil.rmtree(out / rid, ignore_errors=True)

    if not built:
        raise SystemExit(f"[{year}] どの改定も変換できませんでした。")

    # 協会けんぽは実施時期より前に改定を公開することが多いため、「最新」は単に
    # 一番新しく見つかった改定ではなく、実施時期(effective_from)が今月以前の
    # ものの中で一番新しいものとする。まだ実施されていない改定が見つかっている
    # 場合、それが latest/ に反映されるのは実施時期が来てから。
    now = now_month()
    valid = [rid for rid in built if built_effective[rid] <= now]
    latest_rid = max(valid, key=lambda r: built_effective[r]) if valid else min(built, key=lambda r: built_effective[r])
    copy_dir(out / latest_rid, out / "latest")

    index = {
        "generated_at": datetime.now(JST).isoformat(timespec="seconds"),
        "year_page": page,
        "year": year,
        "revisions": built,
        "latest": latest_rid,
    }
    (out / "index.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return index


# --------------------------------------------------------------------- main
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", help="年度 (例: r08)。'all' で全年度。省略時は最新のみ")
    ap.add_argument("--out", default="site")
    ap.add_argument("--cache", default=".cache")
    a = ap.parse_args()

    years = list_years()
    latest = years[-1]
    if a.year is None:
        targets = [latest]
    elif a.year == "all":
        targets = years
    elif a.year in years:
        targets = [a.year]
    else:
        raise SystemExit(f"年度 {a.year} は存在しません (候補: {years})")

    out, cache = Path(a.out), Path(a.cache)
    built = []
    year_indexes: dict[str, dict] = {}
    for year in targets:
        try:
            year_indexes[year] = build_year(year, out / year, cache / year)
            built.append(year)
        except (Exception, SystemExit) as e:
            # 単一年度を明示指定した場合は想定通りの失敗として即座に止める。
            # --year all の場合は、レイアウトが異なる古い年度のために全体を
            # 止めないよう、警告を出してスキップし、残りの年度は処理を続ける。
            if len(targets) == 1:
                raise
            print(f"[警告] {year}: スキップ ({e})", file=sys.stderr)
            shutil.rmtree(out / year, ignore_errors=True)

    if not built:
        raise SystemExit("どの年度も変換できませんでした。")

    # 各年度・各改定で実際に生成されたファイルを列挙する。general.json は複数改定が
    # あっても欠けることはないが、optional.json は改定によって無いことがあるため、
    # 「optional.json を含む最新の改定はどれか」のようにこの一覧をフィルタして
    # 適切な改定を選べるようにする。実施時期の古い順に並べる。
    revisions_manifest = []
    for year in built:
        for rid in year_indexes[year]["revisions"]:
            rev_index = json.loads((out / year / rid / "index.json").read_text(encoding="utf-8"))
            files = [f for kind_files in rev_index["files"].values()
                     for f in (kind_files["file"], kind_files["file_ja"])]
            revisions_manifest.append({
                "year": year,
                "revision": rid,
                "path": f"{year}/{rid}",
                "effective_from": rev_index["effective_from"],
                "files": files,
            })
    revisions_manifest.sort(key=lambda r: r["effective_from"])

    # 全年度・全改定を通して、年度をまたいだ「実施時期が今月以前で一番新しいもの」
    # を latest/ に反映する。協会けんぽは実施時期より前に改定を公開することが多く、
    # 単に一番新しく見つかった年度・改定を使うと、まだ実施されていない将来の改定を
    # 「最新」として配信してしまうため。
    now_str = f"{now_month()[0]}-{now_month()[1]:02d}"
    valid_entries = [r for r in revisions_manifest if r["effective_from"] <= now_str]
    latest_entry = max(valid_entries, key=lambda r: r["effective_from"]) if valid_entries else \
        min(revisions_manifest, key=lambda r: r["effective_from"])
    copy_dir(out / latest_entry["path"], out / "latest")
    latest_year = latest_entry["year"]

    (out / "index.json").write_text(json.dumps({
        "generated_at": datetime.now(JST).isoformat(timespec="seconds"),
        "years": built,
        "latest": latest_year,
        "revisions": revisions_manifest,
    }, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
