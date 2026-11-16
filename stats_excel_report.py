#!/usr/bin/env python3
"""公開統計CSVを取得・整形し、Excel報告書（表＋グラフ）を出力するツール。

想定データ形式: Our World in Data のグラフ用CSVのような
「Entity, Code, Year, 値...」形式の年次データ。
値の列は自動判定する（--value-column で明示も可能）。

使い方の例:
  # URLから取得して報告書を作る
  python stats_excel_report.py --out report.xlsx

  # 手元のCSVから作る（対象と期間を指定）
  python stats_excel_report.py --csv data/temperature-anomaly.csv \
      --entities World --start-year 1950 --out report.xlsx

エラー時（ファイルが無い、CSVが壊れている、列が足りない、条件に合う
データが0件など）は、原因が分かる日本語のメッセージを表示して
終了コード1で終了します。詳しい例外を見たい場合は --verbose を付けてください。
"""

from __future__ import annotations

import argparse
import csv
import io
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DEFAULT_URL = "https://ourworldindata.org/grapher/temperature-anomaly.csv"
USER_AGENT = "stats-excel-report/1.0 (sample tool; single fetch)"

KEY_COLUMNS = {"entity", "code", "year"}


class ToolError(Exception):
    """利用者に見せる日本語メッセージを持つエラー。"""


# ---------------------------------------------------------------- 入力の取得
def fetch_csv_text(url: str) -> str:
    """URLからCSVを1回だけ取得してテキストで返す。"""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            raw = res.read()
    except urllib.error.HTTPError as e:
        raise ToolError(
            f"データの取得に失敗しました（HTTP {e.code}）。"
            f"URLが正しいか確認してください: {url}"
        ) from e
    except urllib.error.URLError as e:
        raise ToolError(
            f"データの取得に失敗しました（{e.reason}）。"
            "ネットワーク接続とURLを確認してください。"
        ) from e
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError as e:
        raise ToolError("取得したデータがUTF-8として読み取れませんでした。") from e


def read_csv_file(path: Path) -> str:
    if not path.exists():
        raise ToolError(f"入力ファイルが見つかりません: {path}")
    if not path.is_file():
        raise ToolError(f"入力パスがファイルではありません: {path}")
    try:
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as e:
        raise ToolError(
            f"入力ファイルをUTF-8として読み取れませんでした: {path}"
        ) from e
    except OSError as e:
        raise ToolError(f"入力ファイルを読み取れませんでした: {path}（{e}）") from e


# ---------------------------------------------------------------- CSVの整形
def detect_value_column(fieldnames: list[str], rows: list[dict], explicit: str | None) -> str:
    """値として使う列を決める。"""
    lowered = {name.lower(): name for name in fieldnames}
    if explicit:
        if explicit in fieldnames:
            return explicit
        if explicit.lower() in lowered:
            return lowered[explicit.lower()]
        raise ToolError(
            f"指定された値の列「{explicit}」がCSVにありません。"
            f"ある列は次のとおりです: {', '.join(fieldnames)}"
        )
    if "average" in lowered:
        return lowered["average"]
    # Entity / Code / Year 以外で、値の大半が数値として読める列を探す
    for name in fieldnames:
        if name.lower() in KEY_COLUMNS:
            continue
        sample = [r.get(name) for r in rows[:200]]
        numeric = 0
        total = 0
        for v in sample:
            if v is None or str(v).strip() == "":
                continue
            total += 1
            try:
                float(str(v).replace(",", ""))
                numeric += 1
            except ValueError:
                pass
        if total > 0 and numeric / total >= 0.8:
            return name
    raise ToolError(
        "値として使える数値の列が見つかりませんでした。"
        f"ある列は次のとおりです: {', '.join(fieldnames)}。"
        "--value-column で列名を指定してください。"
    )


def parse_and_clean(text: str, value_column_arg: str | None):
    """CSVテキストを読み、整形済みレコードと整形サマリを返す。

    戻り値: (records, summary)
      records: [{"entity": str, "year": int, "value": float}, ...] 年順
      summary: dict（元行数、採用行数、除外内訳など）
    """
    if text.strip() == "":
        raise ToolError("入力データが空です。ファイルの内容を確認してください。")
    reader = csv.DictReader(io.StringIO(text))
    if reader.fieldnames is None:
        raise ToolError("CSVのヘッダー行が読み取れませんでした。")
    fieldnames = [str(f).strip() for f in reader.fieldnames if f is not None]
    # DictReaderのキーは元のヘッダー文字列。そのまま使う。
    raw_fieldnames = reader.fieldnames

    def find_col(wanted: str) -> str | None:
        for f in raw_fieldnames:
            if f is not None and str(f).strip().lower() == wanted:
                return f
        return None

    entity_col = find_col("entity")
    year_col = find_col("year")
    if entity_col is None or year_col is None:
        missing = [n for n, c in (("Entity", entity_col), ("Year", year_col)) if c is None]
        raise ToolError(
            "必要な列が足りません: " + "、".join(missing) +
            f"。ある列は次のとおりです: {', '.join(str(f) for f in raw_fieldnames)}"
        )

    try:
        raw_rows = list(reader)
    except csv.Error as e:
        raise ToolError(f"CSVの形式が正しくありません（{e}）。") from e
    if not raw_rows:
        raise ToolError("CSVにデータ行がありません（ヘッダーだけです）。")

    value_col_display = detect_value_column(
        [str(f) for f in raw_fieldnames], raw_rows, value_column_arg
    )
    value_col = find_col(value_col_display.lower()) or value_col_display

    records: list[dict] = []
    seen: set[tuple[str, int]] = set()
    dropped_missing = 0
    dropped_bad = 0
    dropped_dup = 0
    entities_found: set[str] = set()
    for row in raw_rows:
        entity_raw = row.get(entity_col)
        year_raw = row.get(year_col)
        value_raw = row.get(value_col)
        if entity_raw is None or year_raw is None:
            dropped_bad += 1
            continue
        entity = str(entity_raw).strip()
        if entity == "":
            dropped_missing += 1
            continue
        entities_found.add(entity)
        try:
            year = int(str(year_raw).strip())
        except (TypeError, ValueError):
            dropped_bad += 1
            continue
        if value_raw is None or str(value_raw).strip() == "":
            dropped_missing += 1
            continue
        try:
            value = float(str(value_raw).replace(",", "").strip())
        except ValueError:
            dropped_bad += 1
            continue
        key = (entity, year)
        if key in seen:
            dropped_dup += 1
            continue
        seen.add(key)
        records.append({"entity": entity, "year": year, "value": value})

    if not records:
        raise ToolError(
            "有効なデータ行が0件でした。CSVの形式や値の列を確認してください"
            f"（欠損で除外 {dropped_missing} 行、数値として読めず除外 {dropped_bad} 行、"
            f"重複で除外 {dropped_dup} 行）。"
        )

    records.sort(key=lambda r: (r["entity"], r["year"]))
    summary = {
        "raw_rows": len(raw_rows),
        "kept_rows": len(records),
        "dropped_missing": dropped_missing,
        "dropped_bad": dropped_bad,
        "dropped_dup": dropped_dup,
        "value_column": value_col_display,
        "entities_found": sorted(entities_found),
    }
    return records, summary


# ---------------------------------------------------------------- 集計
def linear_trend_per_decade(points: list[tuple[int, float]]) -> float | None:
    """最小二乗法で「10年あたりの変化量」を返す。"""
    n = len(points)
    if n < 2:
        return None
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    mx = sum(xs) / n
    my = sum(ys) / n
    denom = sum((x - mx) ** 2 for x in xs)
    if denom == 0:
        return None
    slope = sum((x - mx) * (y - my) for x, y in points) / denom
    return slope * 10


def summarize(records: list[dict], entities: list[str]):
    """エンティティ別サマリと年代別平均を作る。"""
    per_entity: dict[str, list[dict]] = {e: [] for e in entities}
    for r in records:
        if r["entity"] in per_entity:
            per_entity[r["entity"]].append(r)

    summary_rows = []
    decade_rows: dict[str, dict[int, float]] = {}
    decades: set[int] = set()
    for entity in entities:
        rows = sorted(per_entity[entity], key=lambda r: r["year"])
        if not rows:
            continue
        values = [r["value"] for r in rows]
        vmin = min(rows, key=lambda r: r["value"])
        vmax = max(rows, key=lambda r: r["value"])
        latest = rows[-1]
        trend = linear_trend_per_decade([(r["year"], r["value"]) for r in rows])
        summary_rows.append({
            "entity": entity,
            "count": len(rows),
            "mean": sum(values) / len(values),
            "min": vmin["value"],
            "min_year": vmin["year"],
            "max": vmax["value"],
            "max_year": vmax["year"],
            "latest_year": latest["year"],
            "latest": latest["value"],
            "trend": trend,
        })
        by_decade: dict[int, list[float]] = {}
        for r in rows:
            d = (r["year"] // 10) * 10
            by_decade.setdefault(d, []).append(r["value"])
            decades.add(d)
        decade_rows[entity] = {d: sum(v) / len(v) for d, v in by_decade.items()}
    return summary_rows, decade_rows, sorted(decades)


# ---------------------------------------------------------------- Excel出力
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)
SECTION_FONT = Font(bold=True, size=12)


def style_header_row(ws, row: int, ncols: int) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def build_workbook(meta: dict, records: list[dict], entities: list[str],
                   summary_rows: list[dict], decade_rows: dict, decades: list[int],
                   value_column: str) -> Workbook:
    wb = Workbook()

    # ---- データシート（グラフの参照元になるので先に作る）
    ws_data = wb.active
    ws_data.title = "データ"
    years = sorted({r["year"] for r in records})
    lookup = {(r["entity"], r["year"]): r["value"] for r in records}
    ws_data.cell(row=1, column=1, value="年")
    for j, entity in enumerate(entities, start=2):
        ws_data.cell(row=1, column=j, value=entity)
    style_header_row(ws_data, 1, 1 + len(entities))
    for i, year in enumerate(years, start=2):
        c = ws_data.cell(row=i, column=1, value=year)
        c.number_format = "0"
        for j, entity in enumerate(entities, start=2):
            v = lookup.get((entity, year))
            if v is not None:
                cell = ws_data.cell(row=i, column=j, value=v)
                cell.number_format = "0.00"
    ws_data.column_dimensions["A"].width = 8
    for j in range(2, 2 + len(entities)):
        ws_data.column_dimensions[get_column_letter(j)].width = 22
    ws_data.freeze_panes = "B2"
    last_data_row = 1 + len(years)

    # ---- 概要シート
    ws = wb.create_sheet("概要", 0)
    ws.column_dimensions["A"].width = 26
    for col, w in zip("BCDEFG", (16, 14, 14, 14, 14, 18)):
        ws.column_dimensions[col].width = w

    ws["A1"] = "公開統計データ 報告書"
    ws["A1"].font = TITLE_FONT
    info = [
        ("指標（値の列）", value_column),
        ("入力", meta["input_desc"]),
        ("処理日時", meta["processed_at"]),
        ("対象（Entity）", "、".join(entities)),
        ("期間", meta["period_desc"]),
        ("整形サマリ", meta["clean_desc"]),
        ("出典", meta["source_desc"]),
    ]
    row = 3
    for label, val in info:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=2, value=val)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="対象別の要約").font = SECTION_FONT
    row += 1
    headers = ["対象", "データ数", "平均", "最小", "最小の年", "最大", "最大の年",
               "最新年", "最新の値", "変化傾向（10年あたり）"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=row, column=j, value=h)
    style_header_row(ws, row, len(headers))
    header_row_summary = row
    row += 1
    for sr in summary_rows:
        vals = [sr["entity"], sr["count"], sr["mean"], sr["min"], sr["min_year"],
                sr["max"], sr["max_year"], sr["latest_year"], sr["latest"], sr["trend"]]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(row=row, column=j, value=v)
            if j in (3, 4, 6, 9, 10) and v is not None:
                cell.number_format = "0.00"
            if j in (2, 5, 7, 8):
                cell.number_format = "0"
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="年代別平均").font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=1, value="年代")
    for j, entity in enumerate(entities, start=2):
        ws.cell(row=row, column=j, value=entity)
    style_header_row(ws, row, 1 + len(entities))
    row += 1
    for d in decades:
        ws.cell(row=row, column=1, value=f"{d}年代")
        for j, entity in enumerate(entities, start=2):
            v = decade_rows.get(entity, {}).get(d)
            if v is not None:
                cell = ws.cell(row=row, column=j, value=v)
                cell.number_format = "0.00"
        row += 1

    # ---- グラフ（年次推移の折れ線）
    chart = LineChart()
    chart.title = f"年次推移（{value_column}）"
    chart.y_axis.title = value_column
    chart.x_axis.title = "年"
    chart.width = 30
    chart.height = 13
    data_ref = Reference(ws_data, min_col=2, max_col=1 + len(entities),
                         min_row=1, max_row=last_data_row)
    cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=last_data_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    row += 1
    ws.add_chart(chart, f"A{row}")

    return wb


# ---------------------------------------------------------------- メイン
def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="公開統計CSVを取得・整形し、Excel報告書を出力します。"
    )
    p.add_argument("--url", default=DEFAULT_URL,
                   help=f"取得先CSVのURL（--csv指定時は使いません）。既定: {DEFAULT_URL}")
    p.add_argument("--csv", type=Path, default=None,
                   help="手元のCSVファイルを使う場合に指定します。")
    p.add_argument("--entities", default="World,Northern Hemisphere,Southern Hemisphere",
                   help="対象のEntityをカンマ区切りで指定。既定は World と南北両半球。")
    p.add_argument("--start-year", type=int, default=None, help="開始年（含む）。")
    p.add_argument("--end-year", type=int, default=None, help="終了年（含む）。")
    p.add_argument("--value-column", default=None,
                   help="値として使う列名。省略時は自動判定します。")
    p.add_argument("--out", type=Path, default=Path("report.xlsx"),
                   help="出力先の .xlsx パス。既定: report.xlsx")
    p.add_argument("--verbose", action="store_true",
                   help="エラー時に詳しい例外（traceback）を表示します。")
    return p.parse_args(argv)


def run(args: argparse.Namespace) -> Path:
    if args.csv is not None:
        text = read_csv_file(args.csv)
        input_desc = f"ローカルファイル: {args.csv}"
        source_desc = "手元のCSVファイル（取得元はファイル側で管理）"
    else:
        text = fetch_csv_text(args.url)
        input_desc = f"URL取得: {args.url}"
        source_desc = args.url

    records, clean = parse_and_clean(text, args.value_column)

    wanted = [e.strip() for e in args.entities.split(",") if e.strip() != ""]
    if not wanted:
        raise ToolError("--entities が空です。対象を1つ以上指定してください。")
    unknown = [e for e in wanted if e not in clean["entities_found"]]
    if unknown:
        raise ToolError(
            "指定された対象がデータにありません: " + "、".join(unknown) +
            "。データにある対象は次のとおりです: " + "、".join(clean["entities_found"])
        )

    filtered = [r for r in records if r["entity"] in wanted]
    if args.start_year is not None:
        filtered = [r for r in filtered if r["year"] >= args.start_year]
    if args.end_year is not None:
        filtered = [r for r in filtered if r["year"] <= args.end_year]
    if not filtered:
        raise ToolError(
            "条件に合うデータが0件でした。対象と期間の指定を見直してください。"
        )

    summary_rows, decade_rows, decades = summarize(filtered, wanted)
    years = [r["year"] for r in filtered]
    meta = {
        "input_desc": input_desc,
        "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "period_desc": f"{min(years)}年 〜 {max(years)}年",
        "clean_desc": (
            f"元データ {clean['raw_rows']} 行 → 採用 {clean['kept_rows']} 行"
            f"（欠損で除外 {clean['dropped_missing']} 行、"
            f"数値として読めず除外 {clean['dropped_bad']} 行、"
            f"重複で除外 {clean['dropped_dup']} 行）／ 期間・対象で絞った後は {len(filtered)} 行"
        ),
        "source_desc": source_desc,
    }
    wb = build_workbook(meta, filtered, wanted, summary_rows, decade_rows,
                        decades, clean["value_column"])
    args.out.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(args.out)
    except OSError as e:
        raise ToolError(f"出力ファイルに書き込めませんでした: {args.out}（{e}）") from e

    print(f"入力: {input_desc}")
    print(f"整形: {meta['clean_desc']}")
    print(f"出力: {args.out}（シート: 概要 / データ、折れ線グラフ付き）")
    return args.out


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv if argv is not None else sys.argv[1:])
    try:
        run(args)
    except ToolError as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1
    except Exception as e:  # 想定外も、まずは何が起きたかを伝える
        if args.verbose:
            raise
        print(f"エラー: 予期しない問題が発生しました（{type(e).__name__}: {e}）。"
              "詳しく見る場合は --verbose を付けて再実行してください。", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
