#!/usr/bin/env python3
"""ロト6の抽せん結果CSVを取り込み、data/loto6_draws.csv に不足回を追記する。

みずほ銀行など配布元の列名(回別/抽せん日/本数字1..6/ボーナス数字)を自動判別し、
文字コード(Shift_JIS / UTF-8)も自動判別する。既存の回は上書きせず、
新しい回だけを追加して抽せん回の昇順で保存する。

使い方:
    # 手元にダウンロードしたCSVから取り込む
    python3 fetch_draws.py --input ~/Downloads/loto6.csv

    # URLから直接取り込む(そのホストに到達できるネットワークで実行すること)
    python3 fetch_draws.py --url https://example.com/loto6.csv

    # 取り込まずに差分だけ確認する
    python3 fetch_draws.py --input loto6.csv --dry-run
"""

from __future__ import annotations

import argparse
import csv
import io
import re
import sys
import urllib.request
from pathlib import Path

NUMBER_MIN = 1
NUMBER_MAX = 43
PICK = 6
ENCODINGS = ("utf-8-sig", "cp932", "utf-8")
OUT_FIELDS = ["draw", "date", "n1", "n2", "n3", "n4", "n5", "n6", "bonus"]


def decode(raw: bytes) -> str:
    for enc in ENCODINGS:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"文字コードを判別できません(試したもの: {', '.join(ENCODINGS)})")


def find_columns(header: list[str]) -> dict:
    """配布元ごとに違う列名から、必要な列の位置を割り出す。"""
    names = [h.strip() for h in header]
    cols: dict = {}

    def first(*preds) -> int | None:
        """述語を順に試し、最初に一致した列の位置を返す(位置0も正しく扱う)。"""
        for pred in preds:
            for i, h in enumerate(names):
                if pred(h):
                    return i
        return None

    cols["draw"] = first(
        lambda h: h in ("draw", "回別", "抽せん回", "回号"),
        lambda h: "回" in h and "回数" not in h,
    )
    cols["date"] = first(
        lambda h: h in ("date", "抽せん日", "抽選日"),
        lambda h: "日" in h and "曜" not in h,
    )
    main = [i for i, h in enumerate(names) if re.fullmatch(r"(本数字|数字)\s*[1-6]", h)]
    if len(main) != PICK:
        main = [i for i, h in enumerate(names) if re.fullmatch(r"n[1-6]", h)]
    cols["main"] = main
    cols["bonus"] = first(
        lambda h: h in ("bonus", "ボーナス数字"),
        lambda h: "ボーナス" in h,
    )

    missing = [k for k in ("draw", "date", "bonus") if cols[k] is None]
    if missing or len(cols["main"]) != PICK:
        raise ValueError(
            "列を判別できません。判別できなかった項目: "
            + ", ".join(missing + ([f"本数字x{PICK}"] if len(cols["main"]) != PICK else []))
            + f"\n読み込んだヘッダ: {names}"
            + "\n列名を draw,date,n1..n6,bonus に直してから再実行してください。"
        )
    return cols


def parse_date(value: str) -> str:
    """2026/8/24・2026-8-24・2026年8月24日 のいずれも YYYY-MM-DD に揃える。"""
    nums = re.findall(r"\d+", value)
    if len(nums) < 3:
        raise ValueError(f"抽せん日を解釈できません: {value!r}")
    y, m, d = (int(n) for n in nums[:3])
    return f"{y:04d}-{m:02d}-{d:02d}"


def parse_rows(text: str) -> dict[int, dict]:
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        raise ValueError("CSVが空です")
    cols = find_columns(rows[0])
    draws: dict[int, dict] = {}
    for lineno, row in enumerate(rows[1:], start=2):
        if not any(cell.strip() for cell in row):
            continue
        try:
            digits = re.findall(r"\d+", row[cols["draw"]])
            if not digits:
                continue
            draw = int(digits[0])
            main = sorted(int(row[i]) for i in cols["main"])
            bonus = int(re.findall(r"\d+", row[cols["bonus"]])[0])
            date = parse_date(row[cols["date"]])
        except (IndexError, ValueError) as e:
            raise ValueError(f"{lineno}行目を解釈できません: {e}\n行の内容: {row}") from e
        if len(set(main)) != PICK:
            raise ValueError(f"第{draw}回: 本数字が重複しています {main}")
        for n in main + [bonus]:
            if not NUMBER_MIN <= n <= NUMBER_MAX:
                raise ValueError(f"第{draw}回: 範囲外の数字 {n}")
        if bonus in main:
            raise ValueError(f"第{draw}回: ボーナス数字が本数字と重複しています {bonus}")
        draws[draw] = {
            "draw": draw,
            "date": date,
            **{f"n{i}": n for i, n in enumerate(main, start=1)},
            "bonus": bonus,
        }
    if not draws:
        raise ValueError("取り込める抽せん結果が1件もありません")
    return draws


def load_existing(path: Path) -> dict[int, dict]:
    if not path.exists():
        return {}
    with path.open(encoding="utf-8", newline="") as f:
        return {int(r["draw"]): {k: r[k] for k in OUT_FIELDS} for r in csv.DictReader(f)}


def missing_draws(draws) -> list[int]:
    """最小回〜最大回のうち、データに存在しない回(歯抜け)を返す。"""
    if not draws:
        return []
    return sorted(set(range(min(draws), max(draws) + 1)) - set(draws))


def format_ranges(numbers: list[int], limit: int = 5) -> str:
    """[1198,1199,1200,1300] を '1198〜1200, 1300' のように畳んで表示する。"""
    ranges: list[tuple[int, int]] = []
    for n in numbers:
        if ranges and n == ranges[-1][1] + 1:
            ranges[-1] = (ranges[-1][0], n)
        else:
            ranges.append((n, n))
    shown = [f"第{a}回" if a == b else f"第{a}回〜第{b}回" for a, b in ranges[:limit]]
    if len(ranges) > limit:
        shown.append(f"ほか{len(ranges) - limit}区間")
    return ", ".join(shown)


def read_source(args) -> str:
    if args.url:
        req = urllib.request.Request(args.url, headers={"User-Agent": "loto6-fetch/1.0"})
        with urllib.request.urlopen(req, timeout=60) as res:
            return decode(res.read())
    return decode(args.input.read_bytes())


def main() -> int:
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--input", type=Path, help="取り込むCSVファイル")
    src.add_argument("--url", help="取り込むCSVのURL")
    ap.add_argument("--csv", type=Path, default=here / "data" / "loto6_draws.csv", help="更新先")
    ap.add_argument("--dry-run", action="store_true", help="書き込まずに差分だけ表示")
    args = ap.parse_args()

    try:
        incoming = parse_rows(read_source(args))
    except Exception as e:  # 配布元のフォーマット差異はここで人間に見える形で止める
        print(f"取り込みに失敗しました: {e}", file=sys.stderr)
        return 1

    existing = load_existing(args.csv)
    new = sorted(set(incoming) - set(existing))
    conflicts = [
        d
        for d in sorted(set(incoming) & set(existing))
        if [str(incoming[d][k]) for k in OUT_FIELDS] != [str(existing[d][k]) for k in OUT_FIELDS]
    ]

    print(f"取り込み元: {len(incoming)}回 (第{min(incoming)}回〜第{max(incoming)}回)")
    print(f"既存データ: {len(existing)}回" + (f" (第{min(existing)}回〜第{max(existing)}回)" if existing else ""))
    print(f"追加される回: {len(new)}回" + (f" (第{new[0]}回〜第{new[-1]}回)" if new else ""))
    if conflicts:
        print(f"既存と内容が食い違う回: {len(conflicts)}回 {conflicts[:10]} (既存を優先し、上書きしません)")
    merged = dict(existing)
    for d in new:
        merged[d] = incoming[d]
    gaps = missing_draws(merged)
    if gaps:
        print(f"取り込み後も歯抜けが残る回: {len(gaps)}回 ({format_ranges(gaps)})")
    else:
        print(f"取り込み後は第{min(merged)}回〜第{max(merged)}回が歯抜けなく揃う。" if merged else "")

    if not new:
        print("追加する回はありません。")
        return 0
    if args.dry_run:
        print("--dry-run のため書き込みませんでした。")
        return 0

    with args.csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=OUT_FIELDS)
        w.writeheader()
        for d in sorted(merged):
            w.writerow(merged[d])
    print(f"{args.csv} を更新しました({len(merged)}回)。")
    print("続けて `python3 analyze_loto6.py` を実行するとレポートが再生成されます。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
