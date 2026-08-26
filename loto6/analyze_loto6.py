#!/usr/bin/env python3
"""ロト6の過去当選番号を集計し、出現回数の順位から購入番号を組み立てる。

- ホット枠 (10口): 出現回数の多い順に並べ、上位20個を循環させて6個ずつ切り出す。
  上位20個が各3回ずつ均等に登場する(10口 x 6個 = 60 = 20 x 3)。
- コールド枠 (3口): 出現回数の少ない順に並べ、下位18個を先頭から6個ずつ切り出す。

使い方:
    python3 analyze_loto6.py [--csv data/loto6_draws.csv] [--out report.md]
"""

from __future__ import annotations

import argparse
import csv
import math
from collections import Counter
from pathlib import Path

NUMBER_MIN = 1
NUMBER_MAX = 43
PICK = 6
HOT_POOL = 20      # ホット枠に使う上位番号の数 (10口 x 6 / 3回 = 20)
HOT_TICKETS = 10
COLD_TICKETS = 3


def load_draws(path: Path) -> list[dict]:
    """CSVから抽せん結果を読み込み、値の妥当性を検証して返す。"""
    draws = []
    with path.open(encoding="utf-8", newline="") as f:
        for row in csv.DictReader(f):
            main = [int(row[f"n{i}"]) for i in range(1, PICK + 1)]
            if len(set(main)) != PICK:
                raise ValueError(f"第{row['draw']}回: 本数字が重複しています {main}")
            if not all(NUMBER_MIN <= n <= NUMBER_MAX for n in main):
                raise ValueError(f"第{row['draw']}回: 範囲外の本数字 {main}")
            draws.append(
                {
                    "draw": int(row["draw"]),
                    "date": row["date"],
                    "main": sorted(main),
                    "bonus": int(row["bonus"]),
                }
            )
    if not draws:
        raise ValueError(f"{path} に抽せん結果がありません")
    return draws


def count_main(draws: list[dict]) -> Counter:
    counts = Counter()
    for d in draws:
        counts.update(d["main"])
    for n in range(NUMBER_MIN, NUMBER_MAX + 1):
        counts.setdefault(n, 0)
    return counts


def ranking(counts: Counter, *, descending: bool) -> list[int]:
    """出現回数順の番号リスト。同数のときは番号の小さい順で安定させる。"""
    sign = -1 if descending else 1
    return sorted(counts, key=lambda n: (sign * counts[n], n))


def chi_square(counts: Counter, draws: int) -> tuple[float, int]:
    """一様分布(=完全にランダム)からのズレを測る。"""
    expected = draws * PICK / (NUMBER_MAX - NUMBER_MIN + 1)
    stat = sum((counts[n] - expected) ** 2 / expected for n in counts)
    return stat, len(counts) - 1


def chi_square_p(stat: float, dof: int) -> float:
    """自由度が偶数のときの上側確率。dof=42 なので閉じた式で計算できる。"""
    if dof % 2 != 0:
        raise ValueError("この簡易実装は偶数の自由度のみ対応")
    half = stat / 2
    term = math.exp(-half)
    total = term
    for k in range(1, dof // 2):
        term *= half / k
        total += term
    return min(1.0, total)


def hot_tickets(order: list[int]) -> list[list[int]]:
    """上位HOT_POOL個を循環させ、6個ずつ切り出して10口作る。"""
    pool = order[:HOT_POOL]
    tickets = []
    for i in range(HOT_TICKETS):
        start = (i * PICK) % HOT_POOL
        tickets.append(sorted(pool[(start + j) % HOT_POOL] for j in range(PICK)))
    return tickets


def cold_tickets(order: list[int]) -> list[list[int]]:
    """出現回数の少ない順に、先頭から6個ずつ切り出して3口作る。"""
    pool = order[: COLD_TICKETS * PICK]
    return [sorted(pool[i * PICK : (i + 1) * PICK]) for i in range(COLD_TICKETS)]


def format_ticket(ticket: list[int]) -> str:
    return " ".join(f"{n:02d}" for n in ticket)


def build_report(draws: list[dict], counts: Counter) -> str:
    hot_order = ranking(counts, descending=True)
    cold_order = ranking(counts, descending=False)
    hot = hot_tickets(hot_order)
    cold = cold_tickets(cold_order)
    rank_of = {n: i + 1 for i, n in enumerate(hot_order)}
    expected = len(draws) * PICK / (NUMBER_MAX - NUMBER_MIN + 1)
    stat, dof = chi_square(counts, len(draws))
    p = chi_square_p(stat, dof)

    lines = [
        "# ロト6 出現回数分析と番号生成",
        "",
        f"- 対象: 第{draws[0]['draw']}回 ({draws[0]['date']}) 〜 第{draws[-1]['draw']}回 ({draws[-1]['date']})",
        f"- 抽せん回数: {len(draws)}回 / 本数字の延べ出現数: {len(draws) * PICK}個",
        f"- 1個あたりの期待出現回数: {expected:.1f}回",
        f"- カイ二乗検定: χ²={stat:.1f} (自由度{dof}), p={p:.3f}"
        f" → {'一様分布からのズレは統計的に有意ではない' if p >= 0.05 else '一様分布からのズレが有意'}",
        "",
        "## 出現回数ランキング(本数字)",
        "",
        "| 順位 | 番号 | 出現回数 | 期待値との差 |",
        "|---:|---:|---:|---:|",
    ]
    for rank, n in enumerate(hot_order, start=1):
        diff = counts[n] - expected
        lines.append(f"| {rank} | {n:02d} | {counts[n]} | {diff:+.1f} |")

    lines += [
        "",
        f"## ホット枠 {HOT_TICKETS}口(出現回数の多い順)",
        "",
        f"上位{HOT_POOL}個を出現回数順に並べ、6個ずつ循環させて切り出したもの。"
        f"上位{HOT_POOL}個がそれぞれちょうど3回ずつ使われる。",
        "",
        "| 口 | 番号 | 使用した順位 |",
        "|---:|---|---|",
    ]
    for i, t in enumerate(hot, start=1):
        ranks = ",".join(str(rank_of[n]) for n in sorted(t, key=lambda n: rank_of[n]))
        lines.append(f"| {i} | {format_ticket(t)} | {ranks} |")

    lines += [
        "",
        f"## 逆張り枠 {COLD_TICKETS}口(出現回数の少ない順)",
        "",
        f"下位{COLD_TICKETS * PICK}個を出現回数の少ない順に並べ、先頭から6個ずつ切り出したもの。",
        "",
        "| 口 | 番号 | 使用した順位 |",
        "|---:|---|---|",
    ]
    for i, t in enumerate(cold, start=1):
        ranks = ",".join(str(rank_of[n]) for n in sorted(t, key=lambda n: rank_of[n]))
        lines.append(f"| {i} | {format_ticket(t)} | {ranks} |")

    lines += [
        "",
        "## 注意",
        "",
        "ロト6の抽せんは毎回独立で、過去の出現回数は次回の当せん確率を変えない。"
        "上のカイ二乗検定の通り、出現回数のばらつきはランダムな揺らぎの範囲に収まっている。"
        "この番号表は「出現回数順に機械的に並べた結果」であって、当たりやすさの根拠ではない。",
        "",
    ]
    return "\n".join(lines)


def main() -> None:
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--csv", type=Path, default=here / "data" / "loto6_draws.csv")
    ap.add_argument("--out", type=Path, default=here / "report.md")
    args = ap.parse_args()

    draws = load_draws(args.csv)
    counts = count_main(draws)
    report = build_report(draws, counts)
    args.out.write_text(report, encoding="utf-8")
    print(report)


if __name__ == "__main__":
    main()
