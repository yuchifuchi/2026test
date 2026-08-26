#!/usr/bin/env python3
"""fetch_draws.py の取り込み処理のテスト。`python3 test_fetch_draws.py` で実行する。"""

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import fetch_draws as fd

MIZUHO = "\n".join(
    [
        "回別,抽せん日,曜日,本数字1,本数字2,本数字3,本数字4,本数字5,本数字6,ボーナス数字,1等口数",
        "1197,2017/8/3,木,3,24,28,36,37,39,27,2",
        "1198,2017/8/7,月,5,11,19,23,30,41,8,1",
        "第1199回,2017年8月10日,木,2,9,14,22,33,43,17,0",
        "",
    ]
)

OWN = "\n".join(
    [
        "draw,date,n1,n2,n3,n4,n5,n6,bonus",
        "1197,2017-08-03,3,24,28,36,37,39,27",
        "",
    ]
)


def expect_error(text: str, fragment: str) -> None:
    try:
        fd.parse_rows(text)
    except ValueError as e:
        assert fragment in str(e), f"想定と違うエラー: {e}"
        return
    raise AssertionError(f"エラーになるはずが通ってしまった: {fragment}")


def test_mizuho_format():
    draws = fd.parse_rows(MIZUHO)
    assert set(draws) == {1197, 1198, 1199}
    assert draws[1198] == {
        "draw": 1198, "date": "2017-08-07",
        "n1": 5, "n2": 11, "n3": 19, "n4": 23, "n5": 30, "n6": 41, "bonus": 8,
    }
    # 「第1199回」「2017年8月10日」の表記ゆれを吸収する
    assert draws[1199]["date"] == "2017-08-10"


def test_own_format_roundtrip():
    draws = fd.parse_rows(OWN)
    assert set(draws) == {1197}
    assert draws[1197]["n1"] == 3


def test_shift_jis_is_decoded():
    assert fd.decode(MIZUHO.encode("cp932")).splitlines()[0].startswith("回別")


def test_unsorted_main_numbers_are_sorted():
    text = "draw,date,n1,n2,n3,n4,n5,n6,bonus\n1,2000-10-05,30,2,13,8,27,10,39\n"
    d = fd.parse_rows(text)[1]
    assert [d[f"n{i}"] for i in range(1, 7)] == [2, 8, 10, 13, 27, 30]


def test_validation_errors():
    expect_error(
        "draw,date,n1,n2,n3,n4,n5,n6,bonus\n1,2026-01-01,1,1,3,4,5,6,7\n", "重複"
    )
    expect_error(
        "draw,date,n1,n2,n3,n4,n5,n6,bonus\n1,2026-01-01,1,2,3,4,5,44,7\n", "範囲外"
    )
    expect_error(
        "draw,date,n1,n2,n3,n4,n5,n6,bonus\n1,2026-01-01,1,2,3,4,5,6,6\n", "ボーナス数字が本数字と重複"
    )
    expect_error("a,b,c\n1,2,3\n", "列を判別できません")


def test_merge_adds_only_new_draws():
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / "draws.csv"
        target.write_text(OWN, encoding="utf-8")
        existing = fd.load_existing(target)
        incoming = fd.parse_rows(MIZUHO)
        assert sorted(set(incoming) - set(existing)) == [1198, 1199]


def main() -> int:
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
        print(f"ok  {t.__name__}")
    print(f"{len(tests)} tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())
