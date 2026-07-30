# -*- coding: utf-8 -*-
"""formulas エンジンで全数式を実際に評価し、エラー値が出ないことと
記入例・サマリの計算結果が期待どおりかを確認する。"""
import sys
import formulas

ERRS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "#ERROR!")


def norm(v):
    try:
        import numpy as np
        if isinstance(v, np.ndarray):
            return norm(v.ravel()[0]) if v.size else None
    except Exception:
        pass
    return v


def main(path):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()

    cells = {}
    for k, v in sol.items():
        if "!" not in k or k.endswith("!"):
            continue
        ref = k.split("]", 1)[-1]
        cells[ref.upper()] = norm(v)

    bad = []
    for ref, v in cells.items():
        s = str(v)
        for e in ERRS:
            if e in s:
                bad.append((ref, s[:80]))
                break

    print(f"evaluated cells: {len(cells)}")
    if bad:
        print(f"=== FORMULA ERRORS: {len(bad)} ===")
        for ref, s in bad[:60]:
            print("  ", ref, "->", s)
        return 1
    print("エラー値(#REF!/#VALUE!/#NAME? 等): 0 件")

    def g(sheet, ref):
        # formulas のキーは "'[file.xlsx]シート名'!A1" 形式。"]" で切ると "シート名'!A1"
        return cells.get(f"{sheet}'!{ref}".upper())

    print("\n--- 03_稟議用サマリ ---")
    for r in range(6, 10):
        row = [g("03_稟議用サマリ", f"{c}{r}") for c in "CDEFGHI"]
        print("  ", row)

    print("\n--- 04_記入例 集計 ---")
    for lbl, r in [("Must判定", 56), ("加重総合点", 57), ("5年TCO", 58)]:
        print("  ", lbl, [g("04_記入例", f"{c}{r}") for c in "FGHI"])
    print("   大分類別:")
    for r in range(60, 66):
        print("     ", g("04_記入例", f"C{r}"),
              [g("04_記入例", f"{c}{r}") for c in "FGHI"])

    print("\n--- 01_重み付け設定 合計 ---")
    print("  ", g("01_重み付け設定", "C11"), "|", g("01_重み付け設定", "D11"))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
