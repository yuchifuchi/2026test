"""sql.asp の SQL を 1 本ずつ実行して、期待値と突き合わせる（手順書 9-2）。

・SQL は sql.asp から読み取る（検査側に書き写さない）。Sql で始まる Function を模擬の VBScript で
  実行して、ASP が実際に組み立てる文字列そのものを得る。
・検査用のデータは、この検査が直接入れる（画面の作りに頼らない）。期待値も、そのデータから
  Python で別に数える。
・sql.asp の Sql 関数のうち、ここで試していないものが 1 つでもあれば失敗にする。
"""
import datetime
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import bridge, Fail, ROOT, DIST  # noqa: E402
from vbsim import aspfile, sqlcheck  # noqa: E402
from vbsim.parser import parse  # noqa: E402
from vbsim.interp import Interp  # noqa: E402
from vbsim.builtins import Builtins  # noqa: E402

WORK = os.path.join(ROOT, "build", "sqltest")
D = datetime.datetime


def iso(d):
    return {"t": "date", "v": d.isoformat()}


def i(n):
    return {"t": "int", "v": n}


def s(x):
    return {"t": "str", "v": x}


def b(x):
    return {"t": "bool", "v": x}


NULL = {"t": "null", "v": None}


class DummyHost:
    def intrinsic(self, low):
        return None

    def create_object(self, p):
        raise Fail("sql.asp の中で CreateObject を使っています")


def load_sql_functions():
    path = os.path.join(DIST, "part", "include", "sql.asp")
    prog = aspfile.load_page(path, os.path.dirname(path))
    ast = parse(prog.code)
    it = Interp(prog, ast, DummyHost(), None)
    it.bi = Builtins(DummyHost())
    it.prepare()
    out = {}
    for low, proc in it.procs.items():
        if low.startswith("sql"):
            if proc.params:
                raise Fail(f"{proc.name} に引数があります（Sql 関数は引数なしの決まり）")
            out[proc.name] = it.invoke(proc, [], it.g)
    return out


# ---------------- 検査用のデータ ----------------
PEOPLE = [  # 担当者ID, コード, 姓, 名, 氏名, 職員区分, 表示順, 有効, ログオン名
    (1, "001", "山田", "花子", "山田　花子", "パート", 20, True, None),
    (2, "002", "佐藤", "恵", "佐藤　恵", "パート", 10, True, None),
    (3, "003", "鈴木", "一郎", "鈴木　一郎", "職員", 30, True, "suzuki"),
    (4, "004", "田中", "美咲", "田中　美咲", "パート", 40, True, None),
    (5, "005", "高橋", "", "高橋", "職員", 50, False, "takahashi"),
]
DAY = D(2026, 8, 25)
NEXT = D(2026, 8, 26)
# (日時, 担当者ID, 区分ID, 製品ID, 件数)
CALLS = [
    (DAY, 1, 1, 1, 5), (DAY, 1, 1, 2, 3), (DAY, 1, 3, 3, 2), (DAY, 1, 11, 0, 2),
    (DAY, 2, 4, 1, 1), (DAY, 2, 8, 2, 1), (DAY, 2, 6, 1, 1), (DAY, 2, 14, 0, 1),
    (D(2026, 8, 25, 15, 30), 2, 15, 0, 1),          # 時刻が入っていても、その日に数える
    (DAY, 3, 2, 4, 1), (DAY, 3, 12, 0, 1), (DAY, 3, 17, 0, 1),
    (NEXT, 1, 1, 1, 7), (NEXT, 4, 9, 5, 2),          # 翌日の分は 8/25 に混ざらない
    (D(2026, 8, 24), 4, 10, 0, 3),                   # 前の日
]
TASKS = [(DAY, 1, 1, 10), (DAY, 1, 3, 4), (DAY, 2, 1, 5), (DAY, 2, 13, 2), (NEXT, 1, 1, 9)]
ATT = [(DAY, 1, "9:00-13:00", None), (DAY, 2, None, "午後のみ"), (DAY, 3, None, None), (NEXT, 1, None, None)]
DAILY = [(D(2026, 8, 24), 4, "確定"), (DAY, 5, "作成中")]


def fixture():
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    os.makedirs(WORK)
    db = os.path.join(WORK, "fixture.accdb")
    shutil.copyfile(os.path.join(ROOT, "build", "db", "日報集計_be.accdb"), db)
    br = bridge()
    br.close()
    br.open(db)
    for p in PEOPLE:
        br.execute("INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[職員区分],[表示順],[有効],[ログオン名]) VALUES (?,?,?,?,?,?,?,?,?)",
                   [i(p[0]), s(p[1]), s(p[2]), s(p[3]) if p[3] else NULL, s(p[4]), s(p[5]), i(p[6]), b(p[7]), s(p[8]) if p[8] else NULL])
    for dt, pid, k, pr, n in CALLS:
        br.execute("INSERT INTO [T_受電] ([対象日],[担当者ID],[区分ID],[製品ID],[件数],[更新日時]) VALUES (?,?,?,?,?,?)",
                   [iso(dt), i(pid), i(k), i(pr), i(n), iso(D(2026, 8, 25, 9, 0, 0))])
    for dt, pid, g, n in TASKS:
        br.execute("INSERT INTO [T_業務実績] ([対象日],[担当者ID],[業務項目ID],[件数]) VALUES (?,?,?,?)", [iso(dt), i(pid), i(g), i(n)])
    for dt, pid, h, note in ATT:
        br.execute("INSERT INTO [T_出勤] ([対象日],[担当者ID],[勤務時間],[備考]) VALUES (?,?,?,?)",
                   [iso(dt), i(pid), s(h) if h else NULL, s(note) if note else NULL])
    for dt, lines, st in DAILY:
        br.execute("INSERT INTO [T_日報] ([対象日],[回線数],[状態]) VALUES (?,?,?)", [iso(dt), i(lines), s(st)])
    # その他業務 ⑬ を使わない項目にしておく（数が残っていれば日報に出す、の確かめ）
    br.execute("UPDATE [M_業務項目] SET [有効] = ? WHERE [業務項目ID] = ?", [b(False), i(13)])
    return br


# ---------------- 期待値（Python で別に数える） ----------------
KUBUN_COL = {1: 1, 2: 1, 3: 2, 4: 3, 5: 3, 6: 3, 7: 4, 8: 4, 9: 5, 10: 1, 11: 1, 12: 1, 13: 5, 14: 5, 15: 5, 16: 5, 17: 5}
KUBUN_UCHI = {6: "返金", 8: "交換"}
STAFF = {p[0] for p in PEOPLE if p[5] == "職員"}


def calls_on(day, pid=None):
    return [c for c in CALLS if c[0].date() == day.date() and (pid is None or c[1] == pid)]


def col_totals(rows):
    out = {c: [0, 0] for c in range(1, 6)}
    for _, pid, k, _, n in rows:
        out[KUBUN_COL[k]][0] += n
        if pid in STAFF:
            out[KUBUN_COL[k]][1] += n
    return out


def uchi(rows):
    out = {}
    for _, _, k, _, n in rows:
        if k in KUBUN_UCHI:
            out[KUBUN_UCHI[k]] = out.get(KUBUN_UCHI[k], 0) + n
    return out


def run():
    sqls = load_sql_functions()
    br = fixture()
    tested = set()
    results = []

    def q(name, params):
        sql = sqls[name]
        probs = sqlcheck.check(sql)
        if probs:
            raise Fail(f"{name}: ACE で使えない書き方: {probs}")
        tested.add(name)
        cols, rows = br.query(sql, params)
        return [c["name"] for c in cols], rows

    def x(name, params):
        sql = sqls[name]
        probs = sqlcheck.check(sql)
        if probs:
            raise Fail(f"{name}: ACE で使えない書き方: {probs}")
        tested.add(name)
        return br.execute(sql, params)

    def eq(name, got, want):
        if got != want:
            raise Fail(f"{name}: 期待 {want!r} / 実際 {got!r}")
        results.append(name)

    day = [iso(DAY), iso(NEXT)]
    week = [iso(D(2026, 8, 24)), iso(D(2026, 8, 31))]

    # ---- 数の出どころ ----
    cols, rows = q("SqlCallsBase", day)
    eq("SqlCallsBase（行数・件数）", (len(rows), sum(r[cols.index("件数")] for r in rows)),
       (len(calls_on(DAY)), sum(c[4] for c in calls_on(DAY))))
    staff_sum = sum(r[cols.index("職員件数")] for r in rows)
    eq("SqlCallsBase（職員件数）", staff_sum, sum(c[4] for c in calls_on(DAY) if c[1] in STAFF))
    cols, rows = q("SqlCallsBaseOnePerson", day + [i(2)])
    eq("SqlCallsBaseOnePerson", sum(r[cols.index("件数")] for r in rows), sum(c[4] for c in calls_on(DAY, 2)))

    cols, rows = q("SqlColumnTotals", day)
    got = {r[0]: [int(r[3]), int(r[4])] for r in rows}
    eq("SqlColumnTotals（5 列・職員）", got, col_totals(calls_on(DAY)))
    eq("SqlColumnTotals（並び順）", [r[0] for r in rows], [1, 2, 3, 4, 5])
    total = sum(v[0] for v in got.values())
    eq("例の期待値：合計 20／申込 12／うち職員 2", (total, got[1][0], got[1][1]), (20, 12, 2))
    for pid in (1, 2, 3, 4):
        cols, rows = q("SqlColumnTotalsOnePerson", day + [i(pid)])
        eq(f"SqlColumnTotalsOnePerson（{pid}）", {r[0]: [int(r[3]), int(r[4])] for r in rows}, col_totals(calls_on(DAY, pid)))
    cols, rows = q("SqlBreakdownTotals", day)
    eq("SqlBreakdownTotals", {r[0]: int(r[1]) for r in rows}, uchi(calls_on(DAY)))
    cols, rows = q("SqlBreakdownTotalsOnePerson", day + [i(2)])
    eq("SqlBreakdownTotalsOnePerson", {r[0]: int(r[1]) for r in rows}, uchi(calls_on(DAY, 2)))
    q("SqlColTotalsHead", None) if False else None

    # ---- 週の集計表 ----
    wk = [c for c in CALLS if D(2026, 8, 24) <= c[0] < D(2026, 8, 31)]
    cols, rows = q("SqlDayColumnTotals", week)
    got = {}
    for r in rows:
        if (r[0][:10], r[1]) in got:
            raise Fail("SqlDayColumnTotals: 同じ日・同じ列が 2 行に割れています（時刻つきの行）")
        got[(r[0][:10], r[1])] = (int(r[2]), int(r[3]))
    eq("SqlDayColumnTotals：時刻のある行も日でまとまる", rows[0][0][10:], "T00:00")
    want = {}
    for dt, pid, k, _, n in wk:
        key = (dt.date().isoformat(), KUBUN_COL[k])
        a, st = want.get(key, (0, 0))
        want[key] = (a + n, st + (n if pid in STAFF else 0))
    eq("SqlDayColumnTotals", got, want)
    cols, rows = q("SqlDayBreakdownTotals", week)
    want = {}
    for dt, _, k, _, n in wk:
        if k in KUBUN_UCHI:
            want[(dt.date().isoformat(), KUBUN_UCHI[k])] = want.get((dt.date().isoformat(), KUBUN_UCHI[k]), 0) + n
    eq("SqlDayBreakdownTotals", {(r[0][:10], r[1]): int(r[2]) for r in rows}, want)
    cols, rows = q("SqlDayKubunTotals", week)
    want = {}
    for dt, _, k, _, n in wk:
        want[(dt.date().isoformat(), k)] = want.get((dt.date().isoformat(), k), 0) + n
    eq("SqlDayKubunTotals", {(r[0][:10], r[1]): int(r[2]) for r in rows}, want)
    cols, rows = q("SqlDayPersonTotals", week)
    want = {}
    for dt, pid, _, _, n in wk:
        want[(dt.date().isoformat(), pid)] = want.get((dt.date().isoformat(), pid), 0) + n
    eq("SqlDayPersonTotals", {(r[0][:10], r[1]): int(r[2]) for r in rows}, want)
    cols, rows = q("SqlPersonTotals", day)
    want = {}
    for _, pid, _, _, n in calls_on(DAY):
        want[pid] = want.get(pid, 0) + n
    eq("SqlPersonTotals", {r[0]: int(r[1]) for r in rows}, want)
    _, r1 = q("SqlRawCallsCheck", day)
    _, r2 = q("SqlBaseCallsCheck", day)
    eq("SqlRawCallsCheck = SqlBaseCallsCheck", [int(v) for v in r1[0]], [int(v) for v in r2[0]])
    eq("SqlRawCallsCheck（行数・件数）", [int(v) for v in r1[0]], [len(calls_on(DAY)), sum(c[4] for c in calls_on(DAY))])

    # ---- 受付入力 ----
    cols, rows = q("SqlCallRowsOnePerson", day + [i(1)])
    eq("SqlCallRowsOnePerson", sorted((r[1], r[2], r[3]) for r in rows), sorted((c[2], c[3], c[4]) for c in calls_on(DAY, 1)))
    cols, rows = q("SqlCallVersion", day + [i(1)])
    eq("SqlCallVersion", (int(rows[0][0]), int(rows[0][1]), rows[0][2][:16]), (4, 12, "2026-08-25T09:00"))
    x("SqlCallInsert", [iso(DAY), i(4), i(16), i(0), i(6), iso(DAY), iso(DAY), s("検査")])
    _, rows = br.query("SELECT J.[受電ID], J.[件数] FROM [T_受電] AS J WHERE J.[担当者ID] = 4 AND J.[区分ID] = 16", [])
    eq("SqlCallInsert", rows[0][1], 6)
    new_id = rows[0][0]
    x("SqlCallUpdate", [i(8), iso(DAY), s("検査"), i(new_id)])
    _, rows = br.query("SELECT J.[件数] FROM [T_受電] AS J WHERE J.[受電ID] = ?", [i(new_id)])
    eq("SqlCallUpdate", rows, [[8]])
    n = x("SqlCallDelete", [i(new_id)])
    eq("SqlCallDelete（1 行）", n, 1)
    try:
        x("SqlCallInsert", [iso(DAY), i(1), i(1), i(1), i(1), iso(DAY), iso(DAY), s("検査")])
        raise Fail("同じ（日付・担当者・区分・製品）の行を 2 つ入れられてしまいました（UNIQUE が効いていない）")
    except Exception as e:
        if isinstance(e, Fail):
            raise
        results.append("UQ_受電：同じ欄の二重登録は止まる")
    try:
        x("SqlCallInsert", [iso(DAY), i(99), i(1), i(1), i(1), iso(DAY), iso(DAY), s("検査")])
        raise Fail("いない担当者の行を入れられてしまいました（参照整合性が効いていない）")
    except Exception as e:
        if isinstance(e, Fail):
            raise
        results.append("参照整合性：いない担当者の行は入らない")

    # ---- その他業務 ----
    cols, rows = q("SqlTaskRowsOnePerson", day + [i(1)])
    eq("SqlTaskRowsOnePerson", sorted((r[1], r[2]) for r in rows), [(1, 10), (3, 4)])
    cols, rows = q("SqlTaskTotals", day)
    got = {r[0]: (int(r[6]), r[5]) for r in rows}
    eq("SqlTaskTotals（①）", got[1], (15, True))
    eq("SqlTaskTotals（使っていない ⑬ でも数は出る）", got[13], (2, False))
    eq("SqlTaskTotals（行数・並び）", [r[0] for r in rows], list(range(1, 14)))
    cols, rows = q("SqlTaskPersonTotals", day)
    eq("SqlTaskPersonTotals", {r[0]: int(r[1]) for r in rows}, {1: 14, 2: 7})
    x("SqlTaskInsert", [iso(DAY), i(3), i(5), i(2)])
    _, rows = br.query("SELECT R.[実績ID] FROM [T_業務実績] AS R WHERE R.[担当者ID] = 3", [])
    tid = rows[0][0]
    x("SqlTaskUpdate", [i(9), i(tid)])
    _, rows = br.query("SELECT R.[件数] FROM [T_業務実績] AS R WHERE R.[実績ID] = ?", [i(tid)])
    eq("SqlTaskUpdate", rows, [[9]])
    eq("SqlTaskDelete", x("SqlTaskDelete", [i(tid)]), 1)

    # ---- 出勤 ----
    cols, rows = q("SqlAttendance", day)
    eq("SqlAttendance（表示順で並ぶ・3 人）", [r[1] for r in rows], [2, 1, 3])
    eq("SqlAttendance（勤務時間・備考）", (rows[1][2], rows[0][3]), ("9:00-13:00", "午後のみ"))
    eq("出勤者数 3", len(rows), 3)
    eq("SqlAttendanceDeleteDay", x("SqlAttendanceDeleteDay", day), 3)
    x("SqlAttendanceInsert", [iso(DAY), i(4), s("13:00-17:00"), NULL])
    _, rows = q("SqlAttendance", day)
    eq("SqlAttendanceInsert", [(r[1], r[2]) for r in rows], [(4, "13:00-17:00")])

    # ---- 日報 ----
    cols, rows = q("SqlDaily", day)
    eq("SqlDaily", (rows[0][1], rows[0][5]), (5, "作成中"))
    _, rows = q("SqlLastLines", [iso(DAY)])
    eq("SqlLastLines（前の日までで一番新しい回線数）", int(rows[0][0]), 4)
    x("SqlDailyInsert", [iso(NEXT), i(6), s("特記"), NULL, NULL, s("作成中"), NULL, iso(NEXT)])
    x("SqlDailyUpdate", [i(7), s("特記2"), s("代替"), s("要望"), s("確定"), iso(NEXT), iso(NEXT), iso(NEXT), iso(D(2026, 8, 27))])
    _, rows = q("SqlDaily", [iso(NEXT), iso(D(2026, 8, 27))])
    eq("SqlDailyInsert / SqlDailyUpdate", (rows[0][1], rows[0][2], rows[0][5]), (7, "特記2", "確定"))
    x("SqlDailySetState", [s("作成中"), NULL, iso(NEXT), iso(NEXT), iso(D(2026, 8, 27))])
    _, rows = q("SqlDaily", [iso(NEXT), iso(D(2026, 8, 27))])
    eq("SqlDailySetState", (rows[0][5], rows[0][6]), ("作成中", None))

    # ---- マスタ ----
    cols, rows = q("SqlStaffAll", [])
    eq("SqlStaffAll（表示順）", [r[0] for r in rows], [2, 1, 3, 4, 5])
    _, rows = q("SqlStaffCount", [])
    eq("SqlStaffCount", int(rows[0][0]), 5)
    _, rows = q("SqlStaffByLogon", [s("suzuki")])
    eq("SqlStaffByLogon（有効な職員）", int(rows[0][0]), 1)
    _, rows = q("SqlStaffByLogon", [s("takahashi")])
    eq("SqlStaffByLogon（有効でない職員は数えない）", int(rows[0][0]), 0)
    _, rows = q("SqlStaffMaxId", [])
    eq("SqlStaffMaxId", int(rows[0][0]), 5)
    x("SqlStaffInsert", [i(6), s("006"), s("伊藤"), NULL, s("伊藤"), NULL, NULL, s("パート"), iso(DAY), NULL, i(60), b(True), NULL])
    x("SqlStaffUpdate", [s("006"), s("伊藤"), s("さくら"), s("伊藤　さくら"), NULL, NULL, s("パート"), iso(DAY), iso(NEXT), i(60), b(True), s("備考"), i(6)])
    _, rows = br.query("SELECT T.[氏名], T.[在籍終了日] FROM [M_担当者] AS T WHERE T.[担当者ID] = 6", [])
    eq("SqlStaffInsert / SqlStaffUpdate", (rows[0][0], rows[0][1][:10]), ("伊藤　さくら", "2026-08-26"))
    _, rows = q("SqlColumnsAll", [])
    eq("SqlColumnsAll", [r[1] for r in rows], ["申込", "抽選", "払込用紙", "商品発送", "その他"])
    x("SqlColumnRename", [s("申込み"), i(1)])
    _, rows = q("SqlColumnsAll", [])
    eq("SqlColumnRename", rows[0][1], "申込み")
    _, rows = q("SqlBlocksAll", [])
    eq("SqlBlocksAll", [(r[0], r[2]) for r in rows], [(1, True), (2, False), (3, False)])
    _, rows = q("SqlBlockMaxId", [])
    eq("SqlBlockMaxId", int(rows[0][0]), 3)
    x("SqlBlockInsert", [i(4), s("新しいまとまり"), b(False), i(4)])
    x("SqlBlockUpdate", [s("新しいまとまり2"), b(True), i(4), i(4)])
    _, rows = q("SqlBlocksAll", [])
    eq("SqlBlockInsert / SqlBlockUpdate", (rows[3][1], rows[3][2]), ("新しいまとまり2", True))
    _, rows = q("SqlBlockProductRows", [i(1)])
    eq("SqlBlockProductRows（ブロック 1 の製品つきの行）", int(rows[0][0]), len([c for c in CALLS if c[3] != 0 and c[2] <= 9]))
    _, rows = q("SqlKubunAll", [])
    eq("SqlKubunAll", len(rows), 17)
    _, rows = q("SqlKubunMaxId", [])
    eq("SqlKubunMaxId", int(rows[0][0]), 17)
    x("SqlKubunInsert", [i(18), i(4), s("新しい区分"), i(5), NULL, i(1), b(True), s("旧")])
    x("SqlKubunUpdate", [i(4), s("新しい区分2"), i(2), s("交換"), i(1), b(False), NULL, i(18)])
    _, rows = br.query("SELECT K.[区分名], K.[集計列ID], K.[内訳区分], K.[有効] FROM [M_区分] AS K WHERE K.[区分ID] = 18", [])
    eq("SqlKubunInsert / SqlKubunUpdate", rows[0], ["新しい区分2", 2, "交換", False])
    _, rows = q("SqlKubunUseCount", [i(1)])
    eq("SqlKubunUseCount", int(rows[0][0]), len([c for c in CALLS if c[2] == 1]))
    _, rows = q("SqlProductsAll", [])
    eq("SqlProductsAll（製品ID 0 は出さない）", [r[0] for r in rows], [1, 2, 3, 4, 5])
    _, rows = q("SqlProductMaxId", [])
    eq("SqlProductMaxId", int(rows[0][0]), 5)
    x("SqlProductInsert", [i(6), i(1), s("新製品"), iso(DAY), NULL, i(6), b(True)])
    x("SqlProductUpdate", [i(1), s("新製品2"), iso(DAY), iso(NEXT), i(6), b(True), i(6)])
    _, rows = br.query("SELECT P.[製品名], P.[適用終了日] FROM [M_製品] AS P WHERE P.[製品ID] = 6", [])
    eq("SqlProductInsert / SqlProductUpdate", (rows[0][0], rows[0][1][:10]), ("新製品2", "2026-08-26"))
    _, rows = q("SqlProductUseCount", [i(1)])
    eq("SqlProductUseCount", int(rows[0][0]), len([c for c in CALLS if c[3] == 1]))
    _, rows = q("SqlTaskItemsAll", [])
    eq("SqlTaskItemsAll", len(rows), 13)
    _, rows = q("SqlTaskItemMaxId", [])
    eq("SqlTaskItemMaxId", int(rows[0][0]), 13)
    x("SqlTaskItemInsert", [i(14), s("⑭"), s("新しい業務"), s("新しい業務"), i(14), b(True)])
    x("SqlTaskItemUpdate", [s("⑭"), s("新しい業務2"), s("新業務"), i(14), b(False), i(14)])
    _, rows = br.query("SELECT G.[項目名], G.[帳票表示名], G.[有効] FROM [M_業務項目] AS G WHERE G.[業務項目ID] = 14", [])
    eq("SqlTaskItemInsert / SqlTaskItemUpdate", rows[0], ["新しい業務2", "新業務", False])

    # 部品（ほかの Sql 関数の一部）として使うものは、それを含む関数を実行したことで試している
    parts = {"SqlColTotalsHead", "SqlColTotalsTail"}
    for p in parts:
        if not any(sqls[p] in sqls[t] for t in tested):
            raise Fail(f"{p} を使った SQL を 1 本も試していません")
        tested.add(p)
    untested = sorted(set(sqls) - tested)
    if untested:
        raise Fail("試していない SQL があります: " + ", ".join(untested))
    br.close()
    return results, len(sqls)


if __name__ == "__main__":
    try:
        res, n = run()
    except Fail as e:
        print("NG:", e)
        bridge().shutdown()
        sys.exit(1)
    for r in res:
        print("  OK:", r)
    print(f"OK: sql.asp の SQL {n} 本をすべて実行し、期待値と一致（{len(res)} 項目）")
    bridge().shutdown()
