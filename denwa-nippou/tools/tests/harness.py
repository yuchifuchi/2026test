"""検査で使う共通の道具：組み立てた nippou を模擬の IIS に置き、SQL 中継役につなぐ。"""
import datetime
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
ROOT = os.path.dirname(TOOLS)
sys.path.insert(0, TOOLS)

from vbsim.sim import AspSim  # noqa: E402
from vbsim.sqlbridge import Bridge  # noqa: E402
from vbsim import builtins as B  # noqa: E402
from vbsim import sqlcheck  # noqa: E402

DIST = os.path.join(ROOT, "build", "dist", "nippou")
WORK = os.path.join(ROOT, "build", "simwork")

_bridge = None


def bridge():
    global _bridge
    if _bridge is None:
        _bridge = Bridge()
    return _bridge


class Fail(Exception):
    pass


def new_sim(name="sim", today=datetime.datetime(2026, 8, 25, 10, 0, 0), **kw):
    """毎回まっさらな .accdb（出荷するもの）で始める。"""
    b = bridge()
    b.close()
    sim = AspSim(os.path.join(WORK, name), DIST, b, **kw)
    B.NOW_OVERRIDE[0] = today
    seen = []

    def hook(sql):
        probs = sqlcheck.check(sql)
        if probs:
            raise Fail("ACE で使えない SQL: " + " / ".join(probs) + "\n" + sql)
        seen.append(sql)
    sim.sql_hook = hook
    sim.seen_sql = seen
    return sim


def get(sim, url, **kw):
    r = sim.request("GET", url, **kw)
    check_clean(r)
    return r


def post(sim, url, form, **kw):
    r = sim.request("POST", url, form=form, **kw)
    check_clean(r)
    return r


def check_clean(r, allow_error_page=False):
    if r.strict:
        raise Fail(f"{r.url}: 厳格モードの検出: {r.strict}")
    if r.warnings:
        raise Fail(f"{r.url}: 警告: {r.warnings}")
    if r.error is not None and not allow_error_page:
        raise Fail(f"{r.url}: 実行時エラー: {r.error}\n{r.text[:2000]}")
    if r.code == 200 and ("NIPPOU-FATAL" in r.text or "NIPPOU-ERROR" in r.text) and not allow_error_page:
        raise Fail(f"{r.url}: エラーの説明画面が出ました:\n{strip_tags(r.text)[:1500]}")


def strip_tags(html):
    t = re.sub(r"<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", t)


def follow(sim, r, **kw):
    """302 なら移動先を開く。"""
    n = 0
    while r.code in (301, 302) and n < 5:
        loc = r.headers["Location"]
        if not loc.startswith("/"):
            base = r.url.split("?")[0].rsplit("/", 1)[0]
            loc = base + "/" + loc
        r = get(sim, loc, **kw)
        n += 1
    return r


def add_people(sim, people):
    """マスタ保守の画面から担当者を登録する（画面を通すことで、その画面も検査になる）。"""
    for p in people:
        form = {"id": "new", "code": p["code"], "sei": p["sei"], "mei": p.get("mei", ""), "shimei": "",
                "kana": "", "logon": "", "kind": p["kind"], "start": p.get("start", "2026-04-01"),
                "end": p.get("end", ""), "order": "", "active": "1", "note": ""}
        r = post(sim, "/nippou/staff/master.asp?t=tanto", form)
        if r.code != 302:
            raise Fail("担当者の登録に失敗: " + strip_tags(r.text)[:800])


def ids(sim):
    b = sim.bridge
    cols, rows = b.query("SELECT T.[担当者ID], T.[姓] FROM [M_担当者] AS T ORDER BY T.[担当者ID]", [])
    return {r[1]: r[0] for r in rows}
