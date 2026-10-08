"""帳票（sheet.asp の SheetHtml）を、本物と同じ道筋で HTML と PDF にして、見た目を機械的に確かめる。

  python3 tools/render_sheet.py        → build/render/ に 場面ごとの .html / .pdf / .png を作り、検査する

やっていること:
  1. 模擬の IIS の上で、画面（マスタ保守・受付入力・その他業務・日報）からデータを入れる
  2. printpdf.asp を開く。これが pdf.asp → sheet.asp を動かし、Edge の代わりに Chromium で PDF を作る
     （Edge に渡すのと同じコマンド行。HTML は検査側で書き写さず、sheet.asp が組み立てたものそのもの）
  3. その HTML を Chromium で開き、表ごとに
       ・枠の欠け・重なりが無いか（rowspan / colspan を数えて、すべての升目が 1 回ずつ埋まるか）
       ・すべての欄の上下左右に罫線があるか
       ・空の <td> が無いか（&nbsp; が入っているか）
       ・文字が欄からはみ出していないか
     を調べる
  4. PDF が A4 1 枚か（pdfinfo）
場面:
  normal … ふつうの日
  worst  … 帳票に出る文字を、画面で止めている上限いっぱいまで入れた日（14 人・長い氏名・記述欄いっぱい・大きい数）
  empty  … 何も入っていない日（空の枠も罫線が閉じていること）
"""
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "tests"))
from harness import new_sim, get, post, follow, add_people, ids, strip_tags, Fail, bridge, ROOT  # noqa: E402

OUT = os.path.join(ROOT, "build", "render")
CHROME = "/opt/pw-browsers/chromium"
DAY = "2026-08-25"

# 様式（課の印刷用シート）から測った寸法。画像の 1 ピクセル = 186/527 mm。
# まとまりの高さ（mm）：出勤者・問合せ件数・特記事項・職員に代わった案件・要望・電話対応以外の業務
FORM_BLOCK_MM = [45.7, 11.6, 28.6, 34.2, 39.9, 62.3]
# 枠の上端までの高さ（課内限り・回覧・押印欄・表題）
FORM_FRAME_TOP_MM = 41.6
# 縦線の位置（枠の左端からの mm）
FORM_ATT_X_MM = [39.5, 80.5, 121.1]                       # 出勤者：左の欄｜氏名｜氏名｜備考
FORM_CNT_X_MM = [39.5, 64.2, 88.6, 113.0, 137.3, 161.7]   # 問合せ件数：申込｜抽選｜払込用紙｜商品発送｜その他｜注記
FORM_STAMP_MM = [(0.0, 63.5), (101.0, 165.9)]             # 回覧の押印欄（左・右）の左端と右端
FORM_TOL_MM = 1.2

CHECK_JS = r"""
<script>
(function(){
  var out = {tables: [], problems: [], height_mm: 0, geo: {}};
  var mm = 96 / 25.4;
  var sheet = document.querySelector('.sheet');
  var sr = sheet.getBoundingClientRect();
  out.height_mm = sr.height / mm;
  function bw(cs, side) { var w = parseFloat(cs['border' + side + 'Width']), st = cs['border' + side + 'Style']; return (w > 0 && st !== 'none' && st !== 'hidden') ? st : ''; }
  var tables = sheet.querySelectorAll('table');
  for (var t = 0; t < tables.length; t++) {
    var tb = tables[t], grid = [], rows = tb.rows, isBlk = tb.className.indexOf('s-blk') >= 0, isStamp = tb.className.indexOf('s-stamp') >= 0;
    var tcs = getComputedStyle(tb);
    if (isBlk) ['Top', 'Right', 'Bottom', 'Left'].forEach(function (side) { if (!bw(tcs, side)) out.problems.push('外枠が閉じていない: 表' + t + ' ' + side); });
    for (var r = 0; r < rows.length; r++) {
      var c = 0;
      grid[r] = grid[r] || [];
      var ruled = rows[r].className.indexOf('rl') >= 0;
      for (var k = 0; k < rows[r].cells.length; k++) {
        var cell = rows[r].cells[k];
        while (grid[r][c]) c++;
        var c0 = c, r0 = r;
        for (var dr = 0; dr < cell.rowSpan; dr++) for (var dc = 0; dc < cell.colSpan; dc++) {
          grid[r + dr] = grid[r + dr] || [];
          if (grid[r + dr][c + dc]) out.problems.push('升目の重なり: 表' + t + ' 行' + (r + dr) + ' 列' + (c + dc));
          grid[r + dr][c + dc] = 1;
        }
        c += cell.colSpan;
        var cs = getComputedStyle(cell);
        if (isStamp || cell.className.split(' ').indexOf('ln') >= 0) {
          ['Top', 'Right', 'Bottom', 'Left'].forEach(function (side) {
            if (!bw(cs, side)) out.problems.push('罫線が無い: 表' + t + ' 行' + r0 + ' 欄' + k + ' ' + side);
          });
        }
        if (ruled && !bw(cs, 'Bottom')) out.problems.push('点線の罫線が無い: 表' + t + ' 行' + r0 + ' 欄' + k);
        if (cell.innerHTML.replace(/\s/g, '') === '') out.problems.push('空の欄（&nbsp; が無い）: 表' + t + ' 行' + r0 + ' 欄' + k);
        if (cell.scrollWidth > cell.clientWidth + 1) out.problems.push('文字のはみ出し: 表' + t + ' 行' + r0 + ' 「' + cell.textContent.slice(0, 20) + '」');
      }
    }
    var width = 0;
    for (var i = 0; i < grid.length; i++) width = Math.max(width, grid[i].length);
    for (var r2 = 0; r2 < grid.length; r2++) for (var c2 = 0; c2 < width; c2++) if (!grid[r2][c2]) out.problems.push('升目の欠け: 表' + t + ' 行' + r2 + ' 列' + c2);
    out.tables.push({rows: grid.length, cols: width});
  }
  // 様式との突き合わせに使う寸法
  var blks = sheet.querySelectorAll('table.s-blk');
  var f = blks[0].getBoundingClientRect();
  out.geo.frame_top = (f.top - sr.top) / mm;
  out.geo.blocks = [];
  for (var b = 0; b < blks.length; b++) out.geo.blocks.push(blks[b].getBoundingClientRect().height / mm);
  function xs(row) { var a = []; for (var q = 1; q < row.cells.length; q++) a.push((row.cells[q].getBoundingClientRect().left - f.left) / mm); return a; }
  out.geo.att_x = xs(blks[0].rows[0]);
  out.geo.cnt_x = xs(blks[1].rows[0]);
  out.geo.stamps = [];
  var st = sheet.querySelectorAll('table.s-stamp');
  for (var s2 = 0; s2 < st.length; s2++) { var rr = st[s2].getBoundingClientRect(); out.geo.stamps.push([(rr.left - f.left) / mm, (rr.right - f.left) / mm]); }
  var pre = document.createElement('pre'); pre.id = '__result'; pre.textContent = JSON.stringify(out);
  document.body.appendChild(pre);
})();
</script>
"""


def sheet_const(name):
    src = open(os.path.join(ROOT, "src", "include", "sheet.asp"), encoding="utf-8").read()
    return int(re.search(rf"Const {name} = (\d+)", src).group(1))


def worst_people():
    out = []
    for i in range(14):
        # 氏名は画面で止めている上限（SHEET_NAME_MAX = 9 文字）いっぱい
        out.append({"code": f"{i + 1:03d}", "sei": "勅使河原", "mei": "久美子旧" if i % 2 == 0 else "由紀子旧", "kind": "パート"})
    return out


def build(scenario):
    sim = new_sim("render_" + scenario)
    if scenario == "empty":
        return sim
    if scenario == "normal":
        import test_flow
        add_people(sim, test_flow.PEOPLE)
        pid = ids(sim)
        for name, cells in test_flow.CALLS.items():
            test_flow.entry(sim, "part", pid[name], cells)
        for name, items in test_flow.TASKS.items():
            form = {"act": "save", "d": DAY, "t": str(pid[name])}
            form.update({"g_" + k: str(v) for k, v in items.items()})
            post(sim, "/nippou/part/tasks.asp", form)
        form = {"act": "fix", "d": DAY, "lines": "5",
                "t1": "記念貨幣の申込開始日のため、午前中に申込方法の問合せが集中した。",
                "t2": "払込用紙の再発行の可否について、課長に引き継いだ（1 件）。", "t3": "申込サイトの入力例を増やしてほしいとの声あり。"}
        for name in ("山田", "佐藤", "鈴木"):
            form[f"att_{pid[name]}"] = "1"
        form[f"h_{pid['山田']}"] = "9:00-13:00"
        form[f"n_{pid['佐藤']}"] = "午後のみ"
        r = post(sim, "/nippou/staff/daily.asp", form)
        if r.code != 302:
            raise Fail("日報の確定に失敗: " + strip_tags(r.text)[:800])
        return sim
    # ---- worst ----
    import test_flow
    people = worst_people()
    add_people(sim, people)
    pid = sorted(ids_all(sim))
    # 列の名前・業務の名前を上限いっぱいに
    for cid, nm in enumerate(["申込受付件", "抽選結果照", "払込用紙再", "商品発送状", "その他全般"], start=1):
        r = post(sim, f"/nippou/staff/master.asp?t=col", {"id": str(cid), "name": nm})
        if r.code != 302:
            raise Fail("列の名前を直せません: " + strip_tags(r.text)[:500])
    for g in range(1, 14):
        r = post(sim, "/nippou/staff/master.asp?t=gyomu", {"id": str(g), "no": "⑬", "name": "長い業務の名前",
                                                           "label": ("長い業務の名前" * 3)[:sheet_const("SHEET_TASK_LABEL_MAX")], "order": str(g), "active": "1"})
        if r.code != 302:
            raise Fail("業務項目を直せません: " + strip_tags(r.text)[:500])
    for t in pid:
        test_flow.entry(sim, "part", t, {"1_1": 9999, "3_3": 9999, "4_1": 9999, "9_2": 9999, "16_0": 9999, "11_2": 9999, "50_0": 9999})
        form = {"act": "save", "d": DAY, "t": str(t)}
        form.update({f"g_{g}": "9999" for g in range(1, 14)})
        post(sim, "/nippou/part/tasks.asp", form)
    # 1 人目は、受付の表のすべての欄を 4 けたにし、特殊な問合せの内容も上限いっぱいに（個人別の受付表が 1 枚に収まるかを見る）
    r = sim.request("GET", f"/nippou/part/entry.asp?d={DAY}&t={pid[0]}")
    cells = {k: 9999 for k in re.findall(r'name="c_(\d+_\d+)"', r.text)}
    test_flow.entry(sim, "part", pid[0], cells)
    memo_max = int(re.search(r"Const MEMO_MAX = (\d+)", open(os.path.join(ROOT, "src", "pages", "common", "entry.asp"), encoding="utf-8").read()).group(1))
    r = sim.request("GET", f"/nippou/part/entry.asp?d={DAY}&t={pid[0]}")
    ver = re.search(r'name="ver" value="([^"]*)"', r.text).group(1)
    r = post(sim, "/nippou/part/entry.asp", {"act": "save", "d": DAY, "t": str(pid[0]), "ver": ver, "memo": "う" * memo_max})
    if r.code != 302:
        raise Fail("特殊な問合せの内容を保存できません: " + strip_tags(r.text)[:500])
    # 回覧の欄の名前も上限いっぱいに
    for k in range(1, 9):
        r = post(sim, "/nippou/staff/master.asp?t=stamp", {"id": str(k), "name": "藤本課長補佐"[:sheet_const("SHEET_STAMP_NAME_MAX")]})
        if r.code != 302:
            raise Fail("回覧の欄を直せません: " + strip_tags(r.text)[:500])
    # 記述欄を、帳票の罫線の数と 1 行の幅（sheet.asp の SHEET_*）いっぱいに
    wide = "あ" * (sheet_const("SHEET_LINE_WIDE") // 2)
    narrow = "い" * (sheet_const("SHEET_LINE_NARROW") // 2)
    t1 = [narrow] + [wide] * (sheet_const("SHEET_MEMO1_ROWS") - 1)
    t2 = [wide] * sheet_const("SHEET_MEMO2_ROWS")
    t3 = [wide] * sheet_const("SHEET_MEMO3_ROWS")
    form = {"act": "fix", "d": DAY, "lines": "99", "t1": "\r\n".join(t1), "t2": "\r\n".join(t2), "t3": "\r\n".join(t3)}
    note_max = att_note_max()
    for n, t in enumerate(pid):
        form[f"att_{t}"] = "1"
        if n < sheet_const("SHEET_NOTE_ROWS"):
            # 備考の欄の行数いっぱいの人数に、勤務時間＋備考を上限の文字数まで
            form[f"h_{t}"] = "9:00-13:00"[:note_max]
            form[f"n_{t}"] = "早退あり予定"[:max(0, note_max - len(form[f"h_{t}"]))]
    r = post(sim, "/nippou/staff/daily.asp", form)
    if r.code != 302:
        raise Fail("日報（上限いっぱい）の確定に失敗: " + strip_tags(r.text)[:800])
    return sim


def ids_all(sim):
    _, rows = sim.bridge.query("SELECT T.[担当者ID] FROM [M_担当者] AS T", [])
    return [r[0] for r in rows]


def att_note_max():
    src = open(os.path.join(ROOT, "src", "pages", "staff", "daily.asp"), encoding="utf-8").read()
    return int(re.search(r"Const ATT_NOTE_MAX = (\d+)", src).group(1))


def pdf_pages(path):
    r = subprocess.run(["pdfinfo", path], capture_output=True, text=True)
    m = re.search(r"Pages:\s+(\d+)", r.stdout)
    size = re.search(r"Page size:\s+([\d.]+) x ([\d.]+) pts", r.stdout)
    return int(m.group(1)), (float(size.group(1)), float(size.group(2)))


def check_html(html_path):
    html = open(html_path, encoding="utf-8-sig").read()
    probe = html_path.replace(".html", "_check.html")
    open(probe, "w", encoding="utf-8").write(html.replace("</body>", CHECK_JS + "</body>"))
    r = subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox", "--virtual-time-budget=5000",
                        "--dump-dom", "file://" + probe], capture_output=True, text=True, timeout=120)
    m = re.search(r'<pre id="__result">(.*?)</pre>', r.stdout, re.S)
    if not m:
        raise Fail("検査用のスクリプトが結果を返しませんでした: " + r.stderr[-500:])
    return json.loads(m.group(1).replace("&quot;", '"').replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">"))


PSHEET_JS = r"""
<script>
(function(){
  var out = {sheets: 0, problems: [], heights: []};
  var mm = 96 / 25.4;
  var sheets = document.querySelectorAll('.psheet');
  out.sheets = sheets.length;
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    out.heights.push(sh.getBoundingClientRect().height / mm);
    if (sh.scrollWidth > sh.clientWidth + 1) out.problems.push((i + 1) + ' 人目: 横にはみ出しています');
    var cells = sh.querySelectorAll('th, td');
    for (var j = 0; j < cells.length; j++) {
      var c = cells[j];
      if (c.scrollWidth > c.clientWidth + 1) { out.problems.push((i + 1) + ' 人目: 欄の文字がはみ出しています「' + c.textContent.slice(0, 20) + '」'); break; }
      var cs = getComputedStyle(c);
      if (cs.borderTopStyle === 'none' || cs.borderLeftStyle === 'none') { out.problems.push((i + 1) + ' 人目: 罫線の無い欄があります「' + c.textContent.slice(0, 20) + '」'); break; }
    }
  }
  var pre = document.createElement('pre'); pre.id = '__result'; pre.textContent = JSON.stringify(out); document.body.appendChild(pre);
})();
</script>
"""


def check_psheets(html_path):
    html = open(html_path, encoding="utf-8-sig").read()
    probe = html_path.replace(".html", "_check.html")
    open(probe, "w", encoding="utf-8").write(html.replace("</body>", PSHEET_JS + "</body>"))
    r = subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox", "--virtual-time-budget=5000",
                        "--dump-dom", "file://" + probe], capture_output=True, text=True, timeout=120)
    m = re.search(r'<pre id="__result">(.*?)</pre>', r.stdout, re.S)
    if not m:
        raise Fail("検査用のスクリプトが結果を返しませんでした: " + r.stderr[-500:])
    return json.loads(m.group(1).replace("&quot;", '"').replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">"))


def detail_pdf(sim, scenario):
    """個人別の受付表（1 人 1 枚）。入力のあった人の数だけ、A4 が 1 枚ずつ"""
    probs = []
    r = sim.request("GET", f"/nippou/staff/printpdf.asp?d={DAY}&kind=detail")
    if r.error or r.strict or r.warnings:
        raise Fail(f"{scenario}: 個人別の受付表の PDF で問題: {r.error} {r.strict} {r.warnings}")
    if scenario == "empty":
        if r.body.startswith(b"%PDF") or "入力のあった人がいない" not in r.text:
            probs.append("入力の無い日に、個人別の受付表の PDF を作ろうとしました")
        return 0, 0, [], probs
    if not r.body.startswith(b"%PDF"):
        raise Fail(f"{scenario}: 個人別の受付表の PDF ができませんでした: " + strip_tags(r.text)[:800])
    pdf = os.path.join(OUT, f"{scenario}_detail.pdf")
    open(pdf, "wb").write(r.body)
    work = sim.winfs.to_local("C:\\inetpub\\wwwroot\\nippou\\data\\pdfwork")
    jobs = sorted(os.listdir(work), key=lambda j: os.path.getmtime(os.path.join(work, j)))   # 一番新しい作業
    html = os.path.join(OUT, f"{scenario}_detail.html")
    shutil.copyfile(os.path.join(work, jobs[-1], "sheet.html"), html)
    res = check_psheets(html)
    pages, size = pdf_pages(pdf)
    persons = len(people_with_input(sim))
    probs += res["problems"]
    if res["sheets"] != persons:
        probs.append(f"個人別の受付表が {res['sheets']} 人ぶんです（入力のあった人は {persons} 人）")
    if pages != persons:
        probs.append(f"個人別の受付表の PDF が {pages} 枚です（1 人 1 枚で {persons} 枚のはず。1 人ぶんが 2 枚に分かれていないか）")
    if abs(size[0] - 595.3) > 2 or abs(size[1] - 841.9) > 2:
        probs.append(f"個人別の受付表の用紙が A4 縦ではありません（{size}）")
    subprocess.run(["pdftoppm", "-r", "80", "-png", "-f", "1", "-l", "1", "-singlefile", pdf, os.path.join(OUT, scenario + "_detail")], check=True)
    return persons, pages, res["heights"], probs


def people_with_input(sim):
    _, rows = sim.bridge.query("SELECT J.[担当者ID] FROM [T_受電] AS J", [])
    _, rows2 = sim.bridge.query("SELECT R.[担当者ID] FROM [T_業務実績] AS R", [])
    _, rows3 = sim.bridge.query("SELECT M.[担当者ID] FROM [T_受付メモ] AS M", [])
    return {r[0] for r in rows + rows2 + rows3}


def form_problems(g):
    """様式（課の印刷用シート）の寸法と、作った帳票の寸法を mm で突き合わせる。"""
    p = []
    def near(a, b):
        return abs(a - b) <= FORM_TOL_MM
    if not near(g["frame_top"], FORM_FRAME_TOP_MM):
        p.append(f"枠の上端の位置が様式と違います（{g['frame_top']:.1f}mm、様式 {FORM_FRAME_TOP_MM}mm）")
    if len(g["blocks"]) != len(FORM_BLOCK_MM):
        p.append(f"まとまりの数が様式と違います（{len(g['blocks'])}）")
    for i, (a, b) in enumerate(zip(g["blocks"], FORM_BLOCK_MM)):
        if not near(a, b):
            p.append(f"{i + 1} つ目のまとまりの高さが様式と違います（{a:.1f}mm、様式 {b}mm）")
    for name, got, want in (("出勤者", g["att_x"], FORM_ATT_X_MM), ("問合せ件数", g["cnt_x"], FORM_CNT_X_MM)):
        if len(got) != len(want) or any(not near(a, b) for a, b in zip(got, want)):
            p.append(f"{name}の縦線の位置が様式と違います（{[round(x, 1) for x in got]}、様式 {want}）")
    for (a1, a2), (b1, b2) in zip(g["stamps"], FORM_STAMP_MM):
        if not (near(a1, b1) and near(a2, b2)):
            p.append(f"回覧の押印欄の位置が様式と違います（{a1:.1f}～{a2:.1f}mm、様式 {b1}～{b2}mm）")
    return p


def run():
    if os.path.exists(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)
    results = []
    for scenario in ("normal", "worst", "empty"):
        sim = build(scenario)
        r = sim.request("GET", f"/nippou/staff/printpdf.asp?d={DAY}")
        if r.error or r.strict or r.warnings:
            raise Fail(f"{scenario}: printpdf.asp で問題: {r.error} {r.strict} {r.warnings}")
        if not r.body.startswith(b"%PDF"):
            raise Fail(f"{scenario}: PDF ができませんでした: " + strip_tags(r.text)[:800])
        pdf = os.path.join(OUT, f"{scenario}.pdf")
        open(pdf, "wb").write(r.body)
        # Edge に渡した HTML（sheet.asp が組み立てたもの）をそのまま取り出す
        work = sim.winfs.to_local("C:\\inetpub\\wwwroot\\nippou\\data\\pdfwork")
        jobs = sorted(os.listdir(work), key=lambda j: os.path.getmtime(os.path.join(work, j)))   # 一番新しい作業
        html = os.path.join(OUT, f"{scenario}.html")
        shutil.copyfile(os.path.join(work, jobs[-1], "sheet.html"), html)
        pages, size = pdf_pages(pdf)
        res = check_html(html)
        subprocess.run(["pdftoppm", "-r", "80", "-png", "-singlefile", pdf, os.path.join(OUT, scenario)], check=True)
        probs = list(res["problems"]) + form_problems(res["geo"])
        if pages != 1:
            probs.append(f"PDF が {pages} 枚です（A4 1 枚でなければならない）")
        if abs(size[0] - 595.3) > 2 or abs(size[1] - 841.9) > 2:
            probs.append(f"用紙が A4 縦ではありません（{size}）")
        persons, dpages, heights, dprobs = detail_pdf(sim, scenario)
        probs += dprobs
        if persons:
            print(f"{scenario}: 個人別の受付表 {persons} 人ぶん・PDF {dpages} 枚・1 人ぶんの高さ 最大 {max(heights):.1f}mm（上限 279mm）")
        results.append((scenario, pages, res["height_mm"], res["tables"], probs))
    ok = True
    for scenario, pages, h, tables, probs in results:
        print(f"{scenario}: PDF {pages} 枚・帳票の高さ {h:.1f}mm（上限 279mm）・表 {len(tables)} 個")
        for p in probs:
            print("   NG:", p)
            ok = False
    bridge().shutdown()
    if not ok:
        sys.exit(1)
    print("OK: 帳票（課の様式と寸法が一致・枠が閉じている・はみ出し無し・A4 1 枚）と個人別の受付表（1 人 1 枚）")


if __name__ == "__main__":
    try:
        run()
    except Fail as e:
        print("NG:", e)
        bridge().shutdown()
        sys.exit(1)
