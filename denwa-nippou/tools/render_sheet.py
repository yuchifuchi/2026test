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

CHECK_JS = r"""
<script>
(function(){
  var out = {tables: [], problems: [], height_mm: 0};
  var mm = 96 / 25.4;
  var sheet = document.querySelector('.sheet');
  out.height_mm = sheet.getBoundingClientRect().height / mm;
  var tables = sheet.querySelectorAll('table');
  for (var t = 0; t < tables.length; t++) {
    var tb = tables[t], grid = [], rows = tb.rows;
    for (var r = 0; r < rows.length; r++) {
      var c = 0;
      grid[r] = grid[r] || [];
      for (var k = 0; k < rows[r].cells.length; k++) {
        var cell = rows[r].cells[k];
        while (grid[r][c]) c++;
        for (var dr = 0; dr < cell.rowSpan; dr++) for (var dc = 0; dc < cell.colSpan; dc++) {
          grid[r + dr] = grid[r + dr] || [];
          if (grid[r + dr][c + dc]) out.problems.push('升目の重なり: 表' + t + ' 行' + (r + dr) + ' 列' + (c + dc));
          grid[r + dr][c + dc] = 1;
        }
        c += cell.colSpan;
        var cs = getComputedStyle(cell);
        ['Top', 'Right', 'Bottom', 'Left'].forEach(function (side) {
          var w = parseFloat(cs['border' + side + 'Width']), st = cs['border' + side + 'Style'];
          if (!(w > 0) || st === 'none' || st === 'hidden') out.problems.push('罫線が無い: 表' + t + ' 行' + r + ' 欄' + k + ' ' + side);
        });
        if (cell.innerHTML.replace(/\s/g, '') === '') out.problems.push('空の欄（&nbsp; が無い）: 表' + t + ' 行' + r + ' 欄' + k);
        if (cell.scrollWidth > cell.clientWidth + 1) out.problems.push('文字のはみ出し: 表' + t + ' 行' + r + ' 「' + cell.textContent.slice(0, 20) + '」');
      }
    }
    var width = 0;
    for (var i = 0; i < grid.length; i++) width = Math.max(width, grid[i].length);
    for (var r2 = 0; r2 < grid.length; r2++) for (var c2 = 0; c2 < width; c2++) if (!grid[r2][c2]) out.problems.push('升目の欠け: 表' + t + ' 行' + r2 + ' 列' + c2);
    out.tables.push({rows: grid.length, cols: width});
  }
  var pre = document.createElement('pre'); pre.id = '__result'; pre.textContent = JSON.stringify(out);
  document.body.appendChild(pre);
})();
</script>
"""


def worst_people():
    out = []
    for i in range(14):
        # 氏名は画面で止めている上限（12 文字）いっぱい
        out.append({"code": f"{i + 1:03d}", "sei": "勅使河原", "mei": "久美子（旧姓）" if i % 2 == 0 else "由紀子（旧姓）", "kind": "パート"})
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
    for cid, nm in enumerate(["申込受付件数", "抽選結果照会", "払込用紙再発", "商品発送状況", "その他全般分"], start=1):
        r = post(sim, f"/nippou/staff/master.asp?t=col", {"id": str(cid), "name": nm})
        if r.code != 302:
            raise Fail("列の名前を直せません: " + strip_tags(r.text)[:500])
    for g in range(1, 14):
        r = post(sim, "/nippou/staff/master.asp?t=gyomu", {"id": str(g), "no": "⑬", "name": "長い業務の名前",
                                                           "label": "長い業務の名前長い業務の名前長い", "order": str(g), "active": "1"})
        if r.code != 302:
            raise Fail("業務項目を直せません: " + strip_tags(r.text)[:500])
    for t in pid:
        test_flow.entry(sim, "part", t, {"1_1": 9999, "3_3": 9999, "4_1": 9999, "7_2": 9999, "14_0": 9999, "8_2": 9999, "6_1": 9999})
        form = {"act": "save", "d": DAY, "t": str(t)}
        form.update({f"g_{g}": "9999" for g in range(1, 14)})
        post(sim, "/nippou/part/tasks.asp", form)
    # 記述欄を上限（daily.asp の TEXT_TOTAL_LINES）いっぱいに
    per, total = daily_limits()
    line = "あ" * per
    each = total // 3
    form = {"act": "fix", "d": DAY, "lines": "99",
            "t1": "\r\n".join([line] * each), "t2": "\r\n".join([line] * each), "t3": "\r\n".join([line] * (total - 2 * each))}
    note_max = att_note_max()
    for t in pid:
        form[f"att_{t}"] = "1"
        form[f"h_{t}"] = "9:00-13:00"[:note_max]
        form[f"n_{t}"] = "早退あり予定"[:max(0, note_max - len(form[f"h_{t}"]))]
    r = post(sim, "/nippou/staff/daily.asp", form)
    if r.code != 302:
        raise Fail("日報（上限いっぱい）の確定に失敗: " + strip_tags(r.text)[:800])
    return sim


def ids_all(sim):
    _, rows = sim.bridge.query("SELECT T.[担当者ID] FROM [M_担当者] AS T", [])
    return [r[0] for r in rows]


def daily_limits():
    src = open(os.path.join(ROOT, "src", "pages", "staff", "daily.asp"), encoding="utf-8").read()
    per = int(re.search(r"Const TEXT_PER_LINE = (\d+)", src).group(1))
    total = int(re.search(r"Const TEXT_TOTAL_LINES = (\d+)", src).group(1))
    return per, total


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
        jobs = sorted(os.listdir(work))
        html = os.path.join(OUT, f"{scenario}.html")
        shutil.copyfile(os.path.join(work, jobs[-1], "sheet.html"), html)
        pages, size = pdf_pages(pdf)
        res = check_html(html)
        subprocess.run(["pdftoppm", "-r", "80", "-png", "-singlefile", pdf, os.path.join(OUT, scenario)], check=True)
        probs = list(res["problems"])
        if pages != 1:
            probs.append(f"PDF が {pages} 枚です（A4 1 枚でなければならない）")
        if abs(size[0] - 595.3) > 2 or abs(size[1] - 841.9) > 2:
            probs.append(f"用紙が A4 縦ではありません（{size}）")
        results.append((scenario, pages, res["height_mm"], res["tables"], probs))
    ok = True
    for scenario, pages, h, tables, probs in results:
        print(f"{scenario}: PDF {pages} 枚・帳票の高さ {h:.1f}mm（上限 279mm）・表 {len(tables)} 個 {[(t['rows'], t['cols']) for t in tables]}")
        for p in probs:
            print("   NG:", p)
            ok = False
    bridge().shutdown()
    if not ok:
        sys.exit(1)
    print("OK: 帳票（罫線が閉じている・はみ出し無し・A4 1 枚）")


if __name__ == "__main__":
    try:
        run()
    except Fail as e:
        print("NG:", e)
        bridge().shutdown()
        sys.exit(1)
