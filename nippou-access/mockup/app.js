(function () {
  "use strict";

  // ---- 索引 ---------------------------------------------------------------
  var opById = {}, kbById = {}, prById = {};
  D.ops.forEach(function (o) { opById[o.id] = o; });
  D.kubun.forEach(function (k) { kbById[k.id] = k; });
  D.prods.forEach(function (p) { prById[p.id] = p; });
  var KB_RETURN = (D.kubun.filter(function (k) { return k.old === "返金"; })[0] || {}).id;

  // ---- 状態 (このページの中だけ。保存はしない) ------------------------------
  var S = {
    juden: D.juden.slice(),
    nippou: JSON.parse(JSON.stringify(D.nippou)),
    attend: {},
    retired: {},
    date: "2026-08-25",
    op: 10
  };
  S.juden.forEach(function (r) { S.attend[r.d + "|" + r.o] = true; });

  function $(id) { return document.getElementById(id); }
  function el(tag, cls, txt) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (txt != null) n.textContent = txt;
    return n;
  }
  function fmt(d) {
    var w = "日月火水木金土"[new Date(d + "T00:00:00").getDay()];
    return d.replace(/-/g, "/") + "（" + w + "）";
  }
  function wareki(d) {
    var t = new Date(d + "T00:00:00"), y = t.getFullYear();
    var g = y >= 2019 ? ["令和", y - 2018] : ["平成", y - 1988];
    return g[0] + "　" + g[1] + "年　" + (t.getMonth() + 1) + "月　" + t.getDate() +
      "日（" + "日月火水木金土"[t.getDay()] + "）";
  }
  function rowsOn(d) { return S.juden.filter(function (r) { return r.d === d; }); }

  // 現役の担当者。退職日を入れた人はここから外れるが、過去データは触らない。
  function activeOps(d) {
    return D.ops.filter(function (o) {
      if (!o.on) return false;
      var end = S.retired[o.id];
      return !end || end >= d;
    });
  }

  // ---- 集計 (日報も集計表もここを通る。経路が 1 本しかない) -----------------
  function tally(d) {
    var t = { 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, total: 0, ret: 0, staff: {} };
    [3, 4, 5, 6, 7, 8].forEach(function (c) { t.staff[c] = 0; });
    t.staff.total = 0;
    rowsOn(d).forEach(function (r) {
      var k = kbById[r.k], o = opById[r.o];
      if (!k) return;
      t[k.c] += r.c;
      t.total += r.c;
      if (r.k === KB_RETURN) t.ret += r.c;
      if (o && o.kbn === "職員") { t.staff[k.c] += r.c; t.staff.total += r.c; }
    });
    // 帳票の欄への割り当て。現行 印刷用シート 18 行目の数式から復元したもの。
    t.moushikomi = t[3]; t.chusen = t[4]; t.haraikomi = t[5];
    t.hassou = t[6]; t.sonota = t[7] + t[8]; t.kokan = t[8];
    t.s_moushikomi = t.staff[3]; t.s_chusen = t.staff[4]; t.s_haraikomi = t.staff[5];
    t.s_hassou = t.staff[6]; t.s_sonota = t.staff[7] + t.staff[8];
    return t;
  }

  function attendees(d) {
    return D.ops.filter(function (o) { return S.attend[d + "|" + o.id]; });
  }
  function missing(d) {
    return attendees(d).filter(function (o) {
      return !S.juden.some(function (r) { return r.d === d && r.o === o.id && r.c; });
    });
  }
  function nippou(d) {
    if (!S.nippou[d]) S.nippou[d] = { kaisen: 5, tokki: "", daitai: "", youbou: "", state: "入力中" };
    if (!S.nippou[d].state) S.nippou[d].state = "入力中";
    return S.nippou[d];
  }

  // ---- 画面切替 -----------------------------------------------------------
  function show(name) {
    Array.prototype.forEach.call(document.querySelectorAll(".screen"), function (s) {
      s.classList.toggle("on", s.id === "s-" + name);
    });
    Array.prototype.forEach.call($("nav").children, function (b) {
      if (b.dataset.s === name) b.setAttribute("aria-current", "page");
      else b.removeAttribute("aria-current");
    });
    render();
    window.scrollTo({ top: 0, behavior: "instant" });
  }
  $("nav").addEventListener("click", function (e) {
    var b = e.target.closest("button[data-s]");
    if (b) show(b.dataset.s);
  });
  document.addEventListener("click", function (e) {
    var b = e.target.closest("[data-go]");
    if (b) show(b.dataset.go);
  });

  // ---- テーマ -------------------------------------------------------------
  $("themetog").addEventListener("click", function () {
    var cur = document.documentElement.getAttribute("data-theme");
    var dark = cur ? cur === "dark"
      : window.matchMedia("(prefers-color-scheme: dark)").matches;
    document.documentElement.setAttribute("data-theme", dark ? "light" : "dark");
  });

  // ---- 選択肢 -------------------------------------------------------------
  function fillOps(sel, d, keep) {
    sel.innerHTML = "";
    activeOps(d).forEach(function (o) {
      var op = el("option", null, o.code + "　" + o.name);
      op.value = o.id;
      sel.appendChild(op);
    });
    if (keep && sel.querySelector('option[value="' + keep + '"]')) sel.value = keep;
    else if (sel.options.length) S.op = +sel.value;
  }
  (function fillStatic() {
    var kb = $("e-kb");
    D.kubun.forEach(function (k) {
      var o = el("option", null, D.blocks[k.b] + " / " + k.n);
      o.value = k.id;
      kb.appendChild(o);
    });
    var pr = $("e-pr");
    D.prods.forEach(function (p) {
      var o = el("option", null, p.id === 0 ? p.n : D.blocks[p.b] + " / " + p.n);
      o.value = p.id;
      pr.appendChild(o);
    });
  })();

  // ---- 描画 ---------------------------------------------------------------
  function render() {
    var d = S.date, t = tally(d);

    // 整合性ストリップ: 日報は集計列ごとの合計、集計表は件数の素の合計。
    // 計算の入口を分けても必ず一致することを見せている。
    var byCol = t.moushikomi + t.chusen + t.haraikomi + t.hassou + t.sonota;
    var raw = rowsOn(d).reduce(function (a, r) { return a + r.c; }, 0);
    $("t-nippou").textContent = byCol;
    $("t-sum").textContent = raw;
    $("t-badge").textContent = byCol === raw ? "一致" : "不一致";
    $("t-badge").className = "pill " + (byCol === raw ? "ok" : "wip");
    document.querySelector(".truth-in span").textContent =
      d.slice(5).replace("-", "/") + " の件数 ―";

    renderMenu(d, t);
    renderEntry(d);
    renderDaily(d, t);
    renderPaper($("p-date").value || d);
    renderSum();
    renderCheck($("c-date").value || d);
    renderMaster();
  }

  function renderMenu(d, t) {
    $("m-total").innerHTML = t.total + '<span class="u">件</span>';
    $("m-date").textContent = fmt(d);
    var input = {};
    rowsOn(d).forEach(function (r) { input[r.o] = true; });
    $("m-ops").innerHTML = Object.keys(input).length + '<span class="u">名</span>';
    $("m-opsub").textContent = "出勤登録 " + attendees(d).length + " 名";
    var ms = missing(d).length;
    $("m-miss").innerHTML = ms + '<span class="u">名</span>';
    $("m-miss").style.color = ms ? "var(--warn)" : "";
  }

  function renderEntry(d) {
    var body = $("e-rows"), total = 0;
    body.innerHTML = "";
    var mine = S.juden.filter(function (r) { return r.d === d && r.o === S.op; });
    mine.sort(function (a, b) { return a.k - b.k; });

    mine.forEach(function (r) {
      var k = kbById[r.k], p = prById[r.p] || { n: "" };
      total += r.c;
      var tr = el("tr");
      tr.appendChild(el("td", null, D.blocks[k.b] + " / " + k.n));
      tr.appendChild(el("td", null, p.n));

      var tdc = el("td", "n");
      var inp = el("input");
      inp.type = "number"; inp.min = "0"; inp.value = r.c;
      inp.style.maxWidth = "78px";
      inp.addEventListener("change", function () {
        var v = parseInt(inp.value, 10);
        if (!v || v < 0) S.juden.splice(S.juden.indexOf(r), 1);
        else r.c = v;
        render();
      });
      tdc.appendChild(inp);
      tr.appendChild(tdc);

      var tdb = el("td");
      var bk = el("input");
      bk.type = "text"; bk.value = r.b || ""; bk.placeholder = "―";
      bk.addEventListener("change", function () { r.b = bk.value; });
      tdb.appendChild(bk);
      tr.appendChild(tdb);

      tr.appendChild(el("td", null, D.cols[k.c]));

      var tdx = el("td");
      var del = el("button", "btn x sm", "削除");
      del.type = "button";
      del.addEventListener("click", function () {
        S.juden.splice(S.juden.indexOf(r), 1);
        render();
      });
      tdx.appendChild(del);
      tr.appendChild(tdx);
      body.appendChild(tr);
    });

    if (!mine.length) {
      var tr0 = el("tr");
      var td0 = el("td", "empty", "まだ入力がありません。上のフォームから登録してください。");
      td0.colSpan = 6;
      tr0.appendChild(td0);
      body.appendChild(tr0);
    }
    $("e-foot").textContent = total;
    $("e-total").innerHTML = total + '<span class="u">件</span>';
  }

  function renderDaily(d, t) {
    var n = nippou(d);
    $("d-kaisen").value = n.kaisen;
    $("d-tokki").value = n.tokki || "";
    $("d-daitai").value = n.daitai || "";
    $("d-youbou").value = n.youbou || "";
    $("d-state").textContent = n.state;
    $("d-state").className = "pill " + (n.state === "確定" ? "ok" : "wip");
    $("d-fix").textContent = n.state === "確定" ? "確定を解除" : "確定する";

    var b = $("d-counts");
    b.innerHTML = "";
    var r1 = el("tr");
    var c0 = el("td", "n");
    c0.innerHTML = '<b style="font-size:19px">' + t.total + "</b> 件";
    r1.appendChild(c0);
    [t.moushikomi, t.chusen, t.haraikomi, t.hassou, t.sonota, t.kokan, t.ret]
      .forEach(function (v) { r1.appendChild(el("td", "n", v)); });
    b.appendChild(r1);

    var r2 = el("tr");
    r2.appendChild(el("td", null, "うち職員受電"));
    [t.s_moushikomi, t.s_chusen, t.s_haraikomi, t.s_hassou, t.s_sonota]
      .forEach(function (v) { r2.appendChild(el("td", "n", "(" + v + ")")); });
    r2.appendChild(el("td", "n", ""));
    r2.appendChild(el("td", "n", ""));
    b.appendChild(r2);

    var at = attendees(d);
    $("d-attn-n").textContent = "（" + at.length + " 名）";
    var ab = $("d-attn");
    ab.innerHTML = "";
    at.forEach(function (o) {
      var cnt = S.juden.reduce(function (a, r) {
        return a + (r.d === d && r.o === o.id ? r.c : 0);
      }, 0);
      var tr = el("tr");
      tr.appendChild(el("td", null, o.code));
      tr.appendChild(el("td", null, o.name));
      tr.appendChild(el("td", null, o.kbn));
      var td = el("td", "n", cnt);
      if (!cnt) td.style.color = "var(--warn)";
      tr.appendChild(td);
      ab.appendChild(tr);
    });
    if (!at.length) {
      var tr0 = el("tr"), td0 = el("td", "empty", "出勤者がいません。");
      td0.colSpan = 4; tr0.appendChild(td0); ab.appendChild(tr0);
    }
  }

  function renderPaper(d) {
    var t = tally(d), n = nippou(d), at = attendees(d);
    var names = at.map(function (o) { return o.name; });
    function nm(i) { return names[i] || "　"; }

    var rows = "";
    for (var i = 0; i < 7; i++) {
      rows += '<tr><td class="blank"></td><td class="blank"></td><td class="blank"></td>' +
        "<td>" + nm(i) + "</td><td>" + nm(i + 7) + "</td><td>　</td></tr>";
    }

    var tasks = "";
    for (var j = 0; j < 8; j++) {
      var L = D.tasks[j], R = D.tasks[j + 8];
      tasks += "<tr><td>" + (L ? L.n : "　") + "</td><td>　</td><td>" + (L ? "件" : "") + "</td>" +
        "<td>" + (R ? R.n : "　") + "</td><td>　</td><td>" + (R ? "件" : "") + "</td></tr>";
    }

    function cell(v, s) {
      return '<td class="big">' + v + '<div class="sub">(' + s + ")</div></td>";
    }

    $("paper").innerHTML =
      '<div class="kanai">【課内限り】</div>' +
      '<div class="date">' + wareki(d) + "</div>" +
      '<div class="ttl">電話応対報告書日報集計表</div>' +
      "<table><tr>" +
        '<th style="width:20mm">出勤者</th>' +
        '<td style="width:12mm;text-align:center">' + at.length + "</td>" +
        '<td style="width:9mm">名</td>' +
        '<th style="width:48mm">氏名</th><th style="width:48mm">氏名</th><th>備考</th>' +
      "</tr>" + rows + "</table>" +
      '<p style="margin:9px 0">（　<b>' + n.kaisen + "</b>　回線 ）</p>" +
      "<table><tr>" +
        '<th rowspan="2" style="width:24mm">問合せ件数<br>' +
        '<span style="font-size:10px">（　）内は職員受電数</span></th>' +
        '<td rowspan="2" class="big" style="width:22mm">' + t.total + " 件</td>" +
        '<th style="text-align:center">申込</th><th style="text-align:center">抽選</th>' +
        '<th style="text-align:center">払込用紙</th><th style="text-align:center">商品発送</th>' +
        '<th style="text-align:center">その他</th>' +
      "</tr><tr>" +
        cell(t.moushikomi, t.s_moushikomi) + cell(t.chusen, t.s_chusen) +
        cell(t.haraikomi, t.s_haraikomi) + cell(t.hassou, t.s_hassou) +
        cell(t.sonota, t.s_sonota) +
      "</tr></table>" +
      '<p class="uchi">内　　交換 <b>' + t.kokan + "</b> 件　／　内　　返金 <b>" + t.ret + "</b> 件</p>" +
      '<div class="cap">特記事項（報告書の電話件数だけでは伝わり難い事項など）</div>' +
      '<div class="memo">' + esc(n.tokki) + "</div>" +
      '<div class="cap">職員に代わった案件（概要）</div>' +
      '<div class="memo">' + esc(n.daitai) + "</div>" +
      '<div class="cap">要望（お客様からの問い合わせを減らすための改善提案等）</div>' +
      '<div class="memo">' + esc(n.youbou) + "</div>" +
      '<div class="cap">電話対応以外の業務（データ入力作業等）</div>' +
      "<table>" + tasks + "</table>";
  }
  function esc(s) {
    return String(s || "").replace(/[&<>]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c];
    });
  }

  function renderSum() {
    var d1 = $("q-d1").value, d2 = $("q-d2").value;
    if (!d1 || !d2) return;
    var days = {}, tot = [0, 0, 0, 0, 0, 0, 0];
    S.juden.forEach(function (r) {
      if (r.d < d1 || r.d > d2) return;
      var k = kbById[r.k];
      if (!k) return;
      if (!days[r.d]) days[r.d] = { 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, t: 0 };
      days[r.d][k.c] += r.c;
      days[r.d].t += r.c;
    });

    var b = $("q-days");
    b.innerHTML = "";
    Object.keys(days).sort().forEach(function (d) {
      var v = days[d], tr = el("tr");
      var td = el("td");
      var a = el("button", "btn q sm", fmt(d));
      a.type = "button";
      a.addEventListener("click", function () {
        S.date = d; $("d-date").value = d; $("p-date").value = d; show("daily");
      });
      td.appendChild(a);
      tr.appendChild(td);
      [3, 4, 5, 6, 7, 8].forEach(function (c, i) {
        tr.appendChild(el("td", "n", v[c])); tot[i] += v[c];
      });
      var tt = el("td", "n", v.t);
      tt.style.fontWeight = "700";
      tr.appendChild(tt);
      tot[6] += v.t;
      b.appendChild(tr);
    });
    if (!Object.keys(days).length) {
      var tr0 = el("tr"), td0 = el("td", "empty", "この期間のデータはありません。");
      td0.colSpan = 8; tr0.appendChild(td0); b.appendChild(tr0);
    }
    var f = $("q-foot");
    f.innerHTML = "";
    f.appendChild(el("td", null, "計"));
    tot.forEach(function (v) { f.appendChild(el("td", "n", v)); });

    var rb = $("q-rows");
    rb.innerHTML = "";
    var det = S.juden.filter(function (r) { return r.d >= d1 && r.d <= d2; });
    det.sort(function (a, b2) {
      return a.d < b2.d ? -1 : a.d > b2.d ? 1 : a.o - b2.o || a.k - b2.k;
    });
    det.slice(0, 200).forEach(function (r) {
      var k = kbById[r.k], o = opById[r.o], p = prById[r.p] || { n: "" };
      var tr = el("tr");
      [r.d.replace(/-/g, "/"), o ? o.name : "―", D.blocks[k.b], p.n, k.n, D.cols[k.c]]
        .forEach(function (v) { tr.appendChild(el("td", null, v)); });
      tr.appendChild(el("td", "n", r.c));
      rb.appendChild(tr);
    });
    $("q-count").textContent = det.length > 200
      ? "全 " + det.length + " 件のうち先頭 200 件を表示しています。"
      : "全 " + det.length + " 件。";
  }

  function renderCheck(d) {
    var b = $("c-rows");
    b.innerHTML = "";
    var ms = missing(d);
    ms.forEach(function (o) {
      var tr = el("tr");
      tr.appendChild(el("td", null, o.code));
      tr.appendChild(el("td", null, o.name));
      var td = el("td");
      var go = el("button", "btn q sm", "入力画面を開く");
      go.type = "button";
      go.addEventListener("click", function () {
        S.date = d; S.op = o.id;
        $("e-date").value = d;
        fillOps($("e-op"), d, o.id);
        show("entry");
      });
      td.appendChild(go);
      tr.appendChild(td);
      b.appendChild(tr);
    });
    if (!ms.length) {
      var tr0 = el("tr"), td0 = el("td", "empty", "入力もれはありません。");
      td0.colSpan = 3; tr0.appendChild(td0); b.appendChild(tr0);
    }
  }

  function renderMaster() {
    var b = $("ms-rows"), today = S.date;
    b.innerHTML = "";
    D.ops.forEach(function (o) {
      var cnt = S.juden.reduce(function (a, r) { return a + (r.o === o.id ? r.c : 0); }, 0);
      var end = S.retired[o.id] || "";
      var zai = o.on && (!end || end >= today);
      var tr = el("tr");
      tr.appendChild(el("td", null, o.code));

      var tdn = el("td");
      tdn.appendChild(document.createTextNode(o.name));
      if (o.bad) {
        var w = el("div");
        w.style.cssText = "font-size:11.5px;color:var(--warn);margin-top:2px";
        w.textContent = "現行の Sheet1!B1 が別人のまま";
        tdn.appendChild(w);
      }
      tr.appendChild(tdn);
      tr.appendChild(el("td", null, o.kbn));
      tr.appendChild(el("td", null, "―"));

      var tde = el("td");
      var inp = el("input");
      inp.type = "date"; inp.value = end; inp.style.maxWidth = "150px";
      inp.addEventListener("change", function () {
        if (inp.value) S.retired[o.id] = inp.value;
        else delete S.retired[o.id];
        fillOps($("e-op"), S.date, S.op);
        render();
      });
      tde.appendChild(inp);
      tr.appendChild(tde);

      var tds = el("td");
      tds.appendChild(el("span", "pill " + (zai ? "ok" : "off"), zai ? "在籍" : "退職"));
      tr.appendChild(tds);
      tr.appendChild(el("td", "n", cnt || "―"));
      b.appendChild(tr);
    });
  }

  // ---- 入力の受け口 -------------------------------------------------------
  $("e-date").addEventListener("change", function () {
    S.date = this.value;
    fillOps($("e-op"), S.date, S.op);
    $("d-date").value = S.date; $("p-date").value = S.date; $("c-date").value = S.date;
    render();
  });
  $("e-op").addEventListener("change", function () { S.op = +this.value; render(); });

  $("e-add").addEventListener("click", function () {
    var k = +$("e-kb").value, p = +$("e-pr").value, c = parseInt($("e-cnt").value, 10);
    if (!k || !c || c < 1) return;
    S.attend[S.date + "|" + S.op] = true;
    // 同じ 日付・担当者・区分・製品 は 1 行にまとめる。Access 側の UNIQUE 制約と同じ挙動。
    var hit = S.juden.filter(function (r) {
      return r.d === S.date && r.o === S.op && r.k === k && r.p === p;
    })[0];
    if (hit) hit.c += c;
    else S.juden.push({ d: S.date, o: S.op, k: k, p: p, c: c, b: $("e-bk").value });
    $("e-cnt").value = 1; $("e-bk").value = "";
    render();
  });

  $("d-date").addEventListener("change", function () {
    S.date = this.value;
    $("e-date").value = S.date; $("p-date").value = S.date; $("c-date").value = S.date;
    fillOps($("e-op"), S.date, S.op);
    render();
  });
  ["kaisen", "tokki", "daitai", "youbou"].forEach(function (f) {
    $("d-" + f).addEventListener("input", function () {
      var n = nippou(S.date);
      n[f] = f === "kaisen" ? (parseInt(this.value, 10) || 0) : this.value;
      renderPaper($("p-date").value || S.date);
    });
  });
  $("d-fix").addEventListener("click", function () {
    var n = nippou(S.date);
    if (n.state === "確定") { n.state = "入力中"; render(); return; }
    var ms = missing(S.date);
    if (ms.length && !confirm("実績が 1 件も無い担当者が " + ms.length +
        " 名います。\n\n" + ms.map(function (o) { return "・" + o.name; }).join("\n") +
        "\n\nこのまま確定しますか？")) { show("check"); return; }
    n.state = "確定";
    render();
  });

  $("p-date").addEventListener("change", function () { renderPaper(this.value); });
  $("c-date").addEventListener("change", function () { renderCheck(this.value); });
  $("q-d1").addEventListener("change", renderSum);
  $("q-d2").addEventListener("change", renderSum);
  $("q-w1").addEventListener("click", function () {
    $("q-d1").value = "2026-08-24"; $("q-d2").value = "2026-08-28"; renderSum();
  });
  $("q-w0").addEventListener("click", function () {
    $("q-d1").value = "2026-08-17"; $("q-d2").value = "2026-08-21"; renderSum();
  });

  // ---- 起動 ---------------------------------------------------------------
  ["e-date", "d-date", "p-date", "c-date"].forEach(function (i) { $(i).value = S.date; });
  $("q-d1").value = "2026-08-24";
  $("q-d2").value = "2026-08-28";
  fillOps($("e-op"), S.date, S.op);
  render();
})();
