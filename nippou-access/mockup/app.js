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
  // マスタを足したり無効にしたりしたら呼び直す。
  // 現行 Excel のように「全員にファイルを配り直す」必要が無いことの実演。
  function fillMasterSelects() {
    var kb = $("e-kb"), keepK = kb.value;
    kb.innerHTML = "";
    D.kubun.filter(function (k) { return !k.off; }).forEach(function (k) {
      var o = el("option", null, D.blocks[k.b] + " / " + k.n);
      o.value = k.id;
      kb.appendChild(o);
    });
    if (keepK && kb.querySelector('option[value="' + keepK + '"]')) kb.value = keepK;

    var pr = $("e-pr"), keepP = pr.value;
    pr.innerHTML = "";
    D.prods.filter(function (p) { return !p.off; }).forEach(function (p) {
      var o = el("option", null, p.id === 0 ? p.n : D.blocks[p.b] + " / " + p.n);
      o.value = p.id;
      pr.appendChild(o);
    });
    if (keepP && pr.querySelector('option[value="' + keepP + '"]')) pr.value = keepP;
  }

  // マスタ保守画面の選択肢
  (function fillMasterForms() {
    [["mp-block", true], ["mk-block", false]].forEach(function (pair) {
      var sel = $(pair[0]);
      Object.keys(D.blocks).forEach(function (id) {
        if (+id === 0) return;
        if (pair[1] && +id > 2) return;          // 製品を持つのは上 2 ブロックだけ
        var o = el("option", null, D.blocks[id]);
        o.value = id;
        sel.appendChild(o);
      });
    });
    var col = $("mk-col");
    Object.keys(D.cols).forEach(function (c) {
      var o = el("option", null, c + ": " + D.cols[c]);
      o.value = c;
      if (c === "7") o.selected = true;
      col.appendChild(o);
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
      rows += "<tr><td>" + nm(i) + "</td><td>" + nm(i + 7) + "</td><td>　</td></tr>";
    }

    var tasks = "";
    for (var j = 0; j < 8; j++) {
      var L = D.tasks[j], R = D.tasks[j + 8];
      tasks +=
        '<tr><td style="width:34%">' + (L ? L.n : "　") + "</td>" +
        '<td style="width:9%"></td><td style="width:5%">' + (L ? "件" : "") + "</td>" +
        '<td style="width:34%">' + (R ? R.n : "　") + "</td>" +
        '<td style="width:9%"></td><td style="width:9%">' + (R ? "件" : "") + "</td></tr>";
    }

    function cell(v, s) {
      return '<td class="big">' + v + '<div class="sub">(' + s + ")</div></td>";
    }

    $("paper").innerHTML =
      '<div class="kanai">【課内限り】</div>' +
      '<div class="date">' + wareki(d) + "</div>" +
      '<div class="ttl">電話応対報告書日報集計表</div>' +
      "<table><tr>" +
        '<th rowspan="8" class="mid" style="width:20mm">出勤者</th>' +
        '<td rowspan="8" class="mid num" style="width:12mm">' + at.length + "</td>" +
        '<td rowspan="8" class="mid" style="width:9mm">名</td>' +
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

  function usage(kind, id) {
    return S.juden.reduce(function (a, r) {
      if (kind === "kubun" && r.k === id) return a + 1;
      if (kind === "prod" && r.p === id) return a + 1;
      if (kind === "op" && r.o === id) return a + 1;
      return a;
    }, 0);
  }

  // 実績で使われているマスタは消させない。過去の帳票からその名前が失われるため。
  function guardDelete(kind, id, what) {
    var n = usage(kind, id);
    if (n > 0) {
      alert("この" + what + "は実績 " + n + " 件で使われているため削除できません。\n\n" +
            "代わりに「有効」のチェックを外してください。\n" +
            "入力候補から消えますが、過去の日報・集計表はそのまま残ります。");
      return false;
    }
    return confirm("この" + what + "を削除します。よろしいですか？");
  }

  function delBtn(onOk) {
    var b = el("button", "btn x sm", "削除");
    b.type = "button";
    b.addEventListener("click", onOk);
    return b;
  }
  function chkCell(obj, after) {
    var td = el("td");
    var c = el("input");
    c.type = "checkbox";
    c.checked = !obj.off;
    c.addEventListener("change", function () {
      obj.off = !c.checked;
      fillMasterSelects();
      if (after) after();
    });
    td.appendChild(c);
    return td;
  }

  function renderMasterProd() {
    var b = $("mp-rows");
    b.innerHTML = "";
    D.prods.forEach(function (p) {
      if (p.id === 0) return;
      var n = usage("prod", p.id), tr = el("tr");
      tr.appendChild(el("td", null, p.id));
      tr.appendChild(el("td", null, D.blocks[p.b]));
      tr.appendChild(el("td", null, p.n));
      tr.appendChild(chkCell(p, renderMasterProd));
      var tdn = el("td", "n", n || "―");
      if (n) tdn.style.color = "var(--muted)";
      tr.appendChild(tdn);
      var tdx = el("td");
      tdx.appendChild(delBtn(function () {
        if (!guardDelete("prod", p.id, "製品")) return;
        D.prods.splice(D.prods.indexOf(p), 1);
        fillMasterSelects();
        renderMasterProd();
      }));
      tr.appendChild(tdx);
      b.appendChild(tr);
    });
  }

  function renderMasterKubun() {
    var b = $("mk-rows");
    b.innerHTML = "";
    D.kubun.forEach(function (k) {
      var n = usage("kubun", k.id), tr = el("tr");
      tr.appendChild(el("td", null, k.id));
      tr.appendChild(el("td", null, D.blocks[k.b]));
      tr.appendChild(el("td", null, k.n));

      var tdc = el("td");
      var sel = el("select");
      sel.style.maxWidth = "140px";
      Object.keys(D.cols).forEach(function (c) {
        var o = el("option", null, c + ": " + D.cols[c]);
        o.value = c;
        if (+c === k.c) o.selected = true;
        sel.appendChild(o);
      });
      sel.addEventListener("change", function () { k.c = +sel.value; render(); });
      tdc.appendChild(sel);
      tr.appendChild(tdc);

      tr.appendChild(chkCell(k, renderMasterKubun));
      var tdn = el("td", "n", n || "―");
      if (n) tdn.style.color = "var(--muted)";
      tr.appendChild(tdn);
      var tdo = el("td", null, k.old || "―");
      tdo.style.cssText = "font-size:12px;color:var(--muted)";
      tr.appendChild(tdo);

      var tdx = el("td");
      tdx.appendChild(delBtn(function () {
        if (!guardDelete("kubun", k.id, "区分")) return;
        D.kubun.splice(D.kubun.indexOf(k), 1);
        fillMasterSelects();
        renderMasterKubun();
      }));
      tr.appendChild(tdx);
      b.appendChild(tr);
    });
  }

  function renderMasterTask() {
    var b = $("tk-rows");
    b.innerHTML = "";
    D.tasks.forEach(function (t) {
      var tr = el("tr");
      tr.appendChild(el("td", null, t.id));
      tr.appendChild(el("td", null, t.no || "―"));
      tr.appendChild(el("td", null, t.n));
      tr.appendChild(chkCell(t, renderMasterTask));
      var tdx = el("td");
      tdx.appendChild(delBtn(function () {
        if (!confirm("この業務項目を削除します。よろしいですか？")) return;
        D.tasks.splice(D.tasks.indexOf(t), 1);
        renderMasterTask();
        renderPaper($("p-date").value || S.date);
      }));
      tr.appendChild(tdx);
      b.appendChild(tr);
    });
  }

  function renderMaster() {
    renderMasterProd();
    renderMasterKubun();
    renderMasterTask();
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

      var used = usage("op", o.id);
      var tdu = el("td", "n", used || "―");
      if (used) tdu.style.color = "var(--muted)";
      tr.appendChild(tdu);

      var tdx = el("td");
      tdx.appendChild(delBtn(function () {
        if (!guardDelete("op", o.id, "担当者")) return;
        D.ops.splice(D.ops.indexOf(o), 1);
        fillOps($("e-op"), S.date, S.op);
        render();
      }));
      tr.appendChild(tdx);
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
  $("p-print").addEventListener("click", function () { window.print(); });
  $("c-date").addEventListener("change", function () { renderCheck(this.value); });
  $("q-d1").addEventListener("change", renderSum);
  $("q-d2").addEventListener("change", renderSum);
  $("q-w1").addEventListener("click", function () {
    $("q-d1").value = "2026-08-24"; $("q-d2").value = "2026-08-28"; renderSum();
  });
  $("q-w0").addEventListener("click", function () {
    $("q-d1").value = "2026-08-17"; $("q-d2").value = "2026-08-21"; renderSum();
  });

  // ---- マスタ保守のタブと追加ボタン ----------------------------------------
  $("ms-tabs").addEventListener("click", function (e) {
    var b = e.target.closest("button[data-mt]");
    if (!b) return;
    Array.prototype.forEach.call($("ms-tabs").children, function (x) {
      x.className = "btn" + (x === b ? "" : " q");
    });
    ["op", "pr", "kb", "tk"].forEach(function (k) {
      $("mt-" + k).hidden = k !== b.dataset.mt;
    });
  });

  $("mp-add").addEventListener("click", function () {
    var nm = $("mp-name").value.trim();
    if (!nm) { alert("製品名を入力してください。"); return; }
    var b = +$("mp-block").value;
    if (D.prods.some(function (p) { return p.n === nm && p.b === b; })) {
      alert("同じブロックに同名の製品が既にあります。"); return;
    }
    var id = Math.max.apply(null, D.prods.map(function (p) { return p.id; })) + 1;
    D.prods.push({ id: id, b: b, n: nm, ord: +$("mp-ord").value });
    $("mp-name").value = "";
    fillMasterSelects();
    renderMasterProd();
    alert("製品「" + nm + "」を追加しました。受付入力の選択肢にもう出ています。");
  });

  $("mk-add").addEventListener("click", function () {
    var nm = $("mk-name").value.trim();
    if (!nm) { alert("区分名を入力してください。"); return; }
    var id = Math.max.apply(null, D.kubun.map(function (k) { return k.id; })) + 1;
    var c = +$("mk-col").value;
    D.kubun.push({ id: id, b: +$("mk-block").value, n: nm, c: c, old: "" });
    kbById[id] = D.kubun[D.kubun.length - 1];
    $("mk-name").value = "";
    fillMasterSelects();
    renderMasterKubun();
    alert("区分「" + nm + "」を追加しました。\n" +
          "集計表の「" + D.cols[c] + "」列に積まれます。\n" +
          "受付入力の選択肢にもう出ています。");
  });

  $("tk-add").addEventListener("click", function () {
    var nm = $("tk-name").value.trim();
    if (!nm) { alert("帳票表示名を入力してください。"); return; }
    var id = Math.max.apply(null, D.tasks.map(function (t) { return t.id; })) + 1;
    D.tasks.push({ id: id, no: $("tk-no").value.trim(), n: nm });
    $("tk-name").value = ""; $("tk-no").value = "";
    renderMasterTask();
    renderPaper($("p-date").value || S.date);
    alert("業務項目を追加しました。帳票の欄にも出ています。");
  });

  // ---- 起動 ---------------------------------------------------------------
  ["e-date", "d-date", "p-date", "c-date"].forEach(function (i) { $(i).value = S.date; });
  $("q-d1").value = "2026-08-24";
  $("q-d2").value = "2026-08-28";
  fillOps($("e-op"), S.date, S.op);
  fillMasterSelects();
  render();
})();
