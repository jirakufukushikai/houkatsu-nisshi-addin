/* 業務日誌 入力補助: Outlook の予定の作成画面で、件名と本文をルールどおりに入れる */
(function () {
  var R = window.NisshiRules;
  var $ = function (id) { return document.getElementById(id); };
  var inOutlook = false;

  function item() { return Office.context.mailbox.item; }

  function state() {
    return {
      type: $("type").value,
      person: $("person").value,
      core: $("core").checked,
      office: $("office").value,
      meeting: $("meeting").value,
      count: $("count").value,
      external: $("external").value
    };
  }

  function showFields() {
    var t = R.findType($("type").value);
    var show = t ? R.FIELDS[t.kind] : [];
    var fields = document.querySelectorAll(".field");
    for (var i = 0; i < fields.length; i++) {
      fields[i].hidden = show.indexOf(fields[i].getAttribute("data-field")) < 0;
    }
  }

  function render() {
    showFields();
    var r = R.build(state());
    $("previewSubject").textContent = r.subject || "―";
    $("previewBodyWrap").hidden = r.bodyLines.length === 0;
    $("previewBody").textContent = r.bodyLines.join("\n");
    var ul = $("errors");
    ul.innerHTML = "";
    // 入力途中は種別未選択以外のエラーを控えめに出す（何か入力されてから）
    var touched = ["person", "office", "meeting", "count", "external"].some(function (k) { return $(k).value; });
    var errs = touched ? r.errors : r.errors.filter(function (e) { return e.indexOf("種別") >= 0; });
    errs.forEach(function (e) { var li = document.createElement("li"); li.textContent = e; ul.appendChild(li); });
    ul.hidden = errs.length === 0;
    $("apply").disabled = r.errors.length > 0;
    return r;
  }

  function setStatus(msg, warn) {
    $("status").textContent = msg;
    $("status").className = "status" + (warn ? " warn" : "");
  }

  function escapeHtml(s) {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function apply() {
    var r = render();
    if (r.errors.length) return;
    if (!inOutlook) { setStatus("（プレビュー）件名: " + r.subject); return; }
    var it = item();
    it.subject.setAsync(r.subject, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { setStatus("件名を入れられませんでした: " + res.error.message, true); return; }
      if (!r.bodyLines.length) { setStatus("件名を入れました。"); return; }
      // 本文にすでに同じ種類の行（参加者・外部）があれば、重ねて入れない
      it.body.getAsync(Office.CoercionType.Text, function (b) {
        var text = b.status === Office.AsyncResultStatus.Succeeded ? b.value || "" : "";
        var add = r.bodyLines.filter(function (line) {
          var key = line.split("：")[0];
          var re = key === "参加者" ? /参加(者)?(数|人数)?\s*[:：]?\s*[0-9０-９]+/ : /外部(参加者?)?\s*[:：]/;
          return !re.test(text);
        });
        if (!add.length) { setStatus("件名を入れました。本文にはすでに人数・外部参加者が書かれているので変えていません。"); return; }
        var html = add.map(function (l) { return "<div>" + escapeHtml(l) + "</div>"; }).join("");
        it.body.prependAsync(html, { coercionType: Office.CoercionType.Html }, function (p) {
          if (p.status !== Office.AsyncResultStatus.Succeeded) { setStatus("本文に入れられませんでした: " + p.error.message, true); return; }
          var skipped = add.length < r.bodyLines.length ? "（本文にすでにある項目は入れていません）" : "";
          setStatus("件名と本文の先頭に入れました。" + skipped);
        });
      });
    });
  }

  function init() {
    var sel = $("type");
    var opt0 = document.createElement("option");
    opt0.value = ""; opt0.textContent = "選んでください";
    sel.appendChild(opt0);
    R.TYPES.forEach(function (t) {
      var o = document.createElement("option");
      o.value = t.id; o.textContent = t.label;
      sel.appendChild(o);
    });
    ["type", "person", "core", "office", "meeting", "count", "external"].forEach(function (id) {
      $(id).addEventListener(id === "type" || id === "core" ? "change" : "input", render);
    });
    $("apply").addEventListener("click", apply);
    render();
  }

  function prefill(subject) {
    var p = R.parse(subject);
    if (!p) return;
    $("type").value = p.type;
    ["person", "office", "meeting"].forEach(function (k) { if (p[k] !== undefined) $(k).value = p[k]; });
    if (p.core !== undefined) $("core").checked = p.core;
    render();
  }

  init(); // このスクリプトは body の最後で読むので、画面の部品はそろっている
  if (window.Office) {
    Office.onReady(function (info) {
      inOutlook = !!(info && info.host === Office.HostType.Outlook && Office.context.mailbox && Office.context.mailbox.item);
      if (!inOutlook) { setStatus("Outlook の外で開いています（プレビュー表示）"); return; }
      // すでに件名が入っていれば、その内容を入力欄に戻す
      item().subject.getAsync(function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) prefill(res.value);
      });
    });
  }
})();
