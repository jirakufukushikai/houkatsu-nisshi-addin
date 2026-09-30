/*
 * 業務日誌の自動記入ルール（件名・本文の組み立て）。画面（taskpane.js）とテスト（test_rules.mjs）の両方から使う。
 * ルールを変えるときは、Office スクリプト「業務日誌作成_包括」の TAGS と合わせること。
 */
(function (root) {
  // 種別: 業務日誌のどの欄に入るか（kind）と、入力欄の組み合わせ
  var TYPES = [
    { id: "研修会", tag: "研修会", kind: "kaigi", label: "研修会（主催）" },
    { id: "地域ケア会議", tag: "地域ケア会議", kind: "kaigi", label: "地域ケア会議（主催）" },
    { id: "多職種情報交換会", tag: "多職種情報交換会", kind: "kaigi", label: "多職種情報交換会（主催）" },
    { id: "ケアマネ連絡会", tag: "ケアマネ連絡会", kind: "kaigi", label: "ケアマネ連絡会（主催）" },
    { id: "その他会議", tag: "その他会議", kind: "kaigi", label: "その他会議（主催）" },
    { id: "サ担", tag: "サ担", kind: "tantou", label: "サービス担当者会議" },
    { id: "処遇", tag: "処遇", kind: "shogu", label: "処遇検討会（高齢者虐待）" },
    { id: "会議参加", tag: "会議参加", kind: "sanka", label: "会議参加（他機関の会議）" }
  ];

  // 種別ごとに表示する入力欄
  var FIELDS = {
    kaigi: ["meeting", "count"],
    tantou: ["person", "count"],
    shogu: ["person", "core", "external"],
    sanka: ["office", "meeting"]
  };

  function findType(id) {
    for (var i = 0; i < TYPES.length; i++) if (TYPES[i].id === id) return TYPES[i];
    return null;
  }
  // 入力値の整え: 前後の空白を取り、【】や区切り記号の混入を防ぐ
  function clean(s) {
    return String(s || "").replace(/[【】\[\]［］]/g, "").replace(/\s+/g, " ").trim();
  }
  function cleanNoSep(s) {
    return clean(s).replace(/[／/｜|]/g, "・");
  }
  function toHalfDigits(s) {
    return String(s || "").replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xfee0); });
  }
  function withSama(name) {
    name = clean(name);
    if (!name) return "";
    return /(様|さん|氏)$/.test(name) ? name : name + "様";
  }

  // 入力 → { subject, bodyLines[], errors[] }
  function build(input) {
    var t = findType(input.type);
    var errors = [];
    if (!t) return { subject: "", bodyLines: [], errors: ["種別を選んでください"] };
    var subject = "", lines = [];
    var count = toHalfDigits(input.count).replace(/[^0-9]/g, "");
    if (input.count && !count) errors.push("参加者数は数字で入れてください");

    if (t.kind === "kaigi") {
      var meeting = clean(input.meeting);
      if (!meeting) errors.push("会議名を入れてください");
      subject = "【" + t.tag + "】" + meeting;
      if (count) lines.push("参加者：" + count);
    } else if (t.kind === "tantou") {
      var person = withSama(input.person);
      if (!person) errors.push("利用者名を入れてください");
      subject = "【サ担】" + person;
      if (count) lines.push("参加者：" + count);
    } else if (t.kind === "shogu") {
      var p2 = withSama(input.person);
      if (!p2) errors.push("利用者名を入れてください");
      subject = "【" + (input.core ? "コア会議" : "処遇") + "】" + p2;
      var ext = clean(input.external).replace(/[／/｜|]/g, "・");
      if (ext) lines.push("外部：" + ext);
    } else if (t.kind === "sanka") {
      var office = cleanNoSep(input.office);
      var m2 = clean(input.meeting);
      if (!office) errors.push("事業所名を入れてください");
      if (!m2) errors.push("会議名を入れてください");
      subject = "【会議参加】" + office + "／" + m2;
    }
    return { subject: subject, bodyLines: lines, errors: errors };
  }

  // 既存の件名から入力欄を復元する（開き直したときの初期値）
  function parse(subject) {
    var m = /^\s*[【\[［]\s*([^】\]］]+?)\s*[】\]］]\s*(.*)$/.exec(subject || "");
    if (!m) return null;
    var tag = m[1].replace(/\s+/g, ""), rest = m[2].trim();
    var alias = { "研修": "研修会", "ケア会議": "地域ケア会議", "多職種": "多職種情報交換会", "会議": "その他会議",
      "会議開催": "その他会議", "サービス担当者会議": "サ担", "担当者会議": "サ担", "処遇検討会": "処遇", "参加": "会議参加" };
    var core = tag === "コア会議";
    if (core) tag = "処遇";
    if (alias[tag]) tag = alias[tag];
    var t = findType(tag);
    if (!t) return null;
    var out = { type: t.id };
    if (t.kind === "kaigi") out.meeting = rest;
    if (t.kind === "tantou") out.person = rest;
    if (t.kind === "shogu") { out.person = rest.replace(/（コア会議）$/, ""); out.core = core; }
    if (t.kind === "sanka") {
      var parts = rest.split(/\s*[／/｜|]\s*/);
      out.office = parts.length > 1 ? parts[0] : "";
      out.meeting = parts.length > 1 ? parts.slice(1).join("／") : rest;
    }
    return out;
  }

  var api = { TYPES: TYPES, FIELDS: FIELDS, build: build, parse: parse, findType: findType };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.NisshiRules = api;
})(this);
