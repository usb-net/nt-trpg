/**
 * NT卓 投稿ボード用の窓口（Google Apps Script）
 *
 * 使い方
 * 1. Googleスプレッドシートを新規作成し、「拡張機能」→「Apps Script」を開く
 * 2. このファイルの中身を Code.gs に丸ごと貼り付けて保存
 * 3. 関数一覧から setup を選んで一度だけ実行（シート作成）
 * 4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *    実行ユーザー：自分 ／ アクセスできるユーザー：全員
 * 5. 発行された …/exec のURLを board.html の API に貼る
 *
 * 削除はスプレッドシートを直接編集する。状況の「済」⇔「候補」、各項目の編集はボードからもできる。
 * ボードの「URLから取得」は、BOOTH・pixivのページ情報からシナリオ名・種類・形式を推測して返す（保存はしない）。
 * 「種類」列は TRPG／マダミス／2人用／メモ のいずれか（2人用＝KPC＋PC、または2PC）。それ以外の値は「未分類」の候補として表示される。
 */

var SHEET = "投稿";
var HEAD = ["日時", "種類", "シナリオ名", "URL", "形式・人数など", "メモ", "状況"];
var MEMO = "メモ";
var GENRES = ["TRPG", "マダミス", "2人用"];
var MAX_LIST = 300;

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clip_(v, n) {
  return String(v == null ? "" : v).replace(/[\r\n\t]+/g, " ").trim().slice(0, n);
}

// 先頭が = + - @ の値は数式として解釈されないよう ' を付ける
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function list_() {
  var sh = sheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var start = Math.max(2, sh.getLastRow() - MAX_LIST + 1);
  var rows = sh.getRange(start, 1, sh.getLastRow() - start + 1, HEAD.length).getValues();
  return rows.filter(function (r) { return r[1]; }).map(function (r) {
    return {
      at: r[0] instanceof Date ? r[0].getTime() : 0,
      kind: r[1] === MEMO ? "memo" : "scenario",
      genre: GENRES.indexOf(String(r[1])) >= 0 ? String(r[1]) : "",
      title: String(r[2]), url: String(r[3]), format: String(r[4]),
      memo: String(r[5]), status: String(r[6])
    };
  }).reverse();
}

function doGet() {
  try {
    return json_({ ok: true, data: { items: list_() } });
  } catch (err) {
    return json_({ ok: false, error: "読み込みに失敗しました" });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var b = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (b.action === "lookup") return lookup_(b);
    if (b.action === "status") return setStatus_(b);
    if (b.action === "edit") return edit_(b);
    if (b.action !== "add") return json_({ ok: false, error: "不明な操作です" });

    var kind = b.kind === "memo" ? "memo" : "scenario";
    var genre = GENRES.indexOf(b.genre) >= 0 ? b.genre : "TRPG";
    var title = clip_(b.title, 80);
    var url = clip_(b.url, 300);
    var format = clip_(b.format, 60);
    var memo = clip_(b.memo, 300);

    if (url && !/^https?:\/\/\S+$/.test(url)) return json_({ ok: false, error: "URLは http から始まる形で入れてください" });
    if (kind === "scenario" && !title && !url) return json_({ ok: false, error: "シナリオ名かURLのどちらかを入れてください" });
    if (kind === "memo" && !memo) return json_({ ok: false, error: "メモを入れてください" });

    sheet_().appendRow([new Date(), kind === "memo" ? MEMO : genre, safe_(title || (kind === "scenario" ? "名称未確認" : "")),
      url, safe_(format), safe_(memo), kind === "scenario" ? "候補" : ""]);
    return json_({ ok: true, data: { items: list_() } });
  } catch (err) {
    return json_({ ok: false, error: "保存に失敗しました。少し待ってから試してください" });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// シナリオの行を日時（A列）とシナリオ名（C列）で探し、行番号を返す（見つからなければ 0）
function findRow_(sh, at, title) {
  var n = sh.getLastRow() - 1;
  at = Number(at);
  if (n < 1 || !at) return 0;
  var rows = sh.getRange(2, 1, n, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r[1] === MEMO) continue;
    if (r[0] instanceof Date && r[0].getTime() === at && String(r[2]) === String(title || "")) return i + 2;
  }
  return 0;
}

var NOT_FOUND = "対象が見つかりませんでした。再読み込みしてから試してください";

// 状況を「済」⇔「候補」に切り替える
function setStatus_(b) {
  var sh = sheet_();
  var row = findRow_(sh, b.at, b.title);
  if (!row) return json_({ ok: false, error: NOT_FOUND });
  sh.getRange(row, 7).setValue(b.status === "済" ? "済" : "候補");
  return json_({ ok: true, data: { items: list_() } });
}

// 種類・シナリオ名・URL・形式・メモを書き換える（状況はそのまま）
function edit_(b) {
  var f = b.fields || {};
  var genre = GENRES.indexOf(f.genre) >= 0 ? f.genre : "TRPG";
  var title = clip_(f.title, 80);
  var url = clip_(f.url, 300);
  var format = clip_(f.format, 60);
  var memo = clip_(f.memo, 300);
  if (url && !/^https?:\/\/\S+$/.test(url)) return json_({ ok: false, error: "URLは http から始まる形で入れてください" });
  if (!title && !url) return json_({ ok: false, error: "シナリオ名かURLのどちらかを入れてください" });
  var sh = sheet_();
  var row = findRow_(sh, b.at, b.title);
  if (!row) return json_({ ok: false, error: NOT_FOUND });
  sh.getRange(row, 2, 1, 5).setValues([[genre, safe_(title || "名称未確認"), url, safe_(format), safe_(memo)]]);
  return json_({ ok: true, data: { items: list_() } });
}

// ---- URLから取得 ----

function lookup_(b) {
  var url = String(b.url || "");
  var info = null, m, why = "";
  try {
    if ((m = url.match(/booth\.pm\/(?:[a-z]{2}\/)?items\/(\d+)/))) {
      var r = UrlFetchApp.fetch("https://booth.pm/ja/items/" + m[1] + ".json",
        { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0" } });
      why = "BOOTH " + r.getResponseCode();
      if (r.getResponseCode() === 200) {
        var j = JSON.parse(r.getContentText());
        info = guess_(j.name || "", (j.tags || []).map(function (t) { return t.name; }), j.description || "");
      }
    } else if ((m = url.match(/pixiv\.net\/novel\/show\.php\?(?:[^#]*&)?id=(\d+)/))) {
      var r2 = UrlFetchApp.fetch("https://www.pixiv.net/ajax/novel/" + m[1],
        { muteHttpExceptions: true, headers: { "User-Agent": "Mozilla/5.0", "Referer": "https://www.pixiv.net/" } });
      why = "pixiv " + r2.getResponseCode();
      if (r2.getResponseCode() === 200) {
        var body = (JSON.parse(r2.getContentText()) || {}).body || {};
        var tags = ((body.tags || {}).tags || []).map(function (t) { return t.tag; });
        info = guess_(body.title || "", tags, htmlText_(body.description || ""));
      }
    } else {
      return json_({ ok: false, error: "自動取得できるのはBOOTHとpixiv小説のURLです" });
    }
  } catch (err) {
    info = null;
    why = String(err && err.message || err).slice(0, 120);
  }
  if (!info) return json_({ ok: false, error: "ページの情報を取得できませんでした。手で入力してください（" + why + "）" });
  return json_({ ok: true, data: info });
}

function htmlText_(h) {
  return String(h).replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, " ")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");
}

function guess_(name, tags, desc) {
  var fields = fields_(desc);
  return { title: cleanTitle_(name), genre: genre_(tags, desc, fields.people), format: format_(name, tags, desc, fields) };
}

// 「【CoCシナリオ】」「SPLL:E…」などの飾りを外してシナリオ名だけにする
function cleanTitle_(t0) {
  var t = String(t0).replace(/[（(]?\s*SPLL\s*[:：]?\s*E?\d+\s*[）)]?/gi, "").trim();
  var q = t.match(/『([^』]+)』/) || t.match(/「([^」]+)」/);
  if (q) return q[1].trim();
  var KW = /CoC|シナリオ|TRPG|クトゥルフ|タイマン|[0-9０-９]+版|6th|7th/i;
  t = t.replace(/【([^】]*)】/g, function (all, inner) { return KW.test(inner) ? " " : all; });
  t = t.replace(/^\s*(CoC|新?クトゥルフ神話TRPG)\s*(第?[0-9]+(版|th))?\s*(タイマン)?\s*(シナリオ)?\s*[:：\/／|｜]?/i, "");
  t = t.trim().replace(/[（(][^（()）]*[）)]\s*$/, "").trim();
  var only = t.match(/^【([^】]+)】$/);
  if (only) t = only[1];
  return clip_(t.trim() || t0, 80);
}

// 説明文の「人数：」「プレイ時間：」などを抜き出す
function fields_(desc) {
  var d = String(desc).replace(/\r/g, "");
  function pick(re) {
    var m = d.match(re);
    if (!m) return "";
    var v = m[1].split(/[\n【■]/)[0].trim();
    if (v.length > 14) v = v.split(/[（(]/)[0].trim();
    return v.slice(0, 20);
  }
  return {
    people: pick(/(?:PL人数|プレイ人数|推奨人数|対象人数|PL数|人数)\s*[】:：]\s*([^\n]+)/),
    time: pick(/(?:プレイ時間|所要時間|想定時間|時間)\s*[】:：]\s*([^\n]+)/)
  };
}

function genre_(tags, desc, people) {
  var all = tags.join(" ") + "\n" + desc;
  if (/マーダーミステリー|マダミス/.test(all)) return "マダミス";
  if (tags.some(function (t) { return /タイマン|2PL|２PL|KPレス/.test(t); })) return "2人用";
  if (/^(KP\s*[1１]\s*名?\s*[、,／/]?\s*)?(PL\s*)?([12１２]\s*(人|名|PL)|ひとり|一人|二人|ふたり|タイマン)/.test(people)) return "2人用";
  if (/[2２二]\s*人\s*(限定|用)|[2２]\s*PC|KPPC|KPC\s*[+＋]\s*[1１]?\s*PC|PL\s*[1１一]\s*人|タイマン(?!改変)/.test(desc)) return "2人用";
  return "TRPG";
}

function format_(name, tags, desc, fields) {
  var parts = [];
  var all = name + " " + tags.join(" ") + " " + desc;
  // 「新クトゥルフ」は権利表記にも出てくるので、タイトルとタグだけで判定する
  if (/7版|第7版|7th/i.test(all) || /新クトゥルフ/.test(name + " " + tags.join(" "))) parts.push("7版"); else if (/6版|第6版|6th/i.test(all)) parts.push("6版");
  if (tags.some(function (t) { return /KPレス/.test(t); })) parts.push("KPレス");
  if (fields.people) parts.push(fields.people);
  else if (tags.some(function (t) { return /タイマン/.test(t); })) parts.push("タイマン");
  if (fields.time) parts.push(fields.time);
  return clip_(parts.join("／"), 60);
}

/** 「URLから取得」の確認用：エディタで一度実行すると、外部サイトへの接続の承認画面が出る */
function testLookup() {
  Logger.log(lookup_({ url: "https://booth.pm/ja/items/2274429" }).getContent());
  Logger.log(lookup_({ url: "https://www.pixiv.net/novel/show.php?id=8886254" }).getContent());
}

/** 初回だけ手動で実行：見出し付きの「投稿」シートを作る */
function setup() {
  sheet_();
}
