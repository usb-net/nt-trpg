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
 * 削除・状況の変更はスプレッドシートを直接編集する。
 * 「種類」列は TRPG／マダミス／メモ のいずれか。それ以外の値は「未分類」の候補として表示される。
 */

var SHEET = "投稿";
var HEAD = ["日時", "種類", "シナリオ名", "URL", "形式・人数など", "メモ", "状況"];
var MEMO = "メモ";
var GENRES = ["TRPG", "マダミス"];
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

/** 初回だけ手動で実行：見出し付きの「投稿」シートを作る */
function setup() {
  sheet_();
}
