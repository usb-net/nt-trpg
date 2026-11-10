/**
 * ENTP卓 投稿ボード用の窓口（Google Apps Script）
 *
 * 使い方
 * 1. Googleスプレッドシートを新規作成し、「拡張機能」→「Apps Script」を開く
 * 2. このファイルの中身を Code.gs に丸ごと貼り付けて保存
 * 3. 関数一覧から setup を選んで一度だけ実行（シート作成と既存候補の登録）
 * 4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *    実行ユーザー：自分 ／ アクセスできるユーザー：全員
 * 5. 発行された …/exec のURLを board.html の API に貼る
 *
 * 削除・状況の変更はスプレッドシートを直接編集する。
 */

var SHEET = "投稿";
var HEAD = ["日時", "種類", "シナリオ名", "URL", "形式・人数など", "メモ", "状況"];
var KINDS = { scenario: "シナリオ候補", memo: "メモ" };
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
      kind: r[1] === KINDS.memo ? "memo" : "scenario",
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
    var title = clip_(b.title, 80);
    var url = clip_(b.url, 300);
    var format = clip_(b.format, 60);
    var memo = clip_(b.memo, 300);

    if (url && !/^https?:\/\/\S+$/.test(url)) return json_({ ok: false, error: "URLは http から始まる形で入れてください" });
    if (kind === "scenario" && !title && !url) return json_({ ok: false, error: "シナリオ名かURLのどちらかを入れてください" });
    if (kind === "memo" && !memo) return json_({ ok: false, error: "メモを入れてください" });

    sheet_().appendRow([new Date(), KINDS[kind], safe_(title || (kind === "scenario" ? "名称未確認" : "")),
      url, safe_(format), safe_(memo), kind === "scenario" ? "候補" : ""]);
    return json_({ ok: true, data: { items: list_() } });
  } catch (err) {
    return json_({ ok: false, error: "保存に失敗しました。少し待ってから試してください" });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/** 初回だけ手動で実行：シートを作り、既存のシナリオ候補を登録する */
function setup() {
  var sh = sheet_();
  if (sh.getLastRow() > 1) return;
  var t = new Date();
  var seed = [
    ["孤独の密室", "https://booth.pm/ja/items/3431447", "クローズド／複数人", "次回候補、新規推奨", "日程調整中"],
    ["犯人はHO2", "https://booth.pm/ja/items/1940187", "継続可／約6時間", "ギャグ候補", "候補"],
    ["名称未確認", "https://www.pixiv.net/novel/show.php?id=6678249", "継続可", "PLのSAN値にダメージが入るギャグ", "候補"],
    ["名称未確認", "https://www.pixiv.net/novel/show.php?id=6831442#1", "継続可", "オールドなクトゥルフ神話。發さんは既知", "候補"],
    ["嗤う人形師（要確認）", "https://booth.pm/ja/items/2625282", "秘匿あり", "タイトルとURLの対応は要確認", "候補"],
    ["皆既月食", "https://booth.pm/ja/items/3360155", "マダミス／5PL", "昏さんがGM候補。1月以降を想定", "候補"],
    ["名称未確認", "https://booth.pm/ja/items/2633834", "秘匿あり", "犬のシナリオ候補", "候補"],
    ["薔薇の館", "", "継続可／約4時間", "URLを後で追加", "候補"],
    ["ばにさつ", "", "継続可／約4時間", "URLを後で追加", "候補"]
  ];
  // 表の上から順に並ぶよう、古い順に登録する
  var rows = seed.slice().reverse().map(function (s, i) {
    return [new Date(t.getTime() + i * 1000), KINDS.scenario, s[0], s[1], s[2], s[3], s[4]];
  });
  sh.getRange(2, 1, rows.length, HEAD.length).setValues(rows);
}
