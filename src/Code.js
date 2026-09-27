/**
 * 予約台帳 — サーバー側（Google Apps Script）
 *
 * スプレッドシートに紐づいたスクリプトとして動きます。
 * 画面（index.html）からは google.script.run で、下の「公開する関数」を呼びます。
 * 名前の最後に _ が付いた関数は、画面からは呼べない内部用の関数です。
 */

/* ---------- お店の名前（画面の上部とページのタイトルに出ます） ---------- */
const SHOP_NAME = 'SAMPLE SALON';

/* ---------- シートの形 ---------- */
const SHEETS = { bookings: '予約', customers: 'お客様', staff: '担当者', menus: 'メニュー' };
const HEADERS = {
  bookings:  ['ID', '日付', '開始', '分', '担当ID', '担当名', 'お客様', 'メニュー', 'メモ', '状態', '作成日時', '更新日時'],
  customers: ['名前', '来店回数', '最終予約日時'],
  staff:     ['ID', '名前', '色'],
  menus:     ['名前', '分'],
};
// 予約シートの列番号（1始まり）
const COL = { id: 1, date: 2, start: 3, dur: 4, staff: 5, staffName: 6, name: 7, menus: 8, memo: 9, status: 10, created: 11, updated: 12 };
const MENU_SEP = '、';
const STATUS_OK = '予約';
const STATUS_CANCELED = '取消';   // 行は消さずに「取消」にする（あとから履歴を確認できるように）
const TZ = Session.getScriptTimeZone();

// 最初の担当者（公開用の仮の名前です。実際の名前は画面の「担当者・メニューの設定」から変えてください）
const DEFAULT_STAFF = [['staff-a', 'スタッフA', 0], ['staff-b', 'スタッフB', 1]];
const DEFAULT_MENUS = [
  ['カット', 60], ['カラー', 90], ['パーマ', 120], ['縮毛矯正', 180], ['トリートメント', 30],
  ['ヘッドスパ', 30], ['シャンプーブロー', 30], ['セット', 60], ['メイク', 60],
];

/* =========================================================
   公開する関数（画面から呼ばれる）
   ========================================================= */

/** Webアプリとして開かれたときに画面を返す */
function doGet() {
  const t = HtmlService.createTemplateFromFile('index');
  t.shopName = SHOP_NAME; // index.html の <?= shopName ?> に入る
  return t.evaluate()
    .setTitle(SHOP_NAME + ' 予約台帳')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-title', '予約台帳');
}

/** index.html の中で <?!= include('css/tokens'); ?> のように使い、別ファイルを埋め込む */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** 最初に1回だけ、エディタから実行してシートを作る（getData からも自動で呼ばれます） */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(key => {
    let sh = ss.getSheetByName(SHEETS[key]);
    if (!sh) {
      sh = ss.insertSheet(SHEETS[key]);
      sh.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // 日付「2026-09-28」や時刻「10:30」を、スプレッドシートが勝手に日付型に変えないよう文字列扱いにする
  sheet_('bookings').getRange('B:C').setNumberFormat('@');
  sheet_('staff').getRange('A:B').setNumberFormat('@');
  if (sheet_('staff').getLastRow() < 2) sheet_('staff').getRange(2, 1, DEFAULT_STAFF.length, 3).setValues(DEFAULT_STAFF);
  if (sheet_('menus').getLastRow() < 2) sheet_('menus').getRange(2, 1, DEFAULT_MENUS.length, 2).setValues(DEFAULT_MENUS);
}

/** 画面を開いたとき：from（YYYY-MM-DD）以降の予約と設定をまとめて返す */
function getData(from) {
  if (!SpreadsheetApp.getActive().getSheetByName(SHEETS.bookings)) {
    withLock_(setup);
  }
  return snapshot_(from);
}

/** 定期確認：前回から何か変わっていれば、新しいデータを返す */
function getUpdates(version, from) {
  if (String(version) === currentVersion_()) return { changed: false };
  return snapshot_(from);
}

/** 予約を入れる */
function addBooking(item, from) {
  return withLock_(() => {
    const b = cleanBooking_(item);
    assertFree_(b, null);
    const now = new Date();
    sheet_('bookings').appendRow([
      b.id, b.date, b.start, b.dur, b.staff, safeText_(b.staffName), safeText_(b.name),
      safeText_(b.menus.join(MENU_SEP)), safeText_(b.memo), STATUS_OK, now, now,
    ]);
    touchCustomer_(b.name);
  }, from);
}

/** 予約の日時・担当を変える */
function moveBooking(id, patch, from) {
  return withLock_(() => {
    const sh = sheet_('bookings');
    const row = findRow_(sh, id);
    const cur = rowToBooking_(sh.getRange(row, 1, 1, HEADERS.bookings.length).getValues()[0]);
    if (cur.status === STATUS_CANCELED) throw new Error('この予約はほかの端末で取り消されています。');
    const b = cleanBooking_(Object.assign({}, cur, patch));
    assertFree_(b, id);
    sh.getRange(row, COL.date, 1, 5).setValues([[b.date, b.start, b.dur, b.staff, safeText_(b.staffName)]]);
    sh.getRange(row, COL.updated).setValue(new Date());
  }, from);
}

/** 予約を取り消す（行は残して状態を「取消」にする） */
function cancelBooking(id, from) {
  return withLock_(() => {
    const sh = sheet_('bookings');
    const row = findRow_(sh, id);
    sh.getRange(row, COL.status).setValue(STATUS_CANCELED);
    sh.getRange(row, COL.updated).setValue(new Date());
  }, from);
}

/** 担当者・メニューの設定を保存する */
function saveSettings(v, from) {
  return withLock_(() => {
    const staff = (v.staff || []).map(s => [String(s.id), safeText_(String(s.name).trim()), Number(s.color) || 0]);
    const menus = (v.menus || []).map(m => [safeText_(String(m.name).trim()), Math.max(30, Number(m.min) || 30)]);
    if (!staff.length) throw new Error('担当者は1人以上必要です。');
    if (staff.some(s => !s[1])) throw new Error('名前が空の担当者がいます。');

    // 画面側でも確認していますが、ほかの端末で予約が入った可能性があるので、ここでも確認する
    const keep = staff.map(s => s[0]);
    const today = today_();
    const future = readBookings_(today);
    readStaff_().filter(s => !keep.includes(s.id)).forEach(s => {
      const n = future.filter(b => b.staff === s.id).length;
      if (n) throw new Error(`${s.name}さんには今日以降の予約が${n}件あります。先に移動か取り消しをしてください。`);
    });

    writeRows_(sheet_('staff'), staff, 3);
    writeRows_(sheet_('menus'), menus, 2);
  }, from);
}

/** スプレッドシートを人が直接編集したときにも、各端末へ変更を知らせる（シンプルトリガー） */
function onEdit() {
  bumpVersion_();
}

/* =========================================================
   内部用の関数
   ========================================================= */

function sheet_(key) {
  return SpreadsheetApp.getActive().getSheetByName(SHEETS[key]);
}

/**
 * 書き込みは必ずこの中で行う。
 * ロック：2台の端末が同時に保存しても、1台ずつ順番に処理される
 * 終わったら「版番号」を進めて、ほかの端末に変更を知らせる
 */
function withLock_(fn, from) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('ほかの端末が保存中です。少し待ってからもう一度押してください。');
  try {
    fn();
    SpreadsheetApp.flush();
    bumpVersion_();
  } finally {
    lock.releaseLock();
  }
  return from === undefined ? null : snapshot_(from);
}

function currentVersion_() {
  return PropertiesService.getScriptProperties().getProperty('version') || '0';
}
function bumpVersion_() {
  PropertiesService.getScriptProperties().setProperty('version', String(Date.now()));
}

/** 画面に返すデータ一式。google.script.run は Date 型を返せないので、日時は数値か文字列にする */
function snapshot_(from) {
  from = from || today_();
  return {
    changed: true,
    version: currentVersion_(),
    from: from,
    settings: { staff: readStaff_(), menus: readMenus_() },
    customers: readCustomers_(),
    bookings: readBookings_(from).map(b => ({
      id: b.id, date: b.date, start: b.start, dur: b.dur, staff: b.staff, staffName: b.staffName,
      name: b.name, menus: b.menus, memo: b.memo,
    })),
  };
}

function values_(sh, width) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, width).getValues() : [];
}

function readBookings_(from) {
  return values_(sheet_('bookings'), HEADERS.bookings.length)
    .map(rowToBooking_)
    .filter(b => b.id && b.status !== STATUS_CANCELED && (!from || b.date >= from));
}

function rowToBooking_(r) {
  return {
    id: String(r[COL.id - 1]),
    date: fmtDate_(r[COL.date - 1]),
    start: fmtTime_(r[COL.start - 1]),
    dur: Number(r[COL.dur - 1]) || 30,
    staff: String(r[COL.staff - 1]),
    staffName: unsafe_(r[COL.staffName - 1]),
    name: unsafe_(r[COL.name - 1]),
    menus: unsafe_(r[COL.menus - 1]).split(MENU_SEP).filter(String),
    memo: unsafe_(r[COL.memo - 1]),
    status: String(r[COL.status - 1]),
  };
}

function readStaff_() {
  return values_(sheet_('staff'), 3).filter(r => r[0] !== '')
    .map(r => ({ id: String(r[0]), name: unsafe_(r[1]), color: Number(r[2]) || 0 }));
}
function readMenus_() {
  return values_(sheet_('menus'), 2).filter(r => r[0] !== '')
    .map(r => ({ name: unsafe_(r[0]), min: Number(r[1]) || 30 }));
}
function readCustomers_() {
  return values_(sheet_('customers'), 3).filter(r => r[0] !== '')
    .map(r => ({ name: unsafe_(r[0]), count: Number(r[1]) || 1, last: r[2] instanceof Date ? r[2].getTime() : 0 }))
    .sort((a, b) => b.last - a.last)
    .slice(0, 400);
}

function touchCustomer_(name) {
  const sh = sheet_('customers');
  const cell = findCell_(sh, name);
  if (cell) {
    const row = cell.getRow();
    const count = Number(sh.getRange(row, 2).getValue()) || 0;
    sh.getRange(row, 2, 1, 2).setValues([[count + 1, new Date()]]);
  } else {
    sh.appendRow([safeText_(name), 1, new Date()]);
  }
}

/** 1列目が id と完全に一致する行を探す（TextFinder は全行を読み込むより速い） */
function findCell_(sh, text) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  return sh.getRange(2, 1, n, 1).createTextFinder(text).matchEntireCell(true).findNext();
}
function findRow_(sh, id) {
  const cell = findCell_(sh, String(id));
  if (!cell) throw new Error('予約が見つかりません。ほかの端末で変更された可能性があります。');
  return cell.getRow();
}

/** 設定シートを丸ごと書き直す */
function writeRows_(sh, rows, width) {
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

/** 画面から来た値を信用しすぎないよう、形を整えて確認する */
function cleanBooking_(x) {
  const b = {
    id: String(x.id || Utilities.getUuid()),
    date: String(x.date || ''),
    start: String(x.start || ''),
    dur: Number(x.dur),
    staff: String(x.staff || ''),
    staffName: String(x.staffName || ''),
    name: String(x.name || '').trim().slice(0, 60),
    menus: (x.menus || []).map(String).slice(0, 20),
    memo: String(x.memo || '').slice(0, 200),
  };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.date)) throw new Error('日付の形が正しくありません。');
  if (!/^\d{2}:\d{2}$/.test(b.start)) throw new Error('開始時刻の形が正しくありません。');
  if (!(b.dur > 0 && b.dur <= 600)) throw new Error('かかる時間が正しくありません。');
  if (!b.name) throw new Error('お客様の名前が空です。');
  if (!readStaff_().some(s => s.id === b.staff)) throw new Error('担当者が見つかりません。設定が変わった可能性があります。');
  return b;
}

/** 同じ担当者の予約と時間が重なっていないか（ロックの中で確認するので、同時に押しても二重予約にならない） */
function assertFree_(b, ignoreId) {
  const s = toMin_(b.start), e = s + b.dur;
  const hit = readBookings_(b.date).some(x =>
    x.date === b.date && x.staff === b.staff && x.id !== ignoreId &&
    s < toMin_(x.start) + x.dur && toMin_(x.start) < e);
  if (hit) throw new Error('ちょうど今、ほかの端末でこの時間に予約が入りました。別の枠を選んでください。');
}

function toMin_(hm) { const p = String(hm).split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v); }
function fmtTime_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v).padStart(5, '0'); }

/**
 * 名前が「=」「+」などで始まると、スプレッドシートが数式として扱ってしまう。
 * 先頭に ' を付けて、ただの文字として保存する（読むときに外す）
 */
function safeText_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function unsafe_(s) { s = String(s == null ? '' : s); return s.charAt(0) === "'" ? s.slice(1) : s; }
