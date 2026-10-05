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
const SHEETS = { bookings: '予約', customers: 'お客様', staff: '担当者', menus: 'メニュー', hours: '営業日' };
const HEADERS = {
  bookings:  ['ID', '日付', '開始', '分', '担当ID', '担当名', 'お客様', 'メニュー', 'メモ', '状態', '作成日時', '更新日時'],
  customers: ['名前', '来店回数', '最終予約日時'],
  staff:     ['ID', '名前', '色'],
  menus:     ['名前', '分'],
  hours:     ['種類', '曜日・日付', '開店', '閉店', 'メモ'],
};
// 予約シートの列番号（1始まり）
const COL = { id: 1, date: 2, start: 3, dur: 4, staff: 5, staffName: 6, name: 7, menus: 8, memo: 9, status: 10, created: 11, updated: 12 };
const MENU_SEP = '、';
const STATUS_OK = '予約';
const STATUS_CANCELED = '取消';   // 行は消さずに「取消」にする（あとから履歴を確認できるように）
const TZ = Session.getScriptTimeZone();

// 最初の担当者（公開用の仮の名前です。実際の名前は画面の「お店の設定」から変えてください）
const DEFAULT_STAFF = [['staff-a', 'スタッフA', 0], ['staff-b', 'スタッフB', 1]];
// 曜日ごとの営業時間（最初の値）。'休み' は定休日
const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];
const DEFAULT_HOURS = [['10:00', '19:00'], ['10:00', '20:00'], ['10:00', '20:00'], ['休み', ''],
                       ['10:00', '20:00'], ['10:00', '20:00'], ['10:00', '20:00']]; // 日〜土
const KIND_WEEK = '曜日';
const KIND_CLOSED = '臨時休業';
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
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
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
  sheet_('hours').getRange('A:E').setNumberFormat('@');
  if (sheet_('hours').getLastRow() < 2) {
    sheet_('hours').getRange(2, 1, 7, 5).setValues(WEEKDAYS.map((w, i) => [KIND_WEEK, w, DEFAULT_HOURS[i][0], DEFAULT_HOURS[i][1], '']));
  }
  if (sheet_('staff').getLastRow() < 2) sheet_('staff').getRange(2, 1, DEFAULT_STAFF.length, 3).setValues(DEFAULT_STAFF);
  if (sheet_('menus').getLastRow() < 2) sheet_('menus').getRange(2, 1, DEFAULT_MENUS.length, 2).setValues(DEFAULT_MENUS);
}

/** 画面を開いたとき：from（YYYY-MM-DD）以降の予約と設定をまとめて返す */
function getData(from) {
  // シートが足りなければ作る（あとから「営業日」シートを足したときも、ここで自動で作られる）
  const ss = SpreadsheetApp.getActive();
  if (Object.keys(SHEETS).some(k => !ss.getSheetByName(SHEETS[k]))) {
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

/** 担当者・メニュー・営業日の設定を保存する */
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

    // 営業日：休みにする日・営業時間外になる時間に、すでに予約が入っていないか
    const hours = v.hours ? cleanHours_(v.hours) : readHours_();
    const outside = future.filter(b => !withinHours_(b, hours));
    if (outside.length) {
      const ex = outside.slice(0, 3).map(b => `${Number(b.date.slice(5, 7))}/${Number(b.date.slice(8))} ${b.start} ${b.name}様`).join('、');
      throw new Error(`休みや営業時間外にする時間に予約が${outside.length}件あります（${ex}${outside.length > 3 ? ' など' : ''}）。先に移動か取り消しをしてください。`);
    }

    writeRows_(sheet_('staff'), staff, 3);
    writeRows_(sheet_('menus'), menus, 2);
    writeRows_(sheet_('hours'), hoursToRows_(hours), 5);
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
    settings: { staff: readStaff_(), menus: readMenus_(), hours: readHours_() },
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

/**
 * 営業日を画面と同じ形で返す
 *   week:   { 0: [10, 19], 3: null, ... }  曜日ごとの [開店, 閉店]（時。10.5 は 10:30）。null は定休日
 *   closed: [{ date: '2026-12-31', memo: '年末年始' }, ...]  臨時休業日
 */
function readHours_() {
  const week = {};
  WEEKDAYS.forEach((w, i) => {
    week[i] = DEFAULT_HOURS[i][0] === '休み' ? null : [hmToHour_(DEFAULT_HOURS[i][0]), hmToHour_(DEFAULT_HOURS[i][1])];
  });
  const closed = [];
  values_(sheet_('hours'), 5).forEach(r => {
    const kind = String(r[0]);
    if (kind === KIND_WEEK) {
      const i = WEEKDAYS.indexOf(String(r[1]));
      if (i < 0) return;
      const open = fmtTime_(r[2]), close = fmtTime_(r[3]);
      week[i] = /^\d{2}:\d{2}$/.test(open) && /^\d{2}:\d{2}$/.test(close) ? [hmToHour_(open), hmToHour_(close)] : null;
    } else if (kind === KIND_CLOSED) {
      closed.push({ date: fmtDate_(r[1]), memo: unsafe_(r[4]) });
    }
  });
  closed.sort((a, b) => (a.date < b.date ? -1 : 1));
  return { week: week, closed: closed };
}

/** 画面から来た営業日を確認して整える。過去の臨時休業日は捨てる */
function cleanHours_(h) {
  const week = {};
  for (let i = 0; i < 7; i++) {
    const v = h.week ? h.week[i] : null;
    if (!v) { week[i] = null; continue; }
    const o = Number(v[0]), c = Number(v[1]);
    if (!(o >= 5 && c <= 24 && o < c && o * 2 === Math.round(o * 2) && c * 2 === Math.round(c * 2))) {
      throw new Error(`${WEEKDAYS[i]}曜日の営業時間が正しくありません。`);
    }
    week[i] = [o, c];
  }
  const today = today_(), seen = {};
  const closed = (h.closed || [])
    .map(c => ({ date: String(c.date), memo: String(c.memo || '').slice(0, 40) }))
    .filter(c => /^\d{4}-\d{2}-\d{2}$/.test(c.date) && c.date >= today && !seen[c.date] && (seen[c.date] = true))
    .sort((a, b) => (a.date < b.date ? -1 : 1));
  return { week: week, closed: closed };
}

function hoursToRows_(h) {
  const rows = WEEKDAYS.map((w, i) => h.week[i]
    ? [KIND_WEEK, w, hourToHm_(h.week[i][0]), hourToHm_(h.week[i][1]), '']
    : [KIND_WEEK, w, '休み', '', '']);
  h.closed.forEach(c => rows.push([KIND_CLOSED, c.date, '', '', safeText_(c.memo)]));
  return rows;
}

/** その日の営業時間を [開店分, 閉店分] で返す。休みなら null */
function openMinutes_(date, h) {
  if (h.closed.some(c => c.date === date)) return null;
  const p = date.split('-').map(Number);
  const w = h.week[new Date(p[0], p[1] - 1, p[2]).getDay()];
  return w ? [w[0] * 60, w[1] * 60] : null;
}
function withinHours_(b, h) {
  const m = openMinutes_(b.date, h);
  return !!m && toMin_(b.start) >= m[0] && toMin_(b.start) + b.dur <= m[1];
}
function hmToHour_(hm) { return toMin_(hm) / 60; }
function hourToHm_(h) { const m = Math.round(h * 60); return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }

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
  const open = openMinutes_(b.date, readHours_());
  if (!open) throw new Error('この日はお休みです。別の日を選んでください。');
  const s = toMin_(b.start), e = s + b.dur;
  if (s < open[0] || e > open[1]) throw new Error('営業時間の外です。時間を短くするか、別の枠を選んでください。');
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
