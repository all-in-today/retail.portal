/**
 * ALL IN Retail Portal — Google Apps Script backend (v2)
 * 綁定試算表：1o_7UsAY8HMomdbbp44vXgsyFQ-eTGfCBCqJ2j1flYLI
 *
 * 兩種送出模式，更新同一個 [retail] 店別區塊：
 *   count   庫存盤點（早班）→ 效期、櫃位、倉庫　＋ 標題列「盤點日期／盤點人／更新」
 *   arrival 當日進貨（到貨）→ 進(櫃)、進(倉)　＋ 標題列「進貨日期／收貨人／更新」
 * Bender 第 11–14 列以 SUMIFS 讀取 J 欄合計（櫃位＋倉庫）。
 * 每次送出追加明細至 retail_log。
 *
 * 更新程式後：部署 → 管理部署作業 → 編輯（鉛筆）→ 版本：新版本 → 部署（網址不變）
 * 指令碼屬性：PORTAL_PIN
 */

const SHEET_RETAIL = '[retail]';
const SHEET_LOG = 'retail_log';
const STORES = ['信義', '天母', '竹北', '中港', '大安'];
// 欄位（1-based）
const C = { series: 1, product: 2, tier: 3, expiry: 4, inShelf: 5, inWh: 6, shelf: 7, wh: 8, min: 9, total: 10 };
// 標題列欄位
const T = { countDate: 3, countStaff: 5, countAt: 7, arrDate: 9, arrStaff: 11, arrAt: 13 };

function doGet(e) {
  return handle_(function () {
    const p = e.parameter || {};
    checkPin_(p.pin);
    if (p.action === 'load') return load_(p.store);
    return { stores: STORES };
  });
}

function doPost(e) {
  return handle_(function () {
    const body = JSON.parse(e.postData.contents);
    checkPin_(body.pin);
    const lock = LockService.getDocumentLock();
    lock.waitLock(20000);
    try {
      return save_(body);
    } finally {
      lock.releaseLock();
    }
  });
}

function handle_(fn) {
  let out;
  try {
    out = Object.assign({ ok: true }, fn());
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkPin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('PORTAL_PIN');
  if (expected && String(pin || '') !== String(expected)) throw new Error('PIN 錯誤');
}

/** 店別區塊：標題列 A=店名、下一列 A=系列；品項列直到 A 欄空白 */
function block_(sh, store) {
  if (STORES.indexOf(store) < 0) throw new Error('未知櫃位：' + store);
  const a = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; });
  for (let i = 0; i < a.length - 1; i++) {
    if (a[i] === store && a[i + 1] === '系列') {
      let j = i + 2;
      while (j < a.length && a[j] !== '') j++;
      return { titleRow: i + 1, firstRow: i + 3, n: j - (i + 2) };
    }
  }
  throw new Error('[retail] 找不到櫃位區塊：' + store);
}

function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function load_(store) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RETAIL);
  const b = block_(sh, store);
  const t = sh.getRange(b.titleRow, 1, 1, 13).getValues()[0];
  const rows = sh.getRange(b.firstRow, 1, b.n, C.min).getValues();
  return {
    store: store,
    count: { date: str_(t[T.countDate - 1]), staff: str_(t[T.countStaff - 1]), at: str_(t[T.countAt - 1]) },
    arrival: { date: str_(t[T.arrDate - 1]), staff: str_(t[T.arrStaff - 1]), at: str_(t[T.arrAt - 1]) },
    lines: rows.map(function (r) {
      return {
        series: r[0], product: r[1], tier: r[2], expiry: str_(r[3]),
        inShelf: r[4], inWh: r[5], shelf: r[6], wh: r[7], min: r[8]
      };
    })
  };
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  if (isNaN(n) || n < 0) throw new Error('數量格式錯誤：' + v);
  return n;
}

function save_(body) {
  const mode = body.mode;
  if (mode !== 'count' && mode !== 'arrival') throw new Error('未知模式');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(body.date || '')) throw new Error('日期格式錯誤');
  if (!body.staff) throw new Error('請填寫姓名');

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RETAIL);
  const b = block_(sh, body.store);
  const range = sh.getRange(b.firstRow, C.series, b.n, C.wh);   // A..H
  const rows = range.getValues();
  const input = {};
  (body.lines || []).forEach(function (l) { input[[l.series, l.product, l.tier || ''].join('|')] = l; });

  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  const label = mode === 'count' ? '庫存盤點' : '當日進貨';
  const log = [];

  rows.forEach(function (r) {
    const l = input[[r[0], r[1], r[2]].join('|')] || {};
    if (mode === 'count') {
      r[C.expiry - 1] = l.expiry ? String(l.expiry) : '';
      r[C.shelf - 1] = num_(l.shelf);
      r[C.wh - 1] = num_(l.wh);
      log.push([now, label, body.date, body.store, body.staff, r[0], r[1], r[2], r[3], '', '', r[6], r[7]]);
    } else {
      if (l.expiry) r[C.expiry - 1] = String(l.expiry);   // 進貨只補效期，不清除盤點效期
      r[C.inShelf - 1] = num_(l.inShelf);
      r[C.inWh - 1] = num_(l.inWh);
      if (r[4] !== '' || r[5] !== '') log.push([now, label, body.date, body.store, body.staff, r[0], r[1], r[2], r[3], r[4], r[5], '', '']);
    }
  });

  sh.getRange(b.firstRow, C.expiry, b.n, 1).setNumberFormat('@');
  sh.getRange(b.firstRow, C.inShelf, b.n, 4).setNumberFormat('0');
  range.setValues(rows);

  const t = mode === 'count' ? [T.countDate, T.countStaff, T.countAt] : [T.arrDate, T.arrStaff, T.arrAt];
  sh.getRange(b.titleRow, t[0]).setNumberFormat('@').setValue(body.date);
  sh.getRange(b.titleRow, t[1]).setValue(body.staff);
  sh.getRange(b.titleRow, t[2]).setNumberFormat('@').setValue(now);

  let lg = ss.getSheetByName(SHEET_LOG);
  if (!lg) lg = ss.insertSheet(SHEET_LOG);
  if (lg.getLastRow() === 0) {
    lg.appendRow(['送出時間', '類型', '日期', '櫃位', '姓名', '系列', '商品', '效期別', '效期', '進(櫃)', '進(倉)', '櫃位', '倉庫']);
    lg.setFrozenRows(1);
  }
  if (log.length) lg.getRange(lg.getLastRow() + 1, 1, log.length, log[0].length).setNumberFormat('@').setValues(log);

  return { mode: mode, updated: now, rows: log.length };
}
