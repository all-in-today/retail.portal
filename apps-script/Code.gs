/**
 * ALL IN Retail Portal — Google Apps Script backend
 * 綁定試算表：1o_7UsAY8HMomdbbp44vXgsyFQ-eTGfCBCqJ2j1flYLI
 *
 * 寫入 [retail] 分頁（五店區塊），Bender 第 11–14 列以 SUMIFS 自動彙總。
 * 每次送出同時追加一份明細至 retail_log 分頁，供店長差異核對與歷史追溯。
 *
 * 部署：擴充功能 → Apps Script → 貼上本檔 → 部署 → 新增部署作業 → 網頁應用程式
 *       執行身分：我　／　存取權：任何人
 * 指令碼屬性（專案設定 → 指令碼屬性）：PORTAL_PIN = 櫃位共用 PIN（未設定則不驗證）
 */

const SHEET_RETAIL = '[retail]';
const SHEET_LOG = 'retail_log';
const STORES = ['信義', '天母', '竹北', '中港', '大安'];
const COL = { series: 1, product: 2, tier: 3, expiry: 4, inShelf: 5, inWh: 6, shelf: 7, wh: 8, min: 9, total: 10, tasting: 11, use1: 12 };
const USE_COLS = 10;

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

/** 找到店別區塊：回傳 {titleRow, firstRow, lastRow, usesRow}（1-based） */
function block_(sh, store) {
  if (STORES.indexOf(store) < 0) throw new Error('未知櫃位：' + store);
  const colA = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] === store && colA[i][1] === '歸屬日期') {
      const titleRow = i + 1;
      for (let j = i + 2; j < colA.length; j++) {
        if (colA[j][0] === '用途說明') {
          return { titleRow: titleRow, firstRow: titleRow + 2, lastRow: j, usesRow: j + 1 };
        }
      }
    }
  }
  throw new Error('[retail] 找不到櫃位區塊：' + store);
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  return v ? String(v) : '';
}

function load_(store) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RETAIL);
  const b = block_(sh, store);
  const title = sh.getRange(b.titleRow, 1, 1, 7).getValues()[0];
  const n = b.lastRow - b.firstRow + 1;
  const rows = sh.getRange(b.firstRow, 1, n, COL.use1 + USE_COLS - 1).getValues();
  const reasons = sh.getRange(b.usesRow, COL.use1, 1, USE_COLS).getValues()[0];
  const lines = rows.map(function (r) {
    return {
      series: r[0], product: r[1], tier: r[2],
      expiry: r[3] === '' ? '' : String(r[3]),
      inShelf: r[4], inWh: r[5], shelf: r[6], wh: r[7], min: r[8], tasting: r[10],
      uses: r.slice(COL.use1 - 1, COL.use1 - 1 + USE_COLS)
    };
  });
  return {
    store: store,
    meta: { date: fmtDate_(title[2]), staff: title[4], updated: fmtDate_(title[6]) },
    reasons: reasons,
    lines: lines
  };
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  if (isNaN(n)) throw new Error('非數字：' + v);
  return n;
}

function save_(body) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RETAIL);
  const b = block_(sh, body.store);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(body.date || '')) throw new Error('歸屬日期格式錯誤');
  if (!body.staff) throw new Error('請填寫盤點人');

  // 用途 → 欄位 1–10
  const uses = (body.uses || []).filter(function (u) { return u.reason; });
  if (uses.length > USE_COLS) throw new Error('用途最多 ' + USE_COLS + ' 項');
  const reasons = [];
  for (let i = 0; i < USE_COLS; i++) reasons.push(uses[i] ? uses[i].reason : '');

  const n = b.lastRow - b.firstRow + 1;
  const keys = sh.getRange(b.firstRow, 1, n, 3).getValues().map(function (r) { return r.join('|'); });
  const input = {};
  (body.lines || []).forEach(function (l) { input[[l.series, l.product, l.tier || ''].join('|')] = l; });

  const left = [];   // D..H
  const right = [];  // K..U
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  const log = [];
  keys.forEach(function (k) {
    const l = input[k] || {};
    const useVals = [];
    for (let i = 0; i < USE_COLS; i++) {
      const u = uses[i];
      useVals.push(u && u.items && u.items[k] !== undefined ? num_(u.items[k]) : '');
    }
    const row = [l.expiry ? String(l.expiry) : '', num_(l.inShelf), num_(l.inWh), num_(l.shelf), num_(l.wh)];
    left.push(row);
    right.push([num_(l.tasting)].concat(useVals));
    const parts = k.split('|');
    const useText = useVals.map(function (v, i) { return v === '' ? '' : reasons[i] + '×' + v; }).filter(String).join('、');
    log.push([now, body.date, body.store, body.staff, parts[0], parts[1], parts[2]].concat(row).concat([num_(l.tasting), useText]));
  });

  sh.getRange(b.firstRow, COL.expiry, n, 1).setNumberFormat('@');          // 效期保留 0404 前導零
  sh.getRange(b.firstRow, COL.inShelf, n, 4).setNumberFormat('0');        // 數量欄維持數字，SUM 才算得到
  sh.getRange(b.firstRow, COL.expiry, n, 5).setValues(left);
  sh.getRange(b.firstRow, COL.tasting, n, 1 + USE_COLS).setValues(right);
  sh.getRange(b.usesRow, COL.use1, 1, USE_COLS).setValues([reasons]);
  sh.getRange(b.titleRow, 3).setNumberFormat('@').setValue(body.date);
  sh.getRange(b.titleRow, 5).setValue(body.staff);
  sh.getRange(b.titleRow, 7).setNumberFormat('@').setValue(now);

  let lg = ss.getSheetByName(SHEET_LOG);
  if (!lg) {
    lg = ss.insertSheet(SHEET_LOG);
    lg.appendRow(['送出時間', '歸屬日期', '櫃位', '盤點人', '系列', '商品', '效期別', '效期', '進(櫃)', '進(倉)', '櫃位', '倉庫', '試吃/展示', '用途']);
    lg.setFrozenRows(1);
  }
  lg.getRange(lg.getLastRow() + 1, 1, log.length, log[0].length).setNumberFormat('@').setValues(log);

  return { saved: n, updated: now };
}
