/**
 * ALL IN Retail Portal — Google Apps Script backend (v3)
 * 綁定試算表：補貨行事曆 1ezhFEL3CTx-ToU8HJa3wmJY1dyQPvNP7eJc3Mw4fGBk
 * （2026-11-24 起；舊檔 1o_7UsAY8HMomdbbp44vXgsyFQ-eTGfCBCqJ2j1flYLI 已停用）
 *
 * 兩種送出模式，更新同一個 [retail] 店別區塊：
 *   count   庫存盤點（早班）→ 效期、櫃位、倉庫　＋ 標題列「盤點日期／盤點人／更新」
 *   arrival 當日進貨（到貨）→ 進(櫃)、進(倉)　＋ 標題列「進貨日期／收貨人／更新」
 * Bender 第 11–14 列以 SUMIFS 讀取 J 欄合計（櫃位＋倉庫）。
 * 每次送出追加明細至 retail_log。
 * 庫存盤點送出後，另把各品項合計寫入當週分頁（如 1123-1129）送出當天的欄位區塊（見 writeCalendar_）。
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

  const out = { mode: mode, updated: now, rows: log.length };
  if (mode === 'count') {
    // 寫入週分頁失敗不影響盤點送出，只在回應中註明
    try {
      out.calendar = writeCalendar_(ss, body.store, rows, new Date());
    } catch (err) {
      out.calendarError = String(err && err.message ? err.message : err);
      console.error('writeCalendar_ ' + body.store + ': ' + out.calendarError);
    }
  }
  return out;
}

/* ───────────── 週分頁（補貨行事曆）自動填入 ─────────────
 * 分頁名稱 MMDD-MMDD（如 1123-1129）。第 1 列每個日區塊有星期（一～六），
 * 區塊起點 = 星期欄 − 1：起點欄為櫃位／區塊名稱，+1 系列，+2 效期列標籤，+3 起為品項欄。
 * 品項欄依表頭字母對應（全食飲在櫃位列：S R A B；其他系列在系列列上一列：P E U R B）。
 * 寫入值 = 櫃位＋倉庫（大安分「倉庫」「店面」列：倉庫＝倉庫、店面＝櫃位）。
 * 以 portal 為準，覆蓋既有數字；portal 空白就清空該格。週日沒有區塊，不寫入。
 */
const CAL_STORE = { '信義': '新光A4', '天母': '大葉', '竹北': '竹北', '中港': '台中', '大安': '大安' };
const CAL_SERIES = ['全食飲', '小全頭', '全方麵', '全食圈'];
const CAL_CODE = {
  '全食飲': { S: '銀耳黃豆', R: '紅棗核桃', A: '杏仁糙米', B: '黑五次方' },
  '小全頭': { P: '南瓜', E: '毛豆', U: '紫心甘薯', R: '紅棗亞麻', B: '芝麻黑豆' },
  '全方麵': { P: '南瓜', E: '毛豆' },
  '全食圈': { P: '南瓜', U: '紫心甘薯' }
};
const WEEKDAY = ['日', '一', '二', '三', '四', '五', '六'];

function norm_(v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ''); }

/** 找出包含 date（台北時間）的週分頁 */
function weekSheet_(ss, date) {
  const y = Number(Utilities.formatDate(date, 'Asia/Taipei', 'yyyy'));
  const md = Number(Utilities.formatDate(date, 'Asia/Taipei', 'MMdd'));
  const hits = ss.getSheets().filter(function (s) {
    const m = /^(\d{4})-(\d{4})$/.exec(s.getName().trim());
    if (!m) return false;
    const a = Number(m[1]), b = Number(m[2]);
    return a <= b ? (md >= a && md <= b) : (md >= a || md <= b);   // 跨年週（如 1228-0103）
  });
  if (!hits.length) throw new Error('找不到包含 ' + Utilities.formatDate(date, 'Asia/Taipei', 'MM/dd') + ' 的週分頁（名稱格式 MMDD-MMDD）');
  return hits[0];
}

/** 由 [retail] 店別區塊列（A..H）建立 系列|商品|效期別 → {shelf, wh} */
function portalMap_(rows) {
  const m = {};
  rows.forEach(function (r) { m[[r[0], r[1], r[2]].join('|')] = { shelf: r[C.shelf - 1], wh: r[C.wh - 1] }; });
  return m;
}

function sum_() {
  let n = 0, any = false;
  for (let i = 0; i < arguments.length; i++) {
    const v = arguments[i];
    if (v === '' || v === null || v === undefined) continue;
    n += Number(v); any = true;
  }
  return any ? n : '';
}

/**
 * 計算一個店別要寫入的儲存格（純運算，可在 Node 測試）。
 * grid：週分頁值（二維陣列，0-based），rows：[retail] 店別區塊列，store：portal 櫃位，weekday：0–6。
 * 回傳 [{r, c, v}]（0-based）。
 */
function calendarCells_(grid, rows, store, weekday) {
  const label = CAL_STORE[store];
  if (!label) throw new Error('週分頁沒有對應櫃位：' + store);
  if (weekday === 0) return { skipped: '週日沒有日區塊', cells: [] };

  const top = grid[0] || [];
  let wcol = -1;
  for (let c = 0; c < top.length; c++) if (norm_(top[c]) === WEEKDAY[weekday]) { wcol = c; break; }
  if (wcol < 1) throw new Error('第 1 列找不到星期' + WEEKDAY[weekday] + '的區塊');
  const c0 = wcol - 1, cSeries = c0 + 1, cTier = c0 + 2;

  // 店別範圍：起點欄 = label 的列，到下一個非空白起點欄為止
  let r0 = -1, r1 = grid.length;
  for (let r = 1; r < grid.length; r++) {
    const v = norm_((grid[r] || [])[c0]);
    if (r0 < 0) { if (v === label) r0 = r; }
    else if (v !== '') { r1 = r; break; }
  }
  if (r0 < 0) throw new Error('星期' + WEEKDAY[weekday] + '區塊找不到櫃位「' + label + '」');

  const at = function (r, c) { return (grid[r] || [])[c]; };
  const pm = portalMap_(rows);
  const daan = store === '大安';
  const cells = [];

  // 各系列起始列
  const seriesRow = {};
  for (let r = r0; r < r1; r++) {
    const s = norm_(at(r, cSeries));
    if (CAL_SERIES.indexOf(s) >= 0 && seriesRow[s] === undefined) seriesRow[s] = r;
  }

  // 表頭字母 → 欄
  const codeCols = function (hr, series) {
    const out = [];
    for (let c = c0 + 3; c < c0 + 12; c++) {
      const code = norm_(at(hr, c));
      if (CAL_CODE[series][code]) out.push({ c: c, product: CAL_CODE[series][code] });
    }
    return out;
  };

  CAL_SERIES.forEach(function (series) {
    const sr = seriesRow[series];
    if (sr === undefined) return;
    const end = CAL_SERIES.map(function (s) { return seriesRow[s]; })
      .filter(function (x) { return x !== undefined && x > sr; })
      .reduce(function (a, b) { return Math.min(a, b); }, r1);
    const cols = codeCols(series === '全食飲' ? r0 : sr - 1, series);
    if (!cols.length) return;

    for (let r = sr; r < end; r++) {
      const tier = norm_(at(r, cTier));
      cols.forEach(function (k) {
        const pOld = pm[[series, k.product, '舊'].join('|')] || {};
        const pNew = pm[[series, k.product, '新'].join('|')] || {};
        const p = pm[[series, k.product, ''].join('|')] || {};
        let v;
        if (series === '全食飲') {
          if (tier === 'old') v = sum_(pOld.shelf, pOld.wh);
          else if (tier === 'new') v = sum_(pNew.shelf, pNew.wh);
          else if (daan && tier === '倉庫-old') v = sum_(pOld.wh);
          else if (daan && tier === '倉庫-new') v = sum_(pNew.wh);
          else if (daan && tier === '店面') v = sum_(pOld.shelf, pNew.shelf);
          else return;
        } else {
          if (r === sr && (tier === '' || tier === '·' || tier === '倉庫')) v = daan ? sum_(p.wh) : sum_(p.shelf, p.wh);
          else if (daan && r === sr + 1 && tier === '店面') v = sum_(p.shelf);
          else return;
        }
        cells.push({ r: r, c: k.c, v: v });
      });
    }
  });
  if (!cells.length) throw new Error('「' + label + '」區塊沒有可寫入的欄位');
  return { cells: cells };
}

function writeCalendar_(ss, store, rows, now) {
  const sh = weekSheet_(ss, now);
  const weekday = Number(Utilities.formatDate(now, 'Asia/Taipei', 'u')) % 7;   // u: 1=一 … 7=日
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const grid = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const res = calendarCells_(grid, rows, store, weekday);
  if (res.skipped) return { sheet: sh.getName(), skipped: res.skipped };

  // 同列連續欄合併成一次寫入
  const byRow = {};
  res.cells.forEach(function (x) { (byRow[x.r] = byRow[x.r] || []).push(x); });
  Object.keys(byRow).forEach(function (r) {
    const xs = byRow[r].sort(function (a, b) { return a.c - b.c; });
    let i = 0;
    while (i < xs.length) {
      let j = i;
      while (j + 1 < xs.length && xs[j + 1].c === xs[j].c + 1) j++;
      sh.getRange(Number(r) + 1, xs[i].c + 1, 1, j - i + 1)
        .setValues([xs.slice(i, j + 1).map(function (x) { return x.v; })]);
      i = j + 1;
    }
  });
  return { sheet: sh.getName(), weekday: WEEKDAY[weekday], cells: res.cells.length };
}

/** 手動測試用：在 Apps Script 編輯器執行，把 [retail] 目前各店資料寫入今天的區塊 */
function backfillToday() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RETAIL);
  STORES.forEach(function (store) {
    const b = block_(sh, store);
    const t = sh.getRange(b.titleRow, T.countDate).getDisplayValue();
    const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
    if (t !== today) { console.log(store + '：今天尚未盤點（' + (t || '—') + '），略過'); return; }
    const rows = sh.getRange(b.firstRow, C.series, b.n, C.wh).getValues();
    console.log(store + '：' + JSON.stringify(writeCalendar_(ss, store, rows, new Date())));
  });
}
