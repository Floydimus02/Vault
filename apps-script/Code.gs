/**
 * Vault: a ledger in a letter
 * Reads one Google Sheet and gives you two views of it: a daily email and a private web app.
 * Every figure is read or calculated from the sheet; nothing is typed in here except who
 * receives the email. https://github.com/Floydimus02/Vault
 *
 * Setup (see README.md in the repository for detail)
 *   1. In your Vault Google Sheet: Extensions → Apps Script. Replace all code in Code.gs with this file.
 *   2. Set RECIPIENTS in CONFIG (blank = only you).
 *   3. Run sendSnapshot once and approve the permissions.
 *   4. Email: Triggers (clock icon) → Add trigger → sendSnapshot, "Time-driven", any time you like.
 *   5. Daily close: add one more trigger → recordDailyClose, "Time-driven", "Day timer", "Midnight to 1am".
 *      It records net worth once a day, after mutual fund NAVs update, and sends nothing.
 *   6. Web app (optional): File → + → HTML, name it Index, paste Index.html into it.
 *      Deploy → New deployment → Web app. Execute as: Me. Who has access: Only myself.
 *      After any later change: Deploy → Manage deployments → edit → Version: New version.
 *   7. Import and export (in the app's menu): run exportVault once from the editor to approve Drive access.
 *      To import Excel files as well, add the Drive API: Services → + → Drive API → Add.
 *
 * Changed the sheet's layout? See PROMPT.md.
 *
 * Copyright (C) 2026 Floydimus.
 * Free software under the GNU General Public License v3.0. See LICENSE.
 * Distributed WITHOUT ANY WARRANTY. Not financial or tax advice.
 *
 * The daily close adds one row a day to a hidden tab (CONFIG.LOG_TAB). Every comparison
 * (since the last close, 7 days, 30 days, movers, markets, the daily-change chart) uses
 * those closes only, so all figures share one fixed time of day. Emails and the web view
 * don't write to the log. Run errors are listed under Executions in Apps Script.
 */

const CONFIG = {
  RECIPIENTS: '',               // Comma-separated. Blank = the script owner.
  TIMEZONE: 'Asia/Kolkata',
  SPREADSHEET_ID: '',           // Blank = the sheet this script is bound to.
  LOG_TAB: '_vault_snapshot_log',
  LOG_KEEP_ROWS: 800,
  TREND_DAYS: 30,
  TOP_MOVERS: 3,
  TOP_XIRR: 3,
  ITR_FILING_DATE: '',          // e.g. '2027-07-31'. Blank = 31 July after the year ends (section 424 interest).
  // Holdings left out of net worth and every other figure, matched on name, class, custodian and vehicle.
  // Stock appreciation rights are paper value until a liquidity event, and even then only the gain is paid.
  EXCLUDE: /stock appreciation right/i,
  CLOSE_LABEL: 'Midnight',        // Shown in the email; match it to the recordDailyClose trigger's time.
  // Holdings shown in "Rewards and balances", matched on vehicle and instrument.
  // Points: write Amount: T as =(points*paise)/100, e.g. =(492*20)/100 for 492 points at 20 paise each.
  // Expiry: put the date in square brackets in Goal/Comments, e.g. "Expenses [31 Mar 2027]".
  REWARDS: /reward|voucher|gift card|pay balance/i,
  // Reward points expire this many years after they're earned. The daily close records each top-up
  // as a dated lot in a hidden tab, and uses up the oldest lots first when the balance falls.
  REWARD_EXPIRY_YEARS: 3,
  REWARD_TAB: '_vault_rewards',
  // EMI account check: from DAYS_BEFORE days ahead of the EMI date up to the day itself, warn if the
  // account below holds less than MIN. Salary lands on the last working day, so 4 days before the
  // 5th (the 1st) is the first day the balance includes it.
  // ACCOUNT matches the Custodian of the account the EMI is debited from, if the loan tab doesn't name it.
  EMI_GUARD: { ACCOUNT: /example bank/i, VEHICLE: /savings/i, MIN: 50000, EMI_DAY: 5, DAYS_BEFORE: 4 },
  TABS: {
    worth: /net worth/i,
    portfolio: /portfolio/i,
    expenses: /expense/i,
    coi: /\bcoi\b|computation of income/i,
    locker: /locker/i,
    payouts: /dividend|payout/i,
    loan: /loan schedule/i,
    ledger: /rewards? ledger/i,
    estimations: /income\s*&\s*estimation/i,
    xirr: /^xirr$/i
  },
  // Web app: how many movers to list for a net worth move of under 0.25%, 1% and more.
  APP: { MOVERS: [[0.0025, 3], [0.01, 5], [Infinity, 10]],
    // Tabs the data checks scan. Tabs outside this list (a travel planner, say) are not the app's business.
    CHECK_TABS: /net worth|expense|\bcoi\b|income\s*&\s*estimation|cash flow|portfolio|^xirr$|dividend|payout|^epf$|^nps$|loan schedule|rewards? ledger|locker/i,
    INSURANCE_LOG: /insurance (log|premium)/i }
};

const C = {
  ink: '#0A0F1E', muted: '#56637D', line: '#E2E7EF', surface: '#F0F3F8', bg: '#FFFFFF', accent: '#C4005A',
  green: '#0A7A54', greenBg: '#DCF4EA', red: '#C42020', redBg: '#FDE4E4', amber: '#8A5800', amberBg: '#FDF0C8',
  font: "'Helvetica Neue',Helvetica,Arial,sans-serif"
};
const MON = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

/* ===================== Entry points ===================== */

/** Run it from the editor, or point any time-driven trigger at it. */
function sendSnapshot(e) { sendSnapshot_(e && e.triggerUid ? 'trigger' : 'manual'); }

/**
 * Prints each tab's layout to the Execution log, for PROMPT.md.
 * Numbers, dates and email addresses are masked; text and formulas are kept, because
 * the labels are what the script reads. Check the output for names before sharing it.
 */
function describeWorkbook() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
  const ROWS = 60, colName = c => { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s; };
  ss.getSheets().filter(sh => sh.getName() !== CONFIG.LOG_TAB).forEach(sh => {
    const rng = sh.getDataRange(), v = rng.getValues(), f = rng.getFormulas(), out = [];
    out.push(`## Tab "${sh.getName()}": ${v.length} rows x ${(v[0] || []).length} columns`);
    v.slice(0, ROWS).forEach((row, r) => {
      const cells = row.map((x, c) => {
        if (f[r] && f[r][c]) return `${colName(c)}${r + 1}: ${f[r][c]}`;
        if (x === '' || x == null) return null;
        const t = x instanceof Date ? '<date>' : typeof x === 'number' ? '<number>' : /\S+@\S+\.\S+/.test(x) ? '<email>' : JSON.stringify(String(x).slice(0, 60));
        return `${colName(c)}${r + 1}: ${t}`;
      }).filter(Boolean);
      if (cells.length) out.push(cells.join(' | '));
    });
    if (v.length > ROWS) out.push(`... ${v.length - ROWS} more rows`);
    console.log(out.join('\n'));
  });
}

/* ===================== Orchestration ===================== */

/**
 * Web app. Deploy → New deployment → Web app; Execute as: Me; Who has access: Only myself.
 * Google then asks for your sign-in, so nobody else can open it and crawlers can't index it.
 * The app is read-only: opening it doesn't send an email or add a row to the log.
 * Add ?view=email to the web app URL to see the email instead.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.view === 'email') {
    const s = buildSnapshot_('web');
    return HtmlService.createHtmlOutput(s.html).setTitle('Vault').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Vault')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function sendSnapshot_(slot) {
  const s = buildSnapshot_(slot);
  const to = CONFIG.RECIPIENTS || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to, subject: 'Net worth as of ' + fmtDay_(s.now), htmlBody: s.html, name: 'Vault' });
}

/** Daily close: point a "Day timer, Midnight to 1am" trigger at this. Records one row a day; sends nothing. */
function recordDailyClose() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  const W = readWorth_(ss);
  let M = { list: [] };
  try { M = readMarkets_(ss); } catch (e) { console.error('Markets: ' + e.message); }
  upsertClose_(ss, toLogEntry_(new Date(), 'close', W, M));
  trackRewards_(ss, W, new Date());
}

function buildSnapshot_(slot) {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  const now = new Date();
  const run = (label, fn) => { try { return fn(); } catch (e) { console.error(label + ': ' + e.message); return null; } };

  const W = readWorth_(ss);
  const M = run('Markets', () => readMarkets_(ss)) || { list: [], cpi: null, repo: null };
  const X = run('Spend', () => readSpend_(ss, now));
  const I = run('Salary and tax', () => readIncomeTax_(ss, now));
  const L = run('Locker', () => readLocker_(ss));
  const RL = run('Reward lots', () => readRewardLots_(ss)) || {};

  // Comparisons use daily closes only: one row a day at the same time.
  const closes = closesByDay_(readLog_(ss)).filter(e => e.ts <= now.getTime());
  const cur = toLogEntry_(now, slot, W, M);
  const ref = {
    d1: closes.length ? closes[closes.length - 1] : null,
    d7: closest_(closes, now.getTime() - 7 * 864e5, 1.5 * 864e5),
    d30: closest_(closes, now.getTime() - 30 * 864e5, 3 * 864e5)
  };
  return { ss, now, cur, html: renderEmail_({ ss, now, W, M, X, I, L, RL, closes, cur, ref }) };
}

/* ===================== Readers ===================== */

function findSheet_(ss, re) {
  const sh = ss.getSheets().find(s => re.test(s.getName()));
  if (!sh) throw new Error('no tab matches ' + re);
  return sh;
}
/** The tab for the current financial year; else the latest year's tab. Returns its FY too. */
function findFySheet_(ss, re, now) {
  const cur = fyStart_(now);
  const cands = ss.getSheets().filter(s => re.test(s.getName())).map(s => {
    const nums = (s.getName().match(/\d{2,4}/g) || []).map(Number);
    if (!nums.length) return null;
    const a = nums[0] < 100 ? 2000 + nums[0] : nums[0];
    return { sh: s, start: a };
  }).filter(Boolean);
  if (!cands.length) throw new Error('no tab matches ' + re + ' with a year in its name');
  const pick = cands.find(x => x.start === cur) || cands.sort((a, b) => b.start - a.start)[0];
  return { sh: pick.sh, start: pick.start, fy: fyName_(pick.start), current: pick.start === cur };
}
const num_ = v => {
  if (typeof v === 'number') return v;
  if (v == null || v === '' || v instanceof Date) return NaN;
  const s = String(v).replace(/[₹$€£,\s]/g, '');
  if (/%$/.test(s)) return parseFloat(s) / 100;
  return /^-?\d*\.?\d+(e-?\d+)?$/i.test(s) ? parseFloat(s) : NaN;
};
const str_ = v => v == null ? '' : String(v).trim();

function readWorth_(ss) {
  const sh = findSheet_(ss, CONFIG.TABS.worth);
  const v = sh.getDataRange().getValues();
  const hr = v.findIndex(r => r.some(c => /^L1\b/i.test(str_(c))));
  if (hr < 0) throw new Error('No "L1: Group" header in ' + sh.getName());
  const H = v[hr].map(str_), col = re => H.findIndex(h => re.test(h));
  const K = { g: col(/^L1/i), cls: col(/^L2/i), cust: col(/custodian/i), veh: col(/^L4/i), name: col(/^L5/i),
    val: col(/amount:\s*t$/i), prev: col(/t\s*-\s*x/i), xirr: col(/xirr/i), fdCol: col(/>\s*fd/i), liq: col(/liquidity/i), fn: col(/function/i), note: col(/goal|comment/i),
    cat: col(/^L3/i), tax: col(/^taxation/i), cpi: col(/>\s*cpi/i) };
  const F = sh.getDataRange().getFormulas();
  const holdings = [], seen = {}, excluded = [];
  for (let r = hr + 1; r < v.length; r++) {
    const g = str_(v[r][K.g]);
    if (/^last updated/i.test(g)) break;
    if (!/^(asset|liability)$/i.test(g)) continue;
    const val = num_(v[r][K.val]), prev = num_(v[r][K.prev]);
    const name = str_(v[r][K.name]) || str_(v[r][K.veh]);
    const base = [name, str_(v[r][K.cust]), str_(v[r][K.veh])].join('|');
    if (CONFIG.EXCLUDE && CONFIG.EXCLUDE.test([name, str_(v[r][K.cls]), str_(v[r][K.cust]), str_(v[r][K.veh])].join(' '))) {
      excluded.push({ key: base + '|x', name, cls: str_(v[r][K.cls]) || 'Other', cust: str_(v[r][K.cust]), veh: str_(v[r][K.veh]), cat: K.cat >= 0 ? str_(v[r][K.cat]) : '',
        liq: str_(v[r][K.liq]) || 'Unspecified', fn: str_(v[r][K.fn]) || 'Unspecified', note: K.note >= 0 ? str_(v[r][K.note]) : '', val: isNaN(val) ? null : val, row: r + 1 });
      continue;
    }
    seen[base] = (seen[base] || 0) + 1;
    holdings.push({
      key: base + '|' + seen[base], name, g: /^liab/i.test(g) ? 'Liability' : 'Asset', cls: str_(v[r][K.cls]) || 'Other', cust: str_(v[r][K.cust]),
      val: !isNaN(val) ? val : (!isNaN(prev) ? prev : 0), prev: isNaN(prev) ? null : prev,
      xirr: num_(v[r][K.xirr]), liq: str_(v[r][K.liq]) || 'Unspecified', fn: str_(v[r][K.fn]) || 'Unspecified',
      veh: str_(v[r][K.veh]), inst: str_(v[r][K.name]), note: K.note >= 0 ? str_(v[r][K.note]) : '', formula: K.val >= 0 && F[r] ? str_(F[r][K.val]) : '',
      cat: K.cat >= 0 ? str_(v[r][K.cat]) : '', taxation: K.tax >= 0 ? v[r][K.tax] : '', beatsFd: K.fdCol >= 0 ? str_(v[r][K.fdCol]) : '', beatsCpi: K.cpi >= 0 ? str_(v[r][K.cpi]) : '',
      xirrFormula: K.xirr >= 0 && F[r] ? str_(F[r][K.xirr]) : ''
    });
  }
  let snapDate = null, target = null, targetDate = null;
  v.forEach(r => {
    const a = str_(r[0]);
    if (/^last updated/i.test(a) && r[2] instanceof Date) snapDate = r[2];
    if (/^target net worth/i.test(a)) { target = num_(r[2]); targetDate = parseLooseDate_(a); }
  });
  // FD benchmark: the rate inside the "> FD?" formula
  let fd = null;
  if (K.fdCol >= 0) { const f = sh.getRange(hr + 2, K.fdCol + 1).getFormula(); const m = f && f.match(/([\d.]+)\s*%/); if (m) fd = Number(m[1]) / 100; }
  const sumOf = (g, k) => sum_(holdings.filter(h => h.g === g).map(h => k === 'prev' ? (h.prev != null ? h.prev : h.val) : h.val));
  const assets = sumOf('Asset'), liab = sumOf('Liability');
  return { holdings, assets, liab, nw: assets - liab, sheetPrevNw: sumOf('Asset', 'prev') - sumOf('Liability', 'prev'), snapDate, target, targetDate, fd, excluded };
}

function readMarkets_(ss) {
  const v = findSheet_(ss, CONFIG.TABS.portfolio).getDataRange().getValues();
  const WANT = [
    ['Nifty 50', /^nifty 50$/i, 'idx'], ['Nasdaq 100', /^nasdaq 100$/i, 'idx'],
    ['USD/INR', /^usd$/i, 'fx'], ['EUR/INR', /^eur$/i, 'fx'], ['GBP/INR', /^gbp$/i, 'fx'],
    ['Gold 24K · 10 g', /^gold\s*-\s*24k$/i, 'inr'], ['Gold 22K · 10 g', /^gold\s*-\s*22k$/i, 'inr'], ['Silver 999 · kg', /^silver\s*-\s*999k?$/i, 'inr']
  ];
  const list = []; let cpi = null, repo = null, cpiDate = null, repoDate = null;
  for (let r = 0; r < v.length; r++) for (let c = 0; c < v[r].length; c++) {
    const s = str_(v[r][c]);
    WANT.forEach(([label, re, kind]) => { if (re.test(s)) { const val = num_(v[r][c + 1]); if (!isNaN(val)) list.push({ label, kind, val, dayPct: kind === 'idx' ? num_(v[r][c + 3]) : NaN }); } });
    if (/^inflation$/i.test(s)) { const a = num_((v[r + 1] || [])[c + 2]), b = num_(v[r][c + 2]); cpi = !isNaN(a) && a < 1 ? a : (!isNaN(b) && b < 1 ? b : null);
      const da = (v[r + 1] || [])[c + 1], db = v[r][c + 1]; cpiDate = cpi === a && da instanceof Date ? da : (db instanceof Date ? db : null); }
    if (/^repo rate$/i.test(s)) { const x = num_((v[r + 1] || [])[c]); if (!isNaN(x)) { repo = x; const d = (v[r + 1] || [])[c - 1]; repoDate = d instanceof Date ? d : null; } }
  }
  return { list, cpi, repo, cpiDate, repoDate };
}

/** Finds the row whose cells name at least 10 months; returns that row and each month's column. */
function monthHeader_(v, maxRow) {
  for (let r = 0; r < Math.min(v.length, maxRow); r++) {
    const cs = []; v[r].forEach((x, c) => { const m = monthOf_(x); if (m != null) cs.push({ c, m }); });
    if (cs.length >= 10) return { r, cols: cs };
  }
  return null;
}

function readSpend_(ss, now) {
  const T = findFySheet_(ss, CONFIG.TABS.expenses, now);
  const v = T.sh.getDataRange().getValues();
  const hdr = monthHeader_(v, 8);
  if (!hdr) throw new Error('no month header in ' + T.sh.getName());
  const lc = hdr.cols[0].c - 1, rows = []; let cat = '', totalRow = null;
  for (let r = hdr.r + 1; r < v.length; r++) {
    const left = v[r].slice(0, lc).map(str_);
    if (left.some(x => /tracker/i.test(x))) break;
    const label = str_(v[r][lc]), near = left.slice().reverse().find(Boolean);
    if (/^total$/i.test(label) || left.some(x => /^total$/i.test(x))) { if (totalRow == null) totalRow = r; continue; }
    if (near) cat = near;
    if (!label) continue;
    const m = Array(12).fill(0), filled = Array(12).fill(false);
    hdr.cols.forEach(({ c, m: mi }) => { const n = num_(v[r][c]); if (!isNaN(n)) { m[mi] = n; filled[mi] = v[r][c] !== '' && v[r][c] != null; } });
    rows.push({ cat: cat || 'Other', item: label, m, filled });
  }
  const mi = T.current ? fyMonth_(now) : 11;
  // Monthly totals from the tab's own Total row; a month whose cell errors (#REF!) is summed from its rows.
  const monthTot = MON.map((_, j) => {
    const col = hdr.cols.find(x => x.m === j), n = totalRow != null && col ? num_(v[totalRow][col.c]) : NaN;
    return !isNaN(n) ? n : sum_(rows.map(x => x.m[j]));
  });
  // Monthly average from the tab's own Average column when it has one.
  const avgCol = v[hdr.r].findIndex(x => /average|avg/i.test(str_(x)));
  const sheetAvg = avgCol >= 0 && totalRow != null ? num_(v[totalRow][avgCol]) : NaN;
  const totCol = v[hdr.r].findIndex(x => /^total$/i.test(str_(x)));
  const sheetTotal = totCol >= 0 && totalRow != null ? num_(v[totalRow][totCol]) : NaN;
  const priorIdx = MON.map((_, j) => j).filter(j => j < mi && monthTot[j] > 0);
  const used = MON.map((_, j) => j <= mi && monthTot[j] > 0), nUsed = used.filter(Boolean).length;
  const avgOf = f => nUsed ? sum_(rows.filter(f).map(x => sum_(x.m.filter((_, j) => used[j])))) / nUsed : null;
  const ytdOf = f => sum_(rows.filter(f).map(x => sum_(x.m.slice(0, mi + 1))));
  const isEmi = x => /loan|\bemi\b/i.test(x.item), isTravel = x => /travel/i.test(x.item);
  const cats = months => { const m = {}; rows.forEach(x => { m[x.cat] = (m[x.cat] || 0) + sum_(months.map(j => x.m[j])); }); return Object.keys(m).map(k => ({ k, v: m[k] })).filter(x => x.v > 0).sort((a, b) => b.v - a.v); };
  // Pace: the share of the month gone. A month in progress is compared with that share of a usual month.
  const dim = new Date(T.start + (mi >= 9 ? 1 : 0), (mi + 3) % 12 + 1, 0).getDate();
  const day = T.current ? Number(Utilities.formatDate(now, CONFIG.TIMEZONE, 'd')) : dim;
  const pace = Math.min(1, day / dim);
  // Each category's usual month: its average over the completed months that had spend.
  const off = priorIdx.length ? Object.keys(rows.reduce((o, x) => (o[x.cat] = 1, o), {})).map(k => {
    const inCat = rows.filter(x => x.cat === k);
    const usual = sum_(priorIdx.map(j => sum_(inCat.map(x => x.m[j])))) / priorIdx.length;
    const act = sum_(inCat.map(x => x.m[mi])), exp = usual * pace;
    return { k, act, usual, exp, gap: act - exp, pct: exp > 0 ? (act - exp) / exp : null };
  }).filter(x => x.usual > 0 || x.act > 0) : [];
  return {
    pace, day, dim, off, sheetTotal, sheetAvg, totalCell: totCol >= 0 && totalRow != null ? appA1_(totalRow, totCol) : '', avgCell: avgCol >= 0 && totalRow != null ? appA1_(totalRow, avgCol) : '', tab: T.sh.getName(),
    fy: T.fy, rows, mi, monthTot, thisMonth: monthTot[mi], nUsed,
    avg: !isNaN(sheetAvg) ? sheetAvg : (priorIdx.length ? sum_(priorIdx.map(j => monthTot[j])) / priorIdx.length : null),
    avgNote: !isNaN(sheetAvg) ? 'from the sheet' : priorIdx.length ? MON[priorIdx[0]] + ' to ' + MON[priorIdx[priorIdx.length - 1]] : '',
    ytd: sum_(monthTot.slice(0, mi + 1)), emiYtd: ytdOf(isEmi), travelYtd: ytdOf(isTravel),
    emiMonthly: avgOf(isEmi), housingMonthly: avgOf(x => /household|housing/i.test(x.cat)),
  };
}

function readIncomeTax_(ss, now) {
  const T = findFySheet_(ss, CONFIG.TABS.coi, now);
  const v = T.sh.getDataRange().getValues();
  const Y = T.start;

  /* Payslip register */
  const hdr = monthHeader_(v, v.length);
  if (!hdr) throw new Error('no payslip register in ' + T.sh.getName());
  const hr = hdr.r, lc = hdr.cols[0].c - 1, rows = []; let sec = null;
  for (let r = hr + 1; r < v.length; r++) {
    const label = str_(v[r][lc]); if (!label) continue;
    if (/^net\b/i.test(label)) break;
    if (/^(income|earnings)$/i.test(label)) { sec = 'e'; continue; }
    if (/^deductions$/i.test(label)) { sec = 'd'; continue; }
    if (!sec || /gross|total/i.test(label)) continue;
    const m = Array(12).fill(0); hdr.cols.forEach(({ c, m: mi }) => { const n = num_(v[r][c]); if (!isNaN(n)) m[mi] = n; });
    rows.push({ sec, item: label, m });
  }
  const isTds = r => r.sec === 'd' && /income tax|\btds\b|supplementary/i.test(r.item);
  const isPf = r => r.sec === 'd' && /provident|\bpf\b|\bepf\b/i.test(r.item);
  const isOneOff = r => /bonus|notice|joining|arrear|incentive|encash/i.test(r.item);
  const byMonth = f => MON.map((_, i) => sum_(rows.filter(f).map(r => r.m[i])));
  const gross = byMonth(r => r.sec === 'e'), ded = byMonth(r => r.sec === 'd'), tds = byMonth(isTds), pf = byMonth(isPf);
  const lastIdx = gross.map(g => g > 0).lastIndexOf(true);
  const at = (arr, f) => lastIdx >= 0 ? (f ? sum_(rows.filter(f).map(r => r.m[lastIdx])) : arr[lastIdx]) : 0;
  const runGross = at(null, r => r.sec === 'e' && !isOneOff(r)), runTds = at(tds), runPf = at(pf);
  const ahead = 11 - lastIdx;
  const pay = {
    gross, ded, tds, lastIdx, runGross, ahead,
    grossToDate: sum_(gross), tdsToDate: sum_(tds), takeHomeToDate: sum_(gross) - sum_(ded),
    grossFY: sum_(gross) + ahead * runGross, tdsFY: sum_(tds) + ahead * runTds, pfFY: sum_(pf) + ahead * runPf
  };

  /* Tax inputs above the register: column D = old regime, E = new regime, F = selected */
  const val = (re, cols) => { for (let r = 0; r < hr; r++) if (re.test(str_(v[r][1]))) { for (const c of cols) { const n = num_(v[r][c]); if (!isNaN(n)) return n; } return 0; } return 0; };
  const D = [3], E = [4], F = [5, 4, 3];
  const regimeRow = v.find(r => r.some(c => /select tax regime/i.test(str_(c))));
  const regimePick = regimeRow ? regimeRow.slice(regimeRow.findIndex(c => /select tax regime/i.test(str_(c))) + 1).map(str_).find(Boolean) || '' : '';
  const regime = /old/i.test(regimePick) ? 'old' : 'new';
  // SGB interest, dividends and REIT income come from the payouts tab; savings and FD interest from this tab.
  let PO = null;
  try { PO = readPayouts_(ss, Y); } catch (e) { console.error('Payouts: ' + e.message); }
  const inc = {
    bonus: val(/^performance bonus/i, D), sav: val(/^savings account interest/i, D), fd: val(/^fixed deposit interest/i, D),
    sgb: PO ? PO.sgbYear : val(/^sgb interest/i, D), div: PO ? PO.div : 0, reit: PO ? PO.reitTax : val(/^reit distributions/i, D),
    stcg: val(/^short-term capital gains/i, D), ltcg: val(/^long-term capital gains/i, D)
  };
  const dOld = {
    hra: val(/hra exemption/i, D), hlInt: val(/home loan interest/i, D),
    c80: val(/80c\W+home loan principal/i, D) + pay.pfFY + val(/80c\W+ppf/i, D) + val(/80c\W+lic/i, D) + val(/80c\W+term/i, D),
    medSelf: val(/80d.*self/i, D), medPar: val(/80d.*parents/i, D), nps1b: val(/80ccd\(1b\)/i, D)
  };
  const both = { nps2: val(/80ccd\(2\)/i, E), leave: val(/10\(10aa\)/i, E) };
  const paid = {
    tdsOther: Math.max(val(/^tds\W+other income/i, F), PO ? PO.reitTds : 0),
    adv: [1, 2, 3, 4].map(q => val(new RegExp('^advance tax\\W+q' + q, 'i'), F)),
    sat: val(/^self-assessment tax/i, F)
  };
  const grossInc = pay.grossFY + sum_(Object.keys(inc).map(k => inc[k]));
  const calc = reg => {
    const std = reg === 'old' ? 50000 : 75000; // standard deduction: ₹50,000 old regime, ₹75,000 new regime
    const dd = reg === 'old'
      ? std + dOld.hra + Math.min(dOld.hlInt, 200000) + Math.min(dOld.c80, 150000) + Math.min(dOld.medSelf, 25000) + Math.min(dOld.medPar, 50000) + Math.min(inc.sav, 10000) + Math.min(dOld.nps1b, 50000) + both.nps2 + both.leave
      : std + both.nps2 + both.leave;
    const slabInc = Math.max(0, grossInc - dd - inc.stcg - inc.ltcg);
    const slab = slabTax_(slabInc, reg);
    const rebate = reg === 'old' ? (slabInc <= 500000 ? Math.min(12500, slab) : 0) : (slabInc <= 1200000 ? Math.min(60000, slab) : 0);
    const special = inc.stcg * 0.2 + Math.max(0, inc.ltcg - 125000) * 0.125;
    const total = slabInc + inc.stcg + inc.ltcg;
    const sr = total > 5e7 ? (reg === 'old' ? .37 : .25) : total > 2e7 ? .25 : total > 1e7 ? .15 : total > 5e6 ? .1 : 0;
    const base = slab - rebate, sur = base * sr + special * Math.min(sr, .15), cess = (base + special + sur) * 0.04;
    return { tax: base + special + sur + cess, std, dd, slabInc, slab, rebate, special, stcgTax: inc.stcg * 0.2, ltcgTax: Math.max(0, inc.ltcg - 125000) * 0.125, sr, sur, cess };
  };
  const old = calc('old'), neu = calc('new'), sel = regime === 'old' ? old : neu;

  /* Paid, balance, advance tax and interest.
     Income-tax Act, 2025 (tax year 2026-27 on): s.404 ₹10,000 threshold; s.425 (was 234C) 1% a month for 3/3/3/1 months,
     no interest if 12% is paid by 15 Jun or 36% by 15 Sep; s.424 (was 234B) 1% a month or part from 1 April if below 90%. */
  const advPaid = sum_(paid.adv);
  const paidSoFar = pay.tdsToDate + paid.tdsOther + advPaid + paid.sat;
  const netDue = Math.max(0, sel.tax - pay.tdsFY - paid.tdsOther);
  const applies = netDue >= 10000;
  const pos = x => Math.max(0, x);
  const SCHED = [
    { q: 'Q1', due: new Date(Y, 5, 15), pct: .15, safe: .12, months: 3 },
    { q: 'Q2', due: new Date(Y, 8, 15), pct: .45, safe: .36, months: 3 },
    { q: 'Q3', due: new Date(Y, 11, 15), pct: .75, safe: .75, months: 3 },
    { q: 'Q4', due: new Date(Y + 1, 2, 15), pct: 1, safe: 1, months: 1 }
  ];
  let cum = 0;
  const sched = SCHED.map((s, i) => {
    cum += paid.adv[i];
    const req = netDue * s.pct;
    return Object.assign({}, s, { req, paidCum: cum, gap: pos(req - cum), past: s.due < startOfDay_(now), int: applies && cum < netDue * s.safe ? pos(req - cum) * 0.01 * s.months : 0 });
  });
  const filing = CONFIG.ITR_FILING_DATE ? new Date(CONFIG.ITR_FILING_DATE) : new Date(Y + 1, 6, 31);
  const s424Months = Math.max(0, (filing.getFullYear() - (Y + 1)) * 12 + filing.getMonth() - 3 + 1);
  const s424 = applies && advPaid < 0.9 * netDue ? pos(netDue - advPaid) * 0.01 * s424Months : 0;
  return {
    fy: T.fy, Y, regime, pay, inc, dOld, both, payouts: PO, sheet: T.sh.getName(), advByQ: paid.adv, sat: paid.sat, tdsOther: paid.tdsOther, grossInc, old, new: neu, sel, advPaid, paidSoFar, tdsAhead: pay.tdsFY - pay.tdsToDate, netDue, applies, sched, s424, s424Months, filing,
    balance: sel.tax - paidSoFar, afterTds: sel.tax - pay.tdsFY - paid.tdsOther - advPaid - paid.sat
  };
}
function slabTax_(inc, reg) {
  const s = reg === 'old' ? [[250000, 0], [500000, .05], [1000000, .2], [Infinity, .3]] : [[400000, 0], [800000, .05], [1200000, .1], [1600000, .15], [2000000, .2], [2400000, .25], [Infinity, .3]];
  let t = 0, lo = 0; s.forEach(([hi, r]) => { if (inc > lo) t += (Math.min(inc, hi) - lo) * r; lo = hi; }); return t;
}

/** Payouts tab: SGB coupons (cols after Date), REIT blocks (Particulars), stock dividends (Stocks). */
function readPayouts_(ss, Y) {
  const v = findSheet_(ss, CONFIG.TABS.payouts).getDataRange().getValues();
  const from = new Date(Y, 3, 1), to = new Date(Y + 1, 3, 1);
  const inFy = d => d instanceof Date && d >= from && d < to;
  const mIdx = d => (d.getMonth() + 9) % 12;
  const H0 = v[0].map(str_), H1 = (v[1] || []).map(str_);
  const cash = Array(12).fill(0), sgbCash = Array(12).fill(0), reitCash = Array(12).fill(0), divCash = Array(12).fill(0);
  // SGB coupons: half-yearly and fixed, so a bond paid once this year pays the same again.
  const qc = H0.findIndex(h => /^quarter$/i.test(h));
  let sgbRecv = 0, sgbYear = 0;
  for (let c = 1; c < (qc > 0 ? qc : 1); c++) {
    if (!H0[c]) continue;
    const fy = v.slice(1).filter(r => inFy(r[0]) && num_(r[c]) > 0).map(r => ({ d: r[0], a: num_(r[c]) }));
    const recv = sum_(fy.map(p => p.a));
    fy.forEach(p => { cash[mIdx(p.d)] += p.a; sgbCash[mIdx(p.d)] += p.a; });
    sgbRecv += recv;
    sgbYear += recv + (fy.length === 1 ? fy[0].a : 0);
  }
  // REIT blocks: taxable = interest + other income + dividends that carry TDS. SPV debt repayment returns capital.
  const pc = H0.findIndex(h => /^particulars$/i.test(h)), sc = H0.findIndex(h => /^stocks$/i.test(h));
  let reitTax = 0, reitTds = 0, reitNet = 0;
  if (pc > 0) {
    const end = sc > 0 ? sc : H0.length, comps = [];
    for (let c = pc + 1; c < end; c++) if (H1[c]) comps.push({ c, name: H1[c] });
    let date = null, at = {};
    for (let r = 2; r < v.length; r++) {
      const label = str_(v[r][pc]);
      if (/^per unit/i.test(label)) { date = v[r][pc - 1] instanceof Date ? v[r][pc - 1] : null; at = {}; }
      else if (/for your holding/i.test(label)) at.hold = r;
      else if (/^tds rate/i.test(label)) at.rate = r;
      else if (/^tds amount/i.test(label)) at.tds = r;
      else if (/^net distribution/i.test(label) && inFy(date)) {
        comps.forEach(({ c, name }) => {
          const n = num_(v[r][c]);
          if (/^total$/i.test(name)) { if (!isNaN(n)) { reitNet += n; cash[mIdx(date)] += n; reitCash[mIdx(date)] += n; } return; }
          const gross = at.hold != null ? num_(v[at.hold][c]) : NaN, rate = at.rate != null ? num_(v[at.rate][c]) : NaN, tds = at.tds != null ? num_(v[at.tds][c]) : NaN;
          if (!isNaN(gross) && (/interest|other/i.test(name) || (/dividend/i.test(name) && rate > 0))) reitTax += gross;
          if (!isNaN(tds)) reitTds += tds;
        });
      }
    }
  }
  // Stock dividends
  let div = 0;
  if (sc >= 0) for (let c = sc + 1; c < H1.length; c++) if (H1[c]) v.slice(2).forEach(r => { const a = num_(r[c]); if (inFy(r[sc]) && a > 0) { div += a; cash[mIdx(r[sc])] += a; divCash[mIdx(r[sc])] += a; } });
  return { sgbRecv, sgbYear, reitTax, reitTds, reitNet, div, cash, sgbCash, reitCash, divCash };
}

function readLocker_(ss) {
  const sh = findSheet_(ss, CONFIG.TABS.locker), v = sh.getDataRange().getValues();
  const hr = v.findIndex(r => /access date/i.test(str_(r[0])));
  const visits = v.slice(hr + 1).filter(r => r[0] instanceof Date).map(r => ({ d: r[0], by: str_(r[1]), after: str_(r[3]) })).sort((a, b) => b.d - a.d);
  return { url: ss.getUrl() + '#gid=' + sh.getSheetId(), last: visits[0] || null, items: visits[0] ? visits[0].after.split(/\n+/).map(x => x.trim().replace(/^[•\-\d.)\s]+/, '')).filter(Boolean) : [] };
}

/** Loan tenure left from the sheet: outstanding and rate from the holdings tab, EMI from the spend tab. */
function loanTerms_(h, X) {
  const rate = !isNaN(h.xirr) && h.xirr > 0 ? h.xirr : null;
  let emi = null;
  if (X) {
    const tokens = s => normTokens_(s).filter(t => t.length > 1);
    const want = tokens(h.name + ' ' + h.cust).concat(acronym_(h.cust));
    let best = null, score = 1;
    X.rows.forEach(r => { if (!/loan|\bemi\b/i.test(r.item)) return; const t = tokens(r.item); const s = t.filter(x => want.indexOf(x) >= 0).length; if (s > score) { score = s; best = r; } });
    // The smallest recurring amount is the EMI; larger months carry prepayments.
    if (best) { const nz = best.m.filter(x => x > 0); if (nz.length) emi = Math.min.apply(null, nz); }
  }
  let months = null;
  if (rate && emi) { const r = rate / 12, x = r * h.val / emi; months = x < 1 ? Math.ceil(-Math.log(1 - x) / Math.log(1 + r)) : Infinity; }
  return { rate, emi, months };
}
const normTokens_ = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean);
const acronym_ = s => { const t = normTokens_(s).filter(w => !/^(OF|THE|AND)$/.test(w)); return t.length > 1 ? [t.map(w => w[0]).join('')] : []; };

/* ===================== Reward points ===================== */

const isReward_ = h => !!CONFIG.REWARDS && CONFIG.REWARDS.test(h.veh + ' ' + h.inst);
/** Points and ₹ per point from Amount: T written as =(points*paise)/100 or =points*rupees. */
function rewardPts_(h) {
  const m1 = h.formula.match(/^=\s*\(?\s*([\d.,]+)\s*\*\s*([\d.]+)\s*\)?\s*\/\s*100\s*$/), m2 = h.formula.match(/^=\s*\(?\s*([\d.,]+)\s*\*\s*([\d.]+)\s*\)?\s*$/);
  const m = m1 || m2;
  return m ? { pts: Number(m[1].replace(/,/g, '')), rate: Number(m[2]) / (m1 ? 100 : 1) } : { pts: null, rate: null };
}
const noteDate_ = h => { const br = h.note.match(/\[([^\]]*)\]/); return br ? parseLooseDate_(br[1]) : null; };
const addYears_ = (d, n) => { const x = new Date(d.getTime()); x.setFullYear(x.getFullYear() + n); return x; };
function lotsSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.REWARD_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.REWARD_TAB);
    sh.getRange(1, 1, 1, 6).setValues([['Holding', 'Programme', 'Earned on', 'Points', 'Points left', 'Expires']]);
    sh.hideSheet();
  }
  return sh;
}
/** Lots by holding key: [{ earned, pts, left, exp }], oldest first. */
function readRewardLots_(ss) {
  const sh = ss.getSheetByName(CONFIG.REWARD_TAB);
  if (!sh || sh.getLastRow() < 2) return {};
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(r => {
    if (!r[0]) return;
    (out[r[0]] = out[r[0]] || []).push({ prog: str_(r[1]), earned: new Date(r[2]), pts: Number(r[3]), left: Number(r[4]), exp: new Date(r[5]) });
  });
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.earned - b.earned));
  return out;
}
/**
 * Daily close: compares each programme's points with the points left in its lots.
 * A rise is a new lot earned today, expiring REWARD_EXPIRY_YEARS later. A fall uses up the oldest lots first.
 * The first time a programme is seen, its whole balance becomes one lot that expires on the date
 * in square brackets in Goal/Comments, or REWARD_EXPIRY_YEARS from today if there's no date.
 */
function trackRewards_(ss, W, now) {
  const lots = readRewardLots_(ss), Y = CONFIG.REWARD_EXPIRY_YEARS;
  W.holdings.filter(h => h.g === 'Asset' && isReward_(h)).forEach(h => {
    const { pts } = rewardPts_(h);
    if (pts == null) return;
    const mine = lots[h.key] = lots[h.key] || [], prog = [h.cust, h.veh].filter(Boolean).join(' · ');
    const held = sum_(mine.map(l => l.left));
    if (!mine.length) {
      if (pts <= 0) return;
      const exp = noteDate_(h) || addYears_(now, Y);
      mine.push({ prog, earned: addYears_(exp, -Y), pts, left: pts, exp });
    } else if (pts > held) {
      mine.push({ prog, earned: now, pts: pts - held, left: pts - held, exp: addYears_(now, Y) });
    } else if (pts < held) {
      let use = held - pts;
      mine.forEach(l => { const t = Math.min(l.left, use); l.left -= t; use -= t; });
    }
  });
  const rows = [];
  Object.keys(lots).forEach(k => lots[k].forEach(l => rows.push([k, l.prog, l.earned, l.pts, l.left, l.exp])));
  const sh = lotsSheet_(ss);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
}

/* ===================== Snapshot log ===================== */

function logSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LOG_TAB);
    sh.getRange(1, 1, 1, 8).setValues([['Timestamp', 'Run', 'Net worth', 'Assets', 'Liabilities', 'Holdings', 'Classes', 'Markets']]);
    sh.hideSheet();
  }
  return sh;
}
function readLog_(ss) {
  const sh = logSheet_(ss), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 8).getValues().map(r => {
    try { return dropExcluded_({ ts: new Date(r[0]).getTime(), slot: str_(r[1]), nw: Number(r[2]), h: JSON.parse(r[5] || '{}'), cls: JSON.parse(r[6] || '{}'), mkt: JSON.parse(r[7] || '{}') }); } catch (e) { return null; }
  }).filter(Boolean).sort((a, b) => a.ts - b.ts);
}
/** Older log rows may include holdings now excluded; take them out so deltas compare like for like. */
function dropExcluded_(e) {
  if (!CONFIG.EXCLUDE) return e;
  Object.keys(e.h).forEach(k => {
    if (!CONFIG.EXCLUDE.test(k.replace(/\|/g, ' '))) return;
    e.nw -= e.h[k];
    delete e.h[k];
  });
  return e;
}
function toLogEntry_(now, slot, W, M) {
  const h = {}, cls = {}, mkt = {};
  W.holdings.forEach(x => { h[x.key] = Math.round(x.val * 100) / 100; const k = x.g + '|' + x.cls; cls[k] = (cls[k] || 0) + x.val; });
  M.list.forEach(x => { mkt[x.label] = x.val; });
  return { ts: now.getTime(), slot, nw: W.nw, assets: W.assets, liab: W.liab, h, cls, mkt };
}
function appendLog_(ss, e) {
  const sh = logSheet_(ss);
  sh.appendRow([new Date(e.ts), e.slot, e.nw, e.assets, e.liab, JSON.stringify(e.h), JSON.stringify(e.cls), JSON.stringify(e.mkt)]);
  const extra = sh.getLastRow() - 1 - CONFIG.LOG_KEEP_ROWS;
  if (extra > 0) sh.deleteRows(2, extra);
}
/** The day a close belongs to: a close taken between midnight and 4 AM records the day that just ended. */
function closeDay_(ts) { return Number(Utilities.formatDate(new Date(ts), CONFIG.TIMEZONE, 'H')) < 4 ? dayKey_(ts - 4 * 36e5) : dayKey_(ts); }
/** One close per day (the latest, if a day has two), oldest first. */
function closesByDay_(log) {
  const byDay = {};
  log.filter(e => e.slot === 'close').forEach(e => { const k = closeDay_(e.ts); if (!byDay[k] || byDay[k].ts < e.ts) byDay[k] = Object.assign({ k }, e); });
  return Object.keys(byDay).sort().map(k => byDay[k]);
}
const dayKey_ = ts => Utilities.formatDate(new Date(ts), CONFIG.TIMEZONE, 'yyyy-MM-dd');
/** Writes the day's close, replacing an earlier close from the same day. */
function upsertClose_(ss, e) {
  const sh = logSheet_(ss), last = sh.getLastRow();
  if (last > 1) {
    const r = sh.getRange(last, 1, 1, 2).getValues()[0];
    if (str_(r[1]) === 'close' && closeDay_(new Date(r[0]).getTime()) === closeDay_(e.ts)) {
      sh.getRange(last, 1, 1, 8).setValues([[new Date(e.ts), e.slot, e.nw, e.assets, e.liab, JSON.stringify(e.h), JSON.stringify(e.cls), JSON.stringify(e.mkt)]]);
      return;
    }
  }
  appendLog_(ss, e);
}
function closest_(log, t, tol) {
  let best = null, bd = Infinity;
  log.forEach(e => { const d = Math.abs(e.ts - t); if (d < bd && d <= tol) { bd = d; best = e; } });
  return best;
}

/* ===================== Rendering ===================== */

function renderEmail_(o) {
  const { ss, now, W, M, X, I, L, RL, closes, ref } = o;
  const closeOf = e => 'close at the end of ' + fmtDate_(new Date((e.k || dayKey_(e.ts)) + 'T12:00:00'));
  const A = W.holdings.filter(h => h.g === 'Asset');
  let b = ''; const S_ = {};

  /* Net worth */
  const tiles = [
    ['Since last close', ref.d1 && W.nw - ref.d1.nw, ref.d1 && ref.d1.nw],
    ['7 days', ref.d7 && W.nw - ref.d7.nw, ref.d7 && ref.d7.nw],
    ['30 days', ref.d30 && W.nw - ref.d30.nw, ref.d30 && ref.d30.nw],
    ['Since ' + (W.snapDate ? fmtDate_(W.snapDate) : ''), W.snapDate ? W.nw - W.sheetPrevNw : null, W.sheetPrevNw]
  ].filter(t => t[1] != null && !isNaN(t[1]));
  S_.worth = card_(
    `<div style="font-size:13px;color:${C.muted}">Net worth</div>
     <div style="font-size:40px;font-weight:800;letter-spacing:-0.5px;color:${C.ink};line-height:1.1;margin:4px 0 14px">${inrC_(W.nw)}</div>
     ${tiles.length ? `<div style="margin-bottom:6px">${tiles.map(([l, d, base]) => deltaPill_(l, d, base)).join('')}</div>` : ''}
     <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin-top:6px"><tr>
       <td style="font-size:14px;color:${C.muted}">Assets <b style="color:${C.ink}">${inrC_(W.assets)}</b></td>
       <td align="right" style="font-size:14px;color:${C.muted}">Liabilities <b style="color:${C.ink}">${inrC_(W.liab)}</b></td></tr></table>
     ${W.target ? targetBlock_(W, now) : ''}
`);

  /* Trend */
  const chg = dailyChanges_(closes, CONFIG.TREND_DAYS);
  if (chg.length) S_.trend = section_('Net worth, daily change', trendBars_(chg));

  /* Allocation */
  const alloc = groupSum_(A, h => h.cls).map(x => {
    const p = ref.d1 && ref.d1.cls['Asset|' + x.k];
    return Object.assign(x, p != null ? { d: x.v - p, p } : {});
  });
  const quad = (title, rows) => subhead_(title, C.muted) + miniBars_(rows, W.assets);
  S_.alloc = section_('Allocation',
    twoCol_([quad('ASSET CLASS', alloc), quad('CUSTODIAN', groupSum_(A, h => h.cust || 'Unspecified'))])
    + `<div style="height:12px;border-bottom:1px solid ${C.line};margin-bottom:12px"></div>`
    + twoCol_([quad('LIQUIDITY', groupSum_(A, h => h.liq)), quad('PURPOSE', groupSum_(A, h => h.fn))]));

  /* Movers over 24 hours */
  if (ref.d1) {
    const mv = A.filter(h => ref.d1.h[h.key] != null).map(h => ({ h, d: h.val - ref.d1.h[h.key], p: ref.d1.h[h.key] })).filter(x => Math.abs(x.d) >= 1);
    const up = mv.filter(x => x.d > 0).sort((a, c) => c.d - a.d).slice(0, CONFIG.TOP_MOVERS);
    const dn = mv.filter(x => x.d < 0).sort((a, c) => a.d - c.d).slice(0, CONFIG.TOP_MOVERS);
    const list = (arr, title, col) => `${subhead_(title, col)}<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${arr.map(x => `<tr><td style="padding:6px 0;font-size:14px;color:${C.ink}">${esc_(x.h.name)}<div style="font-size:12px;color:${C.muted}">${esc_(x.h.cls)}</div></td><td align="right" style="padding:6px 0;white-space:nowrap">${deltaText_(x.d, x.p)}</td></tr>`).join('')}</table>`;
    const cols = [up.length && list(up, 'TOP ' + up.length + ' GAINERS', C.green), dn.length && list(dn, 'TOP ' + dn.length + ' LOSERS', C.red)].filter(Boolean);
    if (cols.length) S_.movers = section_('Movers since the last close', `<div style="font-size:13px;color:${C.muted};margin:-6px 0 10px">Against the ${closeOf(ref.d1)}</div>` + twoCol_(cols));
  }

  /* Returns */
  const graded = A.filter(h => !isNaN(h.xirr));
  const beats = h => (M.cpi != null && h.xirr > M.cpi ? 1 : 0) + (W.fd != null && h.xirr > W.fd ? 1 : 0);
  const cnt = [2, 1, 0].map(n => graded.filter(h => beats(h) === n)), gradedVal = sum_(graded.map(h => h.val));
  const byX = graded.slice().sort((a, c) => c.xirr - a.xirr);
  const xList = arr => table_(null, arr.map(h => [esc_(h.name) + `<div style="font-size:12px;color:${C.muted}">${esc_(h.cls)} · ${inrC_(h.val)}</div>`, `<b style="color:${h.xirr < 0 ? C.red : W.fd != null && h.xirr > W.fd ? C.green : C.ink}">${pct_(h.xirr, 1)}</b>`]), [0, 1]);
  S_.returns = section_('Returns',
    `<div style="font-size:13px;color:${C.muted};margin-bottom:10px">Against inflation ${pct_(M.cpi, 2)} and FD ${pct_(W.fd, 2)}</div>
     <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
       ${[['Beat both', cnt[0], C.green, C.greenBg], ['Beat one', cnt[1], C.amber, C.amberBg], ['Beat neither', cnt[2], C.red, C.redBg]].map(([l, arr, fg, bg]) => {
         const val = sum_(arr.map(h => h.val));
         return `<td width="33%" valign="top" style="padding-right:6px"><div style="background:${bg};border-radius:6px;padding:10px 12px;color:${fg}">
           <div style="font-size:12px;font-weight:700">${l}</div>
           <div style="margin-top:4px"><span style="font-size:22px;font-weight:800">${arr.length}</span> <span style="font-size:12px">· ${pct_(graded.length ? arr.length / graded.length : 0, 0)} of holdings</span></div>
           <div><span style="font-size:15px;font-weight:800">${inrC_(val)}</span> <span style="font-size:12px">· ${pct_(gradedVal ? val / gradedVal : 0, 1)} of value</span></div></div></td>`;
       }).join('')}
     </tr></table>
     <div style="height:14px"></div>
     ${twoCol_([subhead_('TOP ' + CONFIG.TOP_XIRR + ' XIRR', C.green) + xList(byX.slice(0, CONFIG.TOP_XIRR)), subhead_('BOTTOM ' + CONFIG.TOP_XIRR + ' XIRR', C.red) + xList(byX.slice(-CONFIG.TOP_XIRR).reverse())])}`);

  /* Liabilities */
  const liabs = W.holdings.filter(h => h.g === 'Liability').sort((a, c) => c.val - a.val);
  if (liabs.length) {
    const rows = liabs.map(h => { const t = loanTerms_(h, X); return [esc_(h.name), inrC_(h.val), t.rate ? pct_(t.rate, 2) : '', t.emi ? inr_(t.emi) : '', t.months == null ? '' : isFinite(t.months) ? t.months + ' months' : 'EMI below interest']; });
    const keep = [0, 1, 2, 3, 4].filter(i => i < 2 || rows.some(r => r[i] !== ''));
    S_.liab = section_('Liabilities', table_(keep.map(i => ['Loan', 'Outstanding', 'Rate', 'EMI', 'Remaining'][i]), rows.map(r => keep.map(i => r[i])), keep.map(i => i > 0)));
  }

  /* Markets */
  if (M.list.length) {
    const rows = M.list.map(x => {
      const p = ref.d1 && ref.d1.mkt[x.label];
      const lvl = x.kind === 'inr' ? inr_(x.val) : x.kind === 'fx' ? '₹' + x.val.toFixed(2) : num2_(x.val);
      return [x.label, lvl, p != null ? pctDelta_((x.val - p) / p) : ''];
    }).concat([['Inflation', pct_(M.cpi, 2), ''], ['RBI repo rate', pct_(M.repo, 2), '']].filter(r => r[1] !== '—'));
    const keep = [0, 1, 2].filter(i => i < 2 || rows.some(r => r[i] !== ''));
    S_.markets = section_('Markets', (ref.d1 ? `<div style="font-size:13px;color:${C.muted};margin:-6px 0 10px">Change against the ${closeOf(ref.d1)}</div>` : '')
      + table_(keep.map(i => ['', 'Level', 'Since close'][i]), rows.map(r => keep.map(i => r[i])), keep.map(i => i > 0)));
  }

  /* Spend */
  if (X) {
    const paceAvg = X.avg ? X.avg * X.pace : null, vsAvg = paceAvg ? (X.thisMonth - paceAvg) / paceAvg : null;
    const mx = Math.max.apply(null, X.monthTot.slice(0, X.mi + 1).concat([1]));
    const span = MON[0] + ' to ' + MON[X.mi];
    S_.spend = section_(`Spend · ${X.fy}`,
      `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
        ${statTd_(MON[X.mi] + ' so far', inrC_(X.thisMonth), vsAvg == null ? '' : `<span style="color:${vsAvg > 0 ? C.red : C.green}">${sign_(vsAvg)}${Math.abs(vsAvg * 100).toFixed(0)}% vs pace</span>`)}
        ${X.avg ? statTd_('Monthly average', inrC_(X.avg), X.avgNote) : ''}
        ${statTd_('Year to date', inrC_(X.ytd), span)}</tr></table>
       <div style="font-size:12px;color:${C.muted};margin-top:8px">Includes home loan EMI ${inrC_(X.emiYtd)} and travel ${inrC_(X.travelYtd)} for ${span}.</div>
       ${monthBars_(X.monthTot.slice(0, X.mi + 1).map((t, j) => ({ v: t, label: MON[j], hi: j === X.mi })), mx)}
       <div style="height:16px"></div>
       ${offPace_(X)}`);
  }

  /* Salary */
  if (I) {
    const P = I.pay, series = MON.map((m, i) => ({ v: i > P.lastIdx ? P.runGross : P.gross[i], label: m, proj: i > P.lastIdx }));
    S_.salary = section_(`Salary · ${I.fy}`,
      `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
        ${statTd_('Gross to date', inrC_(P.grossToDate), (P.lastIdx + 1) + ' payslips, Apr to ' + MON[P.lastIdx])}
        ${statTd_('Take-home to date', inrC_(P.takeHomeToDate), 'TDS ' + inrC_(P.tdsToDate))}
        ${statTd_('Full year, estimated', inrC_(P.grossFY), P.ahead ? P.ahead + ' months at ' + inrC_(P.runGross) : 'all payslips in')}</tr></table>
       ${monthBars_(series, Math.max.apply(null, series.map(s => s.v).concat([1])))}
       ${P.ahead ? `<div style="font-size:12px;color:${C.muted};margin-top:6px"><span style="display:inline-block;width:10px;height:10px;background:${C.accent};border-radius:2px"></span> Actual &nbsp; <span style="display:inline-block;width:10px;height:10px;background:#E4E8EF;border:1px dashed #B8C0CC;border-radius:2px"></span> Estimated from the ${MON[P.lastIdx]} payslip, one-off bonuses excluded</div>` : ''}`);

    /* Tax */
    const S = I.sel, other = I.regime === 'old' ? I.new : I.old, save = other.tax - S.tax;
    const next = I.sched.find(s => !s.past);
    const intTotal = sum_(I.sched.map(s => s.int)) + I.s424;
    const status = s => !I.applies ? pill_('Not required', C.muted, C.surface) : s.gap < 1 ? pill_('Covered', C.green, C.greenBg) : s.past ? pill_('Short', C.red, C.redBg) : pill_('Due', C.amber, C.amberBg);
    S_.tax = section_(`Tax · ${I.fy}`,
      `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
        ${statTd_('Liability, estimated', inrC_(S.tax), (I.regime === 'old' ? 'Old' : 'New') + ' regime · ' + pct_(S.tax / I.grossInc, 1))}
        ${statTd_('Paid so far', inrC_(I.paidSoFar), 'TDS ' + inrC_(I.pay.tdsToDate) + (I.advPaid ? ' · advance ' + inrC_(I.advPaid) : ''))}
        ${statTd_('Balance', inrC_(Math.abs(I.balance)), I.balance >= 0 ? 'still to pay' : 'overpaid')}</tr></table>
       <div style="height:14px"></div>${subhead_('INCOME TAXED · FULL YEAR', C.muted)}
       ${table_(null, [
         ['Salary', inr_(I.pay.grossFY), 'estimated'],
         ['Bonus outside payslips', inr_(I.inc.bonus), 'CoI tab'],
         ['SGB interest', inr_(I.inc.sgb), I.payouts ? inr_(I.payouts.sgbRecv) + ' received' : 'CoI tab'],
         ['Dividends', inr_(I.inc.div), I.payouts ? 'payouts tab' : ''],
         ['REIT, taxable part', inr_(I.inc.reit), I.payouts ? 'payouts tab' : 'CoI tab'],
         ['Savings account interest', inr_(I.inc.sav), 'CoI tab'],
         ['Fixed deposit interest', inr_(I.inc.fd), 'CoI tab'],
         ['Capital gains', inr_(I.inc.stcg + I.inc.ltcg), 'CoI tab'],
         ['<b>Gross income</b>', `<b>${inr_(I.grossInc)}</b>`, '']
       ].filter(r => r[1] !== inr_(0) || /Gross/.test(r[0])).map(r => [r[0] + (r[2] ? `<span style="font-size:12px;color:${C.muted}"> · ${r[2]}</span>` : ''), r[1]]), [0, 1])}
       <div style="height:10px"></div>
       ${table_(null, [
         [`TDS on the ${I.pay.ahead} payslips ahead`, inr_(I.tdsAhead)],
         ['Left after all TDS' + (I.advPaid ? ' and advance tax' : ''), `<b style="color:${I.afterTds > 1 ? C.red : C.green}">${I.afterTds > 1 ? inr_(I.afterTds) + ' payable' : inr_(-I.afterTds) + ' refund'}</b>`],
         [`${I.regime === 'old' ? 'New' : 'Old'} regime would be`, `${inr_(other.tax)} <span style="color:${save >= 0 ? C.green : C.red}">(${save >= 0 ? 'yours saves ' + inr_(save) : 'it saves ' + inr_(-save)})</span>`]
       ].filter(r => I.pay.ahead || !/payslips ahead/.test(r[0])), [0, 1])}
       ${subhead_('ADVANCE TAX · ON ' + inr_(I.netDue) + ' LEFT AFTER TDS', C.muted)}
       ${I.applies ? '' : `<div style="font-size:13px;color:${C.green};margin-bottom:8px">Below ₹10,000, so advance tax and its interest do not apply.</div>`}
       ${table_(['Instalment', 'Required, cumulative', 'Paid, cumulative', 'Not paid', 'Interest · s.425', ''],
         I.sched.map(s => [`${s.q} · ${fmtDate_(s.due)}`, inr_(s.req), inr_(s.paidCum), s.gap >= 1 && I.applies ? `<b style="color:${s.past ? C.red : C.amber}">${inr_(s.gap)}</b>` : inr_(s.gap), s.int ? `<span style="color:${s.past ? C.red : C.amber}">${inr_(s.int)}${s.past ? '' : ' if unpaid'}</span>` : inr_(0), status(s)]),
         [0, 1, 1, 1, 1, 1])}
       ${table_(null, [
         [`Interest · s.424 · if advance tax stays below 90% (${inr_(0.9 * I.netDue)}), ${I.s424Months} months to filing on ${fmtDate_(I.filing)}`, I.s424 ? `<span style="color:${C.red}">${inr_(I.s424)}</span>` : inr_(0)],
         ['<b>Interest if nothing more is paid</b>', `<b style="color:${intTotal ? C.red : C.green}">${inr_(intTotal)}</b>`]
       ], [0, 1])}
       ${I.applies && next && next.gap >= 1 ? `<div style="margin-top:12px;padding:12px 14px;border-radius:6px;background:${C.amberBg};color:${C.amber};font-size:14px;font-weight:600">Pay ${inr_(next.gap)} by ${fmtDate_(next.due)} to avoid ${inr_(next.int)} of interest under section 425.</div>` : ''}`);
  }

  /* EMI account check: shown at the very top, only in the days before the EMI date */
  const G = CONFIG.EMI_GUARD;
  if (G) {
    const today = startOfDay_(now), y = today.getFullYear(), m = today.getMonth();
    const due = today.getDate() <= G.EMI_DAY ? new Date(y, m, G.EMI_DAY) : new Date(y, m + 1, G.EMI_DAY);
    const left = Math.round((due - today) / 864e5);
    const acct = A.filter(h => G.ACCOUNT.test(h.cust) && G.VEHICLE.test(h.veh + ' ' + h.inst));
    const bal = sum_(acct.map(h => h.val));
    if (acct.length && left <= G.DAYS_BEFORE && bal < G.MIN) {
      const loan = W.holdings.filter(h => h.g === 'Liability').map(h => loanTerms_(h, X)).find(t => t.emi);
      const emi = loan ? loan.emi : null, short = G.MIN - bal, critical = emi != null && bal < emi;
      const fg = critical ? C.red : C.amber, bg = critical ? C.redBg : C.amberBg;
      const src = A.filter(h => /tier 1/i.test(h.liq) && acct.indexOf(h) < 0 && !isReward_(h) && h.val >= 1000).sort((a, c) => c.val - a.val).slice(0, 3);
      S_.emiAlert = `<tr><td style="background:${bg};border-radius:8px;padding:16px 18px;border:1px solid ${fg}">
        <div style="font-size:15px;font-weight:800;color:${fg}">${critical ? 'EMI at risk' : 'EMI account below ' + inr_(G.MIN)}: top up ${inr_(short)} by ${fmtDay_(new Date(due.getTime() - 864e5))}</div>
        <div style="font-size:13px;line-height:1.5;color:${C.ink};margin-top:6px">${esc_(acct.map(h => h.cust).filter((x, i, a) => a.indexOf(x) === i).join(', '))} savings holds <b>${inr_(bal)}</b>${emi ? ` against the <b>${inr_(emi)}</b> EMI` : ''} due by NACH on <b>${fmtDay_(due)}</b>, ${left === 0 ? 'today' : left === 1 ? 'tomorrow' : 'in ' + left + ' days'}.${critical ? ' The debit will fail at this balance.' : ' The debit clears, but leaves less than the ' + inr_(G.MIN) + ' buffer.'}</div>
        ${src.length ? `<div style="font-size:13px;color:${C.ink};margin-top:8px">Move money from: ${src.map(h => { const sub = [h.veh, h.inst].filter(x => x && h.cust.indexOf(x) < 0 && x.indexOf(h.cust) < 0)[0] || ''; return `<b>${esc_(h.cust)}</b>${sub ? ' ' + esc_(sub) : ''} (${inrC_(h.val)})`; }).join(' · ')}. Or redeem a liquid holding.</div>` : ''}
      </td></tr><tr><td height="12"></td></tr>`;
    }
  }

  /* Rewards and balances: points, vouchers and wallet balances, soonest expiry first */
  const dayTo = d => d ? Math.round((startOfDay_(d) - startOfDay_(now)) / 864e5) : null;
  const RW = A.filter(isReward_).map(h => {
    const { pts, rate } = rewardPts_(h);
    const lots = pts != null ? (RL[h.key] || []).filter(l => l.left > 0).sort((a, c) => a.exp - c.exp) : [];
    const exp = lots.length ? lots[0].exp : noteDate_(h);
    return { h, pts, rate, lots, exp, days: dayTo(exp) };
  }).sort((a, c) => (a.days == null) - (c.days == null) || (a.days != null ? a.days - c.days : c.h.val - a.h.val));
  if (RW.length) {
    const tot = sum_(RW.map(x => x.h.val)), gone = RW.filter(x => x.days != null && x.days < 0), soon = RW.filter(x => x.days != null && x.days >= 0 && x.days <= 90), noDate = RW.filter(x => x.days == null);
    const when = x => x.days == null ? pill_('No date', C.muted, C.surface)
      : x.days < 0 ? pill_('Expired', C.red, C.redBg)
      : pill_(x.days === 0 ? 'Today' : x.days < 60 ? x.days + ' days' : Math.round(x.days / 30.4) + ' months', x.days <= 30 ? C.red : x.days <= 90 ? C.amber : C.green, x.days <= 30 ? C.redBg : x.days <= 90 ? C.amberBg : C.greenBg);
    const note = [
      gone.length ? `<b style="color:${C.red}">${inrC_(sum_(gone.map(x => x.h.val)))} has expired</b>: ${gone.map(x => esc_(x.h.cust) + ' on ' + fmtDate_(x.exp)).join(', ')}.` : '',
      soon.length ? `<b style="color:${C.amber}">${inrC_(sum_(soon.map(x => x.h.val)))} expires within 90 days</b>: ${soon.map(x => esc_(x.h.cust) + ' on ' + fmtDate_(x.exp)).join(', ')}.` : '',
      noDate.length ? `No expiry date for ${noDate.map(x => esc_(x.h.cust)).join(', ')}. Add it in Goal/Comments as [dd Mmm yyyy].` : ''
    ].filter(Boolean).join('<br>');
    // Points expiring, grouped by month across programmes; the earliest three months.
    const byMonth = {};
    RW.forEach(x => x.lots.forEach(l => {
      const k = Utilities.formatDate(l.exp, CONFIG.TIMEZONE, 'yyyy-MM');
      const g = byMonth[k] = byMonth[k] || { d: l.exp, pts: 0, val: 0, progs: {} };
      g.pts += l.left; g.val += l.left * x.rate; g.progs[x.h.cust] = (g.progs[x.h.cust] || 0) + l.left;
    }));
    const months = Object.keys(byMonth).sort().slice(0, 3).map(k => byMonth[k]);
    const expTable = months.length ? `<div style="height:14px"></div>${subhead_('POINTS EXPIRING · EARLIEST ' + months.length + ' MONTH' + (months.length > 1 ? 'S' : ''), C.muted)}
       ${table_(['Month', 'Points', 'Value'], months.map(g => {
         const dd = dayTo(g.d);
         return [`${Utilities.formatDate(g.d, CONFIG.TIMEZONE, 'MMM yyyy')}${dd < 0 ? ' ' + pill_('Expired', C.red, C.redBg) : dd <= 90 ? ' ' + pill_(dd + ' days', dd <= 30 ? C.red : C.amber, dd <= 30 ? C.redBg : C.amberBg) : ''}<div style="font-size:12px;color:${C.muted}">${Object.keys(g.progs).map(p => esc_(p) + ' ' + g.progs[p].toLocaleString('en-IN')).join(' · ')}</div>`,
           g.pts.toLocaleString('en-IN'), `<b>${inr_(g.val)}</b>`];
       }), [0, 1, 1])}
       <div style="font-size:12px;color:${C.muted};margin-top:8px">Points expire ${CONFIG.REWARD_EXPIRY_YEARS} years after they're earned. Each top-up is dated at the ${CONFIG.CLOSE_LABEL.toLowerCase()} close; redemptions use the oldest points first.</div>` : '';
    S_.rewards = section_('Rewards and balances',
      `<div style="font-size:13px;color:${C.muted};margin-bottom:10px">Worth <b style="color:${C.ink}">${inrC_(tot)}</b> across ${RW.length}, soonest expiry first</div>
       ${table_(['', 'Points', 'Value', 'Expires'], RW.map(x => [
         `${esc_(x.h.cust)}<div style="font-size:12px;color:${C.muted}">${esc_([x.h.veh, x.h.inst].filter(Boolean).join(' · '))}</div>`,
         x.pts != null ? `${x.pts.toLocaleString('en-IN')}<div style="font-size:12px;color:${C.muted}">× ₹${x.rate.toFixed(2)}</div>` : `<span style="color:${C.muted}">—</span>`,
         `<b>${inr_(x.h.val)}</b>`,
         `${x.exp ? fmtDate_(x.exp) + '<div style="margin-top:3px">' : '<div>'}${when(x)}</div>${x.lots.length > 1 ? `<div style="font-size:11px;color:${C.muted};margin-top:2px">earliest of ${x.lots.length} lots</div>` : ''}`
       ]), [0, 1, 1, 1])}
       ${note ? `<div style="font-size:13px;line-height:1.5;color:${C.ink};margin-top:10px">${note}</div>` : ''}
       ${expTable}`);
  }

  /* Personal ratios */
  if (X || I) {
    const net = I ? I.grossInc - I.sel.tax : null;
    const avgSpend = X && X.nUsed ? sum_(X.monthTot.filter((t, j) => j <= X.mi && t > 0)) / X.nUsed : null; // all spend, EMI and travel included
    const yearSpend = avgSpend != null ? avgSpend * 12 : null, yearOut = yearSpend;
    const tier1 = sum_(A.filter(h => /tier 1/i.test(h.liq)).map(h => h.val));
    const cushion = sum_(A.filter(h => /tier 1/i.test(h.liq) || /safety/i.test(h.fn)).map(h => h.val));
    const R = [
      ['Savings rate', net && yearOut != null ? (net - yearOut) / net : null, v => pct_(v, 1), 'Take-home kept after spend', null],
      ['Emergency cover', avgSpend ? cushion / avgSpend : null, v => v.toFixed(1) + ' months', 'Instant and safety holdings ÷ monthly spend · 3 to 6', v => v >= 6 ? 'g' : v >= 3 ? 'a' : 'r'],
      ['Liquidity', yearSpend ? tier1 / yearSpend : null, v => v.toFixed(2), 'Instant-access assets ÷ yearly spend · 0.15 or more', v => v >= .15 ? 'g' : 'r'],
      ['Debt service', net && X && X.emiMonthly != null ? X.emiMonthly * 12 / net : null, v => v.toFixed(2), 'Loan EMIs ÷ take-home · 0.30 or less', v => v <= .3 ? 'g' : 'r'],
      ['Housing cost', I && X && X.housingMonthly != null ? X.housingMonthly * 12 / I.grossInc : null, v => pct_(v, 1), 'Household and utilities ÷ gross income · 28% or less', v => v <= .28 ? 'g' : 'r'],
      ['Debt to assets', W.assets ? W.liab / W.assets : null, v => v.toFixed(2), 'Total debt ÷ total assets · lower is better', null],
      ['Tax rate', I ? I.sel.tax / I.grossInc : null, v => pct_(v, 1), 'Estimated tax ÷ gross income', null]
    ].filter(r => r[1] != null && isFinite(r[1]));
    const pillFor = (f, v) => { if (!f) return ''; const s = f(v); return s === 'g' ? pill_('Healthy', C.green, C.greenBg) : s === 'a' ? pill_('Borderline', C.amber, C.amberBg) : pill_('Off target', C.red, C.redBg); };
    if (R.length) S_.ratios = section_('Personal ratios', table_(null, R.map(([k, v, f, note, judge]) => [`${k}<div style="font-size:12px;color:${C.muted}">${note}</div>`, `<b>${f(v)}</b>`, pillFor(judge, v)]), [0, 1, 1])
      + (X ? `<div style="font-size:12px;color:${C.muted};margin-top:8px">Income and tax for the full year, estimated; spend annualised from ${X.nUsed} months.</div>` : ''));
  }

  /* Locker */
  if (L && L.last) S_.locker = section_('Locker',
    `<div style="font-size:13px;color:${C.muted};margin-bottom:10px">Last visit ${fmtDate_(L.last.d)}${L.last.by ? ' by ' + esc_(L.last.by) : ''}</div>
     <details style="border:1px solid ${C.line};border-radius:6px;padding:10px 14px">
       <summary style="cursor:pointer;font-size:14px;font-weight:700;color:${C.ink}">${L.items.length} items inside<span class="vlk-link" style="font-weight:400"> · <a href="${L.url}" style="color:${C.accent}">View the list</a></span></summary>
       <div class="vlk" style="display:none"><ol style="margin:10px 0 2px;padding-left:20px;font-size:14px;line-height:1.6;color:${C.ink}">${L.items.map(x => `<li>${esc_(x)}</li>`).join('')}</ol></div>
     </details>`);

  /* Cash flow: from 1 April to the last payslip month, income → tax → deductions → take-home → spend → saved */
  if (I && I.pay.lastIdx >= 0) {
    const k = I.pay.lastIdx, upTo = arr => sum_(arr.slice(0, k + 1)), PO = I.payouts;
    const endOfK = new Date(I.Y + (k >= 9 ? 1 : 0), (k + 3) % 12 + 1, 0);
    const advToDate = sum_([0, 1, 2, 3].map(q => I.sched[q].due <= endOfK ? I.advByQ[q] : 0));
    const col = (months, adv) => {
      const g = sum_(months.map(j => I.pay.gross[j])), other = PO ? sum_(months.map(j => PO.cash[j])) : 0;
      const tax = sum_(months.map(j => I.pay.tds[j])) + adv, ded = sum_(months.map(j => I.pay.ded[j] - I.pay.tds[j]));
      const spend = X ? sum_(months.map(j => X.monthTot[j])) : null;
      const take = g + other - tax - ded;
      return { inc: g + other, g, other, tax, ded, take, spend, saved: spend != null ? take - spend : null };
    };
    const ytdM = MON.map((_, j) => j).filter(j => j <= k);
    const Y_ = col(ytdM, advToDate), M_ = col([k], 0);
    const lines = [
      ['Income', Y_.inc, M_.inc, C.accent, `Salary ${inrC_(Y_.g)}${Y_.other ? ' · payouts ' + inrC_(Y_.other) : ''}`],
      ['Tax', -Y_.tax, -M_.tax, C.red, 'TDS' + (advToDate ? ' and advance tax' : '')],
      ['Other deductions', -Y_.ded, -M_.ded, '#9AA4B5', 'PF, professional tax, LWF'],
      ['Take-home', Y_.take, M_.take, C.ink, ''],
      ['Spend', Y_.spend != null ? -Y_.spend : null, M_.spend != null ? -M_.spend : null, C.amber, 'Including EMI and travel'],
      ['Saved', Y_.saved, M_.saved, C.green, '']
    ].filter(l => l[1] != null);
    const span = MON[0] + ' to ' + MON[k];
    S_.cashflow = section_(`Cash flow · ${I.fy}`,
      `<table role="presentation" cellpadding="0" cellspacing="0" width="100%">
        <tr><td></td><td></td><td align="right" style="font-size:11px;font-weight:700;letter-spacing:.5px;color:${C.muted};padding-bottom:6px;white-space:nowrap">${span.toUpperCase()}</td><td align="right" style="font-size:11px;font-weight:700;letter-spacing:.5px;color:${C.muted};padding:0 0 6px 12px">${MON[k].toUpperCase()}</td></tr>
        ${lines.map(([l, y, m, colr, note]) => `<tr>
          <td style="font-size:14px;color:${C.ink};padding:7px 10px 7px 0;white-space:nowrap;${/Take-home|Saved/.test(l) ? 'font-weight:700' : ''}">${l}${note ? `<div style="font-size:11px;color:${C.muted};font-weight:400;white-space:normal">${note}</div>` : ''}</td>
          <td width="40%" style="padding:7px 10px 7px 0"><div style="background:${C.surface};border-radius:2px;height:12px"><div style="width:${Math.max(1, Math.min(100, Math.abs(y) / Y_.inc * 100)).toFixed(1)}%;background:${colr};height:12px;border-radius:2px"></div></div></td>
          <td align="right" style="font-size:14px;color:${C.ink};padding:7px 0;white-space:nowrap;${/Take-home|Saved/.test(l) ? 'font-weight:700' : ''}">${y < 0 ? '−' : ''}${inrC_(Math.abs(y))} <span style="font-size:12px;color:${C.muted};font-weight:400">${(Math.abs(y) / Y_.inc * 100).toFixed(0)}%</span></td>
          <td align="right" style="font-size:13px;color:${C.muted};padding:7px 0 7px 12px;white-space:nowrap">${m < 0 ? '−' : ''}${inrC_(Math.abs(m))}</td></tr>`).join('')}
      </table>
      <div style="font-size:12px;color:${C.muted};margin-top:8px">Up to the last payslip (${MON[k]}). Payouts are SGB interest, dividends and REIT distributions from the payouts tab.</div>`);
  }

  b = ['emiAlert', 'worth', 'trend', 'markets', 'movers', 'cashflow', 'salary', 'tax', 'spend', 'rewards', 'liab', 'alloc', 'returns', 'ratios', 'locker'].map(x => S_[x] || '').join('');

  return `<!doctype html><html><head><style>details[open] .vlk{display:block !important} details[open] .vlk-link{display:none !important}</style></head><body style="margin:0;padding:0;background:${C.surface}">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${C.surface}"><tr><td align="center" style="padding:20px 12px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;font-family:${C.font};color:${C.ink}">
    <tr><td style="padding:0 4px 14px"><table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      <td width="44"><div style="width:36px;height:36px;border-radius:6px;background:${C.accent};color:#fff;font-weight:800;font-size:20px;line-height:36px;text-align:center">V</div></td>
      <td style="font-size:20px;font-weight:800;color:${C.ink}">Vault</td></tr></table></td></tr>
    ${b}
    <tr><td style="font-size:13px;color:${C.muted};padding:4px 4px 0">Source: <a href="${ss.getUrl()}" style="color:${C.accent};font-weight:700">${esc_(ss.getName())}</a></td></tr>
  </table></td></tr></table></body></html>`;
}

/* ---------- Components ---------- */

function card_(inner) { return `<tr><td style="background:${C.bg};border-radius:8px;padding:22px 22px 18px;border:1px solid ${C.line}">${inner}</td></tr><tr><td height="12"></td></tr>`; }
function section_(title, inner) { return card_(`<div style="font-size:16px;font-weight:700;color:${C.ink};margin-bottom:12px">${esc_(title)}</div>${inner}`); }
function subhead_(t, col) { return `<div style="font-size:12px;font-weight:700;color:${col};letter-spacing:.5px;margin:4px 0 4px">${esc_(t)}</div>`; }
function twoCol_(cols) {
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>${cols.map((c, i) => `<td valign="top" width="${cols.length === 2 ? 50 : 100}%" style="${i ? `padding-left:12px;border-left:1px solid ${C.line}` : (cols.length === 2 ? 'padding-right:12px' : '')}">${c}</td>`).join('')}</tr></table>`;
}
function deltaPill_(label, d, base) {
  const fg = d > 0 ? C.green : d < 0 ? C.red : C.muted, bg = d > 0 ? C.greenBg : d < 0 ? C.redBg : C.surface;
  return `<span style="display:inline-block;background:${bg};color:${fg};border-radius:3px;padding:6px 12px;margin:0 6px 6px 0;font-size:13px;font-weight:700;white-space:nowrap">${sign_(d)}${inrC_(Math.abs(d))}${base ? ' (' + pctChange_(d / base) + ')' : ''} <span style="font-weight:500;color:${C.muted}">${esc_(label.toLowerCase().indexOf('since') === 0 ? label.charAt(0).toLowerCase() + label.slice(1) : 'in ' + label)}</span></span>`;
}
function pctChange_(p) { const a = Math.abs(p * 100); return (a > 0 && a < 0.01 ? sign_(p) + '<0.01' : sign_(p) + a.toFixed(2)) + '%'; }
function deltaText_(d, base) {
  if (Math.abs(d) < 1) return `<span style="color:${C.muted}">0</span>`;
  return `<span style="color:${d > 0 ? C.green : C.red};font-weight:600;white-space:nowrap">${sign_(d)}${inrC_(Math.abs(d))}${base ? ` <span style="font-weight:400">(${sign_(d)}${Math.abs(d / base * 100).toFixed(2)}%)</span>` : ''}</span>`;
}
function pctDelta_(p) { return isNaN(p) ? '' : `<span style="color:${p > 0 ? C.green : p < 0 ? C.red : C.muted};font-weight:600">${sign_(p)}${Math.abs(p * 100).toFixed(2)}%</span>`; }
function pill_(t, fg, bg) { return `<span style="display:inline-block;background:${bg};color:${fg};border-radius:3px;padding:3px 10px;font-size:12px;font-weight:700;white-space:nowrap">${esc_(t)}</span>`; }
/** Name, value · share, optional 24-hour delta, and a bar. Used by Allocation, Liquidity and purpose, Spend. */
function miniBars_(rows, base) {
  return rows.map(x => `<div style="padding:6px 0">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      <td style="font-size:13px;color:${C.ink}">${esc_(x.k)}</td>
      <td align="right" style="font-size:13px;color:${C.ink};white-space:nowrap"><b>${inrC_(x.v)}</b> <span style="color:${C.muted}">· ${base ? (x.v / base * 100).toFixed(1) : '0.0'}%</span></td></tr>
      ${x.d != null && Math.abs(x.d) >= 1 ? `<tr><td></td><td align="right" style="font-size:12px">${deltaText_(x.d, x.p)}</td></tr>` : ''}</table>
    <div style="background:${C.surface};border-radius:3px;height:6px;margin-top:4px"><div style="width:${Math.max(1, Math.min(100, base ? x.v / base * 100 : 0)).toFixed(1)}%;background:${C.accent};height:6px;border-radius:3px"></div></div></div>`).join('');
}
/** Categories furthest from their usual month, at today's point in the month. Centre line = usual. */
function offPace_(X) {
  const sig = X.off.filter(x => x.pct == null ? x.act > 0 : Math.abs(x.pct) >= 0.1 && Math.abs(x.gap) >= 500);
  const over = sig.filter(x => x.gap > 0).sort((a, c) => c.gap - a.gap).slice(0, 3);
  const under = sig.filter(x => x.gap < 0).sort((a, c) => a.gap - c.gap).slice(0, 2);
  const when = X.pace < 1 ? `day ${X.day} of ${X.dim}, so ${Math.round(X.pace * 100)}% of a usual month` : 'the full month';
  const head = `${subhead_(MON[X.mi].toUpperCase() + ' · OFF PACE', C.muted)}<div style="font-size:12px;color:${C.muted};margin-bottom:8px">Each category against its usual month, scaled to ${when}. Right of the line in red: spending more than usual. Left in green: less.</div>`;
  const list = over.concat(under);
  if (!list.length) return head + `<div style="font-size:14px;color:${C.green};font-weight:600">Every category is within 10% of its usual pace.</div>`;
  const mx = Math.max.apply(null, list.map(x => Math.abs(x.gap)));
  return head + `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="table-layout:fixed">${list.map(x => {
    const w = Math.max(2, Math.abs(x.gap) / mx * 50).toFixed(1) + '%', up = x.gap > 0, col = up ? C.red : C.green;
    const bar = `<div style="position:relative;height:10px;background:${C.surface};border-radius:2px">
      <div style="position:absolute;left:50%;top:-3px;width:1px;height:16px;background:${C.ink}"></div>
      <div style="position:absolute;${up ? 'left' : 'right'}:50%;top:0;width:${w};height:10px;background:${col};border-radius:2px"></div></div>`;
    return `<tr>
      <td width="30%" style="font-size:14px;color:${C.ink};padding:8px 10px 8px 0;border-bottom:1px solid ${C.line}">${esc_(x.k)}<div style="font-size:12px;color:${C.muted}">${inrC_(x.act)} so far · usual ${inrC_(x.exp)}</div></td>
      <td width="38%" style="padding:8px 10px 8px 0;border-bottom:1px solid ${C.line}">${bar}</td>
      <td width="32%" align="right" style="font-size:14px;font-weight:700;color:${col};padding:8px 0;white-space:nowrap;border-bottom:1px solid ${C.line}">${sign_(x.gap)}${inrC_(Math.abs(x.gap))} · ${x.pct == null ? 'new' : sign_(x.pct) + Math.abs(x.pct * 100).toFixed(0) + '%'}</td></tr>`;
  }).join('')}</table>`;
}
function monthBars_(series, mx) {
  const w = (100 / series.length).toFixed(3) + '%';
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin-top:14px;height:74px;table-layout:fixed"><tr>${series.map(s =>
    `<td valign="bottom" align="center" width="${w}" style="width:${w};padding:0 2px"><div style="height:${Math.max(2, Math.round(s.v / mx * 56))}px;box-sizing:border-box;background:${s.proj ? '#E4E8EF' : s.hi === false ? '#D6DCE6' : C.accent};${s.proj ? 'border:1px dashed #B8C0CC;' : ''}border-radius:2px 2px 0 0"></div><div style="font-size:10px;color:${C.muted};padding-top:4px">${s.label}</div></td>`).join('')}</tr></table>`;
}
function table_(head, rows, right) {
  const th = head ? `<tr>${head.map((h, i) => `<td align="${right[i] ? 'right' : 'left'}" style="font-size:11px;font-weight:700;letter-spacing:.5px;color:${C.muted};padding:0 ${i < head.length - 1 ? '10px' : '0'} 6px 0;border-bottom:1px solid ${C.line};text-transform:uppercase">${esc_(h)}</td>`).join('')}</tr>` : '';
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%">${th}${rows.map(r => `<tr>${r.map((c, i) => `<td align="${right[i] ? 'right' : 'left'}" style="font-size:14px;color:${C.ink};padding:8px ${i < r.length - 1 ? '10px' : '0'} 8px 0;border-bottom:1px solid ${C.line}">${c}</td>`).join('')}</tr>`).join('')}</table>`;
}
function statTd_(label, value, sub) {
  return `<td valign="top" width="33%" style="padding-right:8px"><div style="font-size:12px;color:${C.muted}">${label}</div><div style="font-size:20px;font-weight:800;color:${C.ink};margin:2px 0">${value}</div><div style="font-size:12px;color:${C.muted}">${sub || ''}</div></td>`;
}
function targetBlock_(W, now) {
  const p = Math.max(0, Math.min(1, W.nw / W.target)), gap = W.target - W.nw;
  const days = W.targetDate ? Math.ceil((W.targetDate - startOfDay_(now)) / 864e5) : null;
  // Time runs from 1 April of the current financial year to the target date.
  const start = W.targetDate ? new Date(fyStart_(now), 3, 1) : null;
  const tp = start ? Math.max(0, Math.min(1, (startOfDay_(now) - start) / (W.targetDate - start))) : null;
  return `<div style="margin-top:16px;padding-top:14px;border-top:1px solid ${C.line}">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      <td style="font-size:15px;color:${C.ink}"><b>Target ${inrC_(W.target)}</b></td>
      <td align="right" style="font-size:15px;font-weight:700;color:${gap > 0 ? C.ink : C.green}">${gap > 0 ? inrC_(gap) + ' to go' : 'Reached'}</td></tr></table>
    ${bar2_('Amount reached', p, C.accent)}
    ${W.targetDate ? bar2_('Time elapsed', tp, C.muted) : ''}
    ${W.targetDate ? `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      <td style="font-size:13px;color:${C.muted}">${fmtDate_(W.targetDate)}</td>
      <td align="right" style="font-size:13px;color:${C.muted}">${days > 0 ? days + ' days remaining' : days === 0 ? 'Today' : Math.abs(days) + ' days past'}</td></tr></table>` : ''}</div>`;
}
function bar2_(label, p, col) {
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin:8px 0"><tr>
    <td width="112" style="font-size:12px;color:${C.muted};white-space:nowrap">${label}</td>
    <td><div style="background:${C.surface};border-radius:3px;height:10px"><div style="width:${(p * 100).toFixed(1)}%;background:${col};height:10px;border-radius:3px"></div></div></td>
    <td width="52" align="right" style="font-size:13px;font-weight:700;color:${C.ink}">${(p * 100).toFixed(1)}%</td></tr></table>`;
}
/** Day-on-day change between consecutive daily closes: T against T−1. */
function dailyChanges_(closes, days) {
  const out = [];
  for (let i = 1; i < closes.length; i++) {
    const a = closes[i - 1], c = closes[i];
    out.push({ k: c.k, v: c.nw, d: c.nw - a.nw, p: a.nw ? (c.nw - a.nw) / Math.abs(a.nw) : 0 });
  }
  return out.slice(-days);
}
/** Bars above zero for a rise, below zero for a fall; height is the size of the day's change. */
function trendBars_(s) {
  const H = 44, mx = Math.max.apply(null, s.map(x => Math.abs(x.d)).concat([1]));
  const w = (100 / s.length).toFixed(3) + '%', ht = x => Math.max(2, Math.round(Math.abs(x.d) / mx * H));
  const col = x => x.d > 0 ? C.green : x.d < 0 ? C.red : '#B8C0CC';
  const day = k => Utilities.formatDate(new Date(k + 'T12:00:00'), CONFIG.TIMEZONE, 'dd MMM');
  const tip = x => `${day(x.k)}: ${x.d < 0 ? '−' : '+'}${inrC_(Math.abs(x.d))} (${pctChange_(x.p)})`;
  const up = s.map(x => `<td width="${w}" valign="bottom" style="width:${w};height:${H}px;padding:0 1px" title="${esc_(tip(x))}">${x.d >= 0 ? `<div style="height:${ht(x)}px;background:${col(x)};border-radius:2px 2px 0 0"></div>` : ''}</td>`).join('');
  const dn = s.map(x => `<td width="${w}" valign="top" style="width:${w};height:${H}px;padding:0 1px;border-top:1px solid ${C.ink}">${x.d < 0 ? `<div style="height:${ht(x)}px;background:${col(x)};border-radius:0 0 2px 2px"></div>` : ''}</td>`).join('');
  const axis = (t, va) => `<td width="56" valign="${va}" style="width:56px;font-size:10px;color:${C.muted};text-align:right;padding-right:6px">${t}</td>`;
  const last = s[s.length - 1], first = s[0];
  const total = sum_(s.map(x => x.d)), base = first.v - first.d;
  const best = s.reduce((a, x) => x.d > a.d ? x : a, s[0]), worst = s.reduce((a, x) => x.d < a.d ? x : a, s[0]);
  const ups = s.filter(x => x.d > 0).length, downs = s.filter(x => x.d < 0).length;
  const signed = x => `${x.d < 0 ? '−' : '+'}${inrC_(Math.abs(x.d))}`;
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="table-layout:fixed">
      <tr>${axis('+' + inrC_(mx), 'top')}${up}</tr>
      <tr>${axis('−' + inrC_(mx), 'bottom')}${dn}</tr></table>
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin-top:4px"><tr>
      <td width="56"></td>
      <td style="font-size:11px;color:${C.muted}">${day(first.k)}</td>
      <td align="right" style="font-size:11px;color:${C.muted}">${day(last.k)}</td></tr></table>
    <div style="font-size:12px;color:${C.muted};margin:6px 0 12px">Each bar is one day's change in net worth, ${CONFIG.CLOSE_LABEL.toLowerCase()} close against the ${CONFIG.CLOSE_LABEL.toLowerCase()} close the day before. Green is a rise, red a fall; taller means a bigger move.</div>
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      ${statTd_('Last day · ' + day(last.k), `<span style="color:${col(last)}">${signed(last)}</span>`, pctChange_(last.p))}
      ${statTd_('Over ' + s.length + ' days', `<span style="color:${total > 0 ? C.green : total < 0 ? C.red : C.ink}">${total < 0 ? '−' : '+'}${inrC_(Math.abs(total))}</span>`, base ? pctChange_(total / Math.abs(base)) : '')}
      ${statTd_('Up / down days', ups + ' / ' + downs, `Best ${signed(best)} on ${day(best.k)}<br>Worst ${signed(worst)} on ${day(worst.k)}`)}</tr></table>`;
}
function groupSum_(list, fn) {
  const m = {}; list.forEach(h => { const k = fn(h); m[k] = (m[k] || 0) + h.val; });
  return Object.keys(m).map(k => ({ k, v: m[k] })).sort((a, c) => c.v - a.v);
}

/* ---------- Formatting and dates ---------- */

function inr_(n) {
  if (n == null || isNaN(n)) return '—';
  const s = String(Math.round(Math.abs(n))), last3 = s.slice(-3), rest = s.slice(0, -3);
  return (n < 0 ? '−' : '') + '₹' + (rest ? rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' : '') + last3;
}
function inrC_(n) {
  if (n == null || isNaN(n)) return '—';
  const a = Math.abs(n), s = n < 0 ? '−' : '';
  if (a >= 1e7) return s + '₹' + (a / 1e7).toFixed(2) + ' Cr';
  if (a >= 1e5) return s + '₹' + (a / 1e5).toFixed(2) + ' L';
  return inr_(n);
}
const num2_ = n => n.toLocaleString('en-US', { maximumFractionDigits: 2 });
const pct_ = (p, d) => p == null || isNaN(p) ? '—' : (p * 100).toFixed(d == null ? 1 : d) + '%';
const sign_ = d => d > 0 ? '+' : d < 0 ? '−' : '';
const sum_ = a => a.reduce((s, x) => s + (Number(x) || 0), 0);
const esc_ = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmtDate_ = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd MMM yyyy');
const fmtDay_ = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'EEE, dd MMM yyyy');
const startOfDay_ = d => new Date(Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd') + 'T00:00:00');
const fyMonth_ = d => (Number(Utilities.formatDate(d, CONFIG.TIMEZONE, 'M')) - 1 + 9) % 12;
function fyStart_(d) { const y = Number(Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy')), m = Number(Utilities.formatDate(d, CONFIG.TIMEZONE, 'M')); return m >= 4 ? y : y - 1; }
const fyName_ = y => 'FY ' + String(y % 100).padStart(2, '0') + '-' + String((y + 1) % 100).padStart(2, '0');
function monthOf_(x) {
  if (x instanceof Date) return (x.getMonth() + 9) % 12;
  const m = str_(x).match(/^([A-Za-z]{3})[a-z]*\.?(?:[-\s']*\d{2,4})?$/);
  if (!m) return null;
  const i = MONTHS.indexOf(m[1].toLowerCase());
  return i < 0 ? null : (i + 9) % 12;
}
function parseLooseDate_(v) {
  if (v instanceof Date) return v;
  const m = str_(v).match(/(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3})[a-z]*\s+(\d{4})/);
  return m ? new Date(Number(m[3]), MONTHS.indexOf(m[2].toLowerCase()), Number(m[1])) : null;
}
/* ===================== Web app: data ===================== */

/**
 * Called by the page through google.script.run. Returns the whole app as one JSON string.
 * Every call reads the sheet afresh, so the app always shows the values as they are now.
 */
function getAppData(refresh) {
  // Always a fresh read of the sheet: no cache, so every open and every refresh shows the values as they are now.
  SpreadsheetApp.flush();
  return JSON.stringify(buildAppData_());
}

function buildAppData_() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
  const now = new Date(), issues = [];
  const run = (label, fn, dflt) => { try { return fn(); } catch (e) { console.error(label + ': ' + (e.stack || e.message)); issues.push({ part: label, msg: e.message }); return dflt; } };

  const W = readWorth_(ss);
  const A = W.holdings.filter(h => h.g === 'Asset');
  const M = run('Markets', () => readMarkets_(ss), { list: [], cpi: null, repo: null });
  const X = run('Spend', () => readSpend_(ss, now), null);
  const I = run('Income and tax', () => readIncomeTax_(ss, now), null);
  const closes = appCloses_(run('Daily closes', () => appReadLog_(ss), []), now);
  // "Since last close" compares with the latest close from an earlier day, never with a reading taken earlier today.
  const todayKey = dayKey_(now.getTime()), past = closes.filter(c => c.k < todayKey);
  const d1 = past.length ? past[past.length - 1] : null;
  const ref = { d1, past, d7: closest_(past, now.getTime() - 7 * 864e5, 1.5 * 864e5), d30: closest_(past, now.getTime() - 30 * 864e5, 3 * 864e5) };
  const loan = run('Home loan', () => appLoan_(ss, W, now), null);
  const ledger = run('Rewards ledger', () => appRewards_(ss, W, now), null);

  return {
    at: now.getTime(), today: dayKey_(now.getTime()), closeLabel: CONFIG.CLOSE_LABEL, sheetName: ss.getName(), sheetUrl: ss.getUrl(),
    home: run('Home', () => appHome_(W, A, M, closes, ref, loan, now), null),
    income: run('Income', () => appIncome_(I, run('Payout ledger', () => appPayoutLedger_(ss), null)), null),
    tax: run('Tax', () => appTax_(ss, I, now), []),
    payslips: run('Payslips', () => appPayslips_(ss, now), []),
    spend: run('Spend', () => appSpend_(X, loan, ledger, ss, now), null),
    invest: run('Invest', () => appInvest_(ss, W, A, ref), null),
    holdings: run('Holdings', () => appHoldings_(ss, W, A, I, ref, now), {}),
    epf: run('EPF', () => appEpf_(ss), null),
    loan,
    check: run('Check', () => appCheck_(ss, W, A, X, I, closes, now, ledger), null),
    issues
  };
}

/** The log tab without creating it: the app never writes to the sheet. */
function appReadLog_(ss) {
  const sh = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().map(r => {
    try { return dropExcluded_({ ts: new Date(r[0]).getTime(), slot: str_(r[1]), nw: Number(r[2]), assets: Number(r[3]), h: JSON.parse(r[5] || '{}'), cls: JSON.parse(r[6] || '{}'), mkt: JSON.parse(r[7] || '{}') }); } catch (e) { return null; }
  }).filter(Boolean).sort((a, b) => a.ts - b.ts);
}
/** One close a day, from the midnight closes; until there are two, falls back to the last row of each day. */
function appCloses_(log, now) {
  let c = closesByDay_(log);
  if (c.length < 2) {
    const byDay = {};
    log.forEach(e => { const k = dayKey_(e.ts); if (!byDay[k] || byDay[k].ts < e.ts) byDay[k] = Object.assign({ k }, e); });
    c = Object.keys(byDay).sort().map(k => byDay[k]);
  }
  return c.filter(e => e.ts <= now.getTime());
}

/* ---------- Home ---------- */

function appHome_(W, A, M, closes, ref, loan, now) {
  const { d1, d7, d30, past } = ref;
  const deltas = [['since last close', d1], ['in 7 days', d7], ['in 30 days', d30]]
    .filter(x => x[1]).map(([l, e]) => ({ l, d: W.nw - e.nw, base: e.nw, k: e.k }));
  // Trend: net worth at each daily close for the last year, ending with today's live value. The page draws 7 days, 30 days or 12 month-ends from it.
  const todayKey = dayKey_(now.getTime()), yearAgo = dayKey_(now.getTime() - 400 * 864e5);
  const series = past.filter(c => c.k >= yearAgo).map(c => ({ k: c.k, nw: Math.round(c.nw) })).concat([{ k: todayKey, nw: Math.round(W.nw), live: true }]);
  const markets = M.list.map(x => {
    const p = d1 && d1.mkt ? d1.mkt[x.label] : null;
    return { label: x.label, kind: x.kind, val: x.val, pct: p != null && p ? (x.val - p) / p : (isNaN(x.dayPct) ? null : x.dayPct) };
  });
  if (M.cpi != null) markets.push({ label: 'Inflation', kind: 'pct', val: M.cpi, pct: null, asOf: M.cpiDate ? dayKey_(M.cpiDate.getTime()) : null });
  if (M.repo != null) markets.push({ label: 'RBI repo rate', kind: 'pct', val: M.repo, pct: null, asOf: M.repoDate ? dayKey_(M.repoDate.getTime()) : null });
  // Movers: each holding's value now against its value in the last close.
  let movers = [], moveShare = null, shown = 0;
  if (d1) {
    moveShare = d1.nw ? (W.nw - d1.nw) / Math.abs(d1.nw) : 0;
    shown = (CONFIG.APP.MOVERS.find(([lim]) => Math.abs(moveShare) < lim) || [0, 10])[1];
    movers = A.filter(h => d1.h[h.key] != null).map(h => ({ id: h.key, name: h.name, cls: h.cls, d: h.val - d1.h[h.key], p: d1.h[h.key] ? (h.val - d1.h[h.key]) / d1.h[h.key] : null }))
      .filter(x => Math.abs(x.d) >= 1).sort((a, b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, shown);
  }
  // EMI account check, as in the email: shown on Home in the days before the EMI date.
  let emiAlert = null;
  const G = CONFIG.EMI_GUARD;
  if (G && loan && loan.emi) {
    const today = startOfDay_(now), left = Math.round((new Date(loan.nextDue) - today) / 864e5);
    if (left <= G.DAYS_BEFORE && loan.bank && loan.bank.bal < G.MIN) emiAlert = { left, bal: loan.bank.bal, min: G.MIN, emi: loan.emi, due: loan.nextDue, bank: loan.bank.name, critical: loan.bank.bal < loan.emi };
  }
  return { nw: W.nw, assets: W.assets, liab: W.liab, deltas, series, firstClose: closes.length ? closes[0].k : null, lastClose: d1 ? d1.k : null,
    markets, marketsAgainst: d1 ? d1.k : null, movers, moveShare, shown, emiAlert };
}

/* ---------- Export and import ---------- */

/** Menu → Export: a new Google Sheet in Drive. "Vault ·" tabs hold tidy tables of everything the app works out; the rest are copies of your own tabs, ready to import or to start from. */
function exportVault() {
  const ss = SpreadsheetApp.getActive(), now = new Date(), stamp = Utilities.formatDate(now, CONFIG.TIMEZONE, 'd MMM yyyy, HH:mm');
  const D = JSON.parse(getAppData(true)), out = SpreadsheetApp.create('Vault export · ' + stamp), blank = out.getSheets()[0];
  // 1. Tidy tables
  const T = exportTables_(D, now);
  T.forEach(t => {
    const sh = out.insertSheet('Vault · ' + t.name);
    const rows = [t.head].concat(t.rows.length ? t.rows : [t.head.map(() => '')]);
    sh.getRange(1, 1, rows.length, t.head.length).setValues(rows.map(r => r.map(cellOut_)));
    sh.getRange(1, 1, 1, t.head.length).setFontWeight('bold').setBackground('#EEF0FF');
    sh.setFrozenRows(1);
    (t.money || []).forEach(c => sh.getRange(2, c + 1, rows.length - 1, 1).setNumberFormat('#,##0'));
    (t.pct || []).forEach(c => sh.getRange(2, c + 1, rows.length - 1, 1).setNumberFormat('0.00%'));
    if (t.rows.length) sh.getRange(1, 1, rows.length, t.head.length).createFilter();
    sh.autoResizeColumns(1, t.head.length);
  });
  // 2. Your own tabs: values, formulas and formats, so the export imports straight back
  const src = ss.getSheets();
  const dst = src.map(sh => out.insertSheet(sh.getName()));
  src.forEach((sh, i) => { copySheet_(sh, dst[i]); if (sh.isSheetHidden()) dst[i].hideSheet(); });
  out.deleteSheet(blank);
  out.setActiveSheet(out.getSheetByName('Vault · Summary')); out.moveActiveSheet(1);
  return JSON.stringify({ url: out.getUrl(), name: out.getName(), tables: T.length, tabs: src.length });
}

/** The tidy tables. Dates go in as dates, money as numbers. */
function exportTables_(D, now) {
  const H = D.home || {}, T = [], add = (name, head, rows, money, pct) => T.push({ name, head, rows, money, pct });
  const C = D.check || {}, sum = a => a.reduce((x, y) => x + (Number(y) || 0), 0);
  add('Summary', ['Item', 'Value', 'Note'], [
    ['Exported', now, ''], ['Net worth', H.nw, ''], ['Assets', H.assets, ''], ['Liabilities', H.liab, '']
  ].concat((H.deltas || []).slice(0, 1).map(x => ['Change ' + x.l, x.d, x.base ? (x.d / x.base * 100).toFixed(2) + '%' : '']))
    .concat((C.ratios || []).map(r => [r.n, r.v, r.sub || ''])), [1]);
  add('Net worth log', ['Date', 'Net worth'], (H.series || []).map(x => [x.k, x.nw]), [1]);
  add('Holdings', ['Holding', 'Type', 'Asset class', 'Category', 'Custodian', 'Liquidity', 'Purpose', 'Value', 'Previous', 'Change', 'XIRR', 'Taxation', 'Note'],
    Object.keys(D.holdings || {}).map(k => { const h = D.holdings[k]; return [h.name, h.type, h.cls, h.cat, h.cust, h.liq, h.fn, h.val, h.prev, h.d, h.xirr, h.taxation, h.note]; }), [7, 8, 9], [10]);
  const I = D.income;
  if (I) add('Income', ['Source'].concat(I.months).concat(['Year so far']), I.sources.map(x => { const m = I.months.map((_, i) => (x.m || [])[i] || 0); return [x.n].concat(m).concat([sum(m)]); }), I.months.map((m, i) => i + 1).concat([I.months.length + 1]));
  (D.payslips || []).forEach(y => {
    const heads = []; y.months.forEach(m => (m.earn || []).concat(m.ded || []).forEach(e => { if (heads.indexOf(e[0]) < 0) heads.push(e[0]); }));
    const FYM = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];
    add('Payslips ' + y.fy.replace(/^FY\s*/, ''), ['Month'].concat(heads).concat(['Gross', 'Deductions', 'Net pay']), y.months.map(m => {
      const all = (m.earn || []).concat(m.ded || []), get = h => { const e = all.find(x => x[0] === h); return e ? e[1] : ''; };
      return [FYM[m.i]].concat(heads.map(get)).concat([m.gross, m.ded_, m.net]);
    }), heads.map((h, i) => i + 1).concat([heads.length + 1, heads.length + 2, heads.length + 3]));
  });
  (D.tax || []).forEach(t => add('Tax ' + t.fy.replace(/^FY\s*/, ''), ['Line', 'Amount'], (t.working || []).map(w => Array.isArray(w) ? [w[0], w[1]] : [w.n || w.l || '', w.v != null ? w.v : w.a]), [1]));
  const Sp = D.spend;
  if (Sp) add('Spend ' + String(Sp.fy || '').replace(/^FY\s*/, ''), ['Category', 'This month', 'Usual month', 'Year so far', 'Full year'], Sp.cats.map(c => [c.k, c.act, c.usual, c.ytd, c.year]), [1, 2, 3, 4]);
  const L = D.loan;
  if (L) add('Loan schedule', ['Date', 'Entry', 'Principal', 'Interest', 'Prepayment', 'Rate from', 'Rate to', 'Balance after', 'Status'],
    (L.events || []).map(e => [e.k, e.type === 'emi' ? 'EMI' : e.type === 'prepay' ? 'Prepayment' : 'Rate change', e.principal || '', e.interest || '', e.amt || '', e.from || '', e.to || '', e.bal || '', 'Paid'])
      .concat((L.projection || []).map(e => [e.k, 'EMI', e.principal, e.interest, '', '', '', e.bal, 'Projected'])), [2, 3, 4, 7], [5, 6]);
  (D.income ? D.income.sources : []).filter(x => x.lines && x.lines.length).forEach(x => add(x.n, ['Date', 'Paid by', 'Period', 'Gross', 'TDS', 'Net', 'Parts'],
    x.lines.map(l => [l.k, l.what, l.q || l.fy || '', l.gross, l.tds || 0, l.net, (l.parts || []).map(q => q.n + ' ' + Math.round(q.gross)).join(' · ')]), [3, 4, 5]));
  add('Insurance', ['Person', 'Category', 'Type', 'Cover', 'Paid by'], (C.insurance || []).map(x => [x.person, x.cat, x.type, x.cover, x.by]), [3]);
  add('Premiums', ['Person', 'Type', 'Provider', 'Premium date', 'Premium', 'Cover', 'Start', 'End', 'Status'],
    [].concat.apply([], (C.policies || []).map(p => p.lines.map(l => [p.person, p.type, p.provider, l.k, l.amt, l.cover, l.start, l.end, l.paid ? 'Paid' : 'Upcoming']))), [4, 5]);
  const Lk = C.locker;
  if (Lk) {
    add('Locker items', ['Item'], Lk.items.map(x => [x]));
    add('Locker visits', ['Date', 'By', 'Put in', 'Taken out', 'Items after'], Lk.visits.map(v => [v.d, v.by, v.added.join(', '), v.removed.join(', '), v.after.length]));
  }
  add('Definitions', ['Tab', 'What it holds'], [
    ['Summary', 'Net worth, assets and liabilities at the time of export, the change since the last midnight log, and every ratio.'],
    ['Net worth log', 'Net worth at each midnight log, oldest first. The last row is the live value at export.'],
    ['Holdings', 'One row per holding. Change is against the last midnight log. XIRR is the annualised return across every purchase and sale.'],
    ['Income', 'Each source by month of the financial year, before tax.'],
    ['Payslips', 'Each month\'s payslip: every earning and deduction, gross, total deductions and net pay.'],
    ['Tax', 'The computation of income for each financial year, line by line.'],
    ['Spend', 'Each category: this month, a usual month, the year so far and the full-year estimate.'],
    ['Loan schedule', 'Every EMI, prepayment and rate change paid so far, then the projected EMIs to the end of the loan.'],
    ['Dividends, REIT and SGB tabs', 'Every payout: when, from whom, the period, gross, TDS, net and, for REITs, each part.'],
    ['Insurance', 'Every cover by person and category, and who pays for it.'],
    ['Premiums', 'Every premium paid or due, with cover and policy dates.'],
    ['Locker items and visits', 'What the locker holds now, and every visit with what went in and out.'],
    ['Other tabs', 'Copies of your own tabs, formulas included. Import this file into a Vault sheet to restore them, or use it as a starting point.']
  ]);
  return T;
}
function cellOut_(v) { if (v == null) return ''; if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return new Date(v + 'T00:00:00'); return v; }

/** Copies one tab into another, cell for cell: formulas where there are formulas, values elsewhere, and the formats that matter. */
function copySheet_(from, to) {
  to.clear();
  const R = from.getMaxRows(), Cn = from.getMaxColumns();
  if (to.getMaxRows() < R) to.insertRowsAfter(to.getMaxRows(), R - to.getMaxRows());
  if (to.getMaxColumns() < Cn) to.insertColumnsAfter(to.getMaxColumns(), Cn - to.getMaxColumns());
  const lr = Math.max(1, from.getLastRow()), lc = Math.max(1, from.getLastColumn()), src = from.getRange(1, 1, lr, lc), dst = to.getRange(1, 1, lr, lc);
  const f = src.getFormulas(), v = src.getValues();
  dst.setValues(v.map((r, i) => r.map((x, j) => f[i][j] || x)));
  dst.setNumberFormats(src.getNumberFormats()); dst.setFontWeights(src.getFontWeights()); dst.setFontColors(src.getFontColors());
  dst.setBackgrounds(src.getBackgrounds()); dst.setHorizontalAlignments(src.getHorizontalAlignments()); dst.setWraps(src.getWraps());
  for (let c = 1; c <= lc; c++) to.setColumnWidth(c, from.getColumnWidth(c));
  to.setFrozenRows(from.getFrozenRows()); to.setFrozenColumns(from.getFrozenColumns());
}

/**
 * Menu → Import: brings tabs in from a Google Sheet (a link) or an Excel file (.xlsx).
 * Tabs with the same name are overwritten in place, so formulas elsewhere keep pointing at them; new tabs are added; nothing else is touched.
 * A full copy of this sheet is saved to Drive first. Excel needs the Drive API service turned on (Apps Script → Services → + → Drive API).
 */
function importVault(src) {
  const ss = SpreadsheetApp.getActive(), stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'd MMM yyyy, HH:mm');
  let from, temp = null;
  if (src && src.b64) {
    if (typeof Drive === 'undefined') throw new Error('To import an Excel file, turn on the Drive API: in Apps Script, click Services, then +, choose Drive API and Add.');
    const blob = Utilities.newBlob(Utilities.base64Decode(src.b64), MimeType.MICROSOFT_EXCEL, src.name || 'Vault.xlsx');
    const made = Drive.Files.create({ name: 'Vault import · ' + (src.name || stamp), mimeType: MimeType.GOOGLE_SHEETS }, blob);
    temp = made.id; from = SpreadsheetApp.openById(made.id);
  } else {
    const m = String((src && src.url) || '').match(/[-\w]{25,}/);
    if (!m) throw new Error('Paste the full link to a Google Sheet.');
    from = SpreadsheetApp.openById(m[0]);
  }
  if (from.getId() === ss.getId()) throw new Error('That link is this sheet.');
  const backup = DriveApp.getFileById(ss.getId()).makeCopy('Vault backup before import · ' + stamp);
  const tabs = from.getSheets().filter(sh => !/^Vault · /.test(sh.getName())), done = { replaced: [], added: [] };
  // Make every tab first, then fill them, so formulas that point across tabs find their targets
  const pairs = tabs.map(sh => { const name = sh.getName(), old = ss.getSheetByName(name); (old ? done.replaced : done.added).push(name); return [sh, old || ss.insertSheet(name, ss.getSheets().length)]; });
  pairs.forEach(p => { copySheet_(p[0], p[1]); if (p[0].isSheetHidden() || p[1].getName() === CONFIG.LOG_TAB) p[1].hideSheet(); });
  if (temp) DriveApp.getFileById(temp).setTrashed(true);
  return JSON.stringify({ replaced: done.replaced, added: done.added, backup: backup.getUrl() });
}

/* ---------- Income ---------- */

function appIncome_(I, L) {
  if (!I) return null;
  const P = I.pay, PO = I.payouts || { divCash: Array(12).fill(0), reitCash: Array(12).fill(0), sgbCash: Array(12).fill(0) };
  const monthly = [
    { id: 'salary', n: 'Salary', m: P.gross, src: I.sheet + ' · payslip register' },
    { id: 'div', n: 'Dividends', m: PO.divCash, src: 'Dividends & interests payouts' },
    { id: 'reit', n: 'REIT distributions', m: PO.reitCash, src: 'Dividends & interests payouts' },
    { id: 'sgb', n: 'SGB interest', m: PO.sgbCash, src: 'Dividends & interests payouts' }
  ];
  const yearly = [
    { id: 'sav', n: 'Savings interest', fy: I.inc.sav, src: I.sheet },
    { id: 'fd', n: 'FD interest', fy: I.inc.fd, src: I.sheet },
    { id: 'cg', n: 'Capital gains, realised', fy: I.inc.stcg + I.inc.ltcg, src: I.sheet },
    { id: 'bonus', n: 'Bonus outside payslips', fy: I.inc.bonus, src: I.sheet }
  ].filter(x => x.fy > 0);
  const last = P.lastIdx;
  const monthTot = MON.map((_, j) => sum_(monthly.map(s => s.m[j] || 0)));
  const ledger = L || { sgb: [], reit: [], div: [] };
  return {
    fy: I.fy, fyStart: fyStart_(new Date()), months: MON, last, monthTot, est: P.runGross, ahead: P.ahead,
    toDate: sum_(monthTot.slice(0, last + 1)) + sum_(yearly.map(x => x.fy)),
    fyEst: I.grossInc,
    sources: monthly.map(s => ({ id: s.id, n: s.n, m: s.m, src: s.src, monthly: true, lines: s.id === 'salary' ? null : (ledger[s.id] || []) }))
      .filter(s => sum_(s.m) > 0 || (s.lines && s.lines.length))
      .concat(yearly.map(s => ({ id: s.id, n: s.n, fy: s.fy, src: s.src, monthly: false, lines: [] }))),
    netLast: last >= 0 ? P.gross[last] - P.ded[last] : null
  };
}

/** Every payout line on the payouts tab, all years: SGB coupons, REIT distributions (with their parts and TDS) and stock dividends. */
function appPayoutLedger_(ss) {
  const v = findSheet_(ss, CONFIG.TABS.payouts).getDataRange().getValues();
  const H0 = v[0].map(str_), H1 = (v[1] || []).map(str_), out = { sgb: [], reit: [], div: [] };
  const fyOf = d => fyName_(d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1);
  // SGB coupons: one column per bond, dates in column A
  const qc = H0.findIndex(h => /^quarter$/i.test(h));
  for (let c = 1; c < (qc > 0 ? qc : 1); c++) {
    if (!H0[c]) continue;
    v.slice(1).forEach(r => { const a = num_(r[c]); if (r[0] instanceof Date && a > 0) out.sgb.push({ k: dayKey_(r[0].getTime()), fy: fyOf(r[0]), what: H0[c], net: a, gross: a, tds: 0 }); });
  }
  // REIT blocks: a "Per unit" row starts each distribution; each REIT has its own columns, ending in Total
  const pc = H0.findIndex(h => /^particulars$/i.test(h)), sc = H0.findIndex(h => /^stocks$/i.test(h));
  if (pc > 0) {
    const end = sc > 0 ? sc : H0.length, reits = [];
    let cur = null;
    for (let c = pc + 1; c < end; c++) {
      if (H0[c]) { cur = { name: H0[c], cols: [] }; reits.push(cur); }
      if (cur && H1[c]) cur.cols.push({ c, part: H1[c] });
    }
    let date = null, q = '', at = {};
    const flush = () => {
      if (!date || at.hold == null) return;
      reits.forEach(R => {
        const get = (row, re) => { const x = R.cols.find(k => re.test(k.part)); return x && row != null ? num_(v[row][x.c]) : NaN; };
        const total = get(at.hold, /^total$/i);
        if (isNaN(total) || !total) return;
        const parts = R.cols.filter(k => !/^total$/i.test(k.part)).map(k => ({ n: k.part, gross: num_(v[at.hold][k.c]) || 0, tds: at.tds != null ? num_(v[at.tds][k.c]) || 0 : 0, rate: at.rate != null ? num_(v[at.rate][k.c]) : NaN }))
          .filter(x => x.gross);
        const net = get(at.net, /^total$/i);
        out.reit.push({ k: dayKey_(date.getTime()), fy: fyOf(date), q, what: R.name, gross: total, tds: sum_(parts.map(x => x.tds)), net: isNaN(net) ? total - sum_(parts.map(x => x.tds)) : net, parts });
      });
    };
    for (let r = 2; r < v.length; r++) {
      const label = str_(v[r][pc]);
      if (/^per unit/i.test(label)) { flush(); date = v[r][pc - 1] instanceof Date ? v[r][pc - 1] : null; q = str_(v[r][pc - 2]); at = {}; }
      else if (/for your holding/i.test(label)) at.hold = r;
      else if (/^tds rate/i.test(label)) at.rate = r;
      else if (/^tds amount/i.test(label)) at.tds = r;
      else if (/^net distribution/i.test(label)) at.net = r;
    }
    flush();
  }
  // Stock dividends: a date column then one column per stock
  if (sc >= 0) for (let c = sc + 1; c < H1.length; c++) if (H1[c]) v.slice(2).forEach(r => { const a = num_(r[c]); if (r[sc] instanceof Date && a > 0) out.div.push({ k: dayKey_(r[sc].getTime()), fy: fyOf(r[sc]), what: H1[c], gross: a, tds: 0, net: a }); });
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.k < b.k ? 1 : -1));
  return out;
}

/* ---------- Tax ---------- */

function appTax_(ss, I, now) {
  const out = [];
  if (I) {
    const S = I.sel, P = I.pay, reg = I.regime, R = [];
    const add = (k, l, v) => R.push({ k, l, v: v == null ? null : v });
    add('h', 'Income · full year');
    add('i', 'Salary, estimated', P.grossFY);
    [['Bonus outside payslips', I.inc.bonus], ['Savings interest', I.inc.sav], ['FD interest', I.inc.fd], ['SGB interest', I.inc.sgb], ['Dividends', I.inc.div],
      ['REIT, taxable part', I.inc.reit], ['Short-term capital gains', I.inc.stcg], ['Long-term capital gains', I.inc.ltcg]].forEach(([l, v]) => { if (v) add('i', l, v); });
    add('s', 'Gross income', I.grossInc);
    add('h', 'Deductions · ' + (reg === 'old' ? 'old' : 'new') + ' regime');
    add('i', 'Standard deduction', -S.std);
    if (reg === 'old') {
      [['HRA exemption', I.dOld.hra], ['Home loan interest, s.24(b)', Math.min(I.dOld.hlInt, 200000)], ['80C, capped', Math.min(I.dOld.c80, 150000)], ['80D, self', Math.min(I.dOld.medSelf, 25000)],
        ['80D, parents', Math.min(I.dOld.medPar, 50000)], ['80TTA', Math.min(I.inc.sav, 10000)], ['80CCD(1B)', Math.min(I.dOld.nps1b, 50000)]].forEach(([l, v]) => { if (v) add('i', l, -v); });
    }
    if (I.both.nps2) add('i', 'Employer NPS, 80CCD(2)', -I.both.nps2);
    if (I.both.leave) add('i', 'Leave encashment', -I.both.leave);
    add('s', 'Total deductions', -S.dd);
    add('t', 'Income taxed at slab rates', S.slabInc);
    add('h', 'Tax');
    add('i', 'Slab tax', S.slab);
    if (S.rebate) add('i', 'Rebate, 87A', -S.rebate);
    if (S.stcgTax) add('i', 'STCG at 20%', S.stcgTax);
    if (S.ltcgTax) add('i', 'LTCG at 12.5% above ₹1,25,000', S.ltcgTax);
    add('s', 'Tax before surcharge', S.slab - S.rebate + S.special);
    if (S.sur) add('i', 'Surcharge at ' + Math.round(S.sr * 100) + '%', S.sur);
    add('i', 'Health and education cess at 4%', S.cess);
    add('t', 'Tax liability, estimated', S.tax);
    add('h', 'Paid');
    add('i', 'TDS on salary to date', -P.tdsToDate);
    if (I.tdsOther) add('i', 'TDS on other income', -I.tdsOther);
    if (I.advPaid) add('i', 'Advance tax', -I.advPaid);
    if (I.sat) add('i', 'Self-assessment tax', -I.sat);
    add('t', I.balance >= 0 ? 'Balance to pay' : 'Overpaid', Math.abs(I.balance));
    out.push({
      fy: I.fy, current: true, regime: reg, liab: S.tax, grossInc: I.grossInc, paid: I.paidSoFar, tds: P.tdsToDate, adv: I.advPaid, balance: I.balance,
      oldTax: I.old.tax, newTax: I.new.tax, netDue: I.netDue, applies: I.applies, tdsAhead: I.tdsAhead, afterTds: I.afterTds,
      sched: I.sched.map(s => ({ q: s.q, due: dayKey_(s.due.getTime()), pct: s.pct, req: s.req, paidCum: s.paidCum, gap: s.gap, past: s.past, int: s.int })),
      s424: I.s424, s424Months: I.s424Months, filing: dayKey_(I.filing.getTime()), marginal: 0.3 * (1 + S.sr) * 1.04, sheet: I.sheet, working: R
    });
  }
  // Earlier years: the CoI tabs as the sheet has them.
  const cur = fyStart_(now);
  ss.getSheets().filter(s => CONFIG.TABS.coi.test(s.getName())).map(s => ({ s, y: appFyOfName_(s.getName()) }))
    .filter(x => x.y != null && x.y !== cur).sort((a, b) => b.y - a.y)
    .forEach(({ s, y }) => { try { out.push(appPastCoi_(s, y)); } catch (e) { console.error(s.getName() + ': ' + e.message); } });
  return out;
}
function appFyOfName_(name) { const m = String(name).match(/(\d{2,4})\s*-\s*(\d{2,4})/); if (!m) return null; const a = Number(m[1]); return a < 100 ? 2000 + a : a; }

/** An earlier year's CoI tab, from its "Income heads" block down. Labels from A and B; the value is the last number in B to D. */
function appPastCoi_(sh, y) {
  const v = sh.getDataRange().getValues();
  let start = v.findIndex(r => /^income heads/i.test(str_(r[0])));
  if (start < 0) start = 0;
  const rows = [], key = {};
  let group = '';
  for (let r = start + 1; r < v.length; r++) {
    const a = str_(v[r][0]), b = v[r][1];
    if (/ca fees|rover|moon/i.test(a)) break;
    const nums = [1, 2, 3].map(c => num_(v[r][c])).filter(n => !isNaN(n));
    const bText = typeof b === 'string' && isNaN(num_(b)) ? str_(b) : '';
    if (!a && !bText && !nums.length) continue;
    if (!nums.length) { if (a && !/^section$/i.test(a)) { rows.push({ k: 'h', l: a, v: null }); group = a; } continue; }
    if (a) group = a;
    const label = (a || group) + (bText && !/^\[/.test(bText) ? ' · ' + bText : '');
    const val = nums[nums.length - 1];
    const k = /^total income tax|^balance tax/i.test(a) ? 't' : /income to tax ratio/i.test(a) ? 'r' : /^total|^gross income|^applicable/i.test(a) ? 's' : 'i';
    // Keep the computation; leave out the HRA and leave-encashment calculators and empty lines.
    const calc = /allocated hra|actual rent|hra tds|hra exemption|^delta$|tax loss|tds refund|additional tax via|number of leaves|monthly leaves|actual amount received|basic salary monthly|maximum limit|unutilized|exempted amount/i.test(a);
    if (!calc && (val !== 0 || k !== 'i')) rows.push({ k, l: label.replace(/\s*\[[^\]]*\]\s*$/, ''), v: val });
    [['gross', /^gross income/i], ['taxable', /^total taxable income/i], ['tax', /^total income tax/i], ['tds', /^tds paid/i], ['sat', /^self assessment/i], ['balance', /^balance tax/i], ['ratio', /income to tax ratio/i]]
      .forEach(([k2, re]) => { if (re.test(a) && key[k2] == null) key[k2] = val; });
  }
  return { fy: fyName_(y), current: false, liab: key.tax, grossInc: key.gross, taxable: key.taxable, tds: key.tds, sat: key.sat || 0, balance: key.balance, rate: key.ratio, sheet: sh.getName(), working: rows };
}

/* ---------- Payslips ---------- */

/** Every CoI tab's payslip register: earnings, deductions, gross, total deductions and net for each month with pay. */
function appPayslips_(ss, now) {
  return ss.getSheets().filter(s => CONFIG.TABS.coi.test(s.getName())).map(s => ({ s, y: appFyOfName_(s.getName()) })).filter(x => x.y != null)
    .sort((a, b) => b.y - a.y).map(({ s, y }) => {
      const v = s.getDataRange().getValues(), hdr = monthHeader_(v, v.length);
      if (!hdr) return null;
      const lc = hdr.cols[0].c - 1, rows = []; let sec = null;
      for (let r = hdr.r + 1; r < v.length; r++) {
        const label = str_(v[r][lc]); if (!label) continue;
        if (/^net\b/i.test(label)) break;
        if (/^(income|earnings)$/i.test(label)) { sec = 'e'; continue; }
        if (/^deductions$/i.test(label)) { sec = 'd'; continue; }
        if (!sec || /gross|total/i.test(label)) continue;
        const m = Array(12).fill(0); hdr.cols.forEach(({ c, m: mi }) => { const n = num_(v[r][c]); if (!isNaN(n)) m[mi] += n; });
        rows.push({ sec, item: label, m });
      }
      const months = MON.map((_, i) => {
        const earn = rows.filter(r => r.sec === 'e' && r.m[i]).map(r => [r.item, r.m[i]]), ded = rows.filter(r => r.sec === 'd' && r.m[i]).map(r => [r.item, r.m[i]]);
        const gross = sum_(earn.map(x => x[1])), dd = sum_(ded.map(x => x[1]));
        return { i, earn, ded, gross, ded_: dd, net: gross - dd, tds: sum_(ded.filter(x => /income tax|\btds\b|supplementary/i.test(x[0])).map(x => x[1])) };
      }).filter(m => m.gross > 0);
      return { fy: fyName_(y), y, sheet: s.getName(), months };
    }).filter(x => x && x.months.length);
}

/* ---------- Spend ---------- */

function appSpend_(X, loan, rewards, ss, now) {
  if (!X) return { loan, rewards };
  const mi = X.mi;
  // The year's figures as the tab has them: Total (all 12 months, planned fixed costs included) and Monthly Avg on the Total row.
  const monthTotAll = MON.map((_, j) => X.monthTot[j] || 0);
  const ytdPlanned = !isNaN(X.sheetTotal) ? X.sheetTotal : sum_(monthTotAll);
  const monthAvg = !isNaN(X.sheetAvg) ? X.sheetAvg : ytdPlanned / 12;
  // Estimate for this month: rows already filled in for later months are fixed and count in full; the rest scale with the share of the month gone.
  const fixed = r => r.filled.some((f, j) => j > mi && f && r.m[j] > 0);
  const fixedNow = sum_(X.rows.filter(fixed).map(r => r.m[mi])), varNow = sum_(X.rows.filter(r => !fixed(r)).map(r => r.m[mi]));
  const estMonth = X.pace > 0 ? fixedNow + varNow / X.pace : X.thisMonth;
  const cats = {};
  X.rows.forEach(r => { (cats[r.cat] = cats[r.cat] || []).push(r); });
  const prior = MON.map((_, j) => j).filter(j => j < mi && X.monthTot[j] > 0);
  const catList = Object.keys(cats).map(k => {
    const rs = cats[k], tot = j => sum_(rs.map(r => r.m[j]));
    const usual = prior.length ? sum_(prior.map(tot)) / prior.length : 0;
    return { k, act: tot(mi), year: sum_(MON.map((_, j) => tot(j))), ytd: sum_(MON.map((_, j) => j).filter(j => j <= mi).map(tot)), usual,
      trend: MON.map((_, j) => j).filter(j => j <= mi).slice(-6).map(j => ({ m: MON[j], v: tot(j) })),
      items: rs.map(r => ({ n: r.item, v: r.m[mi], year: sum_(r.m), usual: prior.length ? sum_(prior.map(j => r.m[j])) / prior.length : 0, fixed: fixed(r) })).filter(x => x.v || x.year) };
  }).filter(c => c.act || c.year);
  return {
    fy: X.fy, tab: X.tab, month: MON[mi], mi, day: X.day, dim: X.dim, pace: X.pace, thisMonth: X.thisMonth, estMonth, fixedNow, varNow,
    monthAvg, avgCell: X.avgCell, ytdPlanned, totalCell: X.totalCell, ytd: X.ytd, monthTot: monthTotAll,
    cats: catList, loan, rewards
  };
}

/** Rewards ledger tab (Provider, date, points, ₹ per point, expiry) plus voucher and wallet balances on the holdings tab. */
function appRewards_(ss, W, now) {
  const out = { programmes: [], balances: [], expiring: [], diag: '' };
  // The ledger tab: named like "Rewards ledger", else any tab with a Provider column next to a points column.
  const isHdr = r => r.some(c => /^\s*provider\s*$/i.test(str_(c))) && r.some(c => /points?/i.test(str_(c)));
  const tabs = ss.getSheets(), named = tabs.filter(s => CONFIG.TABS.ledger.test(s.getName()) || /reward/i.test(s.getName()));
  let sh = null, v = null, hr = -1;
  for (const t of named.concat(tabs.filter(s => named.indexOf(s) < 0))) {
    if (t.getName() === CONFIG.LOG_TAB) continue;
    const vv = t.getDataRange().getValues(), h = vv.slice(0, 40).findIndex(isHdr);
    if (h >= 0) { sh = t; v = vv; hr = h; break; }
  }
  const today = startOfDay_(now);
  const asDate = x => x instanceof Date ? x : (str_(x) ? (parseLooseDate_(x) || (isNaN(Date.parse(str_(x))) ? null : new Date(Date.parse(str_(x))))) : null);
  if (!sh) out.diag = 'No tab with a Provider and a points column. Tabs: ' + tabs.map(t => t.getName()).join(', ');
  if (sh) {
    if (hr >= 0) {
      const H = v[hr].map(str_), col = re => H.findIndex(h => re.test(h));
      const K = { p: col(/^\s*provider\s*$/i), d: col(/date/i), pts: col(/points?/i), rate: col(/multiplier|per point|rate|value/i), amt: col(/^\s*amount\s*$/i), exp: col(/expir/i) };
      const byP = {};
      out.tab = sh.getName();
      v.slice(hr + 1).forEach((r, i) => {
        const p = str_(r[K.p]), d = asDate(r[K.d]), pts = num_(r[K.pts]), e = K.exp >= 0 ? asDate(r[K.exp]) : null;
        if (!p || !d || isNaN(pts)) return;
        (byP[p] = byP[p] || []).push({ row: hr + 2 + i, d, pts, rate: num_(r[K.rate]), amt: K.amt >= 0 ? num_(r[K.amt]) : NaN, exp: e || addYears_(d, CONFIG.REWARD_EXPIRY_YEARS), expFromSheet: !!e });
      });
      Object.keys(byP).forEach(p => {
        const rows = byP[p].sort((a, b) => a.d - b.d || a.row - b.row), lots = [], lines = [];
        rows.forEach(x => {
          const line = { k: dayKey_(x.d.getTime()), pts: x.pts, rate: isNaN(x.rate) ? null : x.rate, amt: !isNaN(x.amt) ? x.amt : (isNaN(x.rate) ? null : x.pts * x.rate), exp: x.pts > 0 ? dayKey_(x.exp.getTime()) : null, row: x.row };
          lines.push(line);
          if (x.pts > 0) { lots.push({ left: x.pts, rate: x.rate, exp: x.exp, earned: x.d, line }); return; }
          let use = -x.pts; // redemptions use up the oldest points first
          lots.forEach(l => { const t = Math.min(l.left, use); l.left -= t; use -= t; });
        });
        // Available = every point earned minus every point redeemed, leaving out points already past their expiry date.
        const lapsed = lots.filter(l => l.exp < today && l.left > 0);
        const pts = sum_(rows.map(x => x.pts)) - sum_(lapsed.map(l => l.left));
        const rate = rows.map(x => x.rate).filter(r => !isNaN(r)).pop();
        const val = pts * (rate || 0);
        // Upcoming expiries as the ledger's own Expiry column has them
        const ahead = rows.filter(x => x.pts > 0 && x.exp >= today).sort((a, b) => a.exp - b.exp);
        const next = ahead[0];
        out.programmes.push({ n: p, pts, rate: rate == null ? null : rate, val, nextExp: next ? dayKey_(next.exp.getTime()) : null, nextPts: next ? next.pts : null,
          earned: sum_(rows.filter(x => x.pts > 0).map(x => x.pts)), redeemed: -sum_(rows.filter(x => x.pts < 0).map(x => x.pts)), lapsed: sum_(lapsed.map(l => l.left)), lines: lines.reverse() });
        ahead.forEach(x => out.expiring.push({ n: p, exp: dayKey_(x.exp.getTime()), month: Utilities.formatDate(x.exp, CONFIG.TIMEZONE, 'MMM yyyy'), pts: x.pts, val: x.pts * (isNaN(x.rate) ? 0 : x.rate) }));
      });
    }
  }
  // Vouchers and wallets on the holdings tab; point programmes come from the ledger instead.
  W.holdings.filter(h => h.g === 'Asset' && isReward_(h) && rewardPts_(h).pts == null).forEach(h => {
    const exp = noteDate_(h);
    out.balances.push({ n: h.cust, sub: [h.veh, h.inst].filter(Boolean).join(' · '), val: h.val, exp: exp ? dayKey_(exp.getTime()) : null, lapsed: exp ? exp < today : false });
    if (exp && exp >= today) out.expiring.push({ n: h.cust, exp: dayKey_(exp.getTime()), month: Utilities.formatDate(exp, CONFIG.TIMEZONE, 'MMM yyyy'), pts: null, val: h.val });
  });
  const months = [];
  out.expiring.sort((a, b) => a.exp < b.exp ? -1 : 1).forEach(e => { if (months.indexOf(e.month) < 0) months.push(e.month); });
  out.months = months.slice(0, 3).map(m => { const es = out.expiring.filter(e => e.month === m); return { m, val: sum_(es.map(e => e.val)), items: es }; });
  out.total = sum_(out.programmes.map(p => p.val)) + sum_(out.balances.filter(b => !b.lapsed).map(b => b.val));
  delete out.expiring;
  // Tie each programme to its points row on the holdings tab, matched on the provider's name (e.g. "Example Bank" → the Example Bank "Rewards" row).
  // That row's balance is the truth. Working back from it: the oldest points earned are taken as used first,
  // and the newest earnings that add up to the balance are the points still available.
  const ptsRows = W.holdings.filter(h => h.g === 'Asset' && isReward_(h) && rewardPts_(h).pts != null);
  out.programmes.forEach(pg => {
    const t = normTokens_(pg.n).join(' ');
    const h = ptsRows.find(x => { const c = normTokens_(x.cust); return c.join(' ').indexOf(t) === 0 || acronym_(x.cust).indexOf(t) >= 0 || (c[0] && c[0] === normTokens_(pg.n)[0]); });
    if (!h) return;
    const rp = rewardPts_(h);
    pg.holdingId = h.key; pg.pts = rp.pts; pg.val = h.val; pg.rate = rp.rate != null ? rp.rate : pg.rate; pg.fromHolding = true;
    const earned = pg.lines.filter(l => l.pts > 0).sort((a, b) => a.k < b.k ? 1 : a.k > b.k ? -1 : b.row - a.row); // newest first
    let need = rp.pts;
    earned.forEach(l => { const keep = Math.max(0, Math.min(l.pts, need)); need -= keep; l.left = keep; l.status = keep <= 0 ? 'Used' : keep < l.pts ? 'Partly used' : 'Available'; });
    pg.lines.filter(l => l.pts < 0).forEach(l => { l.status = 'Redeemed'; });
    pg.used = sum_(earned.map(l => l.pts - l.left));
    const live = earned.filter(l => l.left > 0 && l.exp && l.exp >= dayKey_(today.getTime())).sort((a, b) => a.exp < b.exp ? -1 : 1);
    pg.nextExp = live.length ? live[0].exp : null; pg.nextPts = live.length ? live[0].left : null;
    pg.expiring = live.map(l => ({ n: pg.n, exp: l.exp, pts: l.left, val: l.left * (pg.rate || 0) }));
  });
  // Rebuild the expiry months from what is actually left
  const exp2 = [];
  out.programmes.forEach(pg => { (pg.expiring || []).forEach(e => exp2.push(Object.assign({ month: Utilities.formatDate(new Date(e.exp + 'T00:00:00'), CONFIG.TIMEZONE, 'MMM yyyy') }, e))); delete pg.expiring; });
  out.balances.filter(b => !b.lapsed && b.exp).forEach(b => exp2.push({ n: b.n, exp: b.exp, month: Utilities.formatDate(new Date(b.exp + 'T00:00:00'), CONFIG.TIMEZONE, 'MMM yyyy'), pts: null, val: b.val }));
  if (out.programmes.some(pg => pg.fromHolding)) {
    const ms = [];
    exp2.sort((a, b) => a.exp < b.exp ? -1 : 1).forEach(e => { if (ms.indexOf(e.month) < 0) ms.push(e.month); });
    out.months = ms.slice(0, 3).map(m => { const es = exp2.filter(e => e.month === m); return { m, val: sum_(es.map(e => e.val)), items: es }; });
    out.total = sum_(out.programmes.map(p => p.val)) + sum_(out.balances.filter(b => !b.lapsed).map(b => b.val));
  }
  if (sh) out.diag = sh.getName() + ' · ' + sum_(out.programmes.map(p => p.lines.length)) + ' lines · ' + out.programmes.length + ' programmes';
  return out;
}

/* ---------- Home loan ---------- */

/** Loan schedule tab: the details block at the top, then the bank statement (date, details, debit, credit, balance). */
function appLoan_(ss, W, now) {
  const sh = ss.getSheets().find(s => CONFIG.TABS.loan.test(s.getName()));
  if (!sh) return null;
  const v = sh.getDataRange().getValues();
  const hr = v.findIndex(r => /post date/i.test(str_(r[0])));
  const meta = {};
  v.slice(0, hr < 0 ? v.length : hr).forEach(r => { const k = str_(r[0]); if (k && r[1] !== '' && r[1] != null) meta[k.toLowerCase()] = r[1]; });
  const get = re => { const k = Object.keys(meta).find(x => re.test(x)); return k ? meta[k] : null; };
  const lh = W.holdings.find(h => h.g === 'Liability' && /home loan|housing/i.test(h.name + ' ' + h.note)) || W.holdings.find(h => h.g === 'Liability');
  const rate = num_(get(/current roi|rate/)), emi = num_(get(/^emi amount/)), sanctioned = num_(get(/sanctioned amount/)) || num_(get(/disbursed/));
  const out = { bank: null, provider: str_(get(/provisioning bank/)), sanctioned, disbursed: num_(get(/disbursed/)), tenure: num_(get(/^tenure/)), remainingSheet: num_(get(/remaining tenure/)),
    firstEmi: get(/first emi/) instanceof Date ? dayKey_(get(/first emi/).getTime()) : null, rate: isNaN(rate) ? (lh ? lh.xirr : null) : rate, emi: isNaN(emi) ? null : emi,
    emiDay: CONFIG.EMI_GUARD ? CONFIG.EMI_GUARD.EMI_DAY : 5, debitFrom: str_(get(/nach debit from/)), outstanding: lh ? lh.val : null, name: lh ? lh.name : 'Home loan' };
  // Statement → monthly EMIs, prepayments and rate changes
  const lines = [];
  for (let r = hr + 1; hr >= 0 && r < v.length; r++) {
    const d = v[r][0] instanceof Date ? v[r][0] : v[r][1];
    if (!(d instanceof Date)) continue;
    lines.push({ d, k: dayKey_(d.getTime()), t: str_(v[r][2]).toUpperCase(), dr: num_(v[r][3]) || 0, cr: num_(v[r][4]) || 0, bal: num_(v[r][5]) });
  }
  const byDay = {};
  lines.forEach(l => (byDay[l.k] = byDay[l.k] || []).push(l));
  const events = [];
  Object.keys(byDay).sort().forEach(k => {
    const ls = byDay[k], d = ls[0].d;
    const prin = ls.filter(l => /PRINCIPALREPAYMENT/.test(l.t)), ach = prin.filter(l => /ACH/.test(l.t)), pre = prin.filter(l => !/ACH/.test(l.t) && l.cr > 0);
    const intr = sum_(ls.filter(l => /^INTERESTREPAYMENT/.test(l.t)).map(l => l.cr));
    const bals = ls.filter(l => /REPAYMENT/.test(l.t) && !isNaN(l.bal)).map(l => l.bal);
    const bal = bals.length ? Math.min.apply(null, bals) : null;
    ls.forEach(l => { const m = l.t.match(/RATECHANGEDFM\s*([\d.]+)%\s*TO\s*([\d.]+)%/); if (m) events.push({ k, type: 'rate', from: Number(m[1]) / 100, to: Number(m[2]) / 100 }); });
    if (ach.length) events.push({ k, type: 'emi', principal: sum_(ach.map(l => l.cr)), interest: intr, bal });
    if (pre.length) events.push({ k, type: 'prepay', amt: sum_(pre.map(l => l.cr)), bal });
  });
  out.events = events;
  out.prepaid = sum_(events.filter(e => e.type === 'prepay').map(e => e.amt));
  out.interestPaid = sum_(lines.filter(l => /^INTEREST$|^PARTPERIODINTEREST$/.test(l.t)).map(l => l.dr));
  const rc = events.filter(e => e.type === 'rate'); out.lastRateChange = rc.length ? rc[rc.length - 1] : null;
  // Projection at the current rate and EMI
  out.projection = []; out.monthsLeft = null; out.interestLeft = null;
  if (out.outstanding && out.rate && out.emi) {
    const r = out.rate / 12;
    let bal = out.outstanding, n = 0, intTot = 0;
    const t = startOfDay_(now);
    let due = new Date(t.getFullYear(), t.getMonth() + (t.getDate() > out.emiDay ? 1 : 0), out.emiDay);
    out.nextDue = dayKey_(due.getTime());
    while (bal > 0.5 && n < 600) {
      const i = bal * r, p = Math.min(bal, out.emi - i);
      if (p <= 0) { n = Infinity; break; }
      bal -= p; intTot += i; n++;
      if (out.projection.length < 12) out.projection.push({ k: dayKey_(due.getTime()), principal: p, interest: i, bal });
      due = new Date(due.getFullYear(), due.getMonth() + 1, out.emiDay);
    }
    out.monthsLeft = n; out.interestLeft = isFinite(n) ? intTot : null;
    if (isFinite(n)) { const end = new Date(t.getFullYear(), t.getMonth() + n, out.emiDay); out.ends = Utilities.formatDate(end, CONFIG.TIMEZONE, 'MMM yyyy'); }
  }
  // The account the EMI is debited from
  const G = CONFIG.EMI_GUARD;
  const acct = W.holdings.filter(h => h.g === 'Asset' && (out.debitFrom ? normTokens_(h.cust).join(' ').indexOf(normTokens_(out.debitFrom).join(' ')) === 0 : G.ACCOUNT.test(h.cust)) && G.VEHICLE.test(h.veh + ' ' + h.inst));
  if (acct.length) out.bank = { name: acct[0].cust + ' ' + acct[0].veh.toLowerCase(), bal: sum_(acct.map(h => h.val)), min: G.MIN, daysBefore: G.DAYS_BEFORE };
  return out;
}

/* ---------- Invest and holdings ---------- */

const appType_ = h => /epf/i.test(h.veh) ? 'EPF' : /ppf/i.test(h.veh) ? 'PPF' : /nps/i.test(h.veh) ? 'NPS'
  : /commodit|gold|silver/i.test(h.cls) ? 'Gold and silver' : /equit/i.test(h.cls) ? 'Equity' : /real estate/i.test(h.cls) ? 'Real estate'
  : /fixed income|debt/i.test(h.cls) ? 'Debt' : /cash/i.test(h.cls) ? 'Cash' : h.cls;

function appInvest_(ss, W, A, ref) {
  const d1 = ref.d1, dOf = h => d1 && d1.h[h.key] != null ? h.val - d1.h[h.key] : null;
  const pct = list => list.map(x => ({ k: x.k, v: x.v, p: W.assets ? x.v / W.assets : 0 }));
  const alloc = { cls: pct(groupSum_(A, h => h.cls)), cust: pct(groupSum_(A, h => h.cust || 'Unspecified')), liq: pct(groupSum_(A, h => h.liq)), fn: pct(groupSum_(A, h => h.fn)) };
  const ORDER = ['Equity', 'EPF', 'PPF', 'NPS', 'Gold and silver', 'Debt', 'Real estate', 'Cash'];
  const types = {};
  A.forEach(h => { const t = appType_(h); (types[t] = types[t] || []).push(h); });
  // Holdings left out of net worth are listed with their type, tagged, and kept out of every total.
  (W.excluded || []).forEach(x => { const t = appType_(x); (types[t] = types[t] || []).push(Object.assign({ excl: true, xirr: NaN }, x)); });
  const list = Object.keys(types).sort((a, b) => (ORDER.indexOf(a) + 1 || 99) - (ORDER.indexOf(b) + 1 || 99)).map(t => {
    const all = types[t].sort((a, b) => (b.val || 0) - (a.val || 0)), hs = all.filter(h => !h.excl), ds = hs.map(dOf).filter(x => x != null);
    return { t, val: sum_(hs.map(h => h.val)), d: ds.length ? sum_(ds) : null, n: hs.length, nx: all.length - hs.length,
      holdings: all.map(h => ({ id: h.key, name: h.name, cust: h.cust, veh: h.veh, val: h.val, d: h.excl ? null : dOf(h), xirr: isNaN(h.xirr) ? null : h.xirr, excl: !!h.excl })) };
  });
  let port = null;
  try {
    const v = findSheet_(ss, CONFIG.TABS.portfolio).getDataRange().getValues();
    const r = v.find(row => row.some(c => /^total portfolio$/i.test(str_(c))));
    if (r) { const i = r.findIndex(c => /^total portfolio$/i.test(str_(c))); port = { invested: num_(r[i + 1]), value: num_(r[i + 2]), gain: num_(r[i + 3]), ret: num_(r[i + 4]) }; }
  } catch (e) { console.error('Portfolio totals: ' + e.message); }
  const assetsD = d1 ? sum_(A.map(dOf).filter(x => x != null)) : null;
  return { total: W.assets, d: assetsD, port, alloc, types: list, fd: W.fd };
}

/** Every asset holding's detail: the holdings row, its Portfolio row (cost, units, price) and its XIRR lots. */
function appHoldings_(ss, W, A, I, ref, now) {
  const portV = (() => { try { return findSheet_(ss, CONFIG.TABS.portfolio).getDataRange().getValues(); } catch (e) { return null; } })();
  const xirrSh = ss.getSheets().find(s => CONFIG.TABS.xirr.test(s.getName()));
  const xV = xirrSh ? xirrSh.getDataRange().getValues() : null, xF = xirrSh ? xirrSh.getDataRange().getFormulas() : null;
  const fd = W.fd || 0.0625;
  const sr = I ? I.sel.sr : 0.1, marginal = 0.3 * (1 + sr) * 1.04;
  const ltcgUsed = I ? I.inc.ltcg : 0, exemptLeft = Math.max(0, 125000 - ltcgUsed);
  const d1 = ref.d1, out = {};
  A.forEach(h => {
    const o = { id: h.key, name: h.name, cls: h.cls, cat: h.cat, cust: h.cust, veh: h.veh, liq: h.liq, fn: h.fn, note: h.note, taxation: h.taxation,
      val: h.val, prev: h.prev, d: d1 && d1.h[h.key] != null ? h.val - d1.h[h.key] : null, xirr: isNaN(h.xirr) ? null : h.xirr, beatsFd: h.beatsFd, beatsCpi: h.beatsCpi, type: appType_(h) };
    // Portfolio row from the Amount: T formula, e.g. =Portfolio!H5 or =Portfolio!N27
    const pm = h.formula.match(/portfolio'?!\$?([A-Z]+)\$?(\d+)/i);
    if (pm && portV && !/^D$/i.test(pm[1])) {
      const row = portV[Number(pm[2]) - 1] || [], col = pm[1].toUpperCase();
      if (col === 'N') o.cost = { units: num_(row[3]), unit: 'g', buy: num_(row[4]), price: num_(row[12]), invested: num_(row[11]) };
      else if (col === 'H') o.cost = { units: num_(row[3]), unit: 'units', buy: num_(row[4]), price: num_(row[5]), invested: num_(row[6]) };
      if (o.cost) Object.keys(o.cost).forEach(k => { if (typeof o.cost[k] === 'number' && isNaN(o.cost[k])) o.cost[k] = null; });
    }
    // Lots from the XIRR formula, e.g. =XIRR!H201: the rows above that "Today" row, back to the start of its block.
    const xm = h.xirrFormula.match(/xirr'?!\$?H\$?(\d+)/i);
    if (xm && xV) {
      const end = Number(xm[1]) - 1, sm = xF && xF[end] ? String(xF[end][4]).match(/E(\d+):E(\d+)/i) : null, lots = [];
      let from = sm ? Number(sm[1]) - 1 : 0;
      if (!sm) { for (let r = end - 1; r >= 0; r--) { const c = str_(xV[r][2]); if (/today/i.test(c) || (!(xV[r][1] instanceof Date) && str_(xV[r][0]) && !c)) { from = r + 1; break; } } }
      for (let r = from; r < end; r++) {
        const row = xV[r] || [];
        if (!(row[1] instanceof Date) || !/^(buy|sell|deposit|interest|withdrawal)$/i.test(str_(row[2]))) continue;
        lots.push({ d: row[1], side: str_(row[2]).toLowerCase(), qty: num_(row[4]), price: num_(row[5]), amt: num_(row[6]), what: str_(row[3]), row: r + 1 });
      }
      const today = xV[end] || [];
      o.today = { units: num_(today[4]), price: num_(today[5]), value: num_(today[6]) };
      if (lots.length) appLotMaths_(o, lots, fd, marginal, sr, exemptLeft, now);
    }
    out[h.key] = o;
  });
  (W.excluded || []).forEach(x => { out[x.key] = Object.assign({ id: x.key, excl: true, type: appType_(x), xirr: null }, x); });
  return out;
}

/** How a holding's gains are taxed, from its Taxation column and class. */
function appTaxRule_(o, marginal, sr) {
  const tx = o.taxation, mult = (1 + Math.min(sr, 0.15)) * 1.04;
  if (typeof tx === 'number' || /^\d/.test(str_(tx))) return { kind: 'equity', ltDays: 365, st: 0.2 * mult, lt: 0.125 * mult, exempt: true, note: 'Listed equity rules: long term after 12 months. Short term at 20%, long term at 12.5% above the ₹1,25,000 a year exempt, plus surcharge and cess.' };
  if (/tax\s*free/i.test(str_(tx))) return { kind: 'free', note: 'Marked tax free in the sheet.' };
  if (/commodit|equit/i.test(o.cls)) return { kind: 'slab-lt', ltDays: 730, st: marginal, lt: 0.125 * mult, note: 'Long term after 24 months at 12.5% plus surcharge and cess. Short term at your slab rate, ' + (marginal * 100).toFixed(2) + '%.' };
  return { kind: 'slab', st: marginal, note: 'All gains at your slab rate, ' + (marginal * 100).toFixed(2) + '%, however long it is held.' };
}

/** FD comparison, the lot ledger and exit tax. Sells use up the oldest units first. */
function appLotMaths_(o, lots, fd, marginal, sr, exemptLeft, now) {
  const yrs = d => (now - d) / (365.25 * 864e5), days = d => Math.floor((startOfDay_(now) - startOfDay_(d)) / 864e5);
  const cash = lots.filter(l => !isNaN(l.amt) && l.side !== 'interest');
  const invested = sum_(cash.map(l => -l.amt));
  o.lots = { n: lots.filter(l => /buy|deposit/.test(l.side)).length, first: dayKey_(lots[0].d.getTime()), invested };
  const byYear = {};
  let cum = 0;
  cash.forEach(l => { cum += -l.amt; byYear[l.d.getFullYear()] = cum; });
  o.investedByYear = Object.keys(byYear).map(y => ({ y: Number(y), v: byYear[y] }));
  const after = fd * (1 - marginal);
  o.fdCompare = { rate: fd, marginal, fdValue: sum_(cash.map(l => -l.amt * Math.pow(1 + fd, yrs(l.d)))), fdAfterTax: sum_(cash.map(l => -l.amt * Math.pow(1 + after, yrs(l.d)))), invested };
  const rule = appTaxRule_(o, marginal, sr);
  const unitBased = lots.some(l => l.side === 'buy' && !isNaN(l.qty) && l.qty > 0);
  const ledger = [];
  if (!unitBased) {
    // Deposits and interest (PPF and the like): no units, so the ledger lists the money in and the interest credited.
    lots.forEach(l => ledger.push({ k: dayKey_(l.d.getTime()), side: l.side === 'interest' ? 'Interest' : l.side === 'deposit' ? 'Deposit' : l.side.charAt(0).toUpperCase() + l.side.slice(1),
      amt: l.side === 'interest' ? (isNaN(l.price) ? null : l.price) : (isNaN(l.amt) ? null : Math.abs(l.amt)), days: days(l.d) }));
    const interest = sum_(ledger.filter(x => x.side === 'Interest').map(x => x.amt || 0));
    o.ledger = { kind: 'cash', lines: ledger.reverse(), interest, deposits: invested, rule: rule.kind, ruleNote: rule.note };
    o.exitTax = { st: 0, lt: 0, gross: o.val - invested, tax: rule.kind === 'free' ? 0 : Math.max(0, o.val - invested) * (rule.st || 0), net: 0, lines: [], basis: str_(o.taxation), note: rule.note };
    o.exitTax.net = o.exitTax.gross - o.exitTax.tax;
    return;
  }
  const buys = lots.filter(l => l.side === 'buy' && l.qty > 0).map(l => ({ src: l, d: l.d, qty: l.qty, qty0: l.qty, cost: !isNaN(l.amt) && l.amt ? Math.abs(l.amt) : l.qty * l.price }));
  lots.filter(l => l.side === 'sell' || (l.side === 'buy' && l.qty < 0)).forEach(s => { let q = Math.abs(s.qty); buys.forEach(b => { if (b.d > s.d || q <= 0) return; const t = Math.min(b.qty, q); b.qty -= t; q -= t; }); });
  const units = sum_(buys.map(b => Math.max(0, b.qty)));
  const price = units && o.val ? o.val / units : (o.today && !isNaN(o.today.price) ? o.today.price : null);
  // Each buy lot: what is left, its gain today and its holding period
  const L = buys.map(b => {
    const left = Math.max(0, b.qty), costLeft = b.cost * left / b.qty0, value = price != null ? left * price : null, dd = days(b.d);
    const term = rule.kind === 'free' ? 'Tax free' : rule.kind === 'slab' ? 'Slab' : dd > rule.ltDays ? 'LT' : 'ST';
    return { k: dayKey_(b.d.getTime()), side: 'Buy', qty: b.qty0, unitCost: b.cost / b.qty0, cost: b.cost, left, price, value, gain: value != null && left > 1e-9 ? value - costLeft : null, days: dd, term, tax: 0 };
  });
  // Tax on the whole holding if sold today, then shared out to the lots that carry a gain
  const sumG = t => sum_(L.filter(x => x.term === t && x.gain != null).map(x => x.gain));
  const st = sumG('ST') + sumG('Slab'), lt = sumG('LT');
  let stTax = 0, ltTax = 0, lines = [];
  if (rule.kind === 'equity') {
    stTax = Math.max(0, st) * rule.st;
    ltTax = Math.max(0, lt - exemptLeft) * rule.lt;
    lines = [['Short term · 20% plus surcharge and cess', st, stTax], ['Long term · 12.5% after ₹' + Math.round(exemptLeft).toLocaleString('en-IN') + ' exempt, plus surcharge and cess', lt, ltTax]];
  } else if (rule.kind === 'slab-lt') {
    stTax = Math.max(0, st) * rule.st; ltTax = Math.max(0, lt) * rule.lt;
    lines = [['Short term · slab rate ' + (rule.st * 100).toFixed(2) + '%', st, stTax], ['Long term · 12.5% plus surcharge and cess', lt, ltTax]];
  } else if (rule.kind === 'slab') {
    stTax = Math.max(0, st) * rule.st;
    lines = [['Gain at slab rate ' + (rule.st * 100).toFixed(2) + '%', st, stTax]];
  }
  const share = (term, total) => { const pos = L.filter(x => term.indexOf(x.term) >= 0 && x.gain > 0), g = sum_(pos.map(x => x.gain)); pos.forEach(x => { x.tax = g ? total * x.gain / g : 0; }); };
  share(['ST', 'Slab'], stTax); share(['LT'], ltTax);
  const S = lots.filter(l => l.side === 'sell' || (l.side === 'buy' && l.qty < 0)).map(l => ({ k: dayKey_(l.d.getTime()), side: /billing|charge/i.test(l.what) ? 'Charge' : 'Sell', qty: Math.abs(l.qty), unitCost: l.price, cost: isNaN(l.amt) ? null : Math.abs(l.amt), days: days(l.d), what: l.what }));
  const gross = st + lt, tax = Math.max(0, stTax + ltTax);
  o.ledger = { kind: 'units', unit: /physical gold|physical silver/i.test(o.veh) ? 'g' : 'units', price, units, lines: L.concat(S).sort((a, b) => a.k < b.k ? 1 : -1), rule: rule.kind, ruleNote: rule.note, ltDays: rule.ltDays || null };
  o.exitTax = { st, lt, gross, tax, net: gross - tax, lines: lines.filter(x => x[1] || x[2]), basis: str_(o.taxation), note: rule.note };
}

/** EPF: the balance by employer and share from the Portfolio tab, and the passbook from the EPF tab. */
function appEpf_(ss) {
  const out = { split: [], passbook: null };
  try {
    const v = findSheet_(ss, CONFIG.TABS.portfolio).getDataRange().getValues();
    const r0 = v.findIndex(r => r.some(c => /^employee provident fund$/i.test(str_(c))));
    if (r0 > 0) {
      const H = v[r0 - 1].map(str_), lc = v[r0].findIndex(c => /^employee provident fund$/i.test(str_(c)));
      const ci = re => H.findIndex((h, i) => i > lc + 1 && re.test(h));
      const K = { ee: ci(/^employee$/i), er: ci(/^employer$/i), pen: ci(/^pension$/i) };
      for (let r = r0; r < v.length; r++) {
        const name = str_(v[r][lc + 1]);
        if (/^total/i.test(str_(v[r][lc]))) { out.total = num_(v[r][K.ee]); break; }
        if (!name) continue;
        out.split.push({ n: name, ee: num_(v[r][K.ee]) || 0, er: num_(v[r][K.er]) || 0, pen: num_(v[r][K.pen]) || 0 });
      }
    }
  } catch (e) { console.error('EPF split: ' + e.message); }
  const sh = ss.getSheets().find(s => /^epf$/i.test(s.getName()));
  if (sh) {
    const v = sh.getDataRange().getValues(), H = v[0].map(str_), col = re => H.findIndex(h => re.test(h));
    const K = { m: col(/for month/i), d: col(/transaction date/i), t: col(/credit|debit/i), wages: col(/^epf wages/i), ee: col(/^employee$/i), er: col(/^employer$/i), pen: col(/^pension$/i), sum: col(/^sum$/i) };
    const lines = [];
    for (let r = 1; r < v.length; r++) {
      const d = v[r][K.d], t = str_(v[r][K.t]).toUpperCase(), ee = num_(v[r][K.ee]) || 0, er = num_(v[r][K.er]) || 0, pen = num_(v[r][K.pen]) || 0, sm = num_(v[r][K.sum]);
      if (!(d instanceof Date) || (!ee && !er && !pen)) continue;
      // Credits are monthly contributions; a debit with a positive sum is the year's interest; a negative sum is a withdrawal.
      const kind = t === 'DR' ? (sm < 0 ? 'Withdrawal' : 'Interest') : 'Contribution';
      const m = v[r][K.m] instanceof Date ? v[r][K.m] : null;
      lines.push({ k: dayKey_(d.getTime()), m: m ? Utilities.formatDate(m, CONFIG.TIMEZONE, 'MMM yyyy') : null, kind, wages: num_(v[r][K.wages]) || null, ee, er, pen, row: r + 1,
        fy: fyName_(d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1) });
    }
    const tot = (kind, k) => sum_(lines.filter(l => l.kind === kind).map(l => l[k]));
    const fys = {};
    lines.forEach(l => { const f = fys[l.fy] = fys[l.fy] || { fy: l.fy, ee: 0, er: 0, pen: 0, intEe: 0, intEr: 0, wd: 0, n: 0 };
      if (l.kind === 'Contribution') { f.ee += l.ee; f.er += l.er; f.pen += l.pen; f.n++; } else if (l.kind === 'Interest') { f.intEe += l.ee; f.intEr += l.er; } else f.wd += l.ee + l.er; });
    out.passbook = { tab: sh.getName(), ee: tot('Contribution', 'ee'), er: tot('Contribution', 'er'), pen: tot('Contribution', 'pen'), intEe: tot('Interest', 'ee'), intEr: tot('Interest', 'er'),
      wdEe: tot('Withdrawal', 'ee'), wdEr: tot('Withdrawal', 'er'), months: lines.filter(l => l.kind === 'Contribution').length,
      fys: Object.keys(fys).sort().reverse().map(k => fys[k]), lines: lines.reverse() };
  }
  return out;
}

/* ---------- Check ---------- */

// Gauge for each ratio: [scale min, scale max, healthy from, healthy to, borderline from, borderline to]
const RATIO_BANDS_ = { sr: [0, 1, .2, 1, .1, .2], ef: [0, 12, 6, 12, 3, 6], lq: [0, 1, .15, 1, null, null], ds: [0, .6, 0, .3, null, null], hc: [0, .6, 0, .28, null, null], da: [0, 1, 0, .3, .3, .5], tr: [0, .45, null, null, null, null] };

/**
 * Insurance premiums tab: one or more side-by-side tables, each headed Person, Insurance type, Provider, Premium amount,
 * Premium Date, Coverage amount, Start date, End date, Status. Each person + type + provider is one policy, with a line per year.
 */
function appInsPolicies_(ss, now) {
  const sh = ss.getSheets().find(s => CONFIG.APP.INSURANCE_LOG.test(s.getName()));
  if (!sh) return null;
  const v = sh.getDataRange().getValues(), today = startOfDay_(now);
  // Dates typed as plain numbers (sheet serials) are read as dates too
  const asDate = x => x instanceof Date ? x : (typeof x === 'number' && x > 20000 && x < 80000 ? new Date(Math.round((x - 25569) * 864e5)) : parseLooseDate_(x));
  const blocks = [];
  for (let r = 0; r < Math.min(v.length, 15); r++) for (let c = 0; c < v[r].length; c++) {
    if (!/^person$/i.test(str_(v[r][c])) || !/insurance type|type/i.test(str_(v[r][c + 1]))) continue;
    const H = {}, end = (() => { let k = c + 1; while (k < v[r].length && !/^person$/i.test(str_(v[r][k]))) k++; return k; })();
    for (let k = c; k < end; k++) H[str_(v[r][k]).toLowerCase()] = k;
    const col = re => { const key = Object.keys(H).find(h => re.test(h)); return key != null ? H[key] : -1; };
    blocks.push({ r, K: { p: c, t: col(/type/), pr: col(/provider|insurer/), a: col(/premium amount|^premium$|amount/), d: col(/premium date|^date/), cv: col(/cover/), s: col(/start/), e: col(/^end/), st: col(/status/) } });
  }
  const pol = {};
  blocks.forEach(({ r, K }) => {
    for (let i = r + 1; i < v.length; i++) {
      const row = v[i], person = str_(row[K.p]);
      if (!person || /^person$/i.test(person)) continue;
      const d = asDate(row[K.d]); if (!d) continue;
      const type = str_(row[K.t]), provider = str_(row[K.pr]), key = [person, type, provider].join('|');
      const amt = num_(row[K.a]), st = K.st >= 0 ? str_(row[K.st]) : '';
      const p = pol[key] = pol[key] || { key, person, type, provider, lines: [] };
      const s0 = K.s >= 0 ? asDate(row[K.s]) : null, e0 = K.e >= 0 ? asDate(row[K.e]) : null;
      p.lines.push({ k: dayKey_(d.getTime()), amt: isNaN(amt) ? null : amt, cover: K.cv >= 0 ? (num_(row[K.cv]) || null) : null,
        start: s0 ? dayKey_(s0.getTime()) : null, end: e0 ? dayKey_(e0.getTime()) : null, paid: /paid/i.test(st) || (!st && d <= today) });
    }
  });
  const todayK = dayKey_(today.getTime());
  return Object.keys(pol).map(k => {
    const p = pol[k];
    p.lines.sort((a, b) => a.k < b.k ? -1 : 1);
    const paid = p.lines.filter(l => l.paid && l.amt != null), last = paid[paid.length - 1], first = paid[0];
    const up = p.lines.find(l => !l.paid && l.k >= todayK);
    // Next renewal: the upcoming line if the tab has one, else the day after the latest cover ends
    let due = up ? up.k : null;
    if (!due && last && last.end) { const e = new Date(last.end + 'T00:00:00'); e.setDate(e.getDate() + 1); due = dayKey_(e.getTime()); }
    const yrs = first && last ? (new Date(last.k) - new Date(first.k)) / (365.25 * 864e5) : 0;
    p.last = last || null; p.first = first || null; p.due = due; p.dueAmt = up && up.amt != null ? up.amt : (last ? last.amt : null); p.dueEst = !(up && up.amt != null);
    p.total = sum_(paid.map(l => l.amt)); p.cover = last ? last.cover : null;
    p.cagr = first && last && yrs >= 0.9 && first.amt > 0 ? Math.pow(last.amt / first.amt, 1 / yrs) - 1 : null;
    return p;
  }).sort((a, b) => (a.due || '9') < (b.due || '9') ? -1 : 1);
}

/** Insurance premiums paid: the insurer's row on every Expenses tab, month by month, plus an insurance log tab if one exists. */
function appPremiums_(ss) {
  const out = [];
  ss.getSheets().filter(sh => CONFIG.TABS.expenses.test(sh.getName())).forEach(sh => {
    const v = sh.getDataRange().getValues(), hdr = monthHeader_(v, 8), y = appFyOfName_(sh.getName());
    if (!hdr || !y) return;
    const lc = hdr.cols[0].c - 1;
    for (let r = hdr.r + 1; r < v.length; r++) {
      const item = str_(v[r][lc]);
      if (!item || !/insur|medical|health|term|mediclaim|\blic\b|policy/i.test(item) || /income tax/i.test(item)) continue;
      hdr.cols.forEach(({ c, m }) => {
        const a = num_(v[r][c]);
        if (!(a > 0)) return;
        const yy = m >= 9 ? y + 1 : y, mm = (m + 3) % 12 + 1;
        out.push({ item, k: yy + '-' + String(mm).padStart(2, '0'), amt: a, fy: fyName_(y), src: 'expenses' });
      });
    }
  });
  const log = ss.getSheets().find(sh => CONFIG.APP.INSURANCE_LOG.test(sh.getName()));
  if (log) {
    const v = log.getDataRange().getValues(), hr = v.slice(0, 20).findIndex(r => r.some(c => /date/i.test(str_(c))) && r.some(c => /amount|premium/i.test(str_(c))));
    if (hr >= 0) {
      const H = v[hr].map(str_), col = re => H.findIndex(h => re.test(h));
      const K = { d: col(/date/i), a: col(/amount|premium/i), i: col(/insurer|provider|company|policy/i), n: col(/note|detail|plan|type/i) };
      v.slice(hr + 1).forEach(r => { const d = r[K.d] instanceof Date ? r[K.d] : null, a = num_(r[K.a]); if (d && a > 0) out.push({ item: [K.i >= 0 ? str_(r[K.i]) : '', K.n >= 0 ? str_(r[K.n]) : ''].filter(Boolean).join(' · '), k: dayKey_(d.getTime()), amt: a, fy: fyName_(d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1), src: 'log' }); });
    }
  }
  return out.sort((a, b) => a.k < b.k ? 1 : -1);
}

function appCheck_(ss, W, A, X, I, closes, now, R2) {
  // Personal ratios: the same as the email
  const R = [];
  if (X || I) {
    const net = I ? I.grossInc - I.sel.tax : null;
    const avgSpend = X && X.nUsed ? sum_(X.monthTot.filter((t, j) => j <= X.mi && t > 0)) / X.nUsed : null;
    const yearSpend = avgSpend != null ? avgSpend * 12 : null;
    const tier1 = sum_(A.filter(h => /tier 1/i.test(h.liq)).map(h => h.val));
    const cushion = sum_(A.filter(h => /tier 1/i.test(h.liq) || /safety/i.test(h.fn)).map(h => h.val));
    [
      ['sr', 'Savings rate', net && yearSpend != null ? (net - yearSpend) / net : null, 'pct', 'Take-home kept after spend · 20% or more', v => v >= .2 ? 'g' : v >= .1 ? 'a' : 'r', [['Take-home, full year', net], ['Spend, annualised', yearSpend]]],
      ['ef', 'Emergency cover', avgSpend ? cushion / avgSpend : null, 'months', 'Instant and safety holdings ÷ monthly spend · 3 to 6', v => v >= 6 ? 'g' : v >= 3 ? 'a' : 'r', [['Instant and safety holdings', cushion], ['Monthly spend', avgSpend]]],
      ['lq', 'Liquidity', yearSpend ? tier1 / yearSpend : null, 'x2', 'Instant-access assets ÷ yearly spend · 0.15 or more', v => v >= .15 ? 'g' : 'r', [['Instant-access assets', tier1], ['Yearly spend', yearSpend]]],
      ['ds', 'Debt service', net && X && X.emiMonthly != null ? X.emiMonthly * 12 / net : null, 'x2', 'Loan EMIs ÷ take-home · 0.30 or less', v => v <= .3 ? 'g' : 'r', [['EMIs, yearly', X && X.emiMonthly != null ? X.emiMonthly * 12 : null], ['Take-home', net]]],
      ['hc', 'Housing cost', I && X && X.housingMonthly != null ? X.housingMonthly * 12 / I.grossInc : null, 'pct', 'Household and utilities ÷ gross income · 28% or less', v => v <= .28 ? 'g' : 'r', [['Household and utilities, yearly', X && X.housingMonthly != null ? X.housingMonthly * 12 : null], ['Gross income', I ? I.grossInc : null]]],
      ['da', 'Debt to assets', W.assets ? W.liab / W.assets : null, 'x2', 'Total debt ÷ total assets · 0.30 or less', v => v <= .3 ? 'g' : v <= .5 ? 'a' : 'r', [['Liabilities', W.liab], ['Assets', W.assets]]],
      ['tr', 'Tax rate', I ? I.sel.tax / I.grossInc : null, 'pct', 'Estimated tax ÷ gross income', null, [['Tax, estimated', I ? I.sel.tax : null], ['Gross income', I ? I.grossInc : null]]]
    ].forEach(([id, n, v, fmt, sub, judge, kv]) => { if (v != null && isFinite(v)) R.push({ id, n, v, fmt, sub, st: judge ? judge(v) : null, kv: kv.filter(x => x[1] != null), band: RATIO_BANDS_[id] || null }); });
  }
  // Insurance: the Insurance block on the Income & estimations tab
  const ins = [];
  try {
    const sh = findSheet_(ss, CONFIG.TABS.estimations), v = sh.getDataRange().getValues();
    const hr = v.findIndex(r => /^person$/i.test(str_(r[0])) && r.some(c => /^cover$/i.test(str_(c))));
    if (hr >= 0) {
      const H = v[hr].map(str_), col = re => H.findIndex(h => re.test(h));
      const K = { p: 0, cat: col(/^category$/i), type: col(/^type$/i), cover: col(/^cover$/i), by: col(/^insurer$/i) };
      let p = '', cat = '';
      for (let r = hr + 1; r < v.length; r++) {
        const row = v[r];
        if (str_(row[K.p])) p = str_(row[K.p]);
        if (str_(row[K.cat])) cat = str_(row[K.cat]);
        const cover = num_(row[K.cover]);
        if (isNaN(cover)) { if (!str_(row[K.type]) && !str_(row[K.p])) break; continue; }
        ins.push({ person: p, cat, type: str_(row[K.type]), cover, by: str_(row[K.by]) });
      }
    }
  } catch (e) { console.error('Insurance: ' + e.message); }
  let premiums = [];
  try { premiums = appPremiums_(ss); } catch (e) { console.error('Premiums: ' + e.message); }
  const tok = s => String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(w => w.length > 2 && !/^(insurance|company|ltd|limited|general|life)$/.test(w));
  ins.forEach(x => { const t = tok(x.by)[0]; x.premiums = t ? premiums.filter(p => tok(p.item).indexOf(t) >= 0) : []; });
  // When the sheet has an Insurance premiums tab, each cover is tied to its policy there (same person, kind, provider and base or top-up).
  let policies = null;
  try { policies = appInsPolicies_(ss, now); } catch (e) { console.error('Insurance premiums: ' + e.message); }
  if (policies) {
    const first = s => String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean)[0] || '';
    const kind = s => /term|life/i.test(s) ? 'term' : /health|medical|mediclaim/i.test(s) ? 'health' : 'other';
    const tier = s => /top.?up|super/i.test(s) ? 'topup' : /floater/i.test(s) ? 'floater' : 'base';
    ins.forEach(x => {
      const i = policies.findIndex(p => first(p.person) === first(x.person) && first(p.provider) === first(x.by) && kind(p.type) === kind(x.cat) && (kind(p.type) === 'term' || tier(p.type) === tier(x.type)));
      x.policy = i >= 0 ? i : null;
      if (i >= 0 && policies[i].cover) { x.coverSheet = x.cover; x.cover = policies[i].cover; }
      if (i >= 0) { const p = policies[i]; p.linked = true; x.premiums = p.lines.filter(l => l.paid && l.amt != null).map(l => ({ k: l.k, amt: l.amt, fy: fyName_(new Date(l.k).getMonth() >= 3 ? new Date(l.k).getFullYear() : new Date(l.k).getFullYear() - 1), src: 'tab' })).reverse(); }
    });
  }
  // Locker: details above the log, then each visit
  let locker = null;
  try {
    const sh = findSheet_(ss, CONFIG.TABS.locker), v = sh.getDataRange().getValues();
    const hr = v.findIndex(r => /access date/i.test(str_(r[0])));
    const meta = v.slice(0, hr).filter(r => str_(r[0])).map(r => [str_(r[0]), r[1] instanceof Date ? dayKey_(r[1].getTime()) : (typeof r[1] === 'number' ? r[1] : str_(r[1]))]);
    const list = x => /^empty$/i.test(str_(x)) ? [] : str_(x).split(/\n+/).map(t => t.trim()).filter(Boolean);
    const visits = v.slice(hr + 1).filter(r => r[0] instanceof Date).map(r => {
      const before = list(r[2]), after = list(r[3]);
      return { d: dayKey_(r[0].getTime()), by: str_(r[1]), txn: str_(r[4]), before, after, added: after.filter(x => before.indexOf(x) < 0), removed: before.filter(x => after.indexOf(x) < 0) };
    }).sort((a, b) => a.d < b.d ? 1 : -1);
    locker = { meta, visits, items: visits.length ? visits[0].after : [] };
  } catch (e) { console.error('Locker: ' + e.message); }
  // Data checks: error cells in the tabs the app reads, as the sheet shows them, and the state of the daily closes
  const checks = [];
  const ERR = /^#(REF!|NAME\?|N\/A|VALUE!|DIV\/0!|ERROR!|NUM!)/;
  ss.getSheets().filter(s => CONFIG.APP.CHECK_TABS.test(s.getName()) && s.getName() !== CONFIG.LOG_TAB && s.getName() !== CONFIG.REWARD_TAB).forEach(s => {
    const rg = s.getDataRange(), v = rg.getValues(), shown = rg.getDisplayValues ? rg.getDisplayValues() : null, F = rg.getFormulas(), bad = [];
    v.forEach((row, r) => row.forEach((c, k) => {
      // Only cells that show an error in the sheet itself; a value the script reads mid-recalculation is not flagged.
      if (typeof c === 'string' && ERR.test(c) && (!shown || ERR.test(String(shown[r][k])))) bad.push({ a: appA1_(r, k), e: c, f: F[r] && F[r][k] ? String(F[r][k]).slice(0, 80) : '' });
    }));
    if (bad.length) checks.push({ ok: false, n: s.getName(), sub: bad.length + ' cell' + (bad.length > 1 ? 's' : '') + ' showing an error', cells: bad.slice(0, 20) });
  });
  const realCloses = closes.filter(c => c.slot === 'close').length;
  checks.unshift({ ok: realCloses > 0, n: 'Daily closes', sub: realCloses ? realCloses + ' closes · last ' + closes[closes.length - 1].k : 'No midnight closes yet. Add the recordDailyClose trigger.', cells: [] });
  W.holdings.filter(h => h.g === 'Asset' && isReward_(h) && h.val > 0).forEach(h => { const e = noteDate_(h); if (e && e < startOfDay_(now)) checks.push({ ok: false, n: 'Lapsed balance still counted', sub: h.cust + ' ' + h.veh.toLowerCase() + ' ₹' + Math.round(h.val).toLocaleString('en-IN') + ', expired ' + dayKey_(e.getTime()), cells: [] }); });
  if (R2) checks.push({ ok: R2.programmes.length > 0, n: 'Rewards ledger', sub: R2.diag || 'Not read', cells: [] });

  (W.excluded || []).forEach(x => checks.push({ ok: true, n: 'Left out of net worth', sub: x.name + ' · ' + (x.val != null ? '₹' + Math.round(x.val).toLocaleString('en-IN') : 'value not readable') + ' · tracked on Invest only', val: x.val, id: x.key, cells: [] }));
  return { ratios: R, insurance: ins, policies, locker, checks };
}
function appA1_(r, c) { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s + (r + 1); }
