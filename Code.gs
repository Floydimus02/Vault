/**
 * Vault Snapshot
 * Emails a visual, structured snapshot of the Vault workbook. Every figure is
 * read or calculated from the workbook; nothing is typed in here except who
 * receives the email.
 *
 * Setup
 *   1. In the Vault Google Sheet: Extensions → Apps Script. Replace all code with this file.
 *   2. Set RECIPIENTS in CONFIG.
 *   3. Run sendSnapshot once and approve the permissions.
 *   4. Schedule it under Triggers (clock icon) → Add trigger: function sendSnapshot,
 *      event source "Time-driven", then any type and time. Add as many triggers as you want.
 *
 * Changed the sheet's layout? See PROMPT.md in the project.
 *
 * Copyright (C) 2026 Floydimus.
 * Free software under the GNU General Public License v3.0. See LICENSE.
 * Distributed WITHOUT ANY WARRANTY. Not financial or tax advice.
 *
 * Each run adds one row to a hidden tab (CONFIG.LOG_TAB). Comparisons with
 * 24 hours, 7 days and 30 days ago come from that tab and appear once it holds
 * a run from that long ago. Run errors are listed under Executions in Apps Script.
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
  TABS: {
    worth: /net worth/i,
    portfolio: /portfolio/i,
    expenses: /expense/i,
    coi: /\bcoi\b|computation of income/i,
    locker: /locker/i,
    payouts: /dividend|payout/i
  }
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

function sendSnapshot_(slot) {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
  SpreadsheetApp.flush();
  const now = new Date();
  const run = (label, fn) => { try { return fn(); } catch (e) { console.error(label + ': ' + e.message); return null; } };

  const W = readWorth_(ss);
  const M = run('Markets', () => readMarkets_(ss)) || { list: [], cpi: null, repo: null };
  const X = run('Spend', () => readSpend_(ss, now));
  const I = run('Salary and tax', () => readIncomeTax_(ss, now));
  const L = run('Locker', () => readLocker_(ss));

  const log = readLog_(ss);
  const cur = toLogEntry_(now, slot, W, M);
  const ref = {
    d1: closest_(log, now.getTime() - 864e5, 6 * 36e5),
    d7: closest_(log, now.getTime() - 7 * 864e5, 2 * 864e5),
    d30: closest_(log, now.getTime() - 30 * 864e5, 5 * 864e5)
  };
  const html = renderEmail_({ ss, now, W, M, X, I, L, log, cur, ref });
  const to = CONFIG.RECIPIENTS || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({ to, subject: 'Net worth as of ' + fmtDay_(now), htmlBody: html, name: 'Vault' });
  appendLog_(ss, cur);
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
    val: col(/amount:\s*t$/i), prev: col(/t\s*-\s*x/i), xirr: col(/xirr/i), fdCol: col(/>\s*fd/i), liq: col(/liquidity/i), fn: col(/function/i) };
  const holdings = [], seen = {};
  for (let r = hr + 1; r < v.length; r++) {
    const g = str_(v[r][K.g]);
    if (/^last updated/i.test(g)) break;
    if (!/^(asset|liability)$/i.test(g)) continue;
    const val = num_(v[r][K.val]), prev = num_(v[r][K.prev]);
    const name = str_(v[r][K.name]) || str_(v[r][K.veh]);
    const base = [name, str_(v[r][K.cust]), str_(v[r][K.veh])].join('|');
    seen[base] = (seen[base] || 0) + 1;
    holdings.push({
      key: base + '|' + seen[base], name, g: /^liab/i.test(g) ? 'Liability' : 'Asset', cls: str_(v[r][K.cls]) || 'Other', cust: str_(v[r][K.cust]),
      val: !isNaN(val) ? val : (!isNaN(prev) ? prev : 0), prev: isNaN(prev) ? null : prev,
      xirr: num_(v[r][K.xirr]), liq: str_(v[r][K.liq]) || 'Unspecified', fn: str_(v[r][K.fn]) || 'Unspecified'
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
  return { holdings, assets, liab, nw: assets - liab, sheetPrevNw: sumOf('Asset', 'prev') - sumOf('Liability', 'prev'), snapDate, target, targetDate, fd };
}

function readMarkets_(ss) {
  const v = findSheet_(ss, CONFIG.TABS.portfolio).getDataRange().getValues();
  const WANT = [
    ['Nifty 50', /^nifty 50$/i, 'idx'], ['Nasdaq 100', /^nasdaq 100$/i, 'idx'],
    ['USD/INR', /^usd$/i, 'fx'], ['EUR/INR', /^eur$/i, 'fx'], ['GBP/INR', /^gbp$/i, 'fx'],
    ['Gold 24K · 10 g', /^gold\s*-\s*24k$/i, 'inr'], ['Gold 22K · 10 g', /^gold\s*-\s*22k$/i, 'inr'], ['Silver 999 · kg', /^silver\s*-\s*999k?$/i, 'inr']
  ];
  const list = []; let cpi = null, repo = null;
  for (let r = 0; r < v.length; r++) for (let c = 0; c < v[r].length; c++) {
    const s = str_(v[r][c]);
    WANT.forEach(([label, re, kind]) => { if (re.test(s)) { const val = num_(v[r][c + 1]); if (!isNaN(val)) list.push({ label, kind, val, dayPct: kind === 'idx' ? num_(v[r][c + 3]) : NaN }); } });
    if (/^inflation$/i.test(s)) { const a = num_((v[r + 1] || [])[c + 2]), b = num_(v[r][c + 2]); cpi = !isNaN(a) && a < 1 ? a : (!isNaN(b) && b < 1 ? b : null); }
    if (/^repo rate$/i.test(s)) { const x = num_((v[r + 1] || [])[c]); if (!isNaN(x)) repo = x; }
  }
  return { list, cpi, repo };
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
    const m = Array(12).fill(0); hdr.cols.forEach(({ c, m: mi }) => { const n = num_(v[r][c]); if (!isNaN(n)) m[mi] = n; });
    rows.push({ cat: cat || 'Other', item: label, m });
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
  const priorIdx = MON.map((_, j) => j).filter(j => j < mi && monthTot[j] > 0);
  const used = MON.map((_, j) => j <= mi && monthTot[j] > 0), nUsed = used.filter(Boolean).length;
  const avgOf = f => nUsed ? sum_(rows.filter(f).map(x => sum_(x.m.filter((_, j) => used[j])))) / nUsed : null;
  const ytdOf = f => sum_(rows.filter(f).map(x => sum_(x.m.slice(0, mi + 1))));
  const isEmi = x => /loan|\bemi\b/i.test(x.item), isTravel = x => /travel/i.test(x.item);
  const cats = months => { const m = {}; rows.forEach(x => { m[x.cat] = (m[x.cat] || 0) + sum_(months.map(j => x.m[j])); }); return Object.keys(m).map(k => ({ k, v: m[k] })).filter(x => x.v > 0).sort((a, b) => b.v - a.v); };
  const topBottom = list => ({ top: list.slice(0, 3), bottom: list.slice(3).slice(-3).reverse(), total: sum_(list.map(x => x.v)) });
  return {
    fy: T.fy, rows, mi, monthTot, thisMonth: monthTot[mi], nUsed,
    avg: !isNaN(sheetAvg) ? sheetAvg : (priorIdx.length ? sum_(priorIdx.map(j => monthTot[j])) / priorIdx.length : null),
    avgNote: !isNaN(sheetAvg) ? 'from the sheet' : priorIdx.length ? MON[priorIdx[0]] + ' to ' + MON[priorIdx[priorIdx.length - 1]] : '',
    ytd: sum_(monthTot.slice(0, mi + 1)), emiYtd: ytdOf(isEmi), travelYtd: ytdOf(isTravel),
    emiMonthly: avgOf(isEmi), housingMonthly: avgOf(x => /household|housing/i.test(x.cat)),
    month: topBottom(cats([mi])), soFar: topBottom(cats(MON.map((_, j) => j).filter(j => j <= mi)))
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
    return { tax: base + special + sur + cess };
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
    fy: T.fy, Y, regime, pay, inc, payouts: PO, advByQ: paid.adv, sat: paid.sat, tdsOther: paid.tdsOther, grossInc, old, new: neu, sel, advPaid, paidSoFar, tdsAhead: pay.tdsFY - pay.tdsToDate, netDue, applies, sched, s424, s424Months, filing,
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
  const cash = Array(12).fill(0);
  // SGB coupons: half-yearly and fixed, so a bond paid once this year pays the same again.
  const qc = H0.findIndex(h => /^quarter$/i.test(h));
  let sgbRecv = 0, sgbYear = 0;
  for (let c = 1; c < (qc > 0 ? qc : 1); c++) {
    if (!H0[c]) continue;
    const fy = v.slice(1).filter(r => inFy(r[0]) && num_(r[c]) > 0).map(r => ({ d: r[0], a: num_(r[c]) }));
    const recv = sum_(fy.map(p => p.a));
    fy.forEach(p => { cash[mIdx(p.d)] += p.a; });
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
          if (/^total$/i.test(name)) { if (!isNaN(n)) { reitNet += n; cash[mIdx(date)] += n; } return; }
          const gross = at.hold != null ? num_(v[at.hold][c]) : NaN, rate = at.rate != null ? num_(v[at.rate][c]) : NaN, tds = at.tds != null ? num_(v[at.tds][c]) : NaN;
          if (!isNaN(gross) && (/interest|other/i.test(name) || (/dividend/i.test(name) && rate > 0))) reitTax += gross;
          if (!isNaN(tds)) reitTds += tds;
        });
      }
    }
  }
  // Stock dividends
  let div = 0;
  if (sc >= 0) for (let c = sc + 1; c < H1.length; c++) if (H1[c]) v.slice(2).forEach(r => { const a = num_(r[c]); if (inFy(r[sc]) && a > 0) { div += a; cash[mIdx(r[sc])] += a; } });
  return { sgbRecv, sgbYear, reitTax, reitTds, reitNet, div, cash };
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
    try { return { ts: new Date(r[0]).getTime(), nw: Number(r[2]), h: JSON.parse(r[5] || '{}'), cls: JSON.parse(r[6] || '{}'), mkt: JSON.parse(r[7] || '{}') }; } catch (e) { return null; }
  }).filter(Boolean).sort((a, b) => a.ts - b.ts);
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
function closest_(log, t, tol) {
  let best = null, bd = Infinity;
  log.forEach(e => { const d = Math.abs(e.ts - t); if (d < bd && d <= tol) { bd = d; best = e; } });
  return best;
}

/* ===================== Rendering ===================== */

function renderEmail_(o) {
  const { ss, now, W, M, X, I, L, log, cur, ref } = o;
  const A = W.holdings.filter(h => h.g === 'Asset');
  let b = ''; const S_ = {};

  /* Net worth */
  const tiles = [
    ['24 hours', ref.d1 && W.nw - ref.d1.nw, ref.d1 && ref.d1.nw],
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
     ${W.target ? targetBlock_(W, now) : ''}`);

  /* Trend */
  const trend = dailySeries_(log.concat([cur]), CONFIG.TREND_DAYS);
  if (trend.length >= 2) S_.trend = section_('Net worth, last ' + CONFIG.TREND_DAYS + ' days', trendBars_(trend));

  /* Allocation */
  const alloc = groupSum_(A, h => h.cls).map(x => {
    const p = ref.d1 && ref.d1.cls['Asset|' + x.k];
    return Object.assign(x, p != null ? { d: x.v - p, p } : {});
  });
  S_.alloc = section_('Allocation', miniBars_(alloc, W.assets));

  /* Movers over 24 hours */
  if (ref.d1) {
    const mv = A.filter(h => ref.d1.h[h.key] != null).map(h => ({ h, d: h.val - ref.d1.h[h.key], p: ref.d1.h[h.key] })).filter(x => Math.abs(x.d) >= 1);
    const up = mv.filter(x => x.d > 0).sort((a, c) => c.d - a.d).slice(0, CONFIG.TOP_MOVERS);
    const dn = mv.filter(x => x.d < 0).sort((a, c) => a.d - c.d).slice(0, CONFIG.TOP_MOVERS);
    const list = (arr, title, col) => `${subhead_(title, col)}<table role="presentation" width="100%" cellpadding="0" cellspacing="0">${arr.map(x => `<tr><td style="padding:6px 0;font-size:14px;color:${C.ink}">${esc_(x.h.name)}<div style="font-size:12px;color:${C.muted}">${esc_(x.h.cls)}</div></td><td align="right" style="padding:6px 0;white-space:nowrap">${deltaText_(x.d, x.p)}</td></tr>`).join('')}</table>`;
    const cols = [up.length && list(up, 'TOP ' + up.length + ' GAINERS', C.green), dn.length && list(dn, 'TOP ' + dn.length + ' LOSERS', C.red)].filter(Boolean);
    if (cols.length) S_.movers = section_('Movers, 24 hours', twoCol_(cols));
  }

  /* Liquidity and purpose */
  S_.liq = section_('Liquidity and purpose', twoCol_([miniBars_(groupSum_(A, h => h.liq), W.assets), miniBars_(groupSum_(A, h => h.fn), W.assets)]));

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
         return `<td width="33%" valign="top" style="padding-right:6px"><div style="background:${bg};border-radius:12px;padding:10px 12px;color:${fg}">
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
      return [x.label, lvl, isNaN(x.dayPct) ? '' : pctDelta_(x.dayPct), p != null ? pctDelta_((x.val - p) / p) : ''];
    }).concat([['Inflation', pct_(M.cpi, 2), '', ''], ['RBI repo rate', pct_(M.repo, 2), '', '']].filter(r => r[1] !== '—'));
    const keep = [0, 1, 2, 3].filter(i => i < 2 || rows.some(r => r[i] !== ''));
    S_.markets = section_('Markets', table_(keep.map(i => ['', 'Level', 'Day', '24 hours'][i]), rows.map(r => keep.map(i => r[i])), keep.map(i => i > 0)));
  }

  /* Spend */
  if (X) {
    const vsAvg = X.avg ? (X.thisMonth - X.avg) / X.avg : null;
    const mx = Math.max.apply(null, X.monthTot.slice(0, X.mi + 1).concat([1]));
    const span = MON[0] + ' to ' + MON[X.mi];
    const tb = (t, label) => (t.top.length ? subhead_(label + ' · TOP ' + t.top.length, C.green) + miniBars_(t.top, t.total) : '') + (t.bottom.length ? '<div style="height:8px"></div>' + subhead_(label + ' · BOTTOM ' + t.bottom.length, C.red) + miniBars_(t.bottom, t.total) : '');
    S_.spend = section_(`Spend · ${X.fy}`,
      `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
        ${statTd_(MON[X.mi] + ' so far', inrC_(X.thisMonth), vsAvg == null ? '' : `<span style="color:${vsAvg > 0 ? C.red : C.green}">${sign_(vsAvg)}${Math.abs(vsAvg * 100).toFixed(0)}% vs average</span>`)}
        ${X.avg ? statTd_('Monthly average', inrC_(X.avg), X.avgNote) : ''}
        ${statTd_('Year to date', inrC_(X.ytd), span)}</tr></table>
       <div style="font-size:12px;color:${C.muted};margin-top:8px">Includes home loan EMI ${inrC_(X.emiYtd)} and travel ${inrC_(X.travelYtd)} for ${span}.</div>
       ${monthBars_(X.monthTot.slice(0, X.mi + 1).map((t, j) => ({ v: t, label: MON[j], hi: j === X.mi })), mx)}
       <div style="height:14px"></div>
       ${twoCol_([tb(X.month, MON[X.mi].toUpperCase()), tb(X.soFar, span.toUpperCase())])}`);
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
       ${I.applies && next && next.gap >= 1 ? `<div style="margin-top:12px;padding:12px 14px;border-radius:12px;background:${C.amberBg};color:${C.amber};font-size:14px;font-weight:600">Pay ${inr_(next.gap)} by ${fmtDate_(next.due)} to avoid ${inr_(next.int)} of interest under section 425.</div>` : ''}`);
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
     <details style="border:1px solid ${C.line};border-radius:12px;padding:10px 14px">
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
          <td width="40%" style="padding:7px 10px 7px 0"><div style="background:${C.surface};border-radius:4px;height:12px"><div style="width:${Math.max(1, Math.min(100, Math.abs(y) / Y_.inc * 100)).toFixed(1)}%;background:${colr};height:12px;border-radius:4px"></div></div></td>
          <td align="right" style="font-size:14px;color:${C.ink};padding:7px 0;white-space:nowrap;${/Take-home|Saved/.test(l) ? 'font-weight:700' : ''}">${y < 0 ? '−' : ''}${inrC_(Math.abs(y))} <span style="font-size:12px;color:${C.muted};font-weight:400">${(Math.abs(y) / Y_.inc * 100).toFixed(0)}%</span></td>
          <td align="right" style="font-size:13px;color:${C.muted};padding:7px 0 7px 12px;white-space:nowrap">${m < 0 ? '−' : ''}${inrC_(Math.abs(m))}</td></tr>`).join('')}
      </table>
      <div style="font-size:12px;color:${C.muted};margin-top:8px">Up to the last payslip (${MON[k]}). Payouts are SGB interest, dividends and REIT distributions from the payouts tab.</div>`);
  }

  b = ['worth', 'trend', 'movers', 'markets', 'cashflow', 'salary', 'tax', 'spend', 'liab', 'alloc', 'liq', 'returns', 'ratios', 'locker'].map(x => S_[x] || '').join('');

  return `<!doctype html><html><head><style>details[open] .vlk{display:block !important} details[open] .vlk-link{display:none !important}</style></head><body style="margin:0;padding:0;background:${C.surface}">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${C.surface}"><tr><td align="center" style="padding:20px 12px">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;font-family:${C.font};color:${C.ink}">
    <tr><td style="padding:0 4px 14px"><table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      <td width="44"><div style="width:36px;height:36px;border-radius:10px;background:${C.accent};color:#fff;font-weight:800;font-size:20px;line-height:36px;text-align:center">V</div></td>
      <td style="font-size:20px;font-weight:800;color:${C.ink}">Vault</td></tr></table></td></tr>
    ${b}
    <tr><td style="font-size:13px;color:${C.muted};padding:4px 4px 0">Source: <a href="${ss.getUrl()}" style="color:${C.accent};font-weight:700">${esc_(ss.getName())}</a></td></tr>
  </table></td></tr></table></body></html>`;
}

/* ---------- Components ---------- */

function card_(inner) { return `<tr><td style="background:${C.bg};border-radius:20px;padding:22px 22px 18px;border:1px solid ${C.line}">${inner}</td></tr><tr><td height="12"></td></tr>`; }
function section_(title, inner) { return card_(`<div style="font-size:16px;font-weight:700;color:${C.ink};margin-bottom:12px">${esc_(title)}</div>${inner}`); }
function subhead_(t, col) { return `<div style="font-size:12px;font-weight:700;color:${col};letter-spacing:.5px;margin:4px 0 4px">${esc_(t)}</div>`; }
function twoCol_(cols) {
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>${cols.map((c, i) => `<td valign="top" width="${cols.length === 2 ? 50 : 100}%" style="${i ? `padding-left:12px;border-left:1px solid ${C.line}` : (cols.length === 2 ? 'padding-right:12px' : '')}">${c}</td>`).join('')}</tr></table>`;
}
function deltaPill_(label, d, base) {
  const fg = d > 0 ? C.green : d < 0 ? C.red : C.muted, bg = d > 0 ? C.greenBg : d < 0 ? C.redBg : C.surface;
  return `<span style="display:inline-block;background:${bg};color:${fg};border-radius:999px;padding:6px 12px;margin:0 6px 6px 0;font-size:13px;font-weight:700;white-space:nowrap">${sign_(d)}${inrC_(Math.abs(d))}${base ? ' (' + pctChange_(d / base) + ')' : ''} <span style="font-weight:500;color:${C.muted}">${esc_(label.toLowerCase().indexOf('since') === 0 ? label.charAt(0).toLowerCase() + label.slice(1) : 'in ' + label)}</span></span>`;
}
function pctChange_(p) { const a = Math.abs(p * 100); return (a > 0 && a < 0.01 ? sign_(p) + '<0.01' : sign_(p) + a.toFixed(2)) + '%'; }
function deltaText_(d, base) {
  if (Math.abs(d) < 1) return `<span style="color:${C.muted}">0</span>`;
  return `<span style="color:${d > 0 ? C.green : C.red};font-weight:600;white-space:nowrap">${sign_(d)}${inrC_(Math.abs(d))}${base ? ` <span style="font-weight:400">(${sign_(d)}${Math.abs(d / base * 100).toFixed(2)}%)</span>` : ''}</span>`;
}
function pctDelta_(p) { return isNaN(p) ? '' : `<span style="color:${p > 0 ? C.green : p < 0 ? C.red : C.muted};font-weight:600">${sign_(p)}${Math.abs(p * 100).toFixed(2)}%</span>`; }
function pill_(t, fg, bg) { return `<span style="display:inline-block;background:${bg};color:${fg};border-radius:999px;padding:3px 10px;font-size:12px;font-weight:700;white-space:nowrap">${esc_(t)}</span>`; }
/** Name, value · share, optional 24-hour delta, and a bar. Used by Allocation, Liquidity and purpose, Spend. */
function miniBars_(rows, base) {
  return rows.map(x => `<div style="padding:6px 0">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>
      <td style="font-size:13px;color:${C.ink}">${esc_(x.k)}</td>
      <td align="right" style="font-size:13px;color:${C.ink};white-space:nowrap"><b>${inrC_(x.v)}</b> <span style="color:${C.muted}">· ${base ? (x.v / base * 100).toFixed(1) : '0.0'}%</span></td></tr>
      ${x.d != null && Math.abs(x.d) >= 1 ? `<tr><td></td><td align="right" style="font-size:12px">${deltaText_(x.d, x.p)}</td></tr>` : ''}</table>
    <div style="background:${C.surface};border-radius:999px;height:6px;margin-top:4px"><div style="width:${Math.max(1, Math.min(100, base ? x.v / base * 100 : 0)).toFixed(1)}%;background:${C.accent};height:6px;border-radius:999px"></div></div></div>`).join('');
}
function monthBars_(series, mx) {
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin-top:14px;height:74px"><tr>${series.map(s =>
    `<td valign="bottom" align="center" style="padding:0 2px"><div style="height:${Math.max(2, Math.round(s.v / mx * 56))}px;background:${s.proj ? '#E4E8EF' : s.hi === false ? '#D6DCE6' : C.accent};${s.proj ? 'border:1px dashed #B8C0CC;' : ''}border-radius:4px 4px 0 0"></div><div style="font-size:10px;color:${C.muted};padding-top:4px">${s.label}</div></td>`).join('')}</tr></table>`;
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
    <td><div style="background:${C.surface};border-radius:999px;height:10px"><div style="width:${(p * 100).toFixed(1)}%;background:${col};height:10px;border-radius:999px"></div></div></td>
    <td width="52" align="right" style="font-size:13px;font-weight:700;color:${C.ink}">${(p * 100).toFixed(1)}%</td></tr></table>`;
}
function dailySeries_(entries, days) {
  const byDay = {};
  entries.forEach(e => { const k = Utilities.formatDate(new Date(e.ts), CONFIG.TIMEZONE, 'yyyy-MM-dd'); if (!byDay[k] || byDay[k].ts < e.ts) byDay[k] = e; });
  return Object.keys(byDay).sort().slice(-days).map(k => ({ k, v: byDay[k].nw }));
}
function trendBars_(s) {
  const vals = s.map(x => x.v), lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals), span = hi - lo || Math.abs(hi) * 0.01 || 1;
  const bars = s.map((x, i) => `<td valign="bottom" style="padding:0 1px"><div style="height:${8 + Math.round((x.v - lo) / span * 64)}px;background:${i === s.length - 1 ? C.accent : i > 0 && x.v < s[i - 1].v ? '#E8A3BF' : '#D6DCE6'};border-radius:3px 3px 0 0"></div></td>`).join('');
  return `<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="height:76px"><tr>${bars}</tr></table>
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin-top:6px"><tr>
      <td style="font-size:12px;color:${C.muted}">${fmtDate_(new Date(s[0].k + 'T00:00:00'))}</td>
      <td align="center" style="font-size:12px;color:${C.muted}">Low ${inrC_(lo)} · High ${inrC_(hi)}</td>
      <td align="right" style="font-size:12px">${deltaText_(s[s.length - 1].v - s[0].v, s[0].v)}</td></tr></table>`;
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
