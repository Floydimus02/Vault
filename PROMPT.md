# Adapt the script to your sheet

The script reads your workbook by tab names, column headers and row labels. If your sheet is laid out differently, sections of the email go missing or the script stops. This prompt has an AI assistant (Claude, ChatGPT, Gemini or similar) rewrite only the parts of the script that read the sheet, so every calculation and the email itself stay as they are.

## Steps

1. **Print your sheet's layout.** In Apps Script, make sure the latest [`Code.gs`](Code.gs) is pasted in. At the top, pick `describeWorkbook` from the function list and click Run. Copy everything the Execution log prints.
   - It masks numbers, dates and email addresses. Text and formulas stay, because labels are what the script reads.
   - **Check the output for names of people, places or accounts, and replace them before you share it.**
2. **Open a new chat with an AI assistant.** Paste the prompt below, then fill its three slots:
   - `CODE`: the full contents of `Code.gs`.
   - `LAYOUT`: the output from step 1.
   - `CHANGES`: what you changed or want, in plain words. For example: "My expenses run January to December", "I have no payslip register", "Holdings are on a tab called Assets with the value in column E".
3. **Replace `Code.gs`** in Apps Script with the code the assistant returns. Run `sendSnapshot` and check the email. If a section is missing, the reason is under Executions; paste it back into the same chat.

## The prompt

Copy everything inside the box.

````text
You are updating a Google Apps Script (V8) that reads a personal finance workbook in Google Sheets and emails an HTML snapshot. My workbook's layout differs from the one the script expects. Change only the code that reads the workbook, so that the rest of the script receives exactly the same data as before.

## How the script is built

It has four layers. Only layer 1 may change.

1. Readers: functions that read tabs and return plain objects. findSheet_, findFySheet_, monthHeader_, readWorth_, readMarkets_, readSpend_, readIncomeTax_ (only its reading part, see below), readPayouts_, readLocker_, plus CONFIG.TABS (patterns that match tab names).
2. Calculations: the tax engine inside readIncomeTax_ (the calc function and everything after "Paid, balance, advance tax and interest"), slabTax_, loanTerms_, return grading, ratios, the snapshot log (logSheet_, readLog_, toLogEntry_, appendLog_, closest_).
3. Rendering: renderEmail_ and every helper after "Components" (card_, section_, miniBars_, table_ and so on).
4. Orchestration: sendSnapshot, sendSnapshot_, describeWorkbook.

Rules for the readers:
- If a reader throws, the email leaves out that reader's sections and logs the error. readWorth_ is the only required reader.
- Shared helpers: num_ (parses numbers, ₹ signs, commas and "12%"), str_ (trimmed text), monthOf_ (month cell to index), parseLooseDate_, fyStart_, fyName_, sum_.
- Months are financial-year arrays of 12 numbers: index 0 is April, index 11 is March. Missing months are 0.
- Rates and returns are decimals (7.5% is 0.075). Amounts are numbers in rupees. Dates are Date objects.

## What each reader must return

readWorth_(ss) returns:
- holdings: array of { key, name, g, cls, cust, val, prev, xirr, liq, fn }
  - key: text that identifies the holding the same way on every run (the script uses name|custodian|vehicle|occurrence)
  - g: "Asset" or "Liability"
  - cls: asset class; cust: custodian; name: instrument name
  - val: value now; prev: value at the last review, or null
  - xirr: return as a decimal, or NaN. For a loan, its interest rate.
  - liq: liquidity text. Text containing "tier 1" counts as instant access.
  - fn: purpose text. Text containing "safety" counts towards emergency cover.
- assets, liab, nw: totals; sheetPrevNw: net worth from the prev values
- snapDate: date of the last review, or null
- target: target net worth, or null; targetDate: Date, or null
- fd: FD benchmark rate as a decimal, or null

readMarkets_(ss) returns { list, cpi, repo }:
- list: array of { label, kind, val, dayPct }. kind is "idx" (index level), "fx" (rupees per unit) or "inr" (a rupee price). dayPct is the day's change as a decimal, or NaN.
- cpi: inflation as a decimal, or null; repo: repo rate as a decimal, or null

readSpend_(ss, now) returns:
- fy: label such as "FY 26-27"
- rows: array of { cat, item, m } with m a 12-month array. Items whose name contains "loan" or "EMI" are treated as loan repayments; items containing "travel" as travel; categories containing "household" or "housing" feed the housing ratio.
- mi: index of the month to report (the current month, or 11 for a past year)
- monthTot: 12-month array of total spend; thisMonth: monthTot[mi]
- nUsed: number of months up to mi with spend
- avg: monthly average, or null; avgNote: short text saying where avg came from
- ytd: spend from April to mi
- emiYtd, travelYtd: loan and travel spend from April to mi
- emiMonthly, housingMonthly: average a month over the months with spend
- month, soFar: each { top, bottom, total }. top and bottom are arrays of up to 3 { k: category, v: amount }, for month mi and for April to mi. bottom lists the smallest first.

readIncomeTax_(ss, now): change only the reading part, which ends just before "const grossInc". It must fill these variables, and everything after them must stay unchanged:
- T: from findFySheet_ (gives fy, start, current); Y: T.start, the year the financial year starts
- rows: payslip lines, array of { sec, item, m }. Keep the pay object that follows rows as it is; it is worked out from rows. sec is "e" for earnings or "d" for deductions. Deduction items containing "income tax", "TDS" or "supplementary" count as tax; items containing "PF" or "provident" count as PF. Earnings items containing bonus, notice, joining, arrear, incentive or encash are one-off. Leave months without a payslip at 0.
- regime: "old" or "new"
- inc: { bonus, sav, fd, sgb, div, reit, stcg, ltcg }: yearly other income in rupees
- dOld: { hra, hlInt, c80, medSelf, medPar, nps1b }: old-regime deductions. c80 includes pay.pfFY.
- both: { nps2, leave }: deductions allowed in both regimes
- paid: { tdsOther, adv, sat }. adv is an array of 4 advance tax payments, one per instalment.
- PO: the result of readPayouts_, or null
If a source doesn't exist in my workbook, set its values to 0 instead of removing the variable.

readPayouts_(ss, Y) returns { sgbRecv, sgbYear, reitTax, reitTds, reitNet, div, cash }:
- sgbRecv: SGB interest received this financial year; sgbYear: the full-year estimate
- reitTax: taxable REIT income; reitTds: TDS on REIT payouts; reitNet: REIT cash received
- div: dividends received
- cash: 12-month array of all payout cash received, by payment month

readLocker_(ss) returns { url, last, items }:
- url: link to the locker tab
- last: { d: Date, by: text, after: text } for the latest visit, or null
- items: array of item names from that visit

## My workbook

LAYOUT (output of describeWorkbook; numbers, dates and emails are masked):
<<<
PASTE THE LAYOUT HERE
>>>

CHANGES (what differs or what I want):
<<<
DESCRIBE YOUR CHANGES HERE
>>>

## What to do

1. Map each reader's outputs to my tabs, columns and labels. Match by header and label text, not fixed positions, so the code survives added rows and columns. Update CONFIG.TABS for tab names.
2. If my workbook has no source for a reader, make that reader throw a clear error such as "No payslip register", so its sections are left out.
3. Keep every function name, every return field and its unit (decimals for rates, rupees for amounts, Date objects for dates, April-first month arrays).
4. If my financial year or month order differs, convert it to the April-first arrays inside the reader.
5. Don't change layers 2 to 4, CONFIG values other than TABS, or the tax rules, unless CHANGES asks for it. Don't add libraries, external services or network calls.
6. If something in my layout is ambiguous, state the assumption you made instead of guessing silently.

## Reply with

1. The complete updated Code.gs in one code block, ready to paste.
2. A list of each change: function, what it now reads, and why.
3. Any assumptions, and which sections of the email will be missing because my workbook has no data for them.

## The current script

CODE:
<<<
PASTE THE FULL CONTENTS OF Code.gs HERE
>>>
````

## Tips

- One change at a time is easier to check. Run `sendSnapshot` after each.
- To change a calculation or the email's design rather than the sheet layout, say so in `CHANGES`. The prompt otherwise tells the assistant to leave those alone.
- Before changing the live script, keep a copy of the working `Code.gs`.
