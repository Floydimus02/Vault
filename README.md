<img src="assets/banner-light.png#gh-light-mode-only" alt="Vault: a ledger in a letter">
<img src="assets/banner-dark.png#gh-dark-mode-only" alt="Vault: a ledger in a letter">

### A ledger in a letter.

Know where every rupee is, what it earns and what it costs you. Every single day.

<a href="https://github.com/Floydimus02/Vault/stargazers"><img src="https://img.shields.io/github/stars/Floydimus02/Vault?style=social" alt="GitHub stars"></a>

## Your money is leaking. Quietly.

You earn, you are taxed, you spend, you invest, and you are taxed again. Somewhere in between, money slips away and you never see it go.

- Cash sits idle while inflation eats it.
- Reward points expire before you redeem them.
- A premium creeps up at every renewal, and nobody notices.
- An advance tax instalment is missed, and interest starts ticking.
- The EMI account runs dry two days before the debit.
- A fund trails a plain fixed deposit for years, after tax.

None of it feels urgent. All of it compounds. And you can't fix what you can't see.

## Vault makes you see it

One Google Sheet in. Two views out: a letter in your inbox and a private app on your phone. From 10,000 feet, your net worth tonight, down to 10 mm, one lot of one fund and the tax if you sold it tomorrow.

<details>
<summary><b>See the letter</b></summary>
<p align="center">
  <img src="assets/email-light.png#gh-light-mode-only" width="420" alt="The Vault letter">
  <img src="assets/email-dark.png#gh-dark-mode-only" width="420" alt="The Vault letter">
</p>
</details>

It is built for three kinds of people:

- **The tracker** who logs every detail and wants it to finally pay off.
- **The starter** who wants discipline but doesn't know where to begin.
- **The lost** who has no idea where their money stands.

Vault brings method to the madness, structure to the data and clarity to the ambiguity.

## Why I built it

I tracked my money for years, changing what I captured, how and when. Rows of numbers never told me where I stood. As a visual thinker, I needed to *see* it: where I am, where I want to be, and the decisions that bridge the two.

Version 1 was a daily email. Version 2 is an app that lives on any device, opens only for me, and tells me in seconds what used to take an evening.

<p align="center">
  <img src="assets/app-light-1.png#gh-light-mode-only" alt="Vault: Home, Invest, a holding and Check">
  <img src="assets/app-dark-1.png#gh-dark-mode-only" alt="Vault: Home, Invest, a holding and Check">
</p>

## Function first

### See where you stand
- **Net worth, every night.** A snapshot at midnight. Trends over 7 days, 30 days or 12 month ends.
- **What moved today.** Markets, your biggest movers, and changes since last night's close.
- **Heads-ups that matter.** An event tomorrow, a renewal next week, a thin EMI account.

### Make every rupee work
- **Every holding, down to the lot.** Gain, XIRR, long or short term, and the tax if you sold today.
- **The honest benchmark.** Each investment against a fixed deposit, after tax. Laggards have nowhere to hide.
- **Where it all sits.** Allocation by purpose, asset class, liquidity and custodian.
- **Ratios with a verdict.** Savings, emergency cover, liquidity, debt and housing, each marked healthy or off target.

### Plug the leaks
- **Spend with a pace check.** The month so far against your usual, and where it will land.
- **Tax, worked out.** Both regimes, advance tax due and any interest, from your own payslips and payouts.
- **The loan, decoded.** EMIs and interest left, rate history, and what a prepayment would save.
- **Rewards in rupees.** What your points are worth and when they expire.
- **Insurance with a memory.** Cover per person, renewals 30 days out, premium growth across the years.

### Keep it yours
- **Only you get in.** It runs as you and opens only for you. No server, database, tracker or third party.
- **Private in public.** One tap hides amounts above ₹25 L on Home and ₹1 Cr elsewhere.
- **Live on refresh.** Every open reads the sheet afresh. Prices follow `GOOGLEFINANCE`, or any better source you plug in.
- **Your sheet, your rules.** Any structure you like. Leave anything out of net worth. A missing value is skipped, never shown blank.
- **Nothing to maintain.** A sheet, a script and a page. Free, on Google.

## Then form

- **Minimal and visual.** One accent, one typeface, a chart wherever a number needs a story.
- **Made for the thumb.** Tap search and the bottom bar becomes a search field that suggests from your own holdings, payees, policies and locker items, typos and all.
- **You never lose your place.** Details open as bottom sheets; going back returns you exactly where you were.
- **Light or dark.** It follows your device.

<p align="center">
  <img src="assets/app-light-2.png#gh-light-mode-only" alt="Vault: Income, Spend, Home loan and Search">
  <img src="assets/app-dark-2.png#gh-dark-mode-only" alt="Vault: Income, Spend, Home loan and Search">
</p>

## Every day you don't look, the leaks win

<a href="https://floydimus02.github.io/Vault/" target="_blank" rel="noopener"><b>Try the live demo →</b></a>

It runs on made-up money. With yours in it, it hits differently.

<br>

## Get started

You need a Google account and about 15 minutes.

**What's inside**

```
Vault/
├── apps-script/
│   ├── Code.gs        The engine: reads the sheet, sends the letter and serves the app
│   └── Index.html     The app's page. Only needed for the app
├── template/
│   └── Vault-template.xlsx    The sheet to start from, filled with made-up data
├── PROMPT.md          Adapts Code.gs to a sheet laid out differently
├── docs/              The live demo, served by GitHub Pages
├── assets/            Images for this page
└── LICENSE
```

The letter needs `Code.gs` alone. The app needs `Code.gs` and `Index.html`. Both read the same sheet and share every calculation.

**1. Copy the sheet**
1. Download [`Vault-template.xlsx`](template/Vault-template.xlsx).
2. In Google Drive: **New → File upload**, and choose it.
3. Open it and click **File → Save as Google Sheets**. Use this copy from now on.

**2. Add the script**
1. In the sheet: **Extensions → Apps Script**.
2. Delete everything in `Code.gs`, paste in [`Code.gs`](apps-script/Code.gs) from this repository, and save.
3. Near the top, set `RECIPIENTS` to your email address, or leave it blank to send only to yourself.

**3. Approve it once**
1. In the bar above the code, choose `sendSnapshot` and click **Run**.
2. Allow the permissions: **Review permissions → your account → Advanced → Go to project → Allow**. The script reads your sheet, keeps one hidden history tab and emails you.
3. Check your inbox for the first letter.

**4. Schedule it**

Click the clock icon (**Triggers**) and add two triggers:

| Function | Event source | Type | Time |
|---|---|---|---|
| `recordDailyClose` | Time-driven | Day timer | Midnight to 1am |
| `sendSnapshot` | Time-driven | Any | Whenever you want the letter. Skip it if you only want the app. |

**5. Turn on the app**
1. In Apps Script, click **+** next to Files → **HTML**, and name it `Index`.
2. Delete what's in it, paste in [`Index.html`](apps-script/Index.html), and save.
3. **Deploy → New deployment →** gear icon **→ Web app**. Set **Execute as: Me** and **Who has access: Only myself**. Click **Deploy**.
4. Open the **Web app URL** and bookmark it. To use your own domain, point it at this URL.
5. On your phone, open the same URL signed in to the same Google account, then add it to your home screen: Safari → Share → **Add to Home Screen**, or Chrome → ⋮ → **Add to home screen**.

**6. Put your money in**
1. Replace the example rows with yours, tab by tab. Keep the tab names, headings and row labels; the script finds everything by them. Delete any tab you don't need and its section disappears.
2. For live prices, use formulas such as `=GOOGLEFINANCE("INDEXNSE:NIFTY_50")` in the Markets block on **Portfolio**, and `=GOOGLEFINANCE("NSE:ITC")` in each holding's **Current Price**.

## Day to day

**Vault updates on its own**

| What | When |
|---|---|
| Net worth snapshot | Every night at midnight |
| Prices and NAVs | As `GOOGLEFINANCE` refreshes |
| Every figure, ratio and estimate | Each time you open the app or tap refresh |
| The letter | On your schedule |

**You update**

| Cadence | What | Where |
|---|---|---|
| Monthly | Payslip | **CoI**, the month's column |
| Monthly | Spend | **Expenses**, the month's column |
| Monthly | Loan statement lines | **Loan schedule** |
| Monthly | EPF passbook | **EPF** |
| When it happens | A buy or sell | Its block on **XIRR**, and its units on **Portfolio** |
| When it happens | Dividend, SGB coupon or REIT payout | **Dividends & payouts** |
| When it happens | Reward points earned or redeemed | **Rewards ledger**, and the balance on **Net worth** |
| When it happens | Insurance renewal | **Insurance premiums** |
| When it happens | Locker visit or event | **Locker log**, **Events** |
| At each review | Values at review and the review date | **Net worth**: `Amount: T - X`, `Last updated` |
| Every April | New financial year | Copy **Expenses** and **CoI**, renamed for the new year |

## Make it yours

1. **Change a setting.** Edit `CONFIG` at the top of `Code.gs`: recipients, time zone, what's left out of net worth (`EXCLUDE`), rewards, the EMI alert, and how many days ahead events show.
2. **Change the look.** Edit the colour tokens under `:root` at the top of `Index.html`; `html.dark` holds the dark theme.
3. **Fit a different sheet.** Run `describeWorkbook` in Apps Script, then follow [`PROMPT.md`](PROMPT.md). An AI assistant rewrites only the parts that read your sheet.
4. **Publish it.** After any change: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL stays the same. To check first, open **Deploy → Test deployments**.

## License

[![License: GPL-3.0](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

Copyright (C) 2026 Floydimus.

Vault is free software under the GNU General Public License v3.0. Any distributed version, modified or not, must remain open under the same license; it can never be made closed-source or proprietary. It is distributed without any warranty. See [LICENSE](LICENSE) for the full text.

Vault is not financial or tax advice.
