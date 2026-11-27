"""Builds Vault-template.xlsx: a made-up workbook with every tab the email and the web app read.
   python3 make_template.py <old-template.xlsx> <out.xlsx>   (then recalculate with LibreOffice)"""
import sys, random, datetime as dt, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
random.seed(7)
D = dt.datetime
TODAY = D(2026, 11, 20)
src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
keep = ['CoI FY 26-27', 'Dividends & payouts', 'Locker log']
for s in list(wb.sheetnames):
    if s not in keep: del wb[s]
BOLD = Font(bold=True)
HEAD = PatternFill('solid', fgColor='E9EDF3')
def header(ws, r, c0, labels):
    for i, l in enumerate(labels):
        cell = ws.cell(r, c0 + i, l); cell.font = BOLD; cell.fill = HEAD

# ---------------------------------------------------------------- XIRR lots (built first: units feed Portfolio)
def months(start, n, day=10):
    y, m = start.year, start.month
    for _ in range(n):
        yield D(y, m, min(day, 28)); m += 1
        if m > 12: m = 1; y += 1
def nav_path(p0, p1, n, vol):
    """A noisy path from p0 to p1 in n steps."""
    out, p = [], p0
    for i in range(n):
        target = p0 + (p1 - p0) * (i + 1) / n
        p = target * (1 + random.uniform(-vol, vol)); out.append(round(p, 4))
    return out
LOTS = {}   # name -> dict(lots=[(date, side, qty, price, amount)], price=current)
def sip(name, start, n, amount, p0, p1, vol=0.04, day=10, sells=()):
    navs = nav_path(p0, p1, n, vol); lots = []
    for d, nav in zip(months(start, n, day), navs):
        q = round(amount / nav, 3); lots.append((d, 'Buy', q, nav, -round(q * nav, 2)))
    for (d, q) in sells:
        nav = navs[min(len(navs) - 1, max(0, (d.year - start.year) * 12 + d.month - start.month))]
        lots.append((d, 'Sell', -q, nav, round(q * nav, 2)))
    lots.sort(key=lambda x: x[0])
    LOTS[name] = {'lots': lots, 'price': p1}
sip('Nifty 50 Index Fund - Direct - Growth', D(2021, 4, 1), 68, 25000, 14.2, 22.6, sells=[(D(2024, 3, 12), 9000)])
sip('Flexi Cap Fund - Direct - Growth', D(2022, 1, 1), 58, 20000, 48.0, 86.4)
sip('Nasdaq 100 Fund of Funds - Direct', D(2021, 9, 1), 40, 15000, 10.5, 21.8, vol=0.06)
LOTS['Arbitrage Fund - Direct - Growth'] = {'lots': [(D(2025, 6, 3), 'Buy', 24000.0, 12.5, -300000.0), (D(2026, 2, 9), 'Buy', 15600.0, 12.82, -199992.0)], 'price': 13.21}
LOTS['Example Industries Ltd'] = {'lots': [(D(2022, 8, 16), 'Buy', 60, 410.0, -24600.0), (D(2024, 1, 22), 'Buy', 40, 520.0, -20800.0)], 'price': 612.5}
LOTS['Example REIT A'] = {'lots': [(D(2024, 5, 14), 'Buy', 500, 352.0, -176000.0)], 'price': 381.4}
LOTS['Example REIT B'] = {'lots': [(D(2024, 10, 3), 'Buy', 400, 341.0, -136400.0)], 'price': 362.75}
LOTS['Gold ETF'] = {'lots': [(D(2024, 7, 8), 'Buy', 2200, 61.2, -134640.0), (D(2025, 11, 4), 'Buy', 1000, 98.5, -98500.0)], 'price': 112.6}
LOTS['Silver ETF'] = {'lots': [(D(2025, 2, 17), 'Buy', 1500, 96.0, -144000.0)], 'price': 141.3}
LOTS['SGB 2031 Series II'] = {'lots': [(D(2023, 9, 18), 'Buy', 30, 5923.0, -177690.0)], 'price': 13640.0}
LOTS['SGB 2032 Series IV'] = {'lots': [(D(2024, 2, 19), 'Buy', 20, 6213.0, -124260.0)], 'price': 13640.0}
LOTS['[22K] Gold bangles'] = {'lots': [(D(2020, 11, 14), 'Buy', 24.0, 118320.0, -118320.0)], 'price': 11320.0}
LOTS['[999K] Silver coins'] = {'lots': [(D(2022, 10, 24), 'Buy', 500.0, 31250.0, -31250.0)], 'price': 151.2}
# NPS: contributions with NAV, and small quarterly charges
nps = []
for d, nav in zip(months(D(2022, 1, 1), 15, 12), nav_path(28.0, 41.0, 15, 0.03)):
    if d.month in (1, 5, 9): nps.append((d, 'Buy', round(25000 / nav, 4), nav, -25000.0))
for y in (2023, 2024, 2025, 2026):
    for m in (4, 7, 10):
        if D(y, m, 6) < TODAY: nps.append((D(y, m, 6), 'Sell', -round(18 / 44.0, 4), 44.0, -18.0))
nps.sort(key=lambda x: x[0])
LOTS['NPS Scheme E - Tier I'] = {'lots': nps, 'price': 48.37}
for k, v in LOTS.items():
    v['units'] = round(sum(l[2] for l in v['lots']), 4)
    v['cost'] = -sum(l[4] for l in v['lots'] if l[1] == 'Buy')
    v['value'] = round(v['units'] * v['price'] / (1 if not k.startswith('[') else 1), 2)
# PPF: deposits and 31 March interest
ppf, bal = [], 0.0
for y in range(2017, 2027):
    dep = 150000 if y < 2024 else 50000
    ppf.append((D(y, 1, 8), 'Deposit', None, dep, -dep)); bal += dep
    if D(y, 3, 31) < TODAY:
        i = round(bal * 0.071); ppf.append((D(y, 3, 31), 'Interest', None, i, 0)); bal += i
PPF_VALUE = round(bal)

# ---------------------------------------------------------------- XIRR tab
x = wb.create_sheet('XIRR')
x['I1'] = 0.0625; x['J1'] = 'FD rate'
header(x, 2, 1, ['Status', 'Date', 'Side', 'Company', 'Quantity', 'Price', 'Amount', 'XIRR'])
r = 3; TODAY_ROW = {}
def block(title, name, lots, unitsless=False, price_ref=None):
    global r
    x.cell(r, 1, title).font = BOLD; r += 1
    a = r
    for (d, side, q, p, amt) in lots:
        x.cell(r, 2, d).number_format = 'dd-mmm-yyyy'; x.cell(r, 3, side); x.cell(r, 4, name)
        if q is not None: x.cell(r, 5, q)
        x.cell(r, 6, p); x.cell(r, 7, amt); r += 1
    b = r - 1
    x.cell(r, 1, 'Active'); c = x.cell(r, 2, '=TODAY()'); c.number_format = 'dd-mmm-yyyy'; x.cell(r, 3, 'Today'); x.cell(r, 4, name)
    if unitsless:
        x.cell(r, 6, price_ref); x.cell(r, 7, f'=F{r}')
    else:
        x.cell(r, 5, f'=SUM(E{a}:E{b})'); x.cell(r, 6, price_ref); x.cell(r, 7, f'=E{r}*F{r}')
    x.cell(r, 8, f'=XIRR(G{a}:G{r},B{a}:B{r})').number_format = '0.00%'
    TODAY_ROW[name] = r; r += 2

# ---------------------------------------------------------------- Portfolio tab
p = wb.create_sheet('Portfolio', 1)
header(p, 1, 2, ['Vehicle', 'Name', 'Quantity', 'Buy Price', 'Current Price', 'Total Investment', 'Current Value', 'Difference', 'Return %'])
PROW = {}
row = 2
def ptable(label, names):
    global row
    first = row
    for i, n in enumerate(names):
        L = LOTS[n]
        if i == 0: p.cell(row, 2, label)
        p.cell(row, 3, n); p.cell(row, 4, L['units']); p.cell(row, 5, round(L['cost'] / L['units'], 4) if L['units'] else 0)
        p.cell(row, 6, L['price']); p.cell(row, 7, f'=D{row}*E{row}'); p.cell(row, 8, f'=D{row}*F{row}'); p.cell(row, 9, f'=H{row}-G{row}'); p.cell(row, 10, f'=IFERROR(I{row}/G{row},0)')
        PROW[n] = row; row += 1
    p.cell(row, 2, 'Total ' + label).font = BOLD
    for c, col in ((7, 'G'), (8, 'H'), (9, 'I')): p.cell(row, c, f'=SUM({col}{first}:{col}{row - 1})')
    p.cell(row, 10, f'=IFERROR(I{row}/G{row},0)'); TOT = row; row += 1
    return TOT
totals = []
totals.append(ptable('Mutual funds', ['Nifty 50 Index Fund - Direct - Growth', 'Flexi Cap Fund - Direct - Growth', 'Nasdaq 100 Fund of Funds - Direct', 'Arbitrage Fund - Direct - Growth']))
totals.append(ptable('REITs', ['Example REIT A', 'Example REIT B']))
totals.append(ptable('Stocks', ['Example Industries Ltd']))
totals.append(ptable('Commodity ETFs', ['Gold ETF', 'Silver ETF']))
totals.append(ptable('National pension scheme', ['NPS Scheme E - Tier I']))
totals.append(ptable('Sovereign gold bonds', ['SGB 2031 Series II', 'SGB 2032 Series IV']))
row += 1
header(p, row, 3, ['Commodity', 'Grams', 'Price / gram', 'Total', 'Making / gram', 'Total making', 'Total investment', 'SGST 1.5%', 'CGST 1.5%', 'With bill', 'Current price', 'Current value', 'Difference', 'Return %'])
row += 1; cfirst = row
for kind, n, making in (('Gold', '[22K] Gold bangles', 250), ('Silver', '[999K] Silver coins', 0)):
    L = LOTS[n]; g = L['lots'][0][2]; bill = -L['lots'][0][4]
    base = round((bill / 1.03 - making * g) / g, 2)
    p.cell(row, 2, kind); p.cell(row, 3, n); p.cell(row, 4, g); p.cell(row, 5, base); p.cell(row, 6, f'=D{row}*E{row}'); p.cell(row, 7, making); p.cell(row, 8, f'=D{row}*G{row}')
    p.cell(row, 9, f'=F{row}+H{row}'); p.cell(row, 10, f'=I{row}*1.5%'); p.cell(row, 11, f'=I{row}*1.5%'); p.cell(row, 12, f'=I{row}+J{row}+K{row}')
    p.cell(row, 13, L['price']); p.cell(row, 14, f'=D{row}*M{row}'); p.cell(row, 15, f'=N{row}-L{row}'); p.cell(row, 16, f'=IFERROR(O{row}/L{row},0)')
    PROW[n] = row; row += 1
p.cell(row, 2, 'Total Commodities').font = BOLD; p.cell(row, 12, f'=SUM(L{cfirst}:L{row - 1})'); p.cell(row, 14, f'=SUM(N{cfirst}:N{row - 1})'); p.cell(row, 15, f'=N{row}-L{row}'); CTOT = row; row += 2
# PPF
header(p, row, 3, ['Total invested', 'Current value', 'Difference', 'Return %']); row += 1
ppf_invested = sum(-l[4] for l in ppf if l[1] == 'Deposit')
p.cell(row, 2, 'PPF'); p.cell(row, 3, ppf_invested); p.cell(row, 4, PPF_VALUE); p.cell(row, 5, f'=D{row}-C{row}'); p.cell(row, 6, f'=E{row}/C{row}'); PPF_ROW = row; row += 2

# ---------------------------------------------------------------- EPF passbook (EPF tab) and its split on Portfolio
e = wb.create_sheet('EPF')
header(e, 1, 1, ['For month', 'Transaction date', 'Credit/Debit', 'EPF wages', 'EPS wages', 'Employee', 'Employer', 'Pension', 'Sum'])
er_, tot = 2, {'ee': 0, 'er': 0, 'pen': 0, 'iee': 0, 'ier': 0}
bal_ee = bal_er = 0
for d in months(D(2019, 7, 1), 88, 1):
    if d >= D(2026, 11, 1): break
    wage = 15000
    ee, pen = 1800, 1250; er = 1800 - pen
    cr = D(d.year + (1 if d.month == 12 else 0), 1 if d.month == 12 else d.month + 1, 14)
    e.cell(er_, 1, d).number_format = 'mmm-yyyy'; e.cell(er_, 2, cr).number_format = 'dd-mmm-yyyy'; e.cell(er_, 3, 'CR')
    for c, v in ((4, wage), (5, wage), (6, ee), (7, er), (8, pen)): e.cell(er_, c, v)
    e.cell(er_, 9, f'=F{er_}+G{er_}+H{er_}'); er_ += 1
    tot['ee'] += ee; tot['er'] += er; tot['pen'] += pen; bal_ee += ee; bal_er += er
    if d.month == 2:  # interest credited on 31 March for the year
        iee, ier = round(bal_ee * 0.0825), round(bal_er * 0.0825)
        e.cell(er_, 2, D(d.year, 3, 31)).number_format = 'dd-mmm-yyyy'; e.cell(er_, 3, 'DR'); e.cell(er_, 6, iee); e.cell(er_, 7, ier); e.cell(er_, 8, 0); e.cell(er_, 9, f'=F{er_}+G{er_}')
        er_ += 1; tot['iee'] += iee; tot['ier'] += ier; bal_ee += iee; bal_er += ier
EPF_EE, EPF_ER, EPF_PEN = bal_ee, bal_er, tot['pen']
header(p, row, 3, ['Employer', 'Employee', 'Employer', 'Pension']); row += 1
p.cell(row, 2, 'Employee provident fund'); p.cell(row, 3, 'Example Employer'); p.cell(row, 4, EPF_EE); p.cell(row, 5, EPF_ER); p.cell(row, 6, EPF_PEN); row += 1
p.cell(row, 2, 'Total Employee provident funds').font = BOLD; p.cell(row, 4, f'=SUM(D{row - 1}:F{row - 1})'); EPF_TOT = row; row += 2
# Total portfolio (market-linked holdings)
inv = '+'.join(f'G{t}' for t in totals) + f'+L{CTOT}'; val = '+'.join(f'H{t}' for t in totals) + f'+N{CTOT}'
p.cell(row, 2, 'Total Portfolio').font = BOLD; p.cell(row, 3, '=' + inv); p.cell(row, 4, '=' + val); p.cell(row, 5, f'=D{row}-C{row}'); p.cell(row, 6, f'=E{row}/C{row}'); row += 2
# Markets
header(p, row, 2, ['Entity', 'Close', 'Change', 'Change %']); p.cell(row, 6, 'Replace the numbers in this block with GOOGLEFINANCE formulas to make them live (see README)'); row += 1
MK = [('Nifty 50', 25620.4, 182.3, '0.72%'), ('Nasdaq 100', 25880.1, -210.6, '-0.81%'), None, ('USD', 88.62, 88.71, None), ('EUR', 102.85, 102.6, None), ('GBP', 116.9, 117.2, None), None,
      ('Gold - 24K', 125600, 1200, None), ('Gold - 22K', 115130, None, None), ('Silver - 999K', 151200, 149800, None), ('Crude oil', 63.45, None, None)]
for m in MK:
    if m:
        p.cell(row, 2, m[0]); p.cell(row, 3, m[1])
        if m[2] is not None: p.cell(row, 4, m[2])
        if m[3]: p.cell(row, 5, m[3])
        elif m[0] in ('USD', 'EUR', 'GBP'): p.cell(row, 5, f'=(C{row}-D{row})/D{row}')
    row += 1
p.cell(row, 2, 'Inflation'); p.cell(row, 3, D(2026, 10, 12)).number_format = 'dd-mmm-yyyy'; p.cell(row, 4, 0.0412); row += 1
p.cell(row, 3, D(2026, 11, 12)).number_format = 'dd-mmm-yyyy'; p.cell(row, 4, 0.0455); row += 2
p.cell(row, 2, 'Effective Date'); p.cell(row, 3, 'Repo Rate'); row += 1
p.cell(row, 2, D(2026, 10, 7)).number_format = 'dd-mmm-yyyy'; p.cell(row, 3, 0.055); row += 1
for c, w in zip('ABCDEFGHIJKLMNOP', (10, 26, 40, 12, 12, 12, 16, 16, 14, 12, 12, 12, 14, 14, 14, 10)): p.column_dimensions[c].width = w

# now the XIRR blocks, priced from Portfolio
for n in ['Nifty 50 Index Fund - Direct - Growth', 'Flexi Cap Fund - Direct - Growth', 'Nasdaq 100 Fund of Funds - Direct', 'Arbitrage Fund - Direct - Growth', 'Example Industries Ltd', 'Example REIT A', 'Example REIT B', 'Gold ETF', 'Silver ETF', 'NPS Scheme E - Tier I', 'SGB 2031 Series II', 'SGB 2032 Series IV']:
    block(n, n, LOTS[n]['lots'], price_ref=f'=Portfolio!F{PROW[n]}')
for n in ['[22K] Gold bangles', '[999K] Silver coins']:
    block(n, n, LOTS[n]['lots'], price_ref=f'=Portfolio!M{PROW[n]}')
block('Public provident fund', 'Public provident fund', ppf, unitsless=True, price_ref=f'=Portfolio!D{PPF_ROW}')
for c, w in zip('ABCDEFGH', (12, 13, 10, 36, 12, 12, 14, 10)): x.column_dimensions[c].width = w

# ---------------------------------------------------------------- Loan schedule
L = wb.create_sheet('Loan schedule')
SANCT, RATE0, TENURE = 7500000, 0.085, 240
first = D(2023, 5, 5)
meta = [('Loan details', None), ('Provisioning bank', 'Example Bank'), ('Sanctioned amount', SANCT), ('Disbursed amount', SANCT), ('Tenure (month)', TENURE), ('First EMI start date', first),
        ('Loan sanctioned date', D(2023, 4, 2)), ('EMI Date', '5th of every month'), ('Current RoI', None), ('Remaining tenure (monthy)', None), ('EMI amount', None), ('NACH debit', None), ('NACH debit from', 'Example Bank')]
r0 = 1
for k, v in meta:
    L.cell(r0, 1, k).font = BOLD if k == 'Loan details' else Font()
    if v is not None: L.cell(r0, 2, v)
    r0 += 1
header(L, 15, 1, ['Post Date', 'Value Date', 'Details', 'Debit', 'Credit', 'Balance'])
lr = 16
def line(d, t, dr=None, cr=None, bal=None):
    global lr
    L.cell(lr, 1, d).number_format = 'dd-mmm-yyyy'; L.cell(lr, 2, d).number_format = 'dd-mmm-yyyy'; L.cell(lr, 3, t)
    if dr is not None: L.cell(lr, 4, round(dr, 2))
    if cr is not None: L.cell(lr, 5, round(cr, 2))
    if bal is not None: L.cell(lr, 6, round(bal, 2))
    lr += 1
bal, rate = float(SANCT), RATE0
mr = rate / 12; emi = round(SANCT * mr * (1 + mr) ** TENURE / ((1 + mr) ** TENURE - 1))
line(D(2023, 4, 2), 'ADVANCE:LOANDISBURSAL PRINCIPAL', dr=SANCT, bal=bal)
d = first; changes = {D(2025, 2, 10): 0.0825, D(2025, 6, 15): 0.08}; prepay = {D(2025, 1, 14): 500000, D(2026, 9, 18): 300000}
while d < TODAY:
    for cd, nr in list(changes.items()):
        if cd < d: line(cd, f'RATECHANGEDFM {rate * 100:.2f}% TO {nr * 100:.2f}%'); rate = nr; del changes[cd]
    for pd, amt in list(prepay.items()):
        if pd < d: bal -= amt; line(pd, 'PRINCIPALREPAYMENTGLTO LOANINB', cr=amt, bal=bal); del prepay[pd]
    i = bal * rate / 12; pr = emi - i
    line(D(d.year, d.month, 1) - dt.timedelta(days=1), 'INTEREST', dr=i, bal=bal + i)
    bal -= pr
    line(d, 'PRINCIPALREPAYMENTGLTO LOANACH', cr=pr, bal=bal); line(d, 'INTERESTREPAYMENTGLTO LOANACH', cr=i, bal=bal)
    d = D(d.year + (1 if d.month == 12 else 0), 1 if d.month == 12 else d.month + 1, 5)
LOAN_BAL = round(bal)
L['B9'] = rate; L['B9'].number_format = '0.00%'; L['B11'] = emi
n_left = 0; b2 = bal
while b2 > 1 and n_left < 600: b2 -= emi - b2 * rate / 12; n_left += 1
L['B10'] = n_left; L['B12'] = emi
for c, w in zip('ABCDEF', (24, 18, 36, 14, 14, 16)): L.column_dimensions[c].width = w

# ---------------------------------------------------------------- Net worth tab (holdings)
w = wb.create_sheet('Net worth', 0)
header(w, 1, 1, ['L1: Group', 'L2: Class', 'L3: Category', 'Custodian', 'L4: Vehicle', 'L5: Instrument / Ticker', 'Amount: T', 'Amount: T - X', 'Absolute', '%', 'XIRR', '> FD?', '> CPI?', 'Taxation', 'Liquidity', 'Function', 'Goal / Comments'])
def xr(n): return f'=XIRR!H{TODAY_ROW[n]}'
def pf(n, col='H'): return f'=Portfolio!{col}{PROW[n]}'
H = [
 ('Asset', 'Cash', 'Cash equivalent', 'Example Bank', 'Savings account', 'Primary savings', 462500, 455000, 0.03, 'Slab rate', 'Tier 1 (Instant)', 'Operational', 'Monthly expenses'),
 ('Asset', 'Cash', 'Cash equivalent', 'Example Bank', 'Rewards', 'Credit card points', '=(1240*25)/100', 290, 0, 'Tax free', 'Tier 1 (Instant)', 'Operational', 'Expenses'),
 ('Asset', 'Cash', 'Cash equivalent', 'Example Store', 'Voucher', 'Gift card', 2000, 2000, 0, 'Tax free', 'Tier 1 (Instant)', 'Operational', 'Gifts [30 Jun 2027]'),
 ('Asset', 'Fixed income', 'Cash equivalent', 'Example Bank', 'Term deposit', 'Emergency deposit', 600000, 600000, 0.068, 'Slab rate', 'Tier 1 (Instant)', 'Safety', 'Emergency fund'),
 ('Asset', 'Fixed income', 'Statutory', 'EPFO', 'EPF account', 'Employee Provident Fund', f'=Portfolio!D{EPF_TOT}', None, 0.0825, 'Tax free', 'Tier 3 (Locked)', 'Core', 'Retirement'),
 ('Asset', 'Fixed income', 'Statutory', 'Example Bank', 'PPF account', 'Public Provident Fund', f'=Portfolio!D{PPF_ROW}', None, '=XIRR!H%d' % TODAY_ROW['Public provident fund'], 'Tax free', 'Tier 3 (Locked)', 'Core', 'Retirement'),
 ('Asset', 'Equities', 'Statutory', 'NPS Trust', 'NPS account', 'NPS Scheme E - Tier I', pf('NPS Scheme E - Tier I'), None, xr('NPS Scheme E - Tier I'), 'Tax free', 'Tier 3 (Locked)', 'Core', 'Retirement'),
 ('Asset', 'Equities', 'Domestic index', 'Example Broker', 'Mutual fund', 'Nifty 50 Index Fund - Direct - Growth', pf('Nifty 50 Index Fund - Direct - Growth'), None, xr('Nifty 50 Index Fund - Direct - Growth'), 0.125, 'Tier 2 (T+X)', 'Core', 'Retirement'),
 ('Asset', 'Equities', 'Multi cap', 'Example Broker', 'Mutual fund', 'Flexi Cap Fund - Direct - Growth', pf('Flexi Cap Fund - Direct - Growth'), None, xr('Flexi Cap Fund - Direct - Growth'), 0.125, 'Tier 2 (T+X)', 'Core', 'Wealth'),
 ('Asset', 'Equities', 'International index', 'Example Broker', 'Mutual fund', 'Nasdaq 100 Fund of Funds - Direct', pf('Nasdaq 100 Fund of Funds - Direct'), None, xr('Nasdaq 100 Fund of Funds - Direct'), 'Slab rate', 'Tier 2 (T+X)', 'Satellite', 'Wealth'),
 ('Asset', 'Equities', 'Arbitrage', 'Example Broker', 'Mutual fund', 'Arbitrage Fund - Direct - Growth', pf('Arbitrage Fund - Direct - Growth'), None, xr('Arbitrage Fund - Direct - Growth'), 0.125, 'Tier 2 (T+X)', 'Safety', 'Emergency fund'),
 ('Asset', 'Equities', 'Small cap', 'Example Broker', 'Stocks', 'Example Industries Ltd', pf('Example Industries Ltd'), None, xr('Example Industries Ltd'), 0.125, 'Tier 2 (T+X)', 'Satellite', 'Wealth'),
 ('Asset', 'Real estate', 'Commercial', 'Example Broker', 'REIT', 'Example REIT A', pf('Example REIT A'), None, xr('Example REIT A'), 0.125, 'Tier 2 (T+X)', 'Core', 'Income'),
 ('Asset', 'Real estate', 'Commercial', 'Example Broker', 'REIT', 'Example REIT B', pf('Example REIT B'), None, xr('Example REIT B'), 0.125, 'Tier 2 (T+X)', 'Core', 'Income'),
 ('Asset', 'Commodities', 'Gold (Hedge)', 'Example Broker', 'ETF', 'Gold ETF', pf('Gold ETF'), None, xr('Gold ETF'), 0.125, 'Tier 2 (T+X)', 'Ballast', 'Wealth'),
 ('Asset', 'Commodities', 'Silver (Hedge)', 'Example Broker', 'ETF', 'Silver ETF', pf('Silver ETF'), None, xr('Silver ETF'), 0.125, 'Tier 2 (T+X)', 'Ballast', 'Wealth'),
 ('Asset', 'Commodities', 'Gold (Hedge)', 'Example Broker', 'Direct bond', 'SGB 2031 Series II', pf('SGB 2031 Series II'), None, xr('SGB 2031 Series II'), 'Tax free', 'Tier 2 (T+X)', 'Ballast', 'Wealth'),
 ('Asset', 'Commodities', 'Gold (Hedge)', 'Example Broker', 'Direct bond', 'SGB 2032 Series IV', pf('SGB 2032 Series IV'), None, xr('SGB 2032 Series IV'), 'Tax free', 'Tier 2 (T+X)', 'Ballast', 'Wealth'),
 ('Asset', 'Commodities', 'Gold (Luxury)', 'Bank locker', 'Physical gold', '[22K] Gold bangles', pf('[22K] Gold bangles', 'N'), None, xr('[22K] Gold bangles'), 'Slab rate', 'Tier 2 (T+X)', 'Ballast', 'Family'),
 ('Asset', 'Commodities', 'Silver (Luxury)', 'Bank locker', 'Physical silver', '[999K] Silver coins', pf('[999K] Silver coins', 'N'), None, xr('[999K] Silver coins'), 'Slab rate', 'Tier 2 (T+X)', 'Ballast', 'Family'),
 ('Asset', 'Fixed income', 'Private loan', 'A friend', 'Direct contract', 'Personal loan', 100000, 100000, 0, 'Tax free', 'Tier 2 (T+X)', 'Core', 'Due back by Mar 2027'),
 ('Asset', 'Real estate', 'Residential', 'Example Bank', 'Direct purchase', 'Apartment', 13800000, 13800000, 0.052, 0.125, 'Tier 3 (Locked)', 'Core', 'Home loan hedge'),
 ('Liability', 'Real estate', 'Term Loan', 'Example Bank', 'Bank debt', 'Home loan', LOAN_BAL, LOAN_BAL, 0.08, 'Tax free', 'Tier 2 (T+X)', 'Liability', None),
 ('Asset', 'Equities', 'Emp Comp', 'Example Employer', 'Private equity', 'Stock appreciation rights', 650000, 650000, 0, 'Slab rate', 'Tier 3 (Locked)', 'Satellite', 'Vests 2028; tracked, not counted'),
]
for i, h in enumerate(H):
    rr = i + 2
    g, cls, cat, cust, veh, inst, val, prev, xi, tax, liq, fn, note = h
    for c, v in enumerate([g, cls, cat, cust, veh, inst], 1): w.cell(rr, c, v)
    w.cell(rr, 7, val); w.cell(rr, 8, prev if prev is not None else f'=G{rr}*0.995')
    w.cell(rr, 9, f'=G{rr}-H{rr}'); w.cell(rr, 10, f'=IFERROR(I{rr}/H{rr},0)'); w.cell(rr, 11, xi)
    w.cell(rr, 12, f'=IF(K{rr}>6.25%,"Yes","No")'); w.cell(rr, 13, f'=IF(K{rr}>4.55%,"Yes","No")')
    w.cell(rr, 14, tax); w.cell(rr, 15, liq); w.cell(rr, 16, fn)
    if note: w.cell(rr, 17, note)
    for c in (7, 8, 9): w.cell(rr, c).number_format = '₹#,##0'
    w.cell(rr, 11).number_format = '0.00%'
last = len(H) + 1
w.cell(last + 2, 1, 'Last updated').font = BOLD; w.cell(last + 2, 3, D(2026, 11, 1)).number_format = 'dd-mmm-yyyy'
w.cell(last + 3, 1, 'Target net worth by 31st Mar 2028').font = BOLD; w.cell(last + 3, 3, 40000000).number_format = '₹#,##0'
for c, wd in zip('ABCDEFGHIJKLMNOPQ', (10, 13, 18, 18, 16, 36, 14, 14, 12, 8, 9, 8, 8, 11, 16, 12, 30)): w.column_dimensions[c].width = wd
w.freeze_panes = 'A2'

# ---------------------------------------------------------------- Expenses (this year and last)
CATS = [
 ('Household & utilities', [('Home loan EMI - Example Bank', 'fixed', emi), ('Groceries', 'var', 13000), ('Electricity', 'var', 3200), ('Internet and phone', 'fixed', 1500), ('Society maintenance', 'fixed', 4500), ('Property tax', 'once:5', 14800)]),
 ('Household services', [('House help', 'fixed', 6000), ('Repairs and upkeep', 'var', 2500)]),
 ('Insurances & AMCs', [('Medical - Lotus Health', 'once:1', 31200), ('Term - Banyan Life', 'once:7', 17400), ('Car service', 'once:4', 7800)]),
 ('Wellness & personal care', [('Gym', 'fixed', 2500), ('Grooming', 'var', 1400)]),
 ('Commute', [('Fuel', 'var', 4000), ('Cabs and metro', 'var', 2800)]),
 ('Entertainment & lifestyle', [('Dining out', 'var', 6200), ('Subscriptions', 'fixed', 1200), ('Shopping', 'var', 5200), ('Travel', 'once:6', 68000), ('Gifts & donations', 'var', 2500)]),
]
MONTHS_FY = ['April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'January', 'February', 'March']
def expenses(name, upto, scale):
    s = wb.create_sheet(name)
    header(s, 1, 1, ['Particulars', None] + MONTHS_FY + ['Total', 'Monthly Avg'])
    rr = 2
    for cat, items in CATS:
        for j, (item, kind, amt) in enumerate(items):
            if j == 0: s.cell(rr, 1, cat)
            s.cell(rr, 2, item)
            for m in range(12):
                v = None
                if kind == 'fixed': v = round(amt * scale)
                elif kind == 'var' and m <= upto: v = round(amt * scale * random.uniform(0.6, 1.4) * (0.55 if m == upto and upto < 11 else 1))
                elif kind.startswith('once:') and m == int(kind[5:]): v = round(amt * scale) if m <= upto or 'Term' in item or 'Property' in item else None
                if v is not None: s.cell(rr, 3 + m, v)
            s.cell(rr, 15, f'=SUM(C{rr}:N{rr})'); s.cell(rr, 16, f'=IFERROR(AVERAGE(C{rr}:N{rr}),0)'); rr += 1
    s.cell(rr, 1, 'Total').font = BOLD
    for m in range(12):
        col = openpyxl.utils.get_column_letter(3 + m); s.cell(rr, 3 + m, f'=SUM({col}2:{col}{rr - 1})')
    s.cell(rr, 15, f'=SUM(C{rr}:N{rr})'); s.cell(rr, 16, f'=AVERAGE(C{rr}:N{rr})')
    for c, wd in zip('ABCDEFGHIJKLMNOP', [24, 30] + [10] * 12 + [12, 12]): s.column_dimensions[c].width = wd
expenses('Expenses FY 26 - 27', 7, 1.0)
expenses('Expenses FY 25 - 26', 11, 0.93)

# ---------------------------------------------------------------- CoI: rename this year's tab to match, add last year's
wb['CoI FY 26-27'].title = 'CoI FY 26 - 27'
c = wb['CoI FY 26 - 27']
# payslips up to October, with the register's TDS matching the year
for col in range(3, 10):  # C..I = Apr..Oct
    for rr, v in ((35, 90000), (36, 45000), (37, 60000), (41, 1800), (42, 200), (43, 20000)):
        c.cell(rr, col, v)
c.cell(38, 3, 150000); c.cell(43, 3, 92000)
for col in range(4, 10): c.cell(38, col, 0)
c['D1'] = 'New Regime'
# a short-term gain and lighter monthly TDS, so advance tax applies; two instalments paid, the second a little short
for col in ('D', 'E'):
    c[f'{col}10'] = 120000; c[f'{col}26'] = 7000; c[f'{col}27'] = 15000
c['C41'] = 1800
# last year's computation, in the older layout (register at the top, then the income heads)
pc = wb.create_sheet('CoI FY 25 - 26')
pc.cell(1, 1, 'Wage type').font = BOLD
for m in range(12):
    y = 2025 if m < 9 else 2026; mo = (m + 3) % 12 + 1
    pc.cell(1, 2 + m, D(y, mo, 1)).number_format = 'mmm-yy'
rows = [('Income', None), ('Performance bonus', [140000] + [0] * 11), ('Basic', [85000] * 12), ('HRA', [42500] * 12), ('Special allowance', [55000] * 12), (None, None),
        ('Gross monthly salary', 'gross'), ('Deductions', None), ('Income tax', [40000] + [22000] * 11), ('Employee provident fund', [1800] * 12), ('Professional tax', [200] * 11 + [300]), (None, None),
        ('Gross monthly tax', 'ded'), (None, None), ('Net monthly salary', 'net')]
rr = 2
for label, vals in rows:
    if label: pc.cell(rr, 1, label)
    if isinstance(vals, list):
        for m, v in enumerate(vals): pc.cell(rr, 2 + m, v)
    elif vals == 'gross':
        for m in range(12): col = openpyxl.utils.get_column_letter(2 + m); pc.cell(rr, 2 + m, f'=SUM({col}3:{col}6)')
        GROSS_ROW = rr
    elif vals == 'ded':
        for m in range(12): col = openpyxl.utils.get_column_letter(2 + m); pc.cell(rr, 2 + m, f'=SUM({col}10:{col}12)')
        DED_ROW = rr
    elif vals == 'net':
        for m in range(12): col = openpyxl.utils.get_column_letter(2 + m); pc.cell(rr, 2 + m, f'={col}{GROSS_ROW}-{col}{DED_ROW}')
    rr += 1
rr += 1
salary = 140000 + (85000 + 42500 + 55000) * 12; tds = 40000 + 22000 * 11
gross_inc = salary + 9200 + 8400 + 4300
std = 75000; taxable = gross_inc - std
slab = 0
for lo, hi, rt in ((400000, 800000, .05), (800000, 1200000, .1), (1200000, 1600000, .15), (1600000, 2000000, .2), (2000000, 2400000, .25), (2400000, 10 ** 12, .3)):
    if taxable > lo: slab += (min(taxable, hi) - lo) * rt
cess = slab * 0.04; total = slab + cess; sat = max(0, round(total - tds))
block2 = [('Income heads', 'Entity', 'Amount'), ('Salary', 'Example Employer', salary), ('Savings account', 'Example Bank', 9200), ('Other income', 'Sovereign gold bonds', 8400), (None, 'REIT dividends', 4300),
          ('Gross income [A]', None, gross_inc), (None, None, None), ('Section', 'Exemption', 'Amount'), ('16(ia)', 'Standard deductions', std), ('Total exemptions [B]', None, std), (None, None, None),
          ('Total taxable income', '[C = A-B]', taxable), ('Applicable income tax [D]', None, round(slab)), ('CESS 4% [F]', None, round(cess)), ('Total income tax [G = D+E+F]', None, round(total)),
          ('TDS paid [H]', None, tds), ('Self assessment tax paid [I]', None, sat), ('Balance tax [J = G-(H+I)]', None, round(total - tds - sat)), ('Income to tax ratio [A/G]', None, round(total / gross_inc, 4)), ('CA fees [ITR]', None, 3500)]
for a, b, v in block2:
    if a: pc.cell(rr, 1, a).font = BOLD if a in ('Income heads', 'Section') else Font()
    if b: pc.cell(rr, 2, b)
    if v is not None: pc.cell(rr, 3, v)
    rr += 1
pc.column_dimensions['A'].width = 30; pc.column_dimensions['B'].width = 22

# ---------------------------------------------------------------- Payouts: more history on the same layout
po = wb['Dividends & payouts']
coupon = {'SGB 2031 Series II': (D(2023, 9, 18), 30 * 5923 * 0.025 / 2), 'SGB 2032 Series IV': (D(2024, 4, 19), 20 * 6213 * 0.025 / 2)}
for rr in range(3, 40): po.cell(rr, 1).value = None; po.cell(rr, 2).value = None; po.cell(rr, 3).value = None
pay = []
for col, (n, (start, amt)) in zip((2, 3), coupon.items()):
    d = D(start.year, start.month, start.day)
    while True:
        d = D(d.year + (1 if d.month > 6 else 0), (d.month + 6 - 1) % 12 + 1, min(d.day, 28))
        if d >= TODAY: break
        pay.append((d, col, round(amt, 2)))
pay.sort()
for i, (d, col, a) in enumerate(pay):
    po.cell(3 + i, 1, d).number_format = 'dd-mmm-yyyy'; po.cell(3 + i, col, a)
# Stock dividends: three companies across this year and last
for rr in range(3, 40):
    for cc in range(17, 21): po.cell(rr, cc).value = None
po.cell(2, 19, 'Example Pharma Ltd'); po.cell(2, 20, 'Example Tech Ltd')
DIVS = [(D(2025, 6, 20), 18, 1500), (D(2025, 8, 14), 19, 2400), (D(2025, 11, 10), 20, 900), (D(2026, 2, 12), 19, 1200),
        (D(2026, 6, 18), 18, 1800), (D(2026, 7, 25), 19, 2600), (D(2026, 8, 22), 20, 1100), (D(2026, 10, 30), 18, 650), (D(2026, 11, 12), 20, 450)]
for i, (d, col, a) in enumerate(DIVS):
    po.cell(3 + i, 17, d).number_format = 'dd-mmm-yyyy'; po.cell(3 + i, col, a)
# REIT distributions: add this year's third quarter and last year's fourth, on the same block layout
def reit_block(r0, q, paid, pu_a, pu_b):
    po.cell(r0, 4, q); po.cell(r0, 5, paid).number_format = 'dd-mmm-yyyy'
    rows = [('Per unit', pu_a, pu_b), ('Units held', [500] * 4, [400] * 4)]
    hold_a = [round(x * 500, 2) for x in pu_a]; hold_b = [round(x * 400, 2) for x in pu_b]
    tds_a = [round(hold_a[0] * .1, 2), 0, 0, round(hold_a[3] * .1, 2)]; tds_b = [round(hold_b[0] * .1, 2), round(hold_b[1] * .1, 2), 0, 0]
    rows += [('For your holding', hold_a, hold_b), ('TDS rate', [.1, 0, 0, .1], [.1, .1, 0, 0]), ('TDS amount', tds_a, tds_b), ('Net distribution', [h - t for h, t in zip(hold_a, tds_a)], [h - t for h, t in zip(hold_b, tds_b)])]
    for k, (lab, a, b) in enumerate(rows):
        po.cell(r0 + k, 6, lab)
        for j in range(4): po.cell(r0 + k, 7 + j, a[j]); po.cell(r0 + k, 12 + j, b[j])
        if lab in ('For your holding', 'TDS amount', 'Net distribution'): po.cell(r0 + k, 11, round(sum(a), 2)); po.cell(r0 + k, 16, round(sum(b), 2))
reit_block(17, 'FY 26-27 Q3', D(2026, 10, 28), [2.4, 1.8, 1.3, 0.1], [1.1, 3.0, 2.3, 0])
reit_block(24, 'FY 25-26 Q4', D(2026, 2, 18), [2.1, 1.9, 1.2, 0.1], [1.0, 2.9, 2.2, 0])

# ---------------------------------------------------------------- Rewards ledger
rl = wb.create_sheet('Rewards ledger')
header(rl, 1, 1, ['Provider', 'Transaction date', 'Earned points', 'Multiplier', 'Amount', 'Expiry'])
rr = 2
for d in months(D(2024, 2, 1), 34, 12):
    if d >= TODAY: break
    pts = random.choice([120, 180, 240, 310, 90, 400])
    rl.cell(rr, 1, 'Example Bank'); rl.cell(rr, 2, d).number_format = 'dd-mmm-yyyy'; rl.cell(rr, 3, pts); rl.cell(rr, 4, 0.25); rl.cell(rr, 5, f'=C{rr}*D{rr}')
    rl.cell(rr, 6, f'=EDATE(B{rr},36)').number_format = 'dd-mmm-yyyy'; rr += 1
    if d.month in (6, 12):
        rl.cell(rr, 1, 'Example Bank'); rl.cell(rr, 2, d + dt.timedelta(days=9)).number_format = 'dd-mmm-yyyy'; rl.cell(rr, 3, -1500); rl.cell(rr, 4, 0.25); rl.cell(rr, 5, f'=C{rr}*D{rr}')
        rl.cell(rr, 6, f'=EDATE(B{rr},36)').number_format = 'dd-mmm-yyyy'; rr += 1
for c, wd in zip('ABCDEF', (16, 16, 14, 11, 10, 14)): rl.column_dimensions[c].width = wd

# ---------------------------------------------------------------- Income & estimations: insurance block
ie = wb.create_sheet('Income & estimations')
ie.cell(1, 1, 'Insurance').font = BOLD
header(ie, 2, 1, ['Person', 'Category', 'Type', 'Cover', 'Insurer'])
INS = [('Self', 'Term insurance', 'Base', 15000000, 'Banyan Life'), (None, None, 'Group', 5000000, 'Employer group cover'), (None, 'Health insurance', 'Base', 1000000, 'Lotus Health'),
       (None, None, 'Top-up', 4000000, 'Lotus Health'), ('Spouse', 'Health insurance', 'Floater', 1000000, 'Lotus Health'), (None, None, 'Group', 500000, 'Employer group cover'),
       ('Parent', 'Health insurance', 'Base', 500000, 'Lotus Health'), (None, 'Emergency fund', 'Fixed', 500000, 'Fixed deposit')]
for i, rowv in enumerate(INS):
    for cidx, v in enumerate(rowv, 1):
        if v is not None: ie.cell(3 + i, cidx, v)
    ie.cell(3 + i, 4).number_format = '₹#,##0'
for c, wd in zip('ABCDE', (12, 18, 10, 14, 22)): ie.column_dimensions[c].width = wd

# ---------------------------------------------------------------- Insurance premiums: one line a year per policy, two tables side by side
ip = wb.create_sheet('Insurance premiums')
IP_HEAD = ['Person', 'Insurance type', 'Provider', 'Premium amount', 'Premium Date', 'Coverage amount', 'Start date', 'End date', 'Status']
def policy(person, typ, provider, first, n, prem0, growth, cover, upcoming_amount=True):
    rows = []
    for i in range(n + 1):
        d = D(first.year + i, first.month, first.day)
        if i == n and d <= TODAY: continue
        amt = round(prem0 * (1 + growth) ** i) if (i < n or upcoming_amount) else None
        rows.append((person, typ, provider, amt, d, cover, d + dt.timedelta(days=1), D(d.year + 1, d.month, d.day)))
    return rows
LEFT = policy('Self', 'Health - Base', 'Lotus Health', D(2021, 4, 1), 6, 9800, 0.08, 1000000) + \
       policy('Self', 'Health - Top-up', 'Lotus Health', D(2022, 4, 1), 5, 2100, 0.06, 4000000) + \
       policy('Spouse', 'Health - Floater', 'Lotus Health', D(2021, 7, 15), 6, 11000, 0.07, 1000000)
RIGHT = policy('Self', 'Term', 'Banyan Life', D(2020, 12, 12), 6, 17400, 0.0, 15000000, upcoming_amount=False) + \
        policy('Parent', 'Health - Base', 'Lotus Health', D(2021, 2, 20), 6, 14500, 0.12, 500000)
for c0, rows in ((1, LEFT), (11, RIGHT)):
    header(ip, 1, c0, IP_HEAD)
    for i, r in enumerate(rows):
        rr = 2 + i
        for k, v in enumerate(r):
            if v is None: continue
            cell = ip.cell(rr, c0 + k, v)
            if k in (4, 6, 7): cell.number_format = 'dd-mmm-yyyy'
            if k in (3, 5): cell.number_format = '₹#,##0'
        dcol = openpyxl.utils.get_column_letter(c0 + 4)
        ip.cell(rr, c0 + 8, f'=IF({dcol}{rr}<=TODAY(), "Paid", "Upcoming")')
for k in range(1, 20): ip.column_dimensions[openpyxl.utils.get_column_letter(k)].width = 14

# ---------------------------------------------------------------- Locker log: a fuller example on the same layout
lk = wb['Locker log']
lk.delete_rows(1, lk.max_row)
meta = [('Bank', 'Example Bank'), ('Branch', 'Example branch'), ('Locker number', '000'), ('Locker size', 'Medium'), ('Access type', 'Either hirer'), ('Access limits', '12 visits a year'), ('Annual rent', 3540), ('Allotment date', D(2021, 5, 1))]
for i, (k, v) in enumerate(meta, 1): lk.cell(i, 1, k); lk.cell(i, 2, v)
header(lk, 10, 1, ['Access date', 'Accessed by', 'Holdings before access', 'Holdings after access', 'Transaction (deposited/withdrawn)'])
V1 = ['Gold bangles', 'Silver coins', 'Property papers', 'Family photo albums']
V2 = V1 + ['Foreign currency (USD)', 'Silver puja set']
V3 = V2 + ['Gold chain', 'Will and nominations', 'Spare house keys']
V4 = [x for x in V3 if x != 'Silver coins'] + ['Old passports']
for i, (d, by, before, after, txn) in enumerate([(D(2024, 8, 10), 'Holder 1', 'Empty', V1, 'Deposited: ' + ', '.join(V1)), (D(2025, 3, 14), 'Holder 2', V1, V2, 'Deposited: Foreign currency (USD), Silver puja set'), (D(2026, 6, 2), 'Holder 1 and Holder 2', V2, V3, 'Deposited: Gold chain, Will and nominations, Spare house keys'), (D(2026, 10, 9), 'Holder 2', V3, V4, 'Withdrawn: Silver coins. Deposited: Old passports')]):
    rr = 11 + i
    lk.cell(rr, 1, d).number_format = 'dd-mmm-yyyy'; lk.cell(rr, 2, by); lk.cell(rr, 3, before if isinstance(before, str) else '\n'.join(before)); lk.cell(rr, 4, '\n'.join(after)); lk.cell(rr, 5, txn)
    for col in (3, 4, 5): lk.cell(rr, col).alignment = Alignment(wrap_text=True, vertical='top')

# ---------------------------------------------------------------- Subscriptions: one row per payment, plans, memberships, gifts and one-off buys
sb = wb.create_sheet('Subscriptions')
header(sb, 1, 1, ['Entity', 'Enrollment', 'Plan type', 'Frequency', 'Billing/Premium date', 'Subscription start', 'Subscription expiry', 'Base amount', 'Tax', 'Total', 'Status'])
SUBS = []
def addm(d, n): y, m = divmod(d.month - 1 + n, 12); return D(d.year + y, m + 1, min(d.day, 28))
def plan(entity, kind, ptype, freq, first, last, price, status, gst=True, raise_at=None):
    step = {'Monthly': 1, 'Quarterly': 3, 'Annual': 12}[freq]; d = first
    while d <= last:
        p = price(d) if callable(price) else price
        end = addm(d, step) - dt.timedelta(days=1)
        SUBS.append((entity, kind, ptype, freq, d, d, end, p, 'gst' if gst else 0, status)); d = addm(d, step)
# Ongoing plans, with price rises along the way
plan('Music streaming', 'Subscription', 'Family', 'Monthly', D(2022, 6, 5), D(2026, 11, 5), lambda d: 126.27 if d >= D(2024, 7, 1) else 100.85, 'Ongoing')
plan('Password manager', 'Subscription', 'Personal', 'Monthly', D(2025, 1, 24), D(2026, 10, 24), 100, 'Ongoing')
plan('Video streaming', 'Subscription', 'Premium', 'Annual', D(2021, 3, 12), D(2026, 3, 12), lambda d: 1524.58 if d >= D(2024, 1, 1) else 1270.34, 'Ongoing')
plan('Office suite', 'Subscription', 'Personal', 'Annual', D(2026, 2, 14), D(2026, 2, 14), 4000, 'Ongoing')
plan('Domain name', 'Subscription', 'example.in', 'Annual', D(2019, 5, 1), D(2026, 5, 1), lambda d: 1100 if d >= D(2025, 1, 1) else 860, 'Ongoing')
SUBS.append(('Trading account', 'Subscription', 'Equity', 'Quarterly', D(2026, 1, 11), D(2025, 12, 29), D(2026, 3, 29), 75, 'gst', 'Ongoing'))
SUBS.append(('Trading account', 'Subscription', 'Equity', 'Annual', D(2026, 4, 10), D(2026, 3, 30), D(2027, 3, 29), 300, 'gst', 'Ongoing'))
# Ended plans
plan('Cloud storage', 'Subscription', '200 GB', 'Annual', D(2022, 10, 16), D(2025, 10, 16), 1100, 'Cancelled')
plan('News app', 'Subscription', 'Digital', 'Monthly', D(2018, 1, 8), D(2018, 12, 8), 83.90, 'Cancelled')
for d, amt in ((D(2022, 8, 25), 10100), (D(2023, 8, 25), 10610), (D(2024, 8, 25), 11150)):
    SUBS.append(('City swimming pool', 'Membership', 'General', 'Annual', d, d, addm(d, 12) - dt.timedelta(days=1), amt, 0, 'Cancelled'))
# Gifts and one-off purchases
for d, amt in ((D(2019, 2, 17), 500), (D(2021, 7, 30), 1000), (D(2023, 6, 11), 1500), (D(2025, 6, 3), 2500)):
    SUBS.append(('Open knowledge fund', 'Donation', 'Ad hoc', 'One time', d, 'N/A', 'N/A', amt, 0, 'N/A'))
for d, amt, what in ((D(2023, 9, 23), 505, 'Prasad seva'), (D(2024, 9, 8), 625, 'Naivedya seva'), (D(2025, 8, 31), 555, 'Prasad seva'), (D(2026, 9, 17), 805, 'Annadana seva')):
    SUBS.append(('Neighbourhood temple', 'Seva', what, 'Annual', d, 'N/A', 'N/A', amt, 0, 'Ongoing'))
SUBS.append(('Photo editor Pro', 'App purchase', 'Ad hoc', 'One time', D(2022, 12, 4), 'N/A', 'N/A', 83.90, 'gst', 'N/A'))
for i, (e, k, pt, f, bd, st, ex, base, tax, status) in enumerate(SUBS):
    rr = 2 + i
    for c, v in enumerate((e, k, pt, f, bd, st, ex, base), 1):
        cell = sb.cell(rr, c, v)
        if isinstance(v, dt.datetime): cell.number_format = 'dd-mmm-yyyy'
    sb.cell(rr, 8).number_format = '₹#,##0.00'
    sb.cell(rr, 9, f'=ROUND(H{rr}*18%,2)' if tax == 'gst' else 0).number_format = '₹#,##0.00'
    sb.cell(rr, 10, f'=IF(SUM(H{rr}:I{rr})<>0,SUM(H{rr}:I{rr}),"")').number_format = '₹#,##0.00'
    sb.cell(rr, 11, status)
sb.freeze_panes = 'A2'
for c, wd in zip('ABCDEFGHIJK', (24, 14, 16, 11, 14, 14, 14, 12, 10, 12, 11)): sb.column_dimensions[c].width = wd

order = ['Net worth', 'Portfolio', 'XIRR', 'Expenses FY 26 - 27', 'Expenses FY 25 - 26', 'CoI FY 26 - 27', 'CoI FY 25 - 26', 'Dividends & payouts', 'EPF', 'Loan schedule', 'Rewards ledger', 'Income & estimations', 'Insurance premiums', 'Subscriptions', 'Locker log']
wb._sheets = [wb[n] for n in order]
wb.save(out)
json.dump({'loan_bal': LOAN_BAL, 'emi': emi, 'rate': rate}, open(out + '.meta.json', 'w'))
print('saved', out, 'loan', LOAN_BAL, 'emi', emi)
