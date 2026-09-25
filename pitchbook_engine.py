# %% [markdown]
# # Sell-Side Valuation & Pitchbook Engine
#
# **One ticker in → a banker-style valuation package out.**
#
# This notebook automates the core analyst workflow on a public M&A target:
#
# | Module | Output |
# |---|---|
# | 1. Market data | Price, capital structure, LTM financials, 5 years of history (Yahoo Finance) |
# | 2. Trading comparables | EV/Revenue, EV/EBITDA, P/E with quartiles and implied share price |
# | 3. Precedents & premiums paid | Implied control value from precedent multiples and historical takeover premia |
# | 4. WACC | Regression betas, peer unlevered/relevered beta, synthetic-rating cost of debt |
# | 5. DCF | Driver-based 5-year projections, perpetuity growth & exit multiple, sensitivities |
# | 6. LBO "ability to pay" | Max price a financial sponsor can pay at a target IRR |
# | 7. Football field | All methodologies on one chart vs current price |
# | 8. Buyer universe | Strategic acquirers screened on size, leverage and EPS accretion |
# | 9. Exports | **Excel model with live formulas** + **PowerPoint pitchbook** |
#
# **How to use:** edit the *Configuration* cell, then `Runtime → Run all`. Files download at the end.
#
# > Educational project. Market data from Yahoo Finance can be incomplete or mis-tagged; always
# > reconcile key figures (debt, share count, EBITDA) with the latest 10-K/10-Q before relying on them.

# %%
# Install dependencies (Colab already has pandas, numpy, matplotlib, openpyxl)
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "install", "-q", "yfinance", "xlsxwriter", "python-pptx"], check=False)

# %%
import os, math, warnings, datetime as dt
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

warnings.filterwarnings("ignore")
pd.set_option("display.float_format", lambda x: f"{x:,.2f}")
pd.set_option("display.width", 200)

# %% [markdown]
# ## 0. Configuration
# Change the target, the peer set and the potential acquirers here. Every assumption set to `None`
# is derived from the data; overwrite it with a number to impose your own view.

# %%
TARGET = "CAG"                                   # company being valued
PEERS  = ["GIS", "CPB", "SJM", "HRL", "KHC", "MKC", "POST"]   # trading comparables
BUYERS = ["MDLZ", "PEP", "GIS", "HSY", "KHC"]    # potential strategic acquirers
PROJECT_NAME = "Project Harvest"                  # banks always use a code name
AUTHOR = "Alessandro Radice"                      # written into the Excel and PowerPoint file properties

DATA_MODE = os.environ.get("PITCHBOOK_MODE", "live")   # "live" (Yahoo Finance) or "demo" (synthetic, offline)

A = dict(
    # --- DCF ---
    projection_years = 5,
    terminal_growth  = 0.025,     # perpetuity growth rate
    exit_multiple    = None,      # EV/EBITDA at exit; None -> peer median LTM EV/EBITDA
    revenue_growth   = None,      # list of 5 rates, or None -> fade from 3y CAGR to terminal growth
    ebitda_margin    = None,      # None -> 3y average
    da_pct_rev       = None,      # None -> 3y average
    capex_pct_rev    = None,      # None -> 3y average
    nwc_pct_incr_rev = 0.10,      # net working capital investment as % of incremental revenue
    tax_rate         = None,      # None -> historical effective rate, bounded 15%-30%
    # --- WACC ---
    risk_free        = None,      # None -> current 10y UST (^TNX)
    erp              = 0.045,     # equity risk premium: update from Damodaran (pages.stern.nyu.edu/~adamodar)
    # --- Premiums paid ---
    premium_range    = (0.20, 0.40),   # typical US public-target 1-day premia; refine with your own deal data
    # --- LBO ability to pay ---
    lbo_leverage     = 5.0,       # total debt / LTM EBITDA at entry
    lbo_interest     = 0.085,     # blended cost of debt (SOFR + spread)
    lbo_target_irr   = 0.20,
    lbo_fees         = 0.02,      # transaction fees as % of EV
    # --- Buyer screen ---
    buyer_max_leverage = 4.0,     # pro forma net debt / EBITDA a strategic can tolerate (IG-ish ceiling)
)

# Optional: precedent transactions you collected (EDGAR merger proxies, press releases, Capital IQ...)
# Example row: {"date": "2024-08", "acquirer": "Acquirer Inc.", "target": "Target Inc.", "ev_ebitda": 13.5}
PRECEDENTS = []

OUTPUT_DIR = "pitchbook_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)
TODAY = dt.date.today()

# %% [markdown]
# ## 1. Data layer
# `get_company()` returns one standardized record per ticker (all amounts in **$ millions**, except per-share data).
# If Yahoo Finance fails for the target, the notebook falls back to a synthetic demo universe so the pipeline still runs.

# %%
def _row(df, names):
    """First matching line item from a yfinance statement, oldest -> newest."""
    if df is None or len(df) == 0:
        return None
    for n in names:
        if n in df.index:
            s = df.loc[n]
            if isinstance(s, pd.DataFrame):
                s = s.iloc[0]
            s = pd.to_numeric(s, errors="coerce")
            s.index = [pd.Timestamp(c).year for c in s.index]
            return s.sort_index()
    return None


def fetch_live(tk):
    import yfinance as yf
    t = yf.Ticker(tk)
    info = t.info or {}
    price = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose")
    shares = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding")
    if not price or not shares:
        raise ValueError(f"{tk}: no price/share data")

    inc, cf, bs = t.income_stmt, t.cashflow, t.balance_sheet
    rev = _row(inc, ["Total Revenue", "Operating Revenue"])
    ebitda = _row(inc, ["Normalized EBITDA", "EBITDA"])
    ebit = _row(inc, ["EBIT", "Operating Income"])
    da = _row(cf, ["Depreciation And Amortization", "Depreciation Amortization Depletion"])
    if da is None:
        da = _row(inc, ["Reconciled Depreciation"])
    capex = _row(cf, ["Capital Expenditure"])
    dwc = _row(cf, ["Change In Working Capital"])
    ni = _row(inc, ["Net Income Common Stockholders", "Net Income"])
    tax = _row(inc, ["Tax Rate For Calcs"])
    intexp = _row(inc, ["Interest Expense", "Interest Expense Non Operating"])
    minority = _row(bs, ["Minority Interest"])

    hist = pd.DataFrame({
        "Revenue": rev, "EBITDA": ebitda, "EBIT": ebit, "D&A": da,
        "Capex": None if capex is None else capex.abs(),
        "Change in NWC": None if dwc is None else -dwc,   # CF sign flipped: positive = cash invested
        "Net Income": ni, "Interest Expense": None if intexp is None else intexp.abs(),
    }) / 1e6
    hist["Tax Rate"] = tax
    hist = hist.dropna(subset=["Revenue"])
    if "EBITDA" in hist and hist["EBITDA"].isna().any() and hist["EBIT"].notna().all():
        hist["EBITDA"] = hist["EBITDA"].fillna(hist["EBIT"] + hist["D&A"])

    m = 1e6
    rec = dict(
        ticker=tk, name=info.get("longName") or info.get("shortName") or tk,
        sector=info.get("sector", "n.a."), industry=info.get("industry", "n.a."),
        summary=info.get("longBusinessSummary", ""), currency=info.get("financialCurrency", "USD"),
        price=float(price), shares=shares / m,
        debt=(info.get("totalDebt") or 0) / m, cash=(info.get("totalCash") or 0) / m,
        minority=float(minority.iloc[-1]) if minority is not None and pd.notna(minority.iloc[-1]) else 0.0,
        revenue_ltm=(info.get("totalRevenue") or hist["Revenue"].iloc[-1] * m) / m,
        ebitda_ltm=(info.get("ebitda") or hist["EBITDA"].iloc[-1] * m) / m,
        ni_ltm=(info.get("netIncomeToCommon") or hist["Net Income"].iloc[-1] * m) / m,
        eps_ltm=info.get("trailingEps"), eps_ntm=info.get("forwardEps"),
        low52=info.get("fiftyTwoWeekLow"), high52=info.get("fiftyTwoWeekHigh"),
        tgt_low=info.get("targetLowPrice"), tgt_high=info.get("targetHighPrice"),
        hist=hist,
    )
    if rec["minority"]:
        rec["minority"] /= m
    return rec


def fetch_prices(tickers, index="^GSPC"):
    import yfinance as yf
    px = yf.download(list(tickers) + [index], period="5y", interval="1d",
                     auto_adjust=True, progress=False)["Close"]
    return px.rename(columns={index: "INDEX"})


# ---------------- Synthetic demo universe (offline fallback) ----------------
def make_demo_universe():
    rng = np.random.default_rng(7)
    dates = pd.bdate_range(end=pd.Timestamp(TODAY), periods=1260)
    mkt = rng.normal(0.0004, 0.010, len(dates))
    prices = {"INDEX": 4000 * np.exp(np.cumsum(mkt))}
    base_year = TODAY.year - 1
    specs = {  # ticker: (revenue, margin, growth, leverage, P/E, beta)
        "TGT.DEMO":  (11800, 0.175, 0.020, 3.0, 16.5, 0.80),
        "PEER-A.DEMO": (19500, 0.205, 0.030, 2.9, 15.5, 0.85), "PEER-B.DEMO": (9600, 0.180, 0.045, 3.1, 17.0, 0.75),
        "PEER-C.DEMO": (8700, 0.215, 0.025, 3.6, 12.5, 0.70), "PEER-D.DEMO": (12000, 0.120, 0.010, 1.2, 18.5, 0.65),
        "PEER-E.DEMO": (25800, 0.240, -0.010, 3.3, 11.0, 0.80), "PEER-F.DEMO": (6700, 0.205, 0.050, 2.6, 24.0, 0.95),
        "PEER-G.DEMO": (7900, 0.170, 0.060, 4.6, 16.0, 1.00),
        "BUY-A.DEMO": (36000, 0.210, 0.045, 2.1, 20.0, 0.65), "BUY-B.DEMO": (91000, 0.185, 0.035, 2.0, 19.0, 0.55),
        "BUY-C.DEMO": (11200, 0.265, 0.040, 1.6, 21.0, 0.40), "BUY-D.DEMO": (52000, 0.160, 0.030, 3.0, 17.5, 0.50),
    }
    universe = {}
    for i, (tk, (rev, mg, g, lev, pe, beta)) in enumerate(specs.items()):
        years = list(range(base_year - 3, base_year + 1))
        revs = np.array([rev / (1 + g) ** (base_year - y) for y in years]) * (1 + rng.normal(0, 0.01, 4))
        ebitda = revs * (mg + rng.normal(0, 0.006, 4))
        da, capex = revs * 0.035, revs * 0.037
        debt = lev * ebitda[-1]
        intexp = debt * 0.05
        ni = (ebitda - da - intexp) * (1 - 0.23)
        hist = pd.DataFrame({"Revenue": revs, "EBITDA": ebitda, "EBIT": ebitda - da, "D&A": da, "Capex": capex,
                             "Change in NWC": np.r_[0, np.diff(revs)] * 0.08, "Net Income": ni,
                             "Interest Expense": np.full(4, intexp), "Tax Rate": 0.23}, index=years)
        mcap = pe * ni[-1]
        price_now = 20 + (i * 17) % 90
        shares = mcap / price_now
        idio = rng.normal(0, 0.012, len(dates))
        path = np.exp(np.cumsum(beta * mkt + idio))
        path = path / path[-1] * price_now
        prices[tk] = path
        last = path[-252:]
        universe[tk] = dict(
            ticker=tk, name=("Demo Target Co." if tk.startswith("TGT") else f"{tk.split('.')[0].title()} Corp.") + " (synthetic)",
            sector="Consumer Defensive", industry="Packaged Foods (synthetic)",
            summary="Synthetic company generated for offline demonstration. Switch DATA_MODE to 'live' for real data.",
            currency="USD", price=price_now, shares=shares, debt=debt, cash=0.04 * revs[-1], minority=0.0,
            revenue_ltm=revs[-1] * (1 + g / 2), ebitda_ltm=ebitda[-1] * (1 + g / 2), ni_ltm=ni[-1] * (1 + g / 2),
            eps_ltm=ni[-1] * (1 + g / 2) / shares, eps_ntm=ni[-1] * (1 + g) * 1.02 / shares,
            low52=float(last.min()), high52=float(last.max()), tgt_low=price_now * 0.88, tgt_high=price_now * 1.22,
            hist=hist)
    return universe, pd.DataFrame(prices, index=dates)


# ---------------- Load everything ----------------
IS_DEMO = DATA_MODE != "live"
if not IS_DEMO:
    try:
        target = fetch_live(TARGET)
        peers, buyers = [], []
        for tk in PEERS:
            try:
                peers.append(fetch_live(tk))
            except Exception as e:
                print(f"  skipped peer {tk}: {e}")
        for tk in BUYERS:
            try:
                buyers.append(fetch_live(tk))
            except Exception as e:
                print(f"  skipped buyer {tk}: {e}")
        prices = fetch_prices({TARGET, *PEERS, *BUYERS})
        risk_free = A["risk_free"]
        if risk_free is None:
            import yfinance as yf
            risk_free = float(yf.Ticker("^TNX").history(period="5d")["Close"].iloc[-1]) / 100
        if len(peers) < 3:
            raise ValueError("fewer than 3 peers available")
    except Exception as e:
        print(f"Live data unavailable ({e}). Falling back to synthetic DEMO data.")
        IS_DEMO = True

if IS_DEMO:
    U, prices = make_demo_universe()
    TARGET = "TGT.DEMO"
    PEERS = [k for k in U if k.startswith("PEER")]
    BUYERS = [k for k in U if k.startswith("BUY")] + ["PEER-A.DEMO"]
    target, peers, buyers = U[TARGET], [U[k] for k in PEERS], [U[k] for k in BUYERS]
    risk_free = A["risk_free"] or 0.042
    if not PRECEDENTS:   # fictional deals, demo only
        PRECEDENTS = [
            {"date": "2025-03", "acquirer": "Buyer One (demo)", "target": "Snack Co (demo)", "ev_ebitda": 13.8},
            {"date": "2024-11", "acquirer": "Buyer Two (demo)", "target": "Sauce Co (demo)", "ev_ebitda": 12.1},
            {"date": "2024-06", "acquirer": "Sponsor Three (demo)", "target": "Frozen Co (demo)", "ev_ebitda": 10.4},
            {"date": "2023-09", "acquirer": "Buyer Four (demo)", "target": "Cereal Co (demo)", "ev_ebitda": 14.6},
            {"date": "2023-02", "acquirer": "Buyer Five (demo)", "target": "Spice Co (demo)", "ev_ebitda": 16.2},
            {"date": "2022-07", "acquirer": "Sponsor Six (demo)", "target": "Bakery Co (demo)", "ev_ebitda": 11.3},
        ]

DATA_LABEL = "SYNTHETIC DEMO DATA" if IS_DEMO else f"Yahoo Finance, company filings; market data as of {TODAY:%d %b %Y}"

for c in [target] + peers + buyers:
    c["mcap"] = c["price"] * c["shares"]
    c["net_debt"] = c["debt"] - c["cash"]
    c["ev"] = c["mcap"] + c["net_debt"] + c["minority"]

print(f"Target: {target['name']} ({target['ticker']})  |  mode: {'DEMO' if IS_DEMO else 'LIVE'}")
print(f"Peers loaded: {[p['ticker'] for p in peers]}")
print(f"Price ${target['price']:.2f} | Mkt cap ${target['mcap']:,.0f}mm | EV ${target['ev']:,.0f}mm | "
      f"LTM EBITDA ${target['ebitda_ltm']:,.0f}mm")
target["hist"].round(1)

# %% [markdown]
# ## 2. Trading comparables
# Multiples above 100x or with negative denominators are excluded from the statistics (standard banker practice),
# then the 25th–75th percentile range is applied to the target.

# %%
def multiples(c):
    ok = lambda num, den: num / den if den and den > 0 and num / den < 100 else np.nan
    return dict(Ticker=c["ticker"], Company=c["name"], Price=c["price"], MktCap=c["mcap"], EV=c["ev"],
                Revenue=c["revenue_ltm"], EBITDA=c["ebitda_ltm"], Margin=c["ebitda_ltm"] / c["revenue_ltm"],
                EV_Rev=ok(c["ev"], c["revenue_ltm"]), EV_EBITDA=ok(c["ev"], c["ebitda_ltm"]),
                PE_LTM=ok(c["price"], c["eps_ltm"]), PE_NTM=ok(c["price"], c["eps_ntm"]),
                NetDebt_EBITDA=c["net_debt"] / c["ebitda_ltm"] if c["ebitda_ltm"] > 0 else np.nan)

comps = pd.DataFrame([multiples(p) for p in peers]).set_index("Ticker")
mult_cols = ["EV_Rev", "EV_EBITDA", "PE_LTM", "PE_NTM"]
comp_stats = pd.DataFrame({
    "Mean": comps[mult_cols].mean(), "Median": comps[mult_cols].median(),
    "25th pct": comps[mult_cols].quantile(0.25), "75th pct": comps[mult_cols].quantile(0.75)}).T
target_mult = multiples(target)

def ev_to_price(ev, c=target):
    return (ev - c["net_debt"] - c["minority"]) / c["shares"]

implied = {
    "EV / LTM EBITDA": [ev_to_price(comp_stats.loc[q, "EV_EBITDA"] * target["ebitda_ltm"]) for q in ["25th pct", "75th pct"]],
    "EV / LTM Revenue": [ev_to_price(comp_stats.loc[q, "EV_Rev"] * target["revenue_ltm"]) for q in ["25th pct", "75th pct"]],
    "P / NTM EPS": [comp_stats.loc[q, "PE_NTM"] * (target["eps_ntm"] or np.nan) for q in ["25th pct", "75th pct"]],
}
display_cols = ["Company", "MktCap", "EV", "Margin", "EV_Rev", "EV_EBITDA", "PE_LTM", "PE_NTM", "NetDebt_EBITDA"]
print("Target multiples: EV/EBITDA {:.1f}x | P/E NTM {:.1f}x".format(target_mult["EV_EBITDA"], target_mult["PE_NTM"]))
print(pd.DataFrame(implied, index=["Low", "High"]).T.round(2))
comps[display_cols].round(2)

# %% [markdown]
# ## 3. Precedent transactions & premiums paid
# Control transactions trade at a premium to the unaffected share price. Two lenses:
# * **Precedent multiples** (if you filled `PRECEDENTS`): 25th–75th percentile EV/LTM EBITDA applied to the target.
# * **Premiums paid**: the typical one-day premium range applied to the current (unaffected) price.

# %%
prec = pd.DataFrame(PRECEDENTS)
if len(prec):
    q25, q75 = prec["ev_ebitda"].quantile([0.25, 0.75])
    implied["Precedent transactions"] = [ev_to_price(q25 * target["ebitda_ltm"]), ev_to_price(q75 * target["ebitda_ltm"])]
    print(f"Precedent EV/EBITDA: median {prec['ev_ebitda'].median():.1f}x, IQR {q25:.1f}x-{q75:.1f}x")
lo_p, hi_p = A["premium_range"]
implied["Premiums paid"] = [target["price"] * (1 + lo_p), target["price"] * (1 + hi_p)]
mid_offer = target["price"] * (1 + np.mean(A["premium_range"]))
print(f"Premiums paid {lo_p:.0%}-{hi_p:.0%}: ${implied['Premiums paid'][0]:.2f} - ${implied['Premiums paid'][1]:.2f}")
prec

# %% [markdown]
# ## 4. WACC
# * Raw regression betas (2y weekly, 5y monthly) vs the S&P 500, Blume-adjusted.
# * Industry beta: peer betas unlevered with Hamada, median, relevered at the target's market D/E.
# * Cost of debt: interest coverage → synthetic rating → default spread (Damodaran method).

# %%
def reg_beta(stock, index, freq):
    px = prices[[stock, index]].dropna()
    px = px.resample(freq).last()
    px = px.iloc[-105:] if freq == "W-FRI" else px.iloc[-61:]
    r = px.pct_change().dropna()
    b = np.cov(r[stock], r[index])[0, 1] / np.var(r[index], ddof=1)
    r2 = np.corrcoef(r[stock], r[index])[0, 1] ** 2
    return b, r2

tax_hist = target["hist"]["Tax Rate"].dropna()
TAX = A["tax_rate"] or float(np.clip(tax_hist.tail(3).mean() if len(tax_hist) else 0.21, 0.15, 0.30))

b2w, r2w = reg_beta(TARGET, "INDEX", "W-FRI")
b5m, r2m = reg_beta(TARGET, "INDEX", "ME")
blume = lambda b: 0.67 * b + 0.33

peer_beta = []
for p in peers:
    if p["ticker"] not in prices:
        continue
    b, _ = reg_beta(p["ticker"], "INDEX", "W-FRI")
    de = p["debt"] / p["mcap"]
    peer_beta.append(dict(Ticker=p["ticker"], RawBeta=b, AdjBeta=blume(b), DE=de,
                          Unlevered=blume(b) / (1 + (1 - TAX) * de)))
peer_beta = pd.DataFrame(peer_beta).set_index("Ticker")
unlev_med = peer_beta["Unlevered"].median()
de_target = target["debt"] / target["mcap"]
beta_relev = unlev_med * (1 + (1 - TAX) * de_target)

# Synthetic rating (Damodaran, large non-financial firms). Spreads approximate: refresh each January.
RATING_TABLE = [(8.5, "AAA", .0059), (6.5, "AA", .0078), (5.5, "A+", .0098), (4.25, "A", .0108), (3.0, "A-", .0122),
                (2.5, "BBB", .0156), (2.25, "BB+", .0200), (2.0, "BB", .0240), (1.75, "B+", .0275), (1.5, "B", .0321),
                (1.25, "B-", .0394), (0.8, "CCC", .0500), (0.65, "CC", .0800), (0.2, "C", .1000), (-1e9, "D", .1200)]
h = target["hist"].iloc[-1]
coverage = h["EBIT"] / h["Interest Expense"] if h.get("Interest Expense", 0) and h["Interest Expense"] > 0 else 12.0
rating, spread = next((r, s) for thr, r, s in RATING_TABLE if coverage > thr)
kd = risk_free + spread

ke = risk_free + beta_relev * A["erp"]
wE = target["mcap"] / (target["mcap"] + target["debt"])
WACC = wE * ke + (1 - wE) * kd * (1 - TAX)

wacc_tbl = pd.Series({
    "Risk-free rate (10y UST)": risk_free, "Equity risk premium": A["erp"],
    "Raw beta 2y weekly": b2w, "Raw beta 5y monthly": b5m, "Peer median unlevered beta": unlev_med,
    "Target D/E (market)": de_target, "Relevered beta (selected)": beta_relev, "Cost of equity": ke,
    "Interest coverage (EBIT/interest)": coverage, "Synthetic rating": rating, "Pre-tax cost of debt": kd,
    "Tax rate": TAX, "Equity weight": wE, "WACC": WACC})
print("\n".join(f"{k:<36}{v:>10.2%}" if isinstance(v, float) and k not in ("Raw beta 2y weekly","Raw beta 5y monthly","Peer median unlevered beta","Relevered beta (selected)","Interest coverage (EBIT/interest)") else f"{k:<36}{v:>10.2f}" if isinstance(v, float) else f"{k:<36}{v:>10}" for k, v in wacc_tbl.items()))
peer_beta.round(2)

# %% [markdown]
# ## 5. Projections & DCF
# Drivers default to 3-year historical averages; revenue growth fades linearly from the 3-year CAGR to terminal growth.
# Mid-year discounting; terminal value discounted from the end of year N.

# %%
N = A["projection_years"]
H = target["hist"].tail(4)
last3 = H.tail(3)
base_year = int(H.index[-1])
cagr = (H["Revenue"].iloc[-1] / H["Revenue"].iloc[0]) ** (1 / (len(H) - 1)) - 1
g0 = float(np.clip(cagr, -0.05, 0.15))
drivers = pd.DataFrame(index=[f"FY{base_year + i}E" for i in range(1, N + 1)])
drivers["Revenue growth"] = A["revenue_growth"] or list(np.linspace(g0, A["terminal_growth"], N))
drivers["EBITDA margin"] = A["ebitda_margin"] or float((last3["EBITDA"] / last3["Revenue"]).mean())
drivers["D&A % revenue"] = A["da_pct_rev"] or float((last3["D&A"] / last3["Revenue"]).mean())
drivers["Capex % revenue"] = A["capex_pct_rev"] or float((last3["Capex"] / last3["Revenue"]).mean())
drivers["NWC % incr. revenue"] = A["nwc_pct_incr_rev"]


def project(drv, base_rev):
    p = pd.DataFrame(index=drv.index)
    p["Revenue"] = base_rev * (1 + drv["Revenue growth"]).cumprod()
    p["EBITDA"] = p["Revenue"] * drv["EBITDA margin"]
    p["D&A"] = p["Revenue"] * drv["D&A % revenue"]
    p["EBIT"] = p["EBITDA"] - p["D&A"]
    p["Taxes"] = p["EBIT"] * TAX
    p["NOPAT"] = p["EBIT"] - p["Taxes"]
    p["Capex"] = p["Revenue"] * drv["Capex % revenue"]
    p["Change in NWC"] = p["Revenue"].diff().fillna(p["Revenue"].iloc[0] - base_rev) * drv["NWC % incr. revenue"]
    p["Unlevered FCF"] = p["NOPAT"] + p["D&A"] - p["Capex"] - p["Change in NWC"]
    return p

proj = project(drivers, H["Revenue"].iloc[-1])
EXIT_MULT = A["exit_multiple"] or float(comp_stats.loc["Median", "EV_EBITDA"])
g_term = A["terminal_growth"]


def dcf(wacc, g=None, exit_mult=None, p=proj):
    t = np.arange(1, len(p) + 1)
    pv_fcf = (p["Unlevered FCF"].values / (1 + wacc) ** (t - 0.5)).sum()
    if exit_mult is None:
        tv = p["Unlevered FCF"].iloc[-1] * (1 + g) / (wacc - g)
    else:
        tv = p["EBITDA"].iloc[-1] * exit_mult
    pv_tv = tv / (1 + wacc) ** len(p)
    ev = pv_fcf + pv_tv
    return dict(ev=ev, pv_fcf=pv_fcf, pv_tv=pv_tv, tv=tv, tv_share=pv_tv / ev, price=ev_to_price(ev))

d_pg = dcf(WACC, g=g_term)
d_em = dcf(WACC, exit_mult=EXIT_MULT)
implied_exit = d_pg["tv"] / proj["EBITDA"].iloc[-1]
implied_g = (d_em["tv"] * WACC - proj["Unlevered FCF"].iloc[-1]) / (d_em["tv"] + proj["Unlevered FCF"].iloc[-1])

wacc_grid = np.round(WACC + np.array([-0.01, -0.005, 0, 0.005, 0.01]), 4)
g_grid = np.round(g_term + np.array([-0.01, -0.005, 0, 0.005, 0.01]), 4)
m_grid = np.round(EXIT_MULT + np.array([-2, -1, 0, 1, 2]), 1)
sens_pg = pd.DataFrame([[dcf(w, g=g)["price"] for g in g_grid] for w in wacc_grid], index=wacc_grid, columns=g_grid)
sens_em = pd.DataFrame([[dcf(w, exit_mult=m)["price"] for m in m_grid] for w in wacc_grid], index=wacc_grid, columns=m_grid)
implied["DCF – perpetuity growth"] = [sens_pg.iloc[1:4, 1:4].values.min(), sens_pg.iloc[1:4, 1:4].values.max()]
implied["DCF – exit multiple"] = [sens_em.iloc[1:4, 1:4].values.min(), sens_em.iloc[1:4, 1:4].values.max()]

print(f"WACC {WACC:.2%} | g {g_term:.1%} | exit {EXIT_MULT:.1f}x")
print(f"Perpetuity: EV ${d_pg['ev']:,.0f}mm -> ${d_pg['price']:.2f}/sh (TV {d_pg['tv_share']:.0%} of EV, implied exit {implied_exit:.1f}x)")
print(f"Exit mult.: EV ${d_em['ev']:,.0f}mm -> ${d_em['price']:.2f}/sh (TV {d_em['tv_share']:.0%} of EV, implied g {implied_g:.2%})")
proj.T.round(0)

# %% [markdown]
# ## 6. LBO "ability to pay"
# What is the highest price a financial sponsor could pay and still earn its target IRR?
# Entry debt = leverage × LTM EBITDA, 100% cash sweep, exit after 5 years at the peer median EV/EBITDA (same exit multiple as the DCF).

# %%
def lbo(entry_ev, lev=A["lbo_leverage"], rate=A["lbo_interest"], p=proj, fees=A["lbo_fees"]):
    e0 = target["ebitda_ltm"]
    debt0 = lev * e0
    equity0 = entry_ev * (1 + fees) - debt0
    debt, cash, rows = debt0, 0.0, []
    for yr in p.index:
        ebitda, da, capex, dnwc = p.loc[yr, ["EBITDA", "D&A", "Capex", "Change in NWC"]]
        interest = rate * debt
        taxes = max(0.0, ebitda - da - interest) * TAX
        lfcf = ebitda - interest - taxes - capex - dnwc
        repay = min(debt, max(lfcf, 0.0))
        debt -= repay
        cash += lfcf - repay
        rows.append(dict(Year=yr, EBITDA=ebitda, Interest=interest, Taxes=taxes, LFCF=lfcf, Repayment=repay,
                         Debt=debt, Cash=cash, Leverage=(debt - cash) / ebitda))
    exit_ev = EXIT_MULT * p["EBITDA"].iloc[-1]
    exit_eq = exit_ev - debt + cash
    if equity0 <= 0:
        return dict(irr=np.nan)
    moic = exit_eq / equity0
    irr = moic ** (1 / len(p)) - 1 if moic > 0 else -1.0
    return dict(irr=irr, moic=moic, equity0=equity0, debt0=debt0, exit_ev=exit_ev, exit_eq=exit_eq,
                schedule=pd.DataFrame(rows).set_index("Year"))


def solve_entry_ev(target_irr, lev=A["lbo_leverage"]):
    lo, hi = lev * target["ebitda_ltm"] / (1 + A["lbo_fees"]) * 1.001, 40 * target["ebitda_ltm"]
    f = lambda ev: lbo(ev, lev=lev)["irr"] - target_irr
    if np.isnan(f(lo)) or f(lo) < 0:
        return np.nan
    for _ in range(100):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if f(mid) > 0 else (lo, mid)
    return (lo + hi) / 2

lbo_ev = solve_entry_ev(A["lbo_target_irr"])
lbo_base = lbo(lbo_ev)
lbo_price = ev_to_price(lbo_ev)
irr_grid, lev_grid = [0.25, 0.225, 0.20, 0.175, 0.15], [4.0, 4.5, 5.0, 5.5, 6.0]
sens_lbo = pd.DataFrame([[ev_to_price(solve_entry_ev(i, l)) for l in lev_grid] for i in irr_grid],
                        index=irr_grid, columns=lev_grid)
implied["LBO (20–25% IRR)"] = sorted([ev_to_price(solve_entry_ev(0.25, A["lbo_leverage"] - 0.5)),
                                     ev_to_price(solve_entry_ev(0.20, A["lbo_leverage"] + 0.5))])
print(f"Max entry EV at {A['lbo_target_irr']:.0%} IRR and {A['lbo_leverage']:.1f}x: ${lbo_ev:,.0f}mm "
      f"({lbo_ev / target['ebitda_ltm']:.1f}x LTM EBITDA) -> ${lbo_price:.2f}/sh "
      f"({lbo_price / target['price'] - 1:+.0%} vs current) | MOIC {lbo_base['moic']:.2f}x")
lbo_base["schedule"].round(1)

# %% [markdown]
# ## 7. Football field

# %%
ff = {"52-week trading range": [target["low52"], target["high52"]]}
if target.get("tgt_low") and target.get("tgt_high"):
    ff["Analyst price targets"] = [target["tgt_low"], target["tgt_high"]]
ff.update(implied)
ff = pd.DataFrame(ff, index=["Low", "High"]).T.dropna()
ff = ff[(ff["Low"] > 0)].drop(index="EV / LTM Revenue", errors="ignore")   # revenue multiple shown in comps only

BLUE, ORANGE, INK, MUTED = "#2a78d6", "#eb6834", "#0b0b0b", "#52514e"
plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#b5b3ad", "axes.labelcolor": MUTED, "xtick.color": MUTED, "ytick.color": INK})


def plot_football_field(ff, path):
    fig, ax = plt.subplots(figsize=(11, 0.55 * len(ff) + 1.6))
    y = np.arange(len(ff))[::-1]
    ax.barh(y, ff["High"] - ff["Low"], left=ff["Low"], height=0.55, color=BLUE)
    span = ff["High"].max() - ff["Low"].min()
    for yi, (lo, hi) in zip(y, ff.values):
        ax.text(lo - span * 0.01, yi, f"${lo:,.2f}", va="center", ha="right", fontsize=9, color=INK)
        ax.text(hi + span * 0.01, yi, f"${hi:,.2f}", va="center", ha="left", fontsize=9, color=INK)
    ax.axvline(target["price"], color=INK, lw=1.5, ls="--")
    ax.axvspan(*implied["Premiums paid"], color=ORANGE, alpha=0.10, lw=0)
    ax.text(target["price"], -0.75, f"Current ${target['price']:.2f}", ha="center", va="top", fontsize=9, color=INK, fontweight="bold")
    ax.set_ylim(-1.1, len(ff) - 0.5)
    ax.set_yticks(y, ff.index, fontsize=10)
    ax.set_xlim(ff["Low"].min() - span * 0.12, ff["High"].max() + span * 0.12)
    ax.xaxis.set_major_formatter(mtick.StrMethodFormatter("${x:,.0f}"))
    ax.grid(axis="x", color="#e6e4df", lw=0.8)
    ax.set_axisbelow(True)
    ax.set_title(f"{target['name']} – implied value per share", loc="left", fontsize=13, color=INK, pad=14)
    fig.text(0.01, 0.01, f"Shaded band: premiums-paid range. Source: {DATA_LABEL}", fontsize=8, color=MUTED)
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(path, dpi=200)
    return fig

plot_football_field(ff, f"{OUTPUT_DIR}/football_field.png")
plt.show()
ff.round(2)

# %% [markdown]
# ## 8. Potential strategic buyers
# Screen: can the acquirer fund the deal at the mid-point premium, and would an all-stock deal be EPS accretive?
# * **Pro forma leverage** assumes 100% debt funding: (buyer net debt + target net debt + equity purchase price) / combined EBITDA.
# * **All-stock accretion** is a first-order test: accretive if the buyer's P/E exceeds the P/E paid for the target (before synergies).

# %%
offer_eq = mid_offer * target["shares"]
offer_pe = mid_offer / target["eps_ntm"] if target["eps_ntm"] else np.nan
rows = []
for b in buyers:
    if b["ticker"] == target["ticker"]:
        continue
    pf_lev = (b["net_debt"] + target["net_debt"] + offer_eq) / (b["ebitda_ltm"] + target["ebitda_ltm"])
    pe_b = b["price"] / b["eps_ntm"] if b["eps_ntm"] and b["eps_ntm"] > 0 else np.nan
    rows.append({"Ticker": b["ticker"], "Company": b["name"], "Mkt cap ($mm)": b["mcap"],
                 "Net debt / EBITDA": b["net_debt"] / b["ebitda_ltm"],
                 "Deal size / buyer mkt cap": (offer_eq + target["net_debt"]) / b["mcap"],
                 "PF leverage (all-debt)": pf_lev, "Buyer P/E NTM": pe_b,
                 "All-cash feasible": "Yes" if pf_lev <= A["buyer_max_leverage"] else "Stretch",
                 "All-stock EPS": "Accretive" if pe_b > offer_pe else "Dilutive"})
buyers_df = pd.DataFrame(rows).set_index("Ticker").sort_values("PF leverage (all-debt)")
print(f"Offer at mid premium: ${mid_offer:.2f}/sh, equity value ${offer_eq:,.0f}mm, implied P/E NTM {offer_pe:.1f}x")
buyers_df.round(2)

# %% [markdown]
# ## 9. Excel export: live-formula model
# Blue = hard-coded inputs, black = formulas, green = links to other sheets. Change any blue cell and the model recalculates.

# %%
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as rc, xl_col_to_name

XLSX = f"{OUTPUT_DIR}/{target['ticker'].replace('.', '_')}_valuation_model.xlsx"
wb = xlsxwriter.Workbook(XLSX)
wb.set_properties({"title": f"{PROJECT_NAME}: {target['name']} valuation model", "subject": "Preliminary M&A valuation: comps, precedents, WACC, DCF, LBO",
                   "author": AUTHOR, "manager": AUTHOR, "company": "", "category": "Valuation model",
                   "keywords": "M&A, valuation, DCF, LBO, trading comps, investment banking",
                   "comments": "Generated by the Sell-Side Valuation & Pitchbook Engine", "created": dt.datetime.now()})
F = dict(font_name="Arial", font_size=10)
fmt = lambda **k: wb.add_format({**F, **k})
f_title = fmt(bold=True, font_size=14, font_color="#0B2545")
f_sub = fmt(italic=True, font_color="#52514e")
f_hdr = fmt(bold=True, font_color="white", bg_color="#0B2545", align="center", valign="vcenter", text_wrap=True)
f_lbl, f_bold = fmt(), fmt(bold=True)
f_in_num = fmt(font_color="#0000FF", num_format='#,##0.0;(#,##0.0)')
f_in_pct = fmt(font_color="#0000FF", num_format='0.0%')
f_in_x = fmt(font_color="#0000FF", num_format='0.0"x"')
f_in_px = fmt(font_color="#0000FF", num_format='$#,##0.00')
f_in_txt = fmt(font_color="#0000FF")
f_num = fmt(num_format='#,##0.0;(#,##0.0)')
f_num_b = fmt(num_format='#,##0.0;(#,##0.0)', bold=True, top=1)
f_pct = fmt(num_format='0.0%')
f_x = fmt(num_format='0.0"x"')
f_px = fmt(num_format='$#,##0.00')
f_px_b = fmt(num_format='$#,##0.00', bold=True, bg_color="#FFF2CC", border=1)
f_link_num = fmt(font_color="#008000", num_format='#,##0.0;(#,##0.0)')
f_link_pct = fmt(font_color="#008000", num_format='0.0%')
f_link_x = fmt(font_color="#008000", num_format='0.0"x"')
f_link_px = fmt(font_color="#008000", num_format='$#,##0.00')
f_sens_px = fmt(num_format='$#,##0.00', border=1, align="center")
f_sens_hdr_pct = fmt(num_format='0.0%', bold=True, bg_color="#DCE6F1", border=1, align="center")
f_sens_hdr_x = fmt(num_format='0.0"x"', bold=True, bg_color="#DCE6F1", border=1, align="center")

# ---------------- Cover ----------------
ws = wb.add_worksheet("Cover")
ws.hide_gridlines(2); ws.set_column("B:B", 40); ws.set_column("C:C", 60)
ws.write("B2", f"{PROJECT_NAME}: {target['name']}", f_title)
ws.write("B3", "Preliminary valuation model (illustrative, educational)", f_sub)
ws.write("B4", f"Source: {DATA_LABEL}", f_sub)
for i, (k, v) in enumerate([("Inputs", "Market data and all assumptions (blue cells)"),
                            ("Comps", "Trading comparables with live multiple and quartile formulas"),
                            ("WACC", "Beta unlevering/relevering and cost of capital build"),
                            ("DCF", "5-year projections, perpetuity and exit valuation, live sensitivity grids"),
                            ("LBO", "Sponsor ability-to-pay: entry price input, debt schedule, IRR"),
                            ("Football Field", "Valuation summary and chart")]):
    ws.write(6 + i, 1, k, f_bold); ws.write(6 + i, 2, v)
ws.write("B14", "Colour code", f_bold)
ws.write("B15", "Blue = hard-coded input", fmt(font_color="#0000FF"))
ws.write("B16", "Black = formula"); ws.write("B17", "Green = link to another sheet", fmt(font_color="#008000"))

# ---------------- Inputs ----------------
wi = wb.add_worksheet("Inputs")
wi.hide_gridlines(2); wi.set_column("A:A", 2); wi.set_column("B:B", 36); wi.set_column("C:H", 13)
wi.write("B2", "Inputs & assumptions", f_title)
REF = {}
def inp(row, label, value, f, key):
    wi.write(row, 1, label, f_lbl); wi.write(row, 2, value, f); REF[key] = f"Inputs!$C${row + 1}"

wi.write(3, 1, "Target market data ($mm, except per share)", f_bold)
inp(4, "Share price", target["price"], f_in_px, "price")
inp(5, "Diluted shares (mm)", target["shares"], f_in_num, "shares")
inp(6, "Total debt", target["debt"], f_in_num, "debt")
inp(7, "Cash & equivalents", target["cash"], f_in_num, "cash")
inp(8, "Minority interest", target["minority"], f_in_num, "minority")
inp(9, "LTM revenue", target["revenue_ltm"], f_in_num, "rev_ltm")
inp(10, "LTM EBITDA", target["ebitda_ltm"], f_in_num, "ebitda_ltm")
inp(11, "NTM EPS (consensus)", target["eps_ntm"] or 0, f_in_px, "eps_ntm")
wi.write(12, 1, "Market capitalisation"); wi.write_formula(12, 2, f"={REF['price']}*{REF['shares']}", f_num); REF["mcap"] = "Inputs!$C$13"
wi.write(13, 1, "Enterprise value"); wi.write_formula(13, 2, f"=C13+C7-C8+C9", f_num); REF["ev"] = "Inputs!$C$14"

wi.write(15, 1, "Valuation assumptions", f_bold)
inp(16, "Risk-free rate", risk_free, f_in_pct, "rf")
inp(17, "Equity risk premium", A["erp"], f_in_pct, "erp")
inp(18, "Tax rate", TAX, f_in_pct, "tax")
inp(19, "Terminal growth", g_term, f_in_pct, "g")
inp(20, "Exit multiple (EV/EBITDA)", EXIT_MULT, f_in_x, "exit")
inp(21, "Pre-tax cost of debt (synthetic rating)", kd, f_in_pct, "kd")

wi.write(23, 1, "LBO assumptions", f_bold)
inp(24, "Entry leverage (x LTM EBITDA)", A["lbo_leverage"], f_in_x, "lbo_lev")
inp(25, "Cost of debt", A["lbo_interest"], f_in_pct, "lbo_rate")
inp(26, "Transaction fees (% EV)", A["lbo_fees"], f_in_pct, "lbo_fees")
inp(27, "Offer price per share (solved at target IRR)", lbo_price, f_in_px, "lbo_price")

wi.write(29, 1, "Operating drivers", f_bold)
wi.write_row(29, 2, list(drivers.index), f_hdr)
DRV_ROW = {}
for i, col in enumerate(drivers.columns):
    r = 30 + i
    wi.write(r, 1, col)
    wi.write_row(r, 2, list(drivers[col].values), f_in_pct)
    DRV_ROW[col] = r + 1
inp(36, f"Base year revenue (FY{base_year}A)", H["Revenue"].iloc[-1], f_in_num, "base_rev")

# ---------------- Comps ----------------
wc = wb.add_worksheet("Comps")
wc.hide_gridlines(2); wc.set_column("A:A", 2); wc.set_column("B:B", 12); wc.set_column("C:C", 34); wc.set_column("D:P", 11)
wc.write("B2", "Trading comparables ($mm, except per share)", f_title)
heads = ["Ticker", "Company", "Price", "Shares", "Debt", "Cash", "Minority", "Revenue LTM", "EBITDA LTM", "EPS LTM",
         "EPS NTM", "Mkt cap", "EV", "EV/Rev", "EV/EBITDA", "P/E LTM", "P/E NTM"]
wc.set_column("D:D", 10)
wc.write_row(3, 1, heads, f_hdr); wc.set_row(3, 30)
r0 = 4
for i, p in enumerate(peers):
    r = r0 + i; R = r + 1
    wc.write(r, 1, p["ticker"], f_in_txt); wc.write(r, 2, p["name"], f_in_txt)
    for j, (k, f) in enumerate([("price", f_in_px), ("shares", f_in_num), ("debt", f_in_num), ("cash", f_in_num),
                                ("minority", f_in_num), ("revenue_ltm", f_in_num), ("ebitda_ltm", f_in_num),
                                ("eps_ltm", f_in_px), ("eps_ntm", f_in_px)]):
        v = p[k]
        wc.write(r, 3 + j, v if v is not None and not (isinstance(v, float) and np.isnan(v)) else 0, f)
    wc.write_formula(r, 12, f"=D{R}*E{R}", f_num)
    wc.write_formula(r, 13, f"=M{R}+F{R}-G{R}+H{R}", f_num)
    wc.write_formula(r, 14, f'=IF(I{R}>0,N{R}/I{R},"n.m.")', f_x)
    wc.write_formula(r, 15, f'=IF(AND(J{R}>0,N{R}/J{R}<100),N{R}/J{R},"n.m.")', f_x)
    wc.write_formula(r, 16, f'=IF(AND(K{R}>0,D{R}/K{R}<100),D{R}/K{R},"n.m.")', f_x)
    wc.write_formula(r, 17, f'=IF(AND(L{R}>0,D{R}/L{R}<100),D{R}/L{R},"n.m.")', f_x)
r_last = r0 + len(peers)
stats_rows = {}
for k, (lab, fn) in enumerate([("Mean", "AVERAGE({})"), ("Median", "MEDIAN({})"),
                               ("25th percentile", "QUARTILE({},1)"), ("75th percentile", "QUARTILE({},3)")]):
    r = r_last + 1 + k
    wc.write(r, 2, lab, f_bold)
    for c in range(14, 18):
        col = xl_col_to_name(c)
        wc.write_formula(r, c, "=" + fn.format(f"{col}{r0 + 1}:{col}{r_last}"), f_x)
    stats_rows[lab] = r + 1
r = r_last + 6
wc.write(r, 2, "Implied value per share", f_bold); wc.write_row(r, 3, ["Low (25th)", "High (75th)"], f_hdr)
q1, q3 = stats_rows["25th percentile"], stats_rows["75th percentile"]
nd = f"({REF['debt']}-{REF['cash']}+{REF['minority']})"
wc.write(r + 1, 2, "EV / LTM EBITDA")
wc.write_formula(r + 1, 3, f"=(P{q1}*{REF['ebitda_ltm']}-{nd})/{REF['shares']}", f_px)
wc.write_formula(r + 1, 4, f"=(P{q3}*{REF['ebitda_ltm']}-{nd})/{REF['shares']}", f_px)
wc.write(r + 2, 2, "P / NTM EPS")
wc.write_formula(r + 2, 3, f"=R{q1}*{REF['eps_ntm']}", f_px)
wc.write_formula(r + 2, 4, f"=R{q3}*{REF['eps_ntm']}", f_px)
REF["comps_ebitda"] = (f"Comps!$D${r + 2}", f"Comps!$E${r + 2}")
REF["comps_pe"] = (f"Comps!$D${r + 3}", f"Comps!$E${r + 3}")
REF["median_ev_ebitda"] = f"Comps!$P${stats_rows['Median']}"

# ---------------- WACC ----------------
ww = wb.add_worksheet("WACC")
ww.hide_gridlines(2); ww.set_column("A:A", 2); ww.set_column("B:B", 34); ww.set_column("C:G", 12)
ww.write("B2", "Cost of capital", f_title)
ww.write_row(3, 1, ["Peer", "Adj. beta (Blume)", "Debt", "Mkt cap", "D/E", "Unlevered beta"], f_hdr)
for i, (tk, row) in enumerate(peer_beta.iterrows()):
    p = next(x for x in peers if x["ticker"] == tk)
    r = 4 + i; R = r + 1
    ww.write(r, 1, tk, f_in_txt); ww.write(r, 2, row["AdjBeta"], fmt(font_color="#0000FF", num_format="0.00"))
    ww.write(r, 3, p["debt"], f_in_num); ww.write(r, 4, p["mcap"], f_in_num)
    ww.write_formula(r, 5, f"=D{R}/E{R}", f_pct)
    ww.write_formula(r, 6, f"=C{R}/(1+(1-{REF['tax']})*F{R})", fmt(num_format="0.00"))
rl = 4 + len(peer_beta)
f2 = fmt(num_format="0.00")
lines = [("Median unlevered beta", f"=MEDIAN(G5:G{rl})", f2),
         ("Target D/E (market)", f"={REF['debt']}/{REF['mcap']}", f_pct),
         ("Relevered beta", f"=C{rl + 2}*(1+(1-{REF['tax']})*C{rl + 3})", f2),
         ("Risk-free rate", f"={REF['rf']}", f_link_pct),
         ("Equity risk premium", f"={REF['erp']}", f_link_pct),
         ("Cost of equity", f"=C{rl + 5}+C{rl + 4}*C{rl + 6}", f_pct),
         ("Pre-tax cost of debt", f"={REF['kd']}", f_link_pct),
         ("After-tax cost of debt", f"=C{rl + 8}*(1-{REF['tax']})", f_pct),
         ("Equity weight E/(D+E)", f"={REF['mcap']}/({REF['mcap']}+{REF['debt']})", f_pct),
         ("WACC", f"=C{rl + 10}*C{rl + 7}+(1-C{rl + 10})*C{rl + 9}", fmt(num_format="0.00%", bold=True, bg_color="#FFF2CC", border=1))]
for i, (lab, form, f) in enumerate(lines):
    ww.write(rl + 1 + i, 1, lab, f_bold if lab == "WACC" else f_lbl)
    ww.write_formula(rl + 1 + i, 2, form, f)
REF["wacc"] = f"WACC!$C${rl + 11}"
ww.write(rl + 13, 1, f"Regression betas (raw): 2y weekly {b2w:.2f} (R² {r2w:.2f}), 5y monthly {b5m:.2f} (R² {r2m:.2f})", f_sub)
ww.write(rl + 14, 1, f"Synthetic rating {rating} from interest coverage {coverage:.1f}x (Damodaran table, approximate spreads)", f_sub)

# ---------------- DCF ----------------
wd = wb.add_worksheet("DCF")
wd.hide_gridlines(2); wd.set_column("A:A", 2); wd.set_column("B:B", 30); wd.set_column("C:I", 12)
wd.write("B2", "Discounted cash flow ($mm)", f_title)
cols = [f"FY{base_year}A"] + list(drivers.index)
wd.write_row(3, 2, cols, f_hdr)
wd.write(4, 1, "Period (mid-year)")
for j in range(N):
    wd.write(4, 3 + j, j + 1, fmt(font_color="#0000FF", align="center"))
labels = ["Revenue", "% growth", "EBITDA", "% margin", "D&A", "EBIT", "Taxes", "NOPAT", "(+) D&A", "(-) Capex",
          "(-) Change in NWC", "Unlevered FCF", "Discount factor", "PV of FCF"]
RW = {lab: 6 + i for i, lab in enumerate(labels)}  # 0-indexed rows
for lab, r in RW.items():
    wd.write(r, 1, lab, f_bold if lab in ("Revenue", "EBITDA", "Unlevered FCF") else f_lbl)
wd.write_formula(RW["Revenue"], 2, f"={REF['base_rev']}", f_link_num)
for j in range(N):
    c = 3 + j; C = xl_col_to_name(c); P = xl_col_to_name(c - 1); dcol = xl_col_to_name(2 + j)
    row = lambda lab: RW[lab] + 1
    wd.write_formula(RW["% growth"], c, f"=Inputs!{dcol}{DRV_ROW['Revenue growth']}", f_link_pct)
    wd.write_formula(RW["Revenue"], c, f"={P}{row('Revenue')}*(1+{C}{row('% growth')})", f_num)
    wd.write_formula(RW["% margin"], c, f"=Inputs!{dcol}{DRV_ROW['EBITDA margin']}", f_link_pct)
    wd.write_formula(RW["EBITDA"], c, f"={C}{row('Revenue')}*{C}{row('% margin')}", f_num)
    wd.write_formula(RW["D&A"], c, f"={C}{row('Revenue')}*Inputs!{dcol}{DRV_ROW['D&A % revenue']}", f_num)
    wd.write_formula(RW["EBIT"], c, f"={C}{row('EBITDA')}-{C}{row('D&A')}", f_num)
    wd.write_formula(RW["Taxes"], c, f"=-{C}{row('EBIT')}*{REF['tax']}", f_num)
    wd.write_formula(RW["NOPAT"], c, f"={C}{row('EBIT')}+{C}{row('Taxes')}", f_num)
    wd.write_formula(RW["(+) D&A"], c, f"={C}{row('D&A')}", f_num)
    wd.write_formula(RW["(-) Capex"], c, f"=-{C}{row('Revenue')}*Inputs!{dcol}{DRV_ROW['Capex % revenue']}", f_num)
    wd.write_formula(RW["(-) Change in NWC"], c, f"=-({C}{row('Revenue')}-{P}{row('Revenue')})*Inputs!{dcol}{DRV_ROW['NWC % incr. revenue']}", f_num)
    wd.write_formula(RW["Unlevered FCF"], c, f"={C}{row('NOPAT')}+{C}{row('(+) D&A')}+{C}{row('(-) Capex')}+{C}{row('(-) Change in NWC')}", f_num_b)
    wd.write_formula(RW["Discount factor"], c, f"=1/(1+{REF['wacc']})^({C}5-0.5)", fmt(num_format="0.000"))
    wd.write_formula(RW["PV of FCF"], c, f"={C}{row('Unlevered FCF')}*{C}{row('Discount factor')}", f_num)
LC = xl_col_to_name(2 + N)
fcf_rng = f"$D${RW['Unlevered FCF'] + 1}:${LC}${RW['Unlevered FCF'] + 1}"
per_rng = f"$D$5:${LC}$5"
v0 = RW["PV of FCF"] + 3
wd.write_row(v0 - 1, 1, ["Valuation", "Perpetuity", "Exit multiple"], f_hdr)
val_lines = [
    ("WACC", f"={REF['wacc']}", f"={REF['wacc']}", f_link_pct),
    ("Terminal growth / exit multiple", f"={REF['g']}", f"={REF['exit']}", None),
    ("Sum of PV of FCF", f"=SUM(D{RW['PV of FCF'] + 1}:{LC}{RW['PV of FCF'] + 1})", f"=C{v0 + 3}", f_num),
    ("Terminal value", f"={LC}{RW['Unlevered FCF'] + 1}*(1+C{v0 + 2})/(C{v0 + 1}-C{v0 + 2})", f"={LC}{RW['EBITDA'] + 1}*D{v0 + 2}", f_num),
    ("PV of terminal value", f"=C{v0 + 4}/(1+C{v0 + 1})^{N}", f"=D{v0 + 4}/(1+D{v0 + 1})^{N}", f_num),
    ("Enterprise value", f"=C{v0 + 3}+C{v0 + 5}", f"=D{v0 + 3}+D{v0 + 5}", f_num_b),
    ("(-) Net debt & minority", f"=-{nd}", f"=-{nd}", f_num),
    ("Equity value", f"=C{v0 + 6}+C{v0 + 7}", f"=D{v0 + 6}+D{v0 + 7}", f_num_b),
    ("Value per share", f"=C{v0 + 8}/{REF['shares']}", f"=D{v0 + 8}/{REF['shares']}", f_px_b),
    ("TV as % of EV", f"=C{v0 + 5}/C{v0 + 6}", f"=D{v0 + 5}/D{v0 + 6}", f_pct),
    ("Implied exit multiple / growth", f"=C{v0 + 4}/{LC}{RW['EBITDA'] + 1}",
     f"=(D{v0 + 4}*D{v0 + 1}-{LC}{RW['Unlevered FCF'] + 1})/(D{v0 + 4}+{LC}{RW['Unlevered FCF'] + 1})", None),
]
for i, (lab, fa, fb, f) in enumerate(val_lines):
    r = v0 + i
    wd.write(r, 1, lab, f_lbl)
    if lab.startswith("Terminal growth"):
        wd.write_formula(r, 2, fa, f_link_pct); wd.write_formula(r, 3, fb, f_link_x)
    elif lab.startswith("Implied"):
        wd.write_formula(r, 2, fa, f_x); wd.write_formula(r, 3, fb, f_pct)
    else:
        wd.write_formula(r, 2, fa, f); wd.write_formula(r, 3, fb, f)
REF["dcf_pg"], REF["dcf_em"] = f"DCF!$C${v0 + 9}", f"DCF!$D${v0 + 9}"

# Live sensitivity grids: each cell recomputes the full DCF with SUMPRODUCT
s0 = v0 + len(val_lines) + 2
last_fcf, last_ebitda = f"${LC}${RW['Unlevered FCF'] + 1}", f"${LC}${RW['EBITDA'] + 1}"
def sens_grid(r_top, title, col_vals, col_fmt, tv_expr):
    wd.write(r_top, 1, title, f_bold)
    wd.write(r_top + 1, 1, "WACC ↓ / " + ("terminal growth →" if col_fmt is f_sens_hdr_pct else "exit multiple →"), f_sub)
    for j, v in enumerate(col_vals):
        wd.write(r_top + 1, 2 + j, v, col_fmt)
    cells = []
    for i, w in enumerate(wacc_grid):
        r = r_top + 2 + i
        wd.write(r, 1, w, f_sens_hdr_pct)
        for j in range(len(col_vals)):
            W = f"$B{r + 1}"; X = f"{xl_col_to_name(2 + j)}${r_top + 2}"
            tv = tv_expr(W, X)
            wd.write_formula(r, 2 + j, f"=(SUMPRODUCT({fcf_rng}/(1+{W})^({per_rng}-0.5))+{tv}/(1+{W})^{N}-{nd})/{REF['shares']}", f_sens_px)
        cells.append(r + 1)
    return f"D{cells[1]}:F{cells[3]}"
wd.set_column("B:B", 30)
inner_pg = sens_grid(s0, "Value per share: WACC vs terminal growth", list(g_grid), f_sens_hdr_pct,
                     lambda W, X: f"{last_fcf}*(1+{X})/({W}-{X})")
inner_em = sens_grid(s0 + 9, "Value per share: WACC vs exit multiple", list(m_grid), f_sens_hdr_x,
                     lambda W, X: f"{last_ebitda}*{X}")
REF["dcf_pg_rng"], REF["dcf_em_rng"] = f"DCF!{inner_pg}", f"DCF!{inner_em}"

# ---------------- LBO ----------------
wl = wb.add_worksheet("LBO")
wl.hide_gridlines(2); wl.set_column("A:A", 2); wl.set_column("B:B", 32); wl.set_column("C:I", 12)
wl.write("B2", "LBO ability-to-pay ($mm)", f_title)
wl.write("B3", "Change the offer price on the Inputs sheet (or Goal Seek IRR) to test other prices.", f_sub)
L = [("Offer price per share", f"={REF['lbo_price']}", f_link_px),
     ("Equity purchase price", f"=C5*{REF['shares']}", f_num),
     ("(+) Net debt & minority refinanced", f"={nd}", f_num),
     ("Entry enterprise value", "=C6+C7", f_num_b),
     ("Entry EV / LTM EBITDA", f"=C8/{REF['ebitda_ltm']}", f_x),
     ("Transaction fees", f"=C8*{REF['lbo_fees']}", f_num),
     ("New debt raised", f"={REF['lbo_lev']}*{REF['ebitda_ltm']}", f_num),
     ("Sponsor equity", "=C8+C10-C11", f_num_b),
     ("Premium to current price", f"=C5/{REF['price']}-1", f_pct)]
for i, (lab, form, f) in enumerate(L):
    wl.write(4 + i, 1, lab); wl.write_formula(4 + i, 2, form, f)
wl.write_row(15, 2, ["Entry"] + list(drivers.index), f_hdr)
lrows = ["EBITDA", "D&A", "Interest", "Taxes", "Capex", "Change in NWC", "Levered FCF", "Debt repayment", "Debt (end)", "Cash (end)", "Net debt / EBITDA"]
LR = {lab: 16 + i for i, lab in enumerate(lrows)}
for lab, r in LR.items():
    wl.write(r, 1, lab, f_bold if lab in ("Levered FCF", "Debt (end)") else f_lbl)
wl.write_formula(LR["Debt (end)"], 2, "=C11", f_num)
wl.write(LR["Cash (end)"], 2, 0, f_num)
for j in range(N):
    c = 3 + j; C = xl_col_to_name(c); P = xl_col_to_name(c - 1); dC = xl_col_to_name(3 + j)
    g = lambda lab: LR[lab] + 1
    wl.write_formula(LR["EBITDA"], c, f"=DCF!{dC}{RW['EBITDA'] + 1}", f_link_num)
    wl.write_formula(LR["D&A"], c, f"=DCF!{dC}{RW['D&A'] + 1}", f_link_num)
    wl.write_formula(LR["Interest"], c, f"={P}{g('Debt (end)')}*{REF['lbo_rate']}", f_num)
    wl.write_formula(LR["Taxes"], c, f"=MAX(0,{C}{g('EBITDA')}-{C}{g('D&A')}-{C}{g('Interest')})*{REF['tax']}", f_num)
    wl.write_formula(LR["Capex"], c, f"=-DCF!{dC}{RW['(-) Capex'] + 1}", f_link_num)
    wl.write_formula(LR["Change in NWC"], c, f"=-DCF!{dC}{RW['(-) Change in NWC'] + 1}", f_link_num)
    wl.write_formula(LR["Levered FCF"], c, f"={C}{g('EBITDA')}-{C}{g('Interest')}-{C}{g('Taxes')}-{C}{g('Capex')}-{C}{g('Change in NWC')}", f_num_b)
    wl.write_formula(LR["Debt repayment"], c, f"=MIN({P}{g('Debt (end)')},MAX({C}{g('Levered FCF')},0))", f_num)
    wl.write_formula(LR["Debt (end)"], c, f"={P}{g('Debt (end)')}-{C}{g('Debt repayment')}", f_num)
    wl.write_formula(LR["Cash (end)"], c, f"={P}{g('Cash (end)')}+{C}{g('Levered FCF')}-{C}{g('Debt repayment')}", f_num)
    wl.write_formula(LR["Net debt / EBITDA"], c, f"=({C}{g('Debt (end)')}-{C}{g('Cash (end)')})/{C}{g('EBITDA')}", f_x)
LLC = xl_col_to_name(2 + N)
e0 = LR["Net debt / EBITDA"] + 3
X = [("Exit EV (exit multiple x exit EBITDA)", f"={REF['exit']}*{LLC}{LR['EBITDA'] + 1}", f_num),
     ("(-) Debt", f"=-{LLC}{LR['Debt (end)'] + 1}", f_num),
     ("(+) Cash", f"={LLC}{LR['Cash (end)'] + 1}", f_num),
     ("Exit equity value", f"=SUM(C{e0 + 1}:C{e0 + 3})", f_num_b),
     ("MOIC", f"=C{e0 + 4}/C12", f_x),
     ("IRR", f"=C{e0 + 5}^(1/{N})-1", fmt(num_format="0.0%", bold=True, bg_color="#FFF2CC", border=1))]
for i, (lab, form, f) in enumerate(X):
    wl.write(e0 + i, 1, lab); wl.write_formula(e0 + i, 2, form, f)
REF["lbo_irr"] = f"LBO!$C${e0 + 6}"

# ---------------- Football field ----------------
wf = wb.add_worksheet("Football Field")
wf.hide_gridlines(2); wf.set_column("A:A", 2); wf.set_column("B:B", 30); wf.set_column("C:E", 12)
wf.write("B2", "Valuation summary: value per share", f_title)
wf.write_row(3, 1, ["Methodology", "Low", "High", "Spread"], f_hdr)
link = {"EV / LTM EBITDA": REF["comps_ebitda"], "P / NTM EPS": REF["comps_pe"],
        "DCF – perpetuity growth": (f"MIN({REF['dcf_pg_rng']})", f"MAX({REF['dcf_pg_rng']})"),
        "DCF – exit multiple": (f"MIN({REF['dcf_em_rng']})", f"MAX({REF['dcf_em_rng']})")}
for i, (name, (lo, hi)) in enumerate(ff.iterrows()):
    r = 4 + i
    wf.write(r, 1, name)
    if name in link:
        wf.write_formula(r, 2, f"={link[name][0]}", f_link_px); wf.write_formula(r, 3, f"={link[name][1]}", f_link_px)
    else:
        wf.write(r, 2, lo, f_in_px); wf.write(r, 3, hi, f_in_px)
    wf.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", f_px)
rE = 4 + len(ff)
wf.write(rE + 1, 1, "Current share price"); wf.write_formula(rE + 1, 2, f"={REF['price']}", f_link_px)
ch = wb.add_chart({"type": "bar", "subtype": "stacked"})
ch.add_series({"categories": ["Football Field", 4, 1, rE - 1, 1], "values": ["Football Field", 4, 2, rE - 1, 2],
               "fill": {"none": True}, "border": {"none": True}})
ch.add_series({"categories": ["Football Field", 4, 1, rE - 1, 1], "values": ["Football Field", 4, 4, rE - 1, 4],
               "fill": {"color": BLUE}, "gap": 60})
ch.set_legend({"none": True}); ch.set_title({"name": "Implied value per share ($)", "name_font": {"size": 11}})
ch.set_y_axis({"reverse": True}); ch.set_x_axis({"num_format": "$#,##0", "major_gridlines": {"visible": True, "line": {"color": "#e6e4df"}}})
ch.set_size({"width": 720, "height": 380})
wf.insert_chart("G4", ch)

# ---------------- Buyers ----------------
wbuy = wb.add_worksheet("Buyers")
wbuy.hide_gridlines(2); wbuy.set_column("A:A", 2); wbuy.set_column("B:B", 12); wbuy.set_column("C:C", 34); wbuy.set_column("D:K", 14)
wbuy.write("B2", "Potential strategic acquirers", f_title)
bdf = buyers_df.reset_index()
wbuy.write_row(3, 1, list(bdf.columns), f_hdr); wbuy.set_row(3, 30)
for i, row in bdf.iterrows():
    for j, v in enumerate(row.values):
        f = f_num if j == 2 else f_x if j in (3, 5, 6) else f_pct if j == 4 else f_lbl
        wbuy.write(4 + i, 1 + j, v if not (isinstance(v, float) and np.isnan(v)) else "n.m.", f)
wb.close()
print("Saved", XLSX)

# %% [markdown]
# ## 10. PowerPoint pitchbook
# Native, editable PowerPoint tables and charts in a classic sell-side layout.

# %%
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE

NAVY, GREY, LIGHT = RGBColor(0x0B, 0x25, 0x45), RGBColor(0x52, 0x51, 0x4E), RGBColor(0xF2, 0xF4, 0xF7)
C_BLUE, C_ORANGE = RGBColor(0x2A, 0x78, 0xD6), RGBColor(0xEB, 0x68, 0x34)
PPTX = f"{OUTPUT_DIR}/{target['ticker'].replace('.', '_')}_pitchbook.pptx"
prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
BLANK = prs.slide_layouts[6]
page = [0]

def tb(slide, x, y, w, h, text, size=12, bold=False, color=GREY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        run = para.add_run(); run.text = ln
        run.font.size, run.font.bold, run.font.color.rgb, run.font.name = Pt(size), bold, color, "Arial"
    return box

def new_slide(title, subtitle=None):
    s = prs.slides.add_slide(BLANK); page[0] += 1
    tb(s, 0.5, 0.35, 12.3, 0.6, title, 24, True, NAVY)
    if subtitle:
        tb(s, 0.5, 0.95, 12.3, 0.4, subtitle, 13, False, GREY)
    ln = s.shapes.add_connector(1, Inches(0.5), Inches(1.4), Inches(12.83), Inches(1.4))
    ln.line.color.rgb, ln.line.width = NAVY, Pt(1.5)
    tb(s, 0.5, 7.0, 10.5, 0.3, f"Source: {DATA_LABEL}. Illustrative analysis for educational purposes.", 8)
    tb(s, 12.3, 7.0, 0.53, 0.3, str(page[0]), 9, align=PP_ALIGN.RIGHT)
    if IS_DEMO:
        tb(s, 9.3, 0.4, 3.53, 0.3, "SYNTHETIC DEMO DATA", 10, True, C_ORANGE, PP_ALIGN.RIGHT)
    return s

def table(slide, df, x, y, w, col_w=None, font=10, row_h=0.3, bold_rows=(), first_col_w=None):
    rows, cols = df.shape[0] + 1, df.shape[1] + 1
    shp = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_h * rows))
    t = shp.table
    widths = col_w or ([first_col_w or w * 0.3] + [(w - (first_col_w or w * 0.3)) / (cols - 1)] * (cols - 1))
    for j, cw in enumerate(widths):
        t.columns[j].width = Inches(cw)
    hdr = [df.index.name or ""] + list(df.columns)
    for i in range(rows):
        t.rows[i].height = Inches(row_h)
        for j in range(cols):
            cell = t.cell(i, j)
            txt = hdr[j] if i == 0 else (str(df.index[i - 1]) if j == 0 else str(df.iat[i - 1, j - 1]))
            cell.text = txt
            cell.margin_left = cell.margin_right = Inches(0.06); cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            numeric = txt[:1] in "$-–(0123456789" or txt in ("n.m.", "") or txt[-1:] in "x%"
            p.alignment = PP_ALIGN.RIGHT if (j > 0 and numeric) else PP_ALIGN.LEFT
            f = p.runs[0].font if p.runs else p.font
            f.size, f.name = Pt(font), "Arial"
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = NAVY; f.color.rgb = RGBColor(255, 255, 255); f.bold = True
                p.alignment = PP_ALIGN.CENTER if j else PP_ALIGN.LEFT
            else:
                is_bold = (i - 1) in bold_rows or df.index[i - 1] in bold_rows
                cell.fill.fore_color.rgb = LIGHT if is_bold else RGBColor(255, 255, 255)
                f.color.rgb = RGBColor(0x0B, 0x0B, 0x0B); f.bold = is_bold
    return shp

money = lambda v: "n.m." if v is None or pd.isna(v) else f"{v:,.0f}"
px_ = lambda v: "n.m." if v is None or pd.isna(v) else f"${v:,.2f}"
xx = lambda v: "n.m." if v is None or pd.isna(v) else f"{v:.1f}x"
pc = lambda v: "n.m." if v is None or pd.isna(v) else f"{v:.1%}"

# Reference range: overlap of intrinsic & market methods (excludes 52w range and analyst targets)
core = ff.drop(index=[i for i in ["52-week trading range", "Analyst price targets", "EV / LTM Revenue"] if i in ff.index])
ref_lo, ref_hi = core["Low"].median(), core["High"].median()

# 1. Cover
s = prs.slides.add_slide(BLANK); page[0] += 1
bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
tb(s, 0.8, 2.3, 11.5, 0.8, PROJECT_NAME, 40, True, RGBColor(255, 255, 255))
tb(s, 0.8, 3.2, 11.5, 0.6, f"Preliminary valuation materials: {target['name']}", 22, False, RGBColor(0xDC, 0xE6, 0xF1))
tb(s, 0.8, 4.0, 11.5, 0.5, f"{TODAY:%B %Y}  |  Draft, illustrative and for discussion purposes only", 14, False, RGBColor(0xB7, 0xC4, 0xD6))
if IS_DEMO:
    tb(s, 0.8, 6.6, 11.5, 0.4, "Built on SYNTHETIC DEMO DATA: set DATA_MODE = 'live' for real market data", 12, True, C_ORANGE)

# 2. Executive summary
s = new_slide("Executive summary", f"{target['name']} ({target['ticker']})")
prem_lo, prem_hi = ref_lo / target["price"] - 1, ref_hi / target["price"] - 1
bul = [
    f"{target['name']} trades at ${target['price']:.2f} per share: market cap ${target['mcap']:,.0f}mm and EV ${target['ev']:,.0f}mm, "
    f"or {target_mult['EV_EBITDA']:.1f}x LTM EBITDA vs a peer median of {comp_stats.loc['Median', 'EV_EBITDA']:.1f}x.",
    f"Preliminary reference range of ${ref_lo:.2f}–${ref_hi:.2f} per share ({prem_lo:+.0%} to {prem_hi:+.0%} vs current), "
    f"the median of trading comps, transaction and DCF ranges.",
    f"DCF at a {WACC:.1%} WACC: ${d_pg['price']:.2f} (perpetuity growth {g_term:.1%}) and ${d_em['price']:.2f} "
    f"(exit multiple {EXIT_MULT:.1f}x); terminal value is {d_pg['tv_share']:.0%} of EV.",
    f"A financial sponsor at {A['lbo_leverage']:.1f}x leverage could pay up to ${lbo_price:.2f} per share "
    f"({lbo_price / target['price'] - 1:+.0%}) for a {A['lbo_target_irr']:.0%} IRR, which sets a floor for a competitive process.",
    f"{(buyers_df['All-cash feasible'] == 'Yes').sum()} of {len(buyers_df)} strategic buyers screened could fund an all-cash deal "
    f"at the mid-point premium within {A['buyer_max_leverage']:.1f}x pro forma leverage.",
]
box = tb(s, 0.7, 1.8, 11.9, 4.8, bul, 16, False, RGBColor(0x0B, 0x0B, 0x0B))
for para in box.text_frame.paragraphs:
    para.space_after = Pt(14)
    r = para.runs[0]; r.text = "■  " + r.text

# 3. Company overview
s = new_slide("Company overview", f"{target['sector']} | {target['industry']}")
kd_tbl = pd.DataFrame({"Value": [px_(target["price"]), f"{px_(target['low52'])} – {px_(target['high52'])}",
                                  money(target["mcap"]), money(target["net_debt"]), money(target["ev"]),
                                  money(target["revenue_ltm"]), money(target["ebitda_ltm"]),
                                  pc(target["ebitda_ltm"] / target["revenue_ltm"]), xx(target_mult["EV_EBITDA"]),
                                  xx(target_mult["PE_NTM"])]},
                      index=["Share price", "52-week range", "Market cap ($mm)", "Net debt ($mm)", "Enterprise value ($mm)",
                             "LTM revenue ($mm)", "LTM EBITDA ($mm)", "EBITDA margin", "EV / LTM EBITDA", "P / NTM EPS"])
kd_tbl.index.name = "Key market data"
table(s, kd_tbl, 0.5, 1.7, 4.6, font=11, row_h=0.36, first_col_w=2.6)
cd = CategoryChartData()
cd.categories = [f"FY{y}" for y in H.index]
cd.add_series("Revenue ($mm)", [round(v) for v in H["Revenue"]])
cd.add_series("EBITDA ($mm)", [round(v) for v in H["EBITDA"]])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5.5), Inches(1.6), Inches(7.3), Inches(3.6), cd)
chart = gf.chart
chart.has_legend = True; chart.legend.position = XL_LEGEND_POSITION.TOP; chart.legend.include_in_layout = False
chart.legend.font.size = Pt(10)
chart.has_title = True; chart.chart_title.text_frame.text = "Historical revenue and EBITDA ($mm)"
chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
for ser, col in zip(chart.series, [C_BLUE, C_ORANGE]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
    ser.data_labels.show_value = True; ser.data_labels.font.size = Pt(9); ser.data_labels.number_format = '#,##0'
    ser.data_labels.number_format_is_linked = False; ser.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
chart.plots[0].gap_width = 80; chart.plots[0].overlap = -10
chart.value_axis.visible = False; chart.value_axis.has_major_gridlines = False
chart.category_axis.tick_labels.font.size = Pt(10)
summ = (target["summary"][:600] + "…") if len(target["summary"]) > 600 else target["summary"]
tb(s, 5.5, 5.35, 7.3, 1.5, summ, 10)

# 4. Trading comps
s = new_slide("Trading comparables", "Peers valued on LTM and NTM multiples")
ct = comps.copy()
ct_disp = pd.DataFrame({
    "Mkt cap ($mm)": ct["MktCap"].map(money), "EV ($mm)": ct["EV"].map(money), "EBITDA margin": ct["Margin"].map(pc),
    "EV / Revenue": ct["EV_Rev"].map(xx), "EV / EBITDA": ct["EV_EBITDA"].map(xx), "P/E LTM": ct["PE_LTM"].map(xx),
    "P/E NTM": ct["PE_NTM"].map(xx), "Net debt / EBITDA": ct["NetDebt_EBITDA"].map(xx)}, index=ct.index)
for lab in ["Mean", "Median", "25th pct", "75th pct"]:
    ct_disp.loc[lab] = ["", "", "", *[xx(comp_stats.loc[lab, c]) for c in mult_cols], ""]
ct_disp.loc[f"{target['ticker']} (target)"] = [money(target["mcap"]), money(target["ev"]), pc(target_mult["Margin"]),
    xx(target_mult["EV_Rev"]), xx(target_mult["EV_EBITDA"]), xx(target_mult["PE_LTM"]), xx(target_mult["PE_NTM"]),
    xx(target_mult["NetDebt_EBITDA"])]
ct_disp.index.name = "Company"
table(s, ct_disp, 0.5, 1.65, 12.33, font=10, row_h=0.3, first_col_w=2.3,
      bold_rows=("Median", f"{target['ticker']} (target)"))
imp = pd.DataFrame(implied, index=["Low", "High"]).T.loc[["EV / LTM EBITDA", "P / NTM EPS"]]
tb(s, 0.5, 1.65 + 0.3 * (len(ct_disp) + 1) + 0.2, 12.3, 0.5,
   f"Implied value per share (25th–75th percentile):  EV/LTM EBITDA {px_(imp.iloc[0, 0])} – {px_(imp.iloc[0, 1])}   |   "
   f"P/NTM EPS {px_(imp.iloc[1, 0])} – {px_(imp.iloc[1, 1])}", 12, True, NAVY)

# 5. Precedents & premiums
s = new_slide("Precedent transactions and premiums paid", "Control value benchmarks")
if len(prec):
    pt = prec.rename(columns={"date": "Date", "acquirer": "Acquirer", "target": "Target", "ev_ebitda": "EV / LTM EBITDA"}).copy()
    pt["EV / LTM EBITDA"] = pt["EV / LTM EBITDA"].map(xx)
    pt = pt.set_index("Date")
    table(s, pt, 0.5, 1.7, 7.4, font=10, row_h=0.32, first_col_w=1.0)
    stat_txt = [f"Median EV / LTM EBITDA: {prec['ev_ebitda'].median():.1f}x",
                f"Interquartile range: {q25:.1f}x – {q75:.1f}x",
                f"Implied value per share: {px_(implied['Precedent transactions'][0])} – {px_(implied['Precedent transactions'][1])}"]
else:
    stat_txt = ["No precedent transactions provided.", "Populate PRECEDENTS with deals from EDGAR merger proxies (DEFM14A) or Capital IQ."]
tb(s, 8.3, 1.8, 4.5, 1.8, stat_txt, 13, False, RGBColor(0x0B, 0x0B, 0x0B))
tb(s, 8.3, 3.9, 4.5, 0.4, "Premiums paid analysis", 14, True, NAVY)
pp = pd.DataFrame({"Price": [px_(target["price"] * (1 + p)) for p in [0, lo_p, np.mean(A["premium_range"]), hi_p]]},
                  index=["Unaffected (current)", f"{lo_p:.0%} premium", f"{np.mean(A['premium_range']):.0%} premium", f"{hi_p:.0%} premium"])
pp.index.name = "Premium"
table(s, pp, 8.3, 4.4, 4.5, font=11, row_h=0.36, first_col_w=2.7)

# 6. WACC
s = new_slide("Weighted average cost of capital", f"WACC of {WACC:.2%}")
wt = pd.DataFrame({"Value": [pc(risk_free), pc(A["erp"]), f"{b2w:.2f} / {b5m:.2f}", f"{unlev_med:.2f}", pc(de_target),
                             f"{beta_relev:.2f}", pc(ke), f"{rating} ({coverage:.1f}x)", pc(kd), pc(TAX), pc(wE), f"{WACC:.2%}"]},
                  index=["Risk-free rate (10y UST)", "Equity risk premium", "Raw beta: 2y weekly / 5y monthly", "Peer median unlevered beta",
                         "Target D/E (market)", "Relevered beta", "Cost of equity", "Synthetic rating (coverage)",
                         "Pre-tax cost of debt", "Tax rate", "Equity / total capital", "WACC"])
wt.index.name = "Component"
table(s, wt, 0.5, 1.7, 5.8, font=11, row_h=0.36, first_col_w=3.6, bold_rows=("Cost of equity", "WACC"))
pb = peer_beta.copy()
pb_disp = pd.DataFrame({"Raw beta": pb["RawBeta"].map(lambda v: f"{v:.2f}"), "Adj. beta": pb["AdjBeta"].map(lambda v: f"{v:.2f}"),
                        "D/E": pb["DE"].map(pc), "Unlevered": pb["Unlevered"].map(lambda v: f"{v:.2f}")}, index=pb.index)
pb_disp.loc["Median"] = ["", "", "", f"{unlev_med:.2f}"]
pb_disp.index.name = "Peer"
table(s, pb_disp, 6.9, 1.7, 5.9, font=11, row_h=0.36, first_col_w=1.9, bold_rows=("Median",))
tb(s, 6.9, 1.7 + 0.36 * (len(pb_disp) + 1) + 0.25, 5.9, 1.2,
   ["Betas are Blume-adjusted (0.67 × raw + 0.33) and unlevered with Hamada: βu = βl / (1 + (1 − t) × D/E).",
    "Cost of debt uses Damodaran's interest-coverage synthetic rating table."], 10)

# 7. DCF
s = new_slide("Discounted cash flow analysis", "Unlevered free cash flow projections ($mm)")
pr = proj[["Revenue", "EBITDA", "D&A", "EBIT", "Taxes", "Capex", "Change in NWC", "Unlevered FCF"]].T
pr.insert(0, f"FY{base_year}A", [H["Revenue"].iloc[-1], H["EBITDA"].iloc[-1], H["D&A"].iloc[-1], H["EBIT"].iloc[-1],
                                 np.nan, H["Capex"].iloc[-1], H["Change in NWC"].iloc[-1], np.nan])
pr_disp = pr.map(lambda v: "–" if pd.isna(v) else money(v))
pr_disp.loc["Revenue growth"] = [""] + [pc(v) for v in drivers["Revenue growth"]]
pr_disp.loc["EBITDA margin"] = [pc(H["EBITDA"].iloc[-1] / H["Revenue"].iloc[-1])] + [pc(v) for v in drivers["EBITDA margin"]]
pr_disp = pr_disp.loc[["Revenue", "Revenue growth", "EBITDA", "EBITDA margin", "D&A", "EBIT", "Taxes", "Capex", "Change in NWC", "Unlevered FCF"]]
pr_disp.index.name = "$mm"
table(s, pr_disp, 0.5, 1.65, 8.2, font=10, row_h=0.32, first_col_w=2.0, bold_rows=("Revenue", "EBITDA", "Unlevered FCF"))
vt = pd.DataFrame({"Perpetuity": [pc(WACC), pc(g_term), money(d_pg["pv_fcf"]), money(d_pg["pv_tv"]), money(d_pg["ev"]),
                                  money(-(target["net_debt"] + target["minority"])), px_(d_pg["price"]), pc(d_pg["tv_share"])],
                   "Exit multiple": [pc(WACC), xx(EXIT_MULT), money(d_em["pv_fcf"]), money(d_em["pv_tv"]), money(d_em["ev"]),
                                     money(-(target["net_debt"] + target["minority"])), px_(d_em["price"]), pc(d_em["tv_share"])]},
                  index=["WACC", "g / exit multiple", "PV of FCF", "PV of terminal value", "Enterprise value", "(–) Net debt & minority",
                         "Value per share", "TV % of EV"])
vt.index.name = "Valuation"
table(s, vt, 9.0, 1.65, 3.83, font=10, row_h=0.32, first_col_w=1.75, bold_rows=("Enterprise value", "Value per share"))
tb(s, 0.5, 5.3, 12.3, 0.8, [f"Implied exit multiple from perpetuity method: {implied_exit:.1f}x  |  "
                              f"Implied perpetuity growth from exit method: {implied_g:.2%}",
                              "Mid-year convention; terminal value discounted from the end of the final projection year."], 11)

# 8. DCF sensitivity
s = new_slide("DCF sensitivity analysis", "Value per share")
sp = sens_pg.map(px_); sp.index = [pc(i) for i in sens_pg.index]; sp.columns = [pc(c) for c in sens_pg.columns]; sp.index.name = "WACC \\ g"
se = sens_em.map(px_); se.index = [pc(i) for i in sens_em.index]; se.columns = [xx(c) for c in sens_em.columns]; se.index.name = "WACC \\ exit"
tb(s, 0.5, 1.65, 6, 0.4, "WACC vs terminal growth", 14, True, NAVY)
table(s, sp, 0.5, 2.1, 5.9, font=11, row_h=0.42, first_col_w=1.3, bold_rows=(2,))
tb(s, 6.9, 1.65, 6, 0.4, "WACC vs exit multiple (EV / EBITDA)", 14, True, NAVY)
table(s, se, 6.9, 2.1, 5.9, font=11, row_h=0.42, first_col_w=1.3, bold_rows=(2,))
tb(s, 0.5, 5.0, 12.3, 0.5, f"Current share price: ${target['price']:.2f}. Highlighted row = base-case WACC. Exit grid at multiples rounded to 0.1x; "
                           f"at the exact base-case multiple ({EXIT_MULT:.2f}x) the value is ${d_em['price']:.2f}.", 11)

# 9. LBO
s = new_slide("LBO analysis: financial sponsor ability to pay",
              f"Max price at {A['lbo_target_irr']:.0%} IRR and {A['lbo_leverage']:.1f}x leverage: ${lbo_price:.2f} per share")
sch = lbo_base["schedule"].T.loc[["EBITDA", "Interest", "Taxes", "LFCF", "Repayment", "Debt", "Cash", "Leverage"]]
sch_disp = sch.map(money); sch_disp.loc["Leverage"] = sch.loc["Leverage"].map(xx)
sch_disp.index = ["EBITDA", "Interest", "Taxes", "Levered FCF", "Debt repayment", "Debt (end)", "Cash (end)", "Net debt / EBITDA"]
sch_disp.index.name = "$mm"
table(s, sch_disp, 0.5, 1.65, 7.6, font=10, row_h=0.32, first_col_w=2.0, bold_rows=("Levered FCF", "Debt (end)"))
sl = sens_lbo.map(px_); sl.index = [pc(i) for i in sens_lbo.index]; sl.columns = [xx(c) for c in sens_lbo.columns]
sl.index.name = "IRR \\ leverage"
tb(s, 8.5, 1.65, 4.3, 0.4, "Max price per share", 13, True, NAVY)
table(s, sl, 8.5, 2.05, 4.33, font=10, row_h=0.34, first_col_w=1.1, bold_rows=(2,))
tb(s, 0.5, 4.95, 12.3, 1.6,
   [f"Entry EV ${lbo_ev:,.0f}mm ({lbo_ev / target['ebitda_ltm']:.1f}x LTM EBITDA), sponsor equity ${lbo_base['equity0']:,.0f}mm "
    f"({lbo_base['equity0'] / (lbo_ev * (1 + A['lbo_fees'])):.0%} of total funding), MOIC {lbo_base['moic']:.2f}x over {N} years.",
    f"Assumes {A['lbo_interest']:.1%} blended cost of debt, 100% cash sweep, {A['lbo_fees']:.0%} fees and exit at {EXIT_MULT:.1f}x EBITDA (peer median)."], 11)

# 10. Football field
s = new_slide("Valuation summary", f"Preliminary reference range ${ref_lo:.2f} – ${ref_hi:.2f} per share")
from PIL import Image
iw, ih = Image.open(f"{OUTPUT_DIR}/football_field.png").size
pic_h = min(5.25, 12.1 * ih / iw); pic_w = pic_h * iw / ih
s.shapes.add_picture(f"{OUTPUT_DIR}/football_field.png", Inches((13.333 - pic_w) / 2), Inches(1.6), height=Inches(pic_h))

# 11. Buyers
s = new_slide("Potential strategic acquirers", f"Screened at ${mid_offer:.2f} per share ({np.mean(A['premium_range']):.0%} premium)")
bd = pd.DataFrame({"Company": buyers_df["Company"].str.slice(0, 28), "Mkt cap ($mm)": buyers_df["Mkt cap ($mm)"].map(money),
                   "Net debt / EBITDA": buyers_df["Net debt / EBITDA"].map(xx),
                   "Deal / mkt cap": buyers_df["Deal size / buyer mkt cap"].map(pc),
                   "PF leverage (all-debt)": buyers_df["PF leverage (all-debt)"].map(xx),
                   "Buyer P/E NTM": buyers_df["Buyer P/E NTM"].map(xx),
                   "All-cash feasible": buyers_df["All-cash feasible"], "All-stock EPS": buyers_df["All-stock EPS"]}, index=buyers_df.index)
bd.index.name = "Ticker"
table(s, bd, 0.5, 1.7, 12.33, font=10, row_h=0.4, col_w=[1.2, 2.9, 1.3, 1.3, 1.2, 1.35, 1.1, 1.0, 0.98])
tb(s, 0.5, 1.7 + 0.4 * (len(bd) + 1) + 0.3, 12.3, 1.2,
   [f"Offer P/E NTM at the mid-point premium: {offer_pe:.1f}x. An all-stock deal is accretive before synergies when the buyer's P/E is higher.",
    f"All-cash feasible = pro forma net debt / EBITDA ≤ {A['buyer_max_leverage']:.1f}x assuming 100% debt funding. "
    "Strategic fit and antitrust overlap need qualitative review."], 11)

# 12. Methodology
s = new_slide("Methodology and key assumptions")
tb(s, 0.7, 1.8, 11.9, 5.0, [
    "Trading comparables: LTM multiples from Yahoo Finance; multiples above 100x or with negative denominators excluded; 25th–75th percentile applied.",
    "Premiums paid: one-day premium range applied to the current (unaffected) share price; refine with a deal-specific precedent set.",
    f"DCF: {N}-year driver-based projections, mid-year convention, perpetuity growth {g_term:.1%} and exit multiple {EXIT_MULT:.1f}x (peer median).",
    "WACC: CAPM with peer-derived relevered beta; cost of debt from interest-coverage synthetic rating.",
    f"LBO: max entry price solved by bisection at the target IRR, with exit at {EXIT_MULT:.1f}x EBITDA (peer median).",
    "Limitations: data vendor figures are not reconciled to filings; no stub-period adjustment; share count excludes dilution from options/RSUs.",
], 14, False, RGBColor(0x0B, 0x0B, 0x0B))
for para in list(s.shapes)[-1].text_frame.paragraphs:
    para.space_after = Pt(12); para.runs[0].text = "■  " + para.runs[0].text

cp = prs.core_properties
cp.title = f"{PROJECT_NAME}: preliminary valuation materials ({target['name']})"
cp.subject = "Sell-side valuation pitchbook"
cp.author = cp.last_modified_by = AUTHOR
cp.keywords = "M&A, valuation, pitchbook, DCF, LBO"
cp.comments = "Generated by the Sell-Side Valuation & Pitchbook Engine"
cp.category = "Pitchbook"
cp.created = cp.modified = dt.datetime.now()
cp.revision = 1
prs.save(PPTX)
print("Saved", PPTX)

# %% [markdown]
# ## 11. Download outputs
# In Colab the files download automatically. Locally, they are in `pitchbook_output/`.

# %%
try:
    from google.colab import files
    for fpath in [XLSX, PPTX, f"{OUTPUT_DIR}/football_field.png"]:
        files.download(fpath)
except ImportError:
    print("Outputs saved in:", os.path.abspath(OUTPUT_DIR))
