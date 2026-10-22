"""
Calculation engine for the Enterprise Emissions Analysis Tool.

Design rules
------------
* Columns are identified by MEANING (header text), never by Excel column letter.
* Every number is computed from the raw rows - nothing is copied from the source workbook's
  own formula cells (totals, ranks and "Data Status" are recomputed).
* Anything the tool repairs, skips or cannot trust is written to a data-quality log, never hidden.
* The engine is independent of the user interface (app.py) so it can be tested on its own.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------------------
# 1. Column recognition
# --------------------------------------------------------------------------------------
ROLE_PATTERNS = {
    "name":     [r"^(sme)?name$", r"^smename$", r"^enterprise(name)?$", r"^company(name)?$", r"^factory(name)?$",
                 r"^unit(name)?$", r"^organi[sz]ation(name)?$"],
    "size":     [r"^size(category)?$", r"^enterprisesize$", r"^sizeclass$", r"^category$"],
    "sector":   [r"^sector$", r"^industry$", r"^industrysector$"],
    "location": [r"^location$", r"^city$", r"^area$", r"^region$", r"^district$"],
    "process":  [r"^operations?(process)?$", r"^process(es)?$", r"^operation$"],
    "scope1":   [r"^scope1", r"^direct(emissions)?"],
    "scope2":   [r"^scope2", r"^(purchased|indirect)?electricity"],
    "scope3":   [r"^scope3", r"^valuechain"],
    "total":    [r"^total(carbon)?(footprints?|emissions?|co2e?)", r"^totalcf", r"^totalco2"],
    "solar_kw": [r"solar.*(capacity|kw)", r"^installedsolar"],
    "offset":   [r"^co2offset", r"^offset"],
}
ROLE_LABEL = {"scope1": "Scope 1", "scope2": "Scope 2", "scope3": "Scope 3"}


def norm_header(h) -> str:
    return re.sub(r"[^a-z0-9]", "", str(h).lower()) if h is not None and not (isinstance(h, float) and np.isnan(h)) else ""


def match_role(header) -> str | None:
    n = norm_header(header)
    if not n:
        return None
    for role, pats in ROLE_PATTERNS.items():
        for p in pats:
            if re.search(p, n):
                return role
    return None


def auto_map(columns) -> dict:
    """Return {role: column_name}. First matching column wins; user may override in the UI."""
    m = {}
    for c in columns:
        r = match_role(c)
        if r and r not in m:
            m[r] = c
    return m


# --------------------------------------------------------------------------------------
# 2. Reading workbooks (header row may not be row 1; side tables may sit to the right / below)
# --------------------------------------------------------------------------------------
SUMMARY_WORDS = re.compile(r"^(sum|total|totals|%|median|mean|average|avg|minimum|maximum|min|max|count|small|medium|large|scope\s*\d?)\.?$", re.I)


def detect_header_row(raw: pd.DataFrame, scan: int = 25) -> int | None:
    best, best_n = None, 0
    for i in range(min(scan, len(raw))):
        n = sum(1 for v in raw.iloc[i].tolist() if match_role(v))
        if n > best_n:
            best, best_n = i, n
    return best if best_n >= 3 else None


@dataclass
class SheetTable:
    source: str
    sheet: str
    df: pd.DataFrame
    mapping: dict
    notes: list = field(default_factory=list)


def read_workbook(file, filename: str) -> list[SheetTable]:
    """Read every sheet that looks like an enterprise data table. `file` is a path or file-like object."""
    out = []
    if filename.lower().endswith((".csv", ".txt")):
        sheets = {"CSV": pd.read_csv(file, header=None, dtype=object)}
    else:
        sheets = pd.read_excel(file, sheet_name=None, header=None, dtype=object)
    for sname, raw in sheets.items():
        raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        if raw.empty:
            continue
        h = detect_header_row(raw)
        if h is None:
            out.append(SheetTable(filename, sname, pd.DataFrame(), {}, ["No recognisable header row - sheet skipped"]))
            continue
        header = raw.iloc[h].tolist()
        cols, seen = [], {}
        for j, v in enumerate(header):
            nm = str(v).strip() if (v is not None and not (isinstance(v, float) and np.isnan(v))) else f"(unnamed column {j+1})"
            if nm in seen:
                seen[nm] += 1
                nm = f"{nm} ({seen[nm]})"
            else:
                seen[nm] = 1
            cols.append(nm)
        body = raw.iloc[h + 1:].copy()
        body.columns = cols
        m = auto_map(cols)
        key = m.get("name")
        notes = []
        # stop at the first blank name row (a side/summary table usually starts after a gap)
        if key is not None:
            blank = body[key].isna() | (body[key].astype(str).str.strip() == "")
            stop = int(np.argmax(blank.values)) if blank.any() else len(body)
            body = body.iloc[:stop]
            summ = body[key].astype(str).str.strip().str.match(SUMMARY_WORDS)
            if summ.any():
                notes.append(f"{int(summ.sum())} summary-style row(s) skipped")
                body = body[~summ]
        body = body.reset_index(drop=True)
        out.append(SheetTable(filename, sname, body, m, notes))
    return out


def is_enterprise_table(t: SheetTable) -> bool:
    """A sheet is treated as enterprise data only if it has a name column and at least two emission columns."""
    return (not t.df.empty) and "name" in t.mapping and sum(k in t.mapping for k in ("scope1", "scope2", "scope3")) >= 2


# --------------------------------------------------------------------------------------
# 3. Cleaning numbers
# --------------------------------------------------------------------------------------
MISSING_TOKENS = {"", "n/a", "na", "n.a.", "nan", "none", "null", "-", "--", "—", "not recorded", "nr", "pending", "tbd"}


def parse_number(v):
    """-> (value or NaN, status) where status in {'ok','missing','repaired','invalid'}."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan, "missing"
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return float(v), "ok"
    s = str(v).strip()
    if s.lower() in MISSING_TOKENS:
        return np.nan, "missing"
    try:
        return float(s), "ok"
    except ValueError:
        pass
    # conservative repairs: thousands separators, trailing dot ("4.51."), surrounding spaces
    t = s.replace(",", "").replace(" ", "")
    t2 = t.rstrip(".")
    for cand in (t, t2):
        try:
            return float(cand), "repaired"
        except ValueError:
            continue
    return np.nan, "invalid"


def clean_numeric(series: pd.Series, label: str, names: pd.Series, log: list, repair: bool = True) -> pd.Series:
    vals = []
    for v, nm in zip(series.tolist(), names.tolist()):
        x, st = parse_number(v)
        if st == "repaired":
            if repair:
                log.append(dict(Severity="Warning", Check="Value repaired", Enterprise=nm, Field=label,
                                Detail=f"Raw value '{v}' was read as {x:g}. Please confirm against the source."))
            else:
                log.append(dict(Severity="Error", Check="Unreadable value", Enterprise=nm, Field=label,
                                Detail=f"Raw value '{v}' is not a clean number; treated as missing."))
                x = np.nan
        elif st == "invalid":
            log.append(dict(Severity="Error", Check="Unreadable value", Enterprise=nm, Field=label,
                            Detail=f"Raw value '{v}' could not be read as a number; treated as missing."))
        vals.append(x)
    return pd.Series(vals, index=series.index, dtype=float)


SIZE_ORDER = ["Small", "Medium", "Large"]


def norm_size(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "":
        return "Unspecified"
    s = str(v).strip().title()
    return s


# --------------------------------------------------------------------------------------
# 4. Building the standard dataset
# --------------------------------------------------------------------------------------
@dataclass
class Settings:
    metrics: list                      # list of standard metric column labels, e.g. ["Scope 1","Scope 2","Scope 3"]
    exclude_sizes: tuple = ("Large",)  # groups left out of the "SME" analysis
    missing_as_zero: bool = True       # report convention: N/A counts as 0 in totals AND in per-enterprise means
    repair_values: bool = True
    drop_duplicates: bool = True
    iqr_k: float = 1.5
    top_n: int = 5
    tol: float = 0.011                 # tolerance when comparing reported vs recomputed totals (rounding)


def build_dataset(tables: list[tuple[SheetTable, dict, str | None]], settings: Settings):
    """
    tables: list of (SheetTable, mapping, sector_label_or_None). `mapping` = {role: column} (may be edited by user);
            extra numeric columns can be passed as mapping['extra:<label>'] = column.
    Returns (data, log) where data has standard columns.
    """
    log, frames = [], []
    for st, mp, sector_label in tables:
        d = st.df
        if d.empty or "name" not in mp:
            continue
        out = pd.DataFrame(index=d.index)
        out["Source"] = f"{st.source} / {st.sheet}"
        names = d[mp["name"]].astype(str).str.strip()
        out["Enterprise"] = names
        out["Sector"] = d[mp["sector"]].astype(str).str.strip() if "sector" in mp else (sector_label or st.sheet)
        out["Size"] = d[mp["size"]].map(norm_size) if "size" in mp else "Unspecified"
        out["Location"] = d[mp["location"]].astype(str).str.strip() if "location" in mp else "Unspecified"
        out["Process"] = d[mp["process"]].astype(str).str.strip() if "process" in mp else "Unspecified"
        for key, col in mp.items():
            if key in ("scope1", "scope2", "scope3"):
                lbl = ROLE_LABEL[key]
            elif key.startswith("extra:"):
                lbl = key[6:]
            else:
                continue
            out[lbl] = clean_numeric(d[col], lbl, names, log, settings.repair_values)
        if "total" in mp:
            out["Reported total"] = clean_numeric(d[mp["total"]], "Reported total", names, [], True)
        for opt, lbl in (("solar_kw", "Solar kW"), ("offset", "CO2 offset")):
            if opt in mp:
                out[lbl] = clean_numeric(d[mp[opt]], lbl, names, log, True)
        frames.append(out)
    if not frames:
        return pd.DataFrame(), log
    data = pd.concat(frames, ignore_index=True)
    for m in settings.metrics:
        if m not in data:
            data[m] = np.nan
    return finalise(data, settings, log)


def finalise(data: pd.DataFrame, settings: Settings, log: list):
    metrics = settings.metrics
    # duplicates (same sector + enterprise + identical numbers)
    key = ["Sector", "Enterprise"] + metrics
    dup_exact = data.duplicated(subset=key, keep="first")
    if dup_exact.any():
        names = ", ".join(sorted(set(data.loc[dup_exact, "Enterprise"]))[:8])
        log.append(dict(Severity="Warning", Check="Duplicate records", Enterprise="", Field="",
                        Detail=f"{int(dup_exact.sum())} exact duplicate row(s) ({names}...). "
                               + ("Removed." if settings.drop_duplicates else "Kept (removal switched off).")))
        if settings.drop_duplicates:
            data = data[~dup_exact].reset_index(drop=True)
    dup_name = data.duplicated(subset=["Sector", "Enterprise"], keep=False)
    if dup_name.any():
        for nm in sorted(set(data.loc[dup_name, "Enterprise"])):
            log.append(dict(Severity="Warning", Check="Same enterprise, different values", Enterprise=nm, Field="",
                            Detail="Appears more than once in the same sector with different numbers."))
    # missing counts
    for m in metrics:
        k = int(data[m].isna().sum())
        if k:
            log.append(dict(Severity="Info", Check="Missing values", Enterprise="", Field=m,
                            Detail=f"{k} of {len(data)} enterprise(s) have no value for {m} "
                                   + ("(counted as 0, as in the report)." if settings.missing_as_zero else "(left out of that metric's statistics).")))
    neg = (data[metrics] < 0).any(axis=1)
    for nm in data.loc[neg, "Enterprise"]:
        log.append(dict(Severity="Error", Check="Negative value", Enterprise=nm, Field="", Detail="A negative emission value was found."))
    # computed total
    anyv = data[metrics].notna().any(axis=1)
    data["Total"] = data[metrics].sum(axis=1, min_count=1)  # NaN only if every metric is missing
    data["Status"] = np.where(anyv, "Available", "Pending Verification")
    if (~anyv).any():
        log.append(dict(Severity="Warning", Check="No data", Enterprise=", ".join(data.loc[~anyv, "Enterprise"][:6]), Field="",
                        Detail=f"{int((~anyv).sum())} enterprise(s) have no emission values and are excluded from statistics."))
    # reconcile reported total vs recomputed total
    if "Reported total" in data:
        diff = (data["Reported total"] - data["Total"]).abs()
        bad = data[(diff > settings.tol) & data["Reported total"].notna() & data["Total"].notna()]
        for _, r in bad.iterrows():
            comp = " + ".join(f"{r[m]:g}" if pd.notna(r[m]) else "N/A" for m in metrics)
            log.append(dict(Severity="Error", Check="Total does not equal the sum of its parts", Enterprise=r["Enterprise"], Field="Total",
                            Detail=f"Reported total {r['Reported total']:g}, but {comp} = {r['Total']:g} (difference {r['Reported total']-r['Total']:+.2f})."))
    # unrecognised size labels
    odd = sorted(set(data["Size"]) - set(SIZE_ORDER) - {"Unspecified"})
    if odd:
        log.append(dict(Severity="Info", Check="Size labels", Enterprise="", Field="Size", Detail=f"Non-standard size labels found: {', '.join(odd)}."))
    # location/process cells that contain a size label or look like a column mix-up
    for col in ("Location", "Process"):
        bad = data[data[col].isin(SIZE_ORDER)]
        for _, r in bad.iterrows():
            log.append(dict(Severity="Warning", Check="Suspicious category value", Enterprise=r["Enterprise"], Field=col,
                            Detail=f"{col} is '{r[col]}', which looks like a size label - possible data-entry or column mix-up."))
    # same enterprise listed with different size in different sources
    cs = data.groupby("Enterprise")["Size"].nunique()
    for nm in cs[cs > 1].index:
        log.append(dict(Severity="Warning", Check="Conflicting size", Enterprise=nm, Field="Size", Detail="Listed with different size categories."))
    return data, log


def select_population(data: pd.DataFrame, settings: Settings, only_sizes=None) -> pd.DataFrame:
    d = data[data["Status"] == "Available"].copy()
    if settings.exclude_sizes:
        d = d[~d["Size"].isin(settings.exclude_sizes)]
    if only_sizes is not None:
        d = d[d["Size"].isin(only_sizes)]
    return d


# --------------------------------------------------------------------------------------
# 5. Statistics
# --------------------------------------------------------------------------------------
def _vals(d, col, settings):
    s = d[col]
    return s.fillna(0.0) if settings.missing_as_zero else s.dropna()


def describe(d: pd.DataFrame, settings: Settings) -> pd.DataFrame:
    """Table 3.2 style summary for Total and every metric."""
    rows = []
    for col in ["Total"] + settings.metrics:
        s = _vals(d, col, settings).astype(float)
        if s.empty:
            continue
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        rows.append({"Variable": col, "N": int(s.count()), "Mean": s.mean(), "Median": s.median(),
                     "Std dev (sample)": s.std(ddof=1) if len(s) > 1 else np.nan, "Min": s.min(), "Max": s.max(),
                     "Q1": q1, "Q3": q3, "IQR": q3 - q1, "Sum": s.sum()})
    return pd.DataFrame(rows)


def dominant_scope(d: pd.DataFrame, metrics: list) -> pd.Series:
    """Metric with the largest value per enterprise ('None' when all are zero/missing). Ties -> first listed metric."""
    v = d[metrics].fillna(0.0)
    idx = v.values.argmax(axis=1)
    lab = pd.Series(np.array(metrics, dtype=object)[idx], index=d.index)
    lab[(v.max(axis=1) <= 0)] = "None"
    return lab


def scope_table(d: pd.DataFrame, settings: Settings) -> pd.DataFrame:
    """Tables 3.4 / 3.5 / 3.6: total, share, mean, median, max and 'dominant in X of N' per scope."""
    n = len(d)
    grand = sum(_vals(d, m, settings).sum() for m in settings.metrics)
    dom = dominant_scope(d, settings.metrics).value_counts()
    rows = []
    for m in settings.metrics:
        s = _vals(d, m, settings).astype(float)
        rows.append({"Scope": m, "Total (tCO2e)": s.sum(), "Share (%)": s.sum() / grand * 100 if grand else np.nan,
                     "Mean per enterprise": s.mean() if len(s) else np.nan, "Median": s.median() if len(s) else np.nan,
                     "Maximum": s.max() if len(s) else np.nan, "Dominant in": int(dom.get(m, 0)), "Out of": n})
    t = pd.DataFrame(rows)
    t["Dominant in (text)"] = t["Dominant in"].astype(str) + " of " + t["Out of"].astype(str)
    return t


def group_profile(d: pd.DataFrame, by: str, settings: Settings, order=None) -> pd.DataFrame:
    """Table 3.3 style: count, average, median, total, share of sample, dominant scope (with its %)."""
    grand = d["Total"].sum()
    rows = []
    groups = list(d[by].dropna().unique())
    if order:
        groups = [g for g in order if g in groups] + [g for g in groups if g not in order]
    for g in groups:
        sub = d[d[by] == g]
        tot = sub["Total"]
        mt = {m: _vals(sub, m, settings).sum() for m in settings.metrics}
        sm = sum(mt.values())
        top = max(mt, key=mt.get) if sm else "None"
        rows.append({by: g, "No.": len(sub), "Average (tCO2e)": tot.mean(), "Median (tCO2e)": tot.median(),
                     "Total (tCO2e)": tot.sum(), "Share of sample (%)": tot.sum() / grand * 100 if grand else np.nan,
                     "Dominant scope": f"{top} ({mt[top] / sm * 100:.1f}%)" if sm else "None"})
    return pd.DataFrame(rows)


def ratio_table(prof: pd.DataFrame, by: str) -> pd.DataFrame:
    """Ratio of average footprint between each pair of groups (larger / smaller group, in listed order)."""
    out = []
    g = prof[[by, "Average (tCO2e)"]].values.tolist()
    for i in range(len(g)):
        for j in range(i + 1, len(g)):
            a, b = g[i], g[j]
            if a[1] and b[1]:
                out.append({"Comparison": f"{b[0]} vs {a[0]}", "Ratio of averages": b[1] / a[1]})
    return pd.DataFrame(out)


def ranking(d: pd.DataFrame, settings: Settings) -> pd.DataFrame:
    r = d.sort_values("Total", ascending=False).reset_index(drop=True)
    grand = r["Total"].sum()
    r.insert(0, "Rank", np.arange(1, len(r) + 1))
    r["Share of total (%)"] = r["Total"] / grand * 100 if grand else np.nan
    r["Cumulative share (%)"] = r["Share of total (%)"].cumsum()
    r["Dominant scope"] = dominant_scope(r, settings.metrics)
    return r


def outliers(d: pd.DataFrame, settings: Settings) -> tuple[pd.DataFrame, dict]:
    s = d["Total"]
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    lo, hi = q1 - settings.iqr_k * iqr, q3 + settings.iqr_k * iqr
    flag = d[(s > hi) | (s < lo)].sort_values("Total", ascending=False)
    return flag[["Sector", "Enterprise", "Size", "Total"]], {"Q1": q1, "Q3": q3, "IQR": iqr, "Lower fence": lo, "Upper fence": hi}


def pivot_by(d: pd.DataFrame, by: str, settings: Settings) -> pd.DataFrame:
    """Total/mean/median/count and scope totals for any category (location, process, sector...)."""
    g = d.groupby(by)
    t = pd.DataFrame({"No.": g.size(), "Total (tCO2e)": g["Total"].sum(), "Average (tCO2e)": g["Total"].mean(),
                      "Median (tCO2e)": g["Total"].median()})
    for m in settings.metrics:
        t[f"{m} (tCO2e)"] = g[m].sum()
    t["Share of total (%)"] = t["Total (tCO2e)"] / t["Total (tCO2e)"].sum() * 100
    return t.sort_values("Total (tCO2e)", ascending=False).reset_index()


def sector_comparison(data: pd.DataFrame, settings: Settings) -> pd.DataFrame:
    """Cross-sector summary (the 'Sumary' sheet) computed from the data instead of typed numbers."""
    d = select_population(data, settings)
    t = pivot_by(d, "Sector", settings).rename(columns={"No.": "Number of enterprises"})
    return t


def solar_table(data: pd.DataFrame, factor: float = 0.7) -> pd.DataFrame | None:
    if "Solar kW" not in data or data["Solar kW"].notna().sum() == 0:
        return None
    d = data[data["Solar kW"].notna()].copy()
    if "CO2 offset" not in d:
        d["CO2 offset"] = np.nan
    d["Offset used"] = d["CO2 offset"].where(d["CO2 offset"].notna(), d["Solar kW"] * factor)
    d["Net footprint"] = d["Total"] - d["Offset used"]
    d["Offset (% of footprint)"] = d["Offset used"] / d["Total"] * 100
    return d[["Sector", "Enterprise", "Size", "Solar kW", "Offset used", "Total", "Net footprint", "Offset (% of footprint)"]]


# --------------------------------------------------------------------------------------
# 6. Checks on the results themselves
# --------------------------------------------------------------------------------------
def reconcile(d: pd.DataFrame, settings: Settings, by: str = "Size") -> pd.DataFrame:
    rows = []
    tot = d["Total"].sum()
    sc = scope_table(d, settings)
    rows.append(("Scope totals add up to overall total", sc["Total (tCO2e)"].sum(), tot))
    rows.append(("Scope shares add up to 100%", sc["Share (%)"].sum(), 100.0))
    gp = group_profile(d, by, settings)
    rows.append((f"{by} group totals add up to overall total", gp["Total (tCO2e)"].sum(), tot))
    rows.append((f"{by} group counts add up to N", float(gp["No."].sum()), float(len(d))))
    rows.append((f"{by} shares add up to 100%", gp["Share of sample (%)"].sum(), 100.0))
    rows.append(("Dominant-scope counts add up to N", float(sc["Dominant in"].sum()), float(len(d))))
    r = pd.DataFrame(rows, columns=["Check", "Calculated", "Expected"])
    r["Result"] = np.where((r["Calculated"] - r["Expected"]).abs() < 0.01, "PASS", "FAIL")
    return r


# --------------------------------------------------------------------------------------
# 7. Findings written from the numbers (no fixed text)
# --------------------------------------------------------------------------------------
def fmt(x, nd=2):
    return f"{x:,.{nd}f}"


def findings(d: pd.DataFrame, settings: Settings, label: str = "the dataset") -> list[str]:
    out = []
    if d.empty:
        return ["No enterprises are available for analysis with the current selection."]
    desc = describe(d, settings).set_index("Variable")
    t = desc.loc["Total"]
    n = int(t["N"])
    out.append(f"{n} enterprises in {label} have a combined footprint of {fmt(t['Sum'])} tCO2e; the mean is {fmt(t['Mean'])} "
               f"and the median {fmt(t['Median'])} tCO2e (standard deviation {fmt(t['Std dev (sample)'])}).")
    if t["Mean"] > 1.25 * t["Median"]:
        out.append("The mean is well above the median, so a few high-emitting enterprises raise the average; most enterprises are below the mean.")
    r = ranking(d, settings)
    top, bot = r.iloc[0], r.iloc[-1]
    out.append(f"Highest emitter: {top['Enterprise']} ({fmt(top['Total'])} tCO2e, {fmt(top['Share of total (%)'], 1)}% of the total). "
               f"Lowest: {bot['Enterprise']} ({fmt(bot['Total'])} tCO2e).")
    k = min(settings.top_n, len(r))
    out.append(f"The top {k} enterprises account for {fmt(r['Total'].head(k).sum() / r['Total'].sum() * 100, 1)}% of the total.")
    sc = scope_table(d, settings)
    big = sc.sort_values("Share (%)", ascending=False).iloc[0]
    out.append(f"{big['Scope']} is the largest source overall ({fmt(big['Share (%)'], 1)}% of emissions, {fmt(big['Total (tCO2e)'])} tCO2e) and "
               f"is the dominant scope in {big['Dominant in']} of {big['Out of']} enterprises.")
    gp = group_profile(d, "Size", settings, SIZE_ORDER)
    if len(gp) > 1:
        parts = [f"{r_['Size']}: {r_['No.']} enterprises, average {fmt(r_['Average (tCO2e)'])} tCO2e, {fmt(r_['Share of sample (%)'], 1)}% of total" for _, r_ in gp.iterrows()]
        out.append("By size - " + "; ".join(parts) + ".")
        rt = ratio_table(gp, "Size")
        if len(rt):
            row = rt.iloc[0] if len(rt) == 1 else rt.loc[rt["Ratio of averages"].idxmax()]
            out.append(f"Average footprint ratio ({row['Comparison']}): {fmt(row['Ratio of averages'], 1)} to 1.")
    flag, fences = outliers(d, settings)
    if len(flag):
        out.append(f"Statistical outliers (IQR rule, k={settings.iqr_k}; upper fence {fmt(fences['Upper fence'])} tCO2e): "
                   + ", ".join(f"{a} ({fmt(b)})" for a, b in zip(flag['Enterprise'], flag['Total'])) + ".")
    else:
        out.append(f"No statistical outliers were found (IQR rule, k={settings.iqr_k}).")
    miss = [m for m in settings.metrics if d[m].isna().any()]
    if miss:
        out.append("Some enterprises have no reported value for " + ", ".join(miss) + "; these were "
                   + ("counted as 0, as in the report." if settings.missing_as_zero else "left out of those metrics' statistics."))
    return out


# --------------------------------------------------------------------------------------
# 8. One call that produces everything
# --------------------------------------------------------------------------------------
def analyse(data: pd.DataFrame, settings: Settings) -> dict:
    d = select_population(data, settings)
    res = {"population": d, "describe": describe(d, settings), "scopes": scope_table(d, settings),
           "by_size": group_profile(d, "Size", settings, SIZE_ORDER), "ranking": ranking(d, settings)}
    res["ratios"] = ratio_table(res["by_size"], "Size")
    res["outliers"], res["fences"] = outliers(d, settings)
    res["reconcile"] = reconcile(d, settings)
    res["by_size_scopes"] = {g: scope_table(d[d["Size"] == g], settings) for g in res["by_size"]["Size"]}
    for by in ("Sector", "Location", "Process"):
        if d[by].nunique() > 1 or by == "Sector":
            res[f"by_{by.lower()}"] = pivot_by(d, by, settings)
    res["findings"] = findings(d, settings)
    return res
