"""
Batch mode: Excel/CSV files in  ->  results workbook (tables + Excel charts), chart images and a text summary out.
No web page needed. Works in Google Colab, Jupyter, or from the command line:

    python batch.py file1.xlsx file2.xlsx --out results
"""
from __future__ import annotations

import argparse
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

import engine as E
from export import build_excel

CAT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
SIZE_COL = {"Small": "#9cc2ef", "Medium": "#2a78d6", "Large": "#123f7a", "Unspecified": "#8a8a85"}


def _clean_ax(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color="#dddddd", lw=0.8)
    ax.set_axisbelow(True)


def load_tables(paths):
    tables = []
    for p in paths:
        tables += E.read_workbook(p, os.path.basename(p))
    return tables


def run(paths, out_dir="results", exclude_sizes=("Large",), missing_as_zero=True, repair_values=True,
        drop_duplicates=True, iqr_k=1.5, top_n=5, sheets=None, verbose=True):
    """
    paths   : list of .xlsx/.xls/.csv files
    sheets  : optional list of sheet names to use (default: automatic - a combined sheet with a Sector column if present,
              otherwise every sheet that looks like enterprise data)
    returns : dict with 'results', 'data', 'log', 'settings', 'files', 'tables_used'
    """
    os.makedirs(out_dir, exist_ok=True)
    tables = load_tables(paths)
    cands = [t for t in tables if E.is_enterprise_table(t)]
    if sheets:
        cands = [t for t in cands if t.sheet in sheets]
    else:
        combined = [t for t in cands if "sector" in t.mapping]
        cands = combined or cands
    if not cands:
        raise ValueError("No sheet with an enterprise-name column and at least two emission columns (e.g. Scope 1, Scope 2) was found.")
    metrics = []
    for t in cands:
        for key in ("scope1", "scope2", "scope3"):
            if key in t.mapping and E.ROLE_LABEL[key] not in metrics:
                metrics.append(E.ROLE_LABEL[key])
    settings = E.Settings(metrics=metrics, exclude_sizes=tuple(exclude_sizes), missing_as_zero=missing_as_zero,
                          repair_values=repair_values, drop_duplicates=drop_duplicates, iqr_k=float(iqr_k), top_n=int(top_n))
    data, log = E.build_dataset([(t, t.mapping, t.sheet) for t in cands], settings)
    res = E.analyse(data, settings)
    files = []
    xl = os.path.join(out_dir, "analysis_results.xlsx")
    with open(xl, "wb") as f:
        f.write(build_excel(res, data, log, settings))
    files.append(xl)
    files += make_charts(res, settings, out_dir)
    txt = os.path.join(out_dir, "key_findings.txt")
    with open(txt, "w", encoding="utf-8") as f:
        f.write("KEY FINDINGS (generated from the calculated numbers)\n\n")
        f.write("\n".join(f"- {x}" for x in res["findings"]))
        if log:
            f.write("\n\nDATA QUALITY NOTES\n\n" + "\n".join(f"- [{l['Severity']}] {l['Check']}: {l['Enterprise']} {l['Field']} - {l['Detail']}" for l in log))
    files.append(txt)
    out = dict(results=res, data=data, log=log, settings=settings, files=files,
               tables_used=[f"{t.source} > {t.sheet} ({len(t.df)} rows)" for t in cands])
    if verbose:
        print_report(out)
    return out


def make_charts(res, settings, out_dir):
    pop, metrics = res["population"], settings.metrics
    colour = {m: CAT[i % len(CAT)] for i, m in enumerate(metrics)}
    paths = []

    def save(fig, name):
        p = os.path.join(out_dir, name)
        fig.tight_layout(); fig.savefig(p, dpi=160); plt.close(fig); paths.append(p)

    gp = res["by_size"]
    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    bars = ax.bar(gp["Size"], gp["Average (tCO2e)"], color=[SIZE_COL.get(s, "#8a8a85") for s in gp["Size"]], width=0.55)
    for b, v, n in zip(bars, gp["Average (tCO2e)"], gp["No."]):
        ax.text(b.get_x() + b.get_width() / 2, v, f"{v:,.1f}\n(n={n})", ha="center", va="bottom", fontsize=9)
    ax.set_ylabel("tCO2e"); ax.set_title("Average carbon footprint by enterprise size", loc="left", fontsize=12); _clean_ax(ax)
    ax.set_ylim(0, gp["Average (tCO2e)"].max() * 1.2); save(fig, "chart1_average_by_size.png")

    sc = res["scopes"]
    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    bars = ax.bar(sc["Scope"], sc["Total (tCO2e)"], color=[colour[m] for m in sc["Scope"]], width=0.55)
    for b, v, s in zip(bars, sc["Total (tCO2e)"], sc["Share (%)"]):
        ax.text(b.get_x() + b.get_width() / 2, v, f"{v:,.1f}\n{s:.1f}%", ha="center", va="bottom", fontsize=9)
    ax.set_ylabel("tCO2e"); ax.set_title("Total emissions by scope", loc="left", fontsize=12); _clean_ax(ax)
    ax.set_ylim(0, sc["Total (tCO2e)"].max() * 1.2); save(fig, "chart2_scope_totals.png")

    rk = res["ranking"].head(15).iloc[::-1]
    fig, ax = plt.subplots(figsize=(8, 0.34 * len(rk) + 1.6))
    ax.barh(rk["Enterprise"], rk["Total"], color=[SIZE_COL.get(s, "#8a8a85") for s in rk["Size"]], height=0.65)
    ax.set_xlabel("tCO2e"); ax.set_title(f"Top {len(rk)} enterprises by total footprint (light = Small, mid = Medium, dark = Large)", loc="left", fontsize=11)
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    ax.grid(axis="x", color="#dddddd", lw=0.8); ax.set_axisbelow(True); save(fig, "chart3_top_enterprises.png")

    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    groups = list(res["by_size"]["Size"]); data_ = [pop.loc[pop["Size"] == g, "Total"].values for g in groups]
    bp = ax.boxplot(data_, patch_artist=True, widths=0.5, medianprops=dict(color="black"))
    for patch, g in zip(bp["boxes"], groups): patch.set_facecolor(SIZE_COL.get(g, "#8a8a85"))
    ax.set_xticks(range(1, len(groups) + 1), groups)
    ax.set_ylabel("tCO2e"); ax.set_title("Distribution of total footprint by size", loc="left", fontsize=12); _clean_ax(ax)
    save(fig, "chart4_distribution_by_size.png")

    sec = res["by_sector"]
    if len(sec) > 1:
        tot = sec[[f"{m} (tCO2e)" for m in metrics]].sum(axis=1)
        fig, ax = plt.subplots(figsize=(7.5, 4.4)); bottom = pd.Series(0.0, index=sec.index)
        for m in metrics:
            v = sec[f"{m} (tCO2e)"] / tot * 100
            ax.bar(sec["Sector"], v, bottom=bottom, label=m, color=colour[m], width=0.6, edgecolor="white", linewidth=1.2); bottom += v
        ax.set_ylabel("% of sector emissions"); ax.set_title("Scope composition by sector", loc="left", fontsize=12)
        ax.legend(frameon=False, ncol=len(metrics), loc="upper center", bbox_to_anchor=(0.5, -0.12)); _clean_ax(ax)
        plt.setp(ax.get_xticklabels(), rotation=20, ha="right"); save(fig, "chart5_scope_by_sector.png")
        fig, ax = plt.subplots(figsize=(7.5, 4.2))
        bars = ax.bar(sec["Sector"], sec["Total (tCO2e)"], color=CAT[0], width=0.6)
        for b, v, s in zip(bars, sec["Total (tCO2e)"], sec["Share of total (%)"]):
            ax.text(b.get_x() + b.get_width() / 2, v, f"{v:,.0f}\n{s:.1f}%", ha="center", va="bottom", fontsize=9)
        ax.set_ylabel("tCO2e"); ax.set_title("Total emissions by sector", loc="left", fontsize=12); _clean_ax(ax)
        ax.set_ylim(0, sec["Total (tCO2e)"].max() * 1.2); plt.setp(ax.get_xticklabels(), rotation=20, ha="right"); save(fig, "chart6_total_by_sector.png")
    return paths


def print_report(out):
    res, log = out["results"], out["log"]
    try:
        from IPython.display import display
    except Exception:
        display = lambda x: print(x.to_string(index=False) if hasattr(x, "to_string") else x)
    pd.options.display.float_format = "{:,.2f}".format
    print("Sheets analysed:", "; ".join(out["tables_used"]))
    print(f"Enterprises analysed: {len(res['population'])} (excluded sizes: {', '.join(out['settings'].exclude_sizes) or 'none'})\n")
    print("KEY FINDINGS"); [print(" -", f) for f in res["findings"]]
    for title, key in (("SUMMARY STATISTICS", "describe"), ("SCOPE-WISE", "scopes"), ("BY SIZE", "by_size"), ("BY SECTOR", "by_sector")):
        if key in res:
            print(f"\n{title}"); df = res[key]; display(df.drop(columns=["Dominant in (text)"], errors="ignore"))
    for g, t in res["by_size_scopes"].items():
        print(f"\nSCOPE-WISE: {g.upper()} ENTERPRISES"); display(t.drop(columns=["Dominant in (text)"]))
    print("\nOUTLIERS"); display(res["outliers"] if len(res["outliers"]) else pd.DataFrame({"": ["none"]}))
    print("\nCHECKS ON THE RESULTS"); display(res["reconcile"])
    n_err = sum(l["Severity"] == "Error" for l in log); n_w = sum(l["Severity"] == "Warning" for l in log)
    print(f"\nDATA QUALITY: {n_err} error(s), {n_w} warning(s)")
    for l in log:
        if l["Severity"] in ("Error", "Warning"):
            print(f"  [{l['Severity']}] {l['Check']} - {l['Enterprise']} {l['Field']}: {l['Detail']}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="+"); ap.add_argument("--out", default="results")
    ap.add_argument("--keep-large", action="store_true"); ap.add_argument("--exclude-missing", action="store_true")
    a = ap.parse_args()
    run(a.files, a.out, exclude_sizes=() if a.keep_large else ("Large",), missing_as_zero=not a.exclude_missing)
    print("\nResults written to:", os.path.abspath(a.out))
