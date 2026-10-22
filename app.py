"""
Enterprise Emissions Analysis Tool  -  run with:  streamlit run app.py
Upload one or many Excel/CSV files; get the same statistics, tables and charts as Chapter 3 of the report,
recalculated from the raw rows every time.
"""
import hashlib
import os

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import engine as E
from export import build_excel

st.set_page_config(page_title="Emissions Analysis Tool", page_icon="📊", layout="wide")

# ---- chart colours (validated categorical order; sizes use one blue ramp because Small < Medium < Large is ordered)
CAT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
SIZE_COL = {"Small": "#9cc2ef", "Medium": "#2a78d6", "Large": "#123f7a", "Unspecified": "#8a8a85"}
INK, GRID = "#52514e", "rgba(128,128,128,.25)"
HERE = os.path.dirname(os.path.abspath(__file__))
EXAMPLE = os.path.join(HERE, "sample_data", "example_workbook.xlsx")


def style(fig, title, h=380, ytitle=None, xtitle=None):
    fig.update_layout(title=dict(text=title, x=0, font=dict(size=15)), height=h, margin=dict(l=10, r=10, t=50, b=10),
                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", legend=dict(orientation="h", y=-0.18),
                      font=dict(color=INK), bargap=0.35)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, title=ytitle)
    fig.update_xaxes(showgrid=False, title=xtitle)
    return fig


def show(df, **kw):
    num = df.select_dtypes("number").columns
    st.dataframe(df.style.format({c: "{:,.2f}" for c in num if not c.startswith(("No", "N", "Rank", "Dominant in", "Out of"))}),
                 hide_index=True, use_container_width=True, **kw)


@st.cache_data(show_spinner=False)
def load(name: str, raw: bytes):
    import io
    return E.read_workbook(io.BytesIO(raw), name)


# ============================== sidebar ==============================
st.sidebar.title("📊 Emissions Analysis")
files = st.sidebar.file_uploader("Upload Excel / CSV files", type=["xlsx", "xlsm", "xls", "csv"], accept_multiple_files=True)
use_example = False
if os.path.exists(EXAMPLE):
    use_example = st.sidebar.checkbox("Use the example workbook", value=not files)

sources = [(f.name, f.getvalue()) for f in files]
if use_example and not files:
    sources = [("example_workbook.xlsx", open(EXAMPLE, "rb").read())]
if not sources:
    st.title("Enterprise Emissions Analysis Tool")
    st.info("Upload one or more Excel/CSV files in the sidebar. Each file needs a header row with enterprise name, size and emission columns "
            "(for example Scope 1 / Scope 2 / Scope 3). Column order and header row position do not matter.")
    st.stop()

tables = []
for name, raw in sources:
    try:
        tables += load(name, raw)
    except Exception as ex:
        st.sidebar.error(f"Could not read {name}: {ex}")
cands = [t for t in tables if E.is_enterprise_table(t)]
skipped = [t for t in tables if not E.is_enterprise_table(t)]
if not cands:
    st.error("No sheet with an enterprise-name column and at least two emission columns was found. "
             "Check the header row, or rename columns (e.g. 'SME Name', 'Scope 1', 'Scope 2').")
    st.stop()

labels = {f"{t.source} › {t.sheet}  ({len(t.df)} rows)": t for t in cands}
has_sector_col = [k for k, t in labels.items() if "sector" in t.mapping]
default_sel = has_sector_col or list(labels)
sel = st.sidebar.multiselect("Sheets to analyse", list(labels), default=default_sel,
                             help="If one sheet already combines all sectors, select only that one - otherwise the same enterprises would be counted twice.")
if has_sector_col and len(sel) > len(has_sector_col):
    st.sidebar.warning("You selected a combined sheet AND individual sheets. Exact duplicates are removed, but please check the Data quality tab.")

st.sidebar.subheader("Rules")
all_sizes = sorted({E.norm_size(v) for k in sel for v in labels[k].df.get(labels[k].mapping.get("size"), pd.Series(dtype=object)).tolist()},
                   key=lambda s: E.SIZE_ORDER.index(s) if s in E.SIZE_ORDER else 9) or E.SIZE_ORDER
excl = st.sidebar.multiselect("Leave out these sizes", all_sizes, default=[s for s in ["Large"] if s in all_sizes],
                              help="The report analyses Small + Medium enterprises (SMEs) and leaves out Large.")
miss0 = st.sidebar.radio("Missing values (N/A, blank, '-')", ["Count as 0 (as in the report)", "Leave out of that metric's statistics"], index=0)
repair = st.sidebar.checkbox("Repair obvious typos (e.g. '4.51.' → 4.51) and log them", value=True)
dedupe = st.sidebar.checkbox("Remove exact duplicate rows", value=True)
iqr_k = st.sidebar.number_input("Outlier rule: IQR multiplier", 0.5, 5.0, 1.5, 0.5)
topn = st.sidebar.number_input("Top-N for 'share of total'", 1, 50, 5)

# ============================== column mapping ==============================
st.title("Enterprise Emissions Analysis")
ROLES = [("name", "Enterprise name"), ("size", "Size"), ("sector", "Sector"), ("location", "Location"), ("process", "Operation / process"),
         ("scope1", "Scope 1"), ("scope2", "Scope 2"), ("scope3", "Scope 3"), ("total", "Total (optional, only used for checking)"),
         ("solar_kw", "Solar capacity kW (optional)"), ("offset", "CO2 offset (optional)")]
NONE = "— none —"
chosen = []
with st.expander("① Check how columns were recognised (change anything that is wrong)", expanded=False):
    for k in sel:
        t = labels[k]
        st.markdown(f"**{k}**")
        cols = [NONE] + list(t.df.columns)
        mp, c1 = {}, st.columns(4)
        for i, (role, lab) in enumerate(ROLES):
            default = t.mapping.get(role, NONE)
            pick = c1[i % 4].selectbox(lab, cols, index=cols.index(default), key=f"{k}-{role}")
            if pick != NONE:
                mp[role] = pick
        used = set(mp.values())
        numeric_like = [c for c in t.df.columns if c not in used and t.df[c].map(lambda v: E.parse_number(v)[1] in ("ok", "repaired")).mean() > 0.6]
        extra = st.multiselect("Extra emission columns to include (e.g. Scope 4, other gases)", numeric_like, key=f"{k}-extra")
        for c in extra:
            mp[f"extra:{c}"] = c
        sector_label = None
        if "sector" not in mp:
            sector_label = st.text_input("Sector name for this sheet", value=t.sheet, key=f"{k}-secname")
        chosen.append((t, mp, sector_label))
        st.divider()
if not chosen:
    chosen = [(labels[k], labels[k].mapping, labels[k].sheet) for k in sel]
if not sel:
    st.warning("Select at least one sheet in the sidebar.")
    st.stop()

metrics = []
for t, mp, _ in chosen:
    for key in ("scope1", "scope2", "scope3"):
        if key in mp and E.ROLE_LABEL[key] not in metrics:
            metrics.append(E.ROLE_LABEL[key])
    for key in mp:
        if key.startswith("extra:") and key[6:] not in metrics:
            metrics.append(key[6:])
if len(metrics) < 2:
    st.error("At least two emission columns are needed (e.g. Scope 1 and Scope 2).")
    st.stop()

settings = E.Settings(metrics=metrics, exclude_sizes=tuple(excl), missing_as_zero=miss0.startswith("Count"),
                      repair_values=repair, drop_duplicates=dedupe, iqr_k=float(iqr_k), top_n=int(topn))
data, log = E.build_dataset(chosen, settings)
if data.empty:
    st.error("No usable rows were found.")
    st.stop()

# optional filters
with st.sidebar:
    st.subheader("Filters")
    fs = st.multiselect("Sector", sorted(data["Sector"].unique()), default=sorted(data["Sector"].unique()))
    fl = st.multiselect("Location", sorted(data["Location"].unique()), default=sorted(data["Location"].unique()))
view = data[data["Sector"].isin(fs) & data["Location"].isin(fl)]
res = E.analyse(view, settings)
pop = res["population"]
if pop.empty:
    st.warning("No enterprises left after the size/sector/location filters.")
    st.stop()
errs = [l for l in log if l["Severity"] == "Error"]
warns = [l for l in log if l["Severity"] == "Warning"]
if errs or warns:
    st.warning(f"Data quality: {len(errs)} error(s), {len(warns)} warning(s) - see the **Data quality** tab before using these numbers in a report.")

colour = {m: CAT[i % len(CAT)] for i, m in enumerate(metrics)}
d = res["describe"].set_index("Variable")

tabs = st.tabs(["Overview", "Statistics", "By size & category", "Rankings & outliers", "Sectors", "Solar", "Data quality", "Data"])

# ============================== Overview ==============================
with tabs[0]:
    c = st.columns(5)
    c[0].metric("Enterprises", int(d.loc["Total", "N"]))
    c[1].metric("Total (tCO₂e)", f"{d.loc['Total', 'Sum']:,.2f}")
    c[2].metric("Mean (tCO₂e)", f"{d.loc['Total', 'Mean']:,.2f}")
    c[3].metric("Median (tCO₂e)", f"{d.loc['Total', 'Median']:,.2f}")
    c[4].metric("Highest (tCO₂e)", f"{d.loc['Total', 'Max']:,.2f}")
    st.subheader("Key findings")
    st.caption("Written automatically from the calculated numbers - they change when the data changes.")
    for f in res["findings"]:
        st.markdown(f"- {f}")
    a, b = st.columns(2)
    gp = res["by_size"]
    fig = go.Figure(go.Bar(x=gp["Size"], y=gp["Average (tCO2e)"], marker_color=[SIZE_COL.get(s, "#8a8a85") for s in gp["Size"]],
                           text=[f"{v:,.1f}<br>(n={n})" for v, n in zip(gp["Average (tCO2e)"], gp["No."])], textposition="outside",
                           hovertemplate="%{x}<br>Average %{y:,.2f} tCO₂e<extra></extra>"))
    a.plotly_chart(style(fig, "Average carbon footprint by enterprise size", ytitle="tCO₂e"), use_container_width=True)
    sc = res["scopes"]
    fig = go.Figure(go.Bar(x=sc["Scope"], y=sc["Total (tCO2e)"], marker_color=[colour[m] for m in sc["Scope"]],
                           text=[f"{v:,.1f}<br>{s:.1f}%" for v, s in zip(sc["Total (tCO2e)"], sc["Share (%)"])], textposition="outside",
                           hovertemplate="%{x}<br>%{y:,.2f} tCO₂e<extra></extra>"))
    b.plotly_chart(style(fig, "Total emissions by scope", ytitle="tCO₂e"), use_container_width=True)
    xl = build_excel(res, view, log, settings)
    st.download_button("⬇ Download full results (Excel, with tables and charts)", xl, "emissions_analysis_results.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# ============================== Statistics ==============================
with tabs[1]:
    st.subheader("Summary statistics")
    show(res["describe"])
    st.caption("Standard deviation is the sample standard deviation (n−1), as in Table 3.2. Quartiles use linear interpolation (same as Excel QUARTILE.INC).")
    st.subheader("Scope-wise statistics")
    show(res["scopes"].drop(columns=["Dominant in (text)"]))
    for g, t in res["by_size_scopes"].items():
        st.markdown(f"**{g} enterprises**")
        show(t.drop(columns=["Dominant in (text)"]))
    fig = go.Figure()
    for g in res["by_size"]["Size"]:
        sub = pop[pop["Size"] == g]
        fig.add_trace(go.Box(y=sub["Total"], name=g, marker_color=SIZE_COL.get(g, "#8a8a85"), boxpoints="all", jitter=0.4, pointpos=0,
                             text=sub["Enterprise"], hovertemplate="%{text}<br>%{y:,.2f} tCO₂e<extra></extra>"))
    st.plotly_chart(style(fig, "Distribution of total footprint by size (box = middle 50%, line = median)", ytitle="tCO₂e"), use_container_width=True)

# ============================== By size & category ==============================
with tabs[2]:
    by = st.selectbox("Group enterprises by", [c for c in ["Size", "Sector", "Location", "Process"] if pop[c].nunique() >= 1], index=0)
    pv = E.group_profile(pop, by, settings, E.SIZE_ORDER if by == "Size" else None) if by == "Size" else E.pivot_by(pop, by, settings)
    show(pv)
    if by == "Size":
        if len(res["ratios"]):
            st.markdown("**Ratio of average footprints**")
            show(res["ratios"])
    else:
        p2 = E.pivot_by(pop, by, settings)
        fig = go.Figure()
        for m in metrics:
            fig.add_trace(go.Bar(x=p2[by], y=p2[f"{m} (tCO2e)"], name=m, marker_color=colour[m]))
        fig.update_layout(barmode="stack")
        st.plotly_chart(style(fig, f"Emissions by {by.lower()} and scope (tCO₂e)", h=440, ytitle="tCO₂e"), use_container_width=True)
        if pop[by].nunique() > 1:
            p3 = p2.sort_values("Average (tCO2e)", ascending=False)
            fig = go.Figure(go.Bar(x=p3[by], y=p3["Average (tCO2e)"], marker_color=CAT[0], text=[f"n={n}" for n in p3["No."]],
                                   textposition="outside"))
            st.plotly_chart(style(fig, f"Average footprint per enterprise by {by.lower()}", ytitle="tCO₂e"), use_container_width=True)

# ============================== Rankings & outliers ==============================
with tabs[3]:
    rk = res["ranking"]
    n_show = st.slider("Enterprises to show in the chart", 3, min(50, len(rk)), min(15, len(rk))) if len(rk) > 3 else len(rk)
    top = rk.head(n_show).iloc[::-1]
    fig = go.Figure(go.Bar(y=top["Enterprise"], x=top["Total"], orientation="h", marker_color=[SIZE_COL.get(s, "#8a8a85") for s in top["Size"]],
                           customdata=top[["Size", "Share of total (%)"]].values,
                           hovertemplate="%{y}<br>%{x:,.2f} tCO₂e · %{customdata[0]} · %{customdata[1]:.1f}% of total<extra></extra>"))
    st.plotly_chart(style(fig, f"Top {n_show} enterprises by total footprint (shade = size: light Small, mid Medium, dark Large)",
                          h=max(320, 24 * n_show + 100), xtitle="tCO₂e"), use_container_width=True)
    show(rk[["Rank", "Enterprise", "Sector", "Size", "Location"] + metrics + ["Total", "Share of total (%)", "Cumulative share (%)", "Dominant scope"]])
    st.subheader("Statistical outliers")
    f = res["fences"]
    st.caption(f"IQR rule: Q1 = {f['Q1']:,.2f}, Q3 = {f['Q3']:,.2f}, IQR = {f['IQR']:,.2f}; values above {f['Upper fence']:,.2f} "
               f"(Q3 + {settings.iqr_k}×IQR) or below {f['Lower fence']:,.2f} are flagged.")
    if len(res["outliers"]):
        show(res["outliers"])
    else:
        st.success("No outliers under this rule.")

# ============================== Sectors ==============================
with tabs[4]:
    sec = res["by_sector"]
    if len(sec) < 2:
        st.info("Only one sector in the current selection. Upload several sectors (or a combined sheet) to compare them.")
    show(sec)
    comp = sec[["Sector"] + [f"{m} (tCO2e)" for m in metrics]].copy()
    fig = go.Figure()
    for m in metrics:
        fig.add_trace(go.Bar(x=comp["Sector"], y=comp[f"{m} (tCO2e)"] / comp[[f"{k} (tCO2e)" for k in metrics]].sum(axis=1) * 100,
                             name=m, marker_color=colour[m]))
    fig.update_layout(barmode="stack")
    st.plotly_chart(style(fig, "Scope composition by sector (% of each sector's emissions)", h=420, ytitle="%"), use_container_width=True)
    fig = go.Figure(go.Bar(x=sec["Sector"], y=sec["Total (tCO2e)"], marker_color=CAT[0],
                           text=[f"{s:.1f}%" for s in sec["Share of total (%)"]], textposition="outside"))
    st.plotly_chart(style(fig, "Total emissions by sector (label = share of all sectors)", ytitle="tCO₂e"), use_container_width=True)

# ============================== Solar ==============================
with tabs[5]:
    so = E.solar_table(view)
    if so is None:
        st.info("No solar columns were mapped. If a sheet has installed solar capacity (kW), map it in step ① to enable offset and net-footprint analysis.")
    else:
        fac = st.number_input("Offset factor (tCO₂e per kW) used where no offset value is given", 0.0, 5.0, 0.7, 0.05)
        so = E.solar_table(view, fac)
        c = st.columns(4)
        c[0].metric("Enterprises with solar data", len(so))
        c[1].metric("Total offset (tCO₂e)", f"{so['Offset used'].sum():,.2f}")
        c[2].metric("Net footprint (tCO₂e)", f"{so['Net footprint'].sum():,.2f}")
        c[3].metric("Offset share", f"{so['Offset used'].sum() / so['Total'].sum() * 100:.1f}%")
        show(so)
        top = so.sort_values("Total", ascending=False).head(15)
        fig = go.Figure([go.Bar(x=top["Enterprise"], y=top["Total"], name="Total footprint", marker_color=CAT[0]),
                         go.Bar(x=top["Enterprise"], y=top["Net footprint"], name="Net after solar offset", marker_color=CAT[2])])
        fig.update_layout(barmode="group")
        st.plotly_chart(style(fig, "Footprint before and after solar offset (top 15)", ytitle="tCO₂e"), use_container_width=True)

# ============================== Data quality ==============================
with tabs[6]:
    st.subheader("Checks on the data")
    if log:
        lg = pd.DataFrame(log)
        order = {"Error": 0, "Warning": 1, "Info": 2}
        lg = lg.sort_values("Severity", key=lambda s: s.map(order))
        st.dataframe(lg, hide_index=True, use_container_width=True)
    else:
        st.success("No data-quality issues found.")
    st.subheader("Checks on the results")
    rc = res["reconcile"]
    st.dataframe(rc.style.format({"Calculated": "{:,.2f}", "Expected": "{:,.2f}"}), hide_index=True, use_container_width=True)
    if skipped:
        st.caption("Sheets that were not treated as enterprise data: " + "; ".join(f"{t.source} › {t.sheet}" for t in skipped))

# ============================== Data ==============================
with tabs[7]:
    st.caption("Clean standardised data used for every calculation (totals are recomputed from the scope columns).")
    show(view)
