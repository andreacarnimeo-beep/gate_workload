
import re
from io import BytesIO
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Gate Workload", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
MARKER_RE = re.compile(r"\(\*\)\s*Numero\s+di\s+Operazioni", re.IGNORECASE)

def _read_excel_raw(file) -> pd.DataFrame:
    """Read first sheet, no header assumptions. Returns raw dataframe."""
    return pd.read_excel(file, sheet_name=0, header=None, engine="openpyxl")

def _truncate_at_marker(raw: pd.DataFrame) -> pd.DataFrame:
    """Truncate duplicated table by cutting at the row containing the marker."""
    cut_idx = None
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).str.cat(sep=" | ")
        if MARKER_RE.search(row):
            cut_idx = i
            break
    if cut_idx is None:
        # nothing to truncate
        return raw
    return raw.iloc[:cut_idx].copy()

def _guess_header_row(raw: pd.DataFrame) -> int:
    """Heuristic: pick the first row with >=3 non-null and some non-numeric strings."""
    best_i, best_score = 0, -1
    for i in range(min(50, len(raw))):
        r = raw.iloc[i]
        nonnull = r.notna().sum()
        if nonnull < 2:
            continue
        strs = r.dropna().astype(str)
        # score: count of alpha chars
        alpha = sum(any(ch.isalpha() for ch in s) for s in strs)
        score = nonnull + alpha
        if score > best_score:
            best_score, best_i = score, i
    return best_i

def _materialize_table(raw: pd.DataFrame) -> pd.DataFrame:
    """Turn truncated raw dataframe into a proper table with headers."""
    raw = raw.copy()
    hdr = _guess_header_row(raw)
    headers = raw.iloc[hdr].astype(str).str.strip().tolist()
    df = raw.iloc[hdr+1:].copy()
    df.columns = headers
    # drop empty columns
    df = df.dropna(axis=1, how="all")
    # drop empty rows
    df = df.dropna(axis=0, how="all")
    # remove "Unnamed" style headers
    df.columns = [c if c and c.lower() != "nan" else f"col_{i}" for i,c in enumerate(df.columns)]
    return df

def _normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    return df

def _find_delivery_round_col(df: pd.DataFrame) -> str:
    """Find Giro column."""
    candidates = [c for c in df.columns if re.search(r"\bgiro\b", c, re.IGNORECASE)]
    if candidates:
        return candidates[0]
    # fallback: common names
    for c in df.columns:
        if re.search(r"tour|round|delivery", c, re.IGNORECASE):
            return c
    # last resort: second column
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]

def _find_value_col(df: pd.DataFrame) -> str:
    """Find the numeric measure column (operations/rows/cartons)."""
    # prefer explicit numeric columns
    num_cols = []
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        if s.notna().sum() >= max(3, int(0.2*len(df))):
            num_cols.append((c, s))
    if not num_cols:
        return df.columns[-1]
    # if any column name hints at total/qty/righe/colli/operazioni
    for c,_ in num_cols:
        if re.search(r"righe|colli|operaz|qta|quant|tot", c, re.IGNORECASE):
            return c
    # else take the one with largest sum
    best = max(num_cols, key=lambda t: pd.to_numeric(df[t[0]], errors="coerce").sum())
    return best[0]

def _find_date_col(df: pd.DataFrame) -> str | None:
    """Try to find a date/datetime column."""
    # name hint
    for c in df.columns:
        if re.search(r"\bdata\b|date|giorno|day|timestamp|ora", c, re.IGNORECASE):
            s = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
            if s.notna().sum() >= max(3, int(0.2*len(df))):
                return c
    # content-based: any col that parses well as datetime
    for c in df.columns:
        s = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
        if s.notna().sum() >= max(3, int(0.2*len(df))):
            return c
    return None

def _label_file_kind(df: pd.DataFrame) -> str:
    """Guess whether the file is RIGHE or COLLI based on column names / overall magnitude."""
    cols = " ".join(df.columns).lower()
    if "righe" in cols or "prelev" in cols:
        return "righe"
    if "colli" in cols or "carton" in cols or "collo" in cols:
        return "colli"
    return "unknown"

def load_easymag(file) -> dict:
    """
    Returns dict with:
      df: normalized table
      giro_col, val_col, date_col, kind
    """
    raw = _read_excel_raw(file)
    raw = _truncate_at_marker(raw)
    df = _materialize_table(raw)
    df = _normalize_cols(df)

    giro_col = _find_delivery_round_col(df)
    val_col = _find_value_col(df)
    date_col = _find_date_col(df)
    kind = _label_file_kind(df)

    # cast
    df[giro_col] = df[giro_col].astype(str).str.strip()
    df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0)

    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)

    return {"df": df, "giro_col": giro_col, "val_col": val_col, "date_col": date_col, "kind": kind}

def load_gate_map(file) -> pd.DataFrame:
    """
    Expects: col B = Giro, col J = Gate (as per user description).
    Uses 0-indexed positions (B=1, J=9).
    """
    x = pd.read_excel(file, sheet_name=0, header=None, engine="openpyxl")
    # drop empty rows
    x = x.dropna(how="all")
    giro = x.iloc[:, 1].astype(str).str.strip()
    gate = x.iloc[:, 9].astype(str).str.strip()
    m = pd.DataFrame({"Giro": giro, "Gate": gate})
    m = m[(m["Giro"].str.lower() != "nan") & (m["Gate"].str.lower() != "nan")]
    m = m.dropna()
    m = m[m["Giro"] != ""]
    return m

def parse_date_from_filename(name: str):
    m = re.search(r"(20\d{2})[-_]?(\d{2})[-_]?(\d{2})", name)
    if not m:
        return None
    y,mo,d = map(int, m.groups())
    try:
        return pd.Timestamp(y,mo,d)
    except Exception:
        return None

# -----------------------------
# UI
# -----------------------------
st.title("Gate Workload • EasyMag")

with st.sidebar:
    st.header("1) Carica file")
    gate_file = st.file_uploader("GATE.xlsx (mappa Giro → Gate)", type=["xlsx"])
    f1 = st.file_uploader("Export EasyMag #1 (righe o colli)", type=["xlsx"])
    f2 = st.file_uploader("Export EasyMag #2 (righe o colli)", type=["xlsx"])

    st.divider()
    st.header("2) Impostazioni")
    top_n = st.slider("Top N giri (vista dettaglio)", 5, 50, 15)
    gate_share_thr = st.slider("Alert: Gate > X% del totale", 5, 90, 40)
    ratio_thr = st.slider("Alert: Colli/100 righe > soglia", 10, 500, 120)

if not (gate_file and f1 and f2):
    st.info("Carica **GATE.xlsx** e i **2 export EasyMag** per vedere la dashboard.")
    st.stop()

# Load data
gate_map = load_gate_map(gate_file)

d1 = load_easymag(f1)
d2 = load_easymag(f2)

# Assign righe/colli
pairs = [d1, d2]
kinds = [p["kind"] for p in pairs]
if "righe" in kinds and "colli" in kinds:
    righe = pairs[kinds.index("righe")]
    colli = pairs[kinds.index("colli")]
else:
    # fallback by magnitude (righe tends to be larger)
    sums = [p["df"][p["val_col"]].sum() for p in pairs]
    righe = pairs[int(np.argmax(sums))]
    colli = pairs[int(np.argmin(sums))]

# Ensure date columns
for p, up in [(righe, f1), (colli, f2)]:
    if p["date_col"] is None:
        # create ExportDate from filename or ask user
        inferred = parse_date_from_filename(getattr(up, "name", "") or "")
        if inferred is None:
            inferred = pd.Timestamp.today().normalize()
        p["df"]["ExportDate"] = inferred
        p["date_col"] = "ExportDate"

# Standardize columns and merge with gate
def build_fact(p: dict, metric_name: str) -> pd.DataFrame:
    df = p["df"].copy()
    giro_col, val_col, date_col = p["giro_col"], p["val_col"], p["date_col"]

    fact = df[[giro_col, val_col, date_col]].copy()
    fact.columns = ["Giro", metric_name, "DateTime"]
    # daily grain
    fact["Day"] = pd.to_datetime(fact["DateTime"], errors="coerce").dt.floor("D")
    fact = fact.dropna(subset=["Giro"])
    fact = fact.merge(gate_map, on="Giro", how="left")
    fact["Gate"] = fact["Gate"].fillna("UNMAPPED")
    fact[metric_name] = pd.to_numeric(fact[metric_name], errors="coerce").fillna(0)
    return fact

fact_r = build_fact(righe, "Righe")
fact_c = build_fact(colli, "Colli")

# Combine
fact = fact_r.merge(
    fact_c[["Giro","Gate","Day","Colli"]],
    on=["Giro","Gate","Day"],
    how="outer"
).fillna({"Righe":0, "Colli":0})

# Filters: date, gate, giro
min_day = fact["Day"].min()
max_day = fact["Day"].max()
with st.sidebar:
    st.header("3) Filtri")
    date_range = st.date_input("Periodo", value=(min_day.date(), max_day.date()))
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_d, end_d = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1])
    else:
        start_d, end_d = pd.Timestamp(min_day), pd.Timestamp(max_day)

    all_gates = sorted(fact["Gate"].unique().tolist())
    selected_gates = st.multiselect("Solo Gate selezionati", options=all_gates, default=all_gates)

    # update giri list based on selected gates
    tmp = fact[fact["Gate"].isin(selected_gates)]
    all_giri = sorted(tmp["Giro"].unique().tolist())
    selected_giri = st.multiselect("Solo Giri selezionati", options=all_giri, default=all_giri)

    metric = st.radio("Metrica principale", ["Righe", "Colli"], horizontal=True)

# apply filters
flt = fact[
    (fact["Day"] >= start_d) & (fact["Day"] <= end_d) &
    (fact["Gate"].isin(selected_gates)) &
    (fact["Giro"].isin(selected_giri))
].copy()

if flt.empty:
    st.warning("Nessun dato nel periodo/filtri selezionati.")
    st.stop()

# Aggregations
gate_tot = flt.groupby("Gate", as_index=False)[["Righe","Colli"]].sum()
gate_tot["Colli_per_100_Righe"] = np.where(gate_tot["Righe"]>0, gate_tot["Colli"]*100.0/gate_tot["Righe"], np.nan)

total_metric = gate_tot[metric].sum()
gate_tot["Share_%"] = np.where(total_metric>0, gate_tot[metric]*100.0/total_metric, 0)

# Alerts
alerts = []
big_gates = gate_tot[gate_tot["Share_%"] > gate_share_thr].sort_values("Share_%", ascending=False)
if not big_gates.empty:
    alerts.append(f"🚨 Gate sopra soglia {gate_share_thr}%: " + ", ".join([f"{r.Gate} ({r.Share_:.1f}%)".replace("Share_","Share_%") for _,r in big_gates.iterrows()]))

bad_ratio = gate_tot[gate_tot["Colli_per_100_Righe"] > ratio_thr].sort_values("Colli_per_100_Righe", ascending=False)
if not bad_ratio.empty:
    alerts.append(f"⚠️ Colli/100 righe sopra {ratio_thr}: " + ", ".join([f"{r.Gate} ({r.Colli_per_100_Righe:.1f})" for _,r in bad_ratio.iterrows()]))

# KPI row
c1, c2, c3, c4 = st.columns(4)
c1.metric("Tot Righe", f"{int(flt['Righe'].sum()):,}".replace(",", "."))
c2.metric("Tot Colli", f"{int(flt['Colli'].sum()):,}".replace(",", "."))
ratio = (flt["Colli"].sum()*100.0/flt["Righe"].sum()) if flt["Righe"].sum()>0 else np.nan
c3.metric("Colli / 100 righe", "-" if np.isnan(ratio) else f"{ratio:.1f}")
c4.metric("Gate attivi", f"{flt['Gate'].nunique()}")

if alerts:
    for a in alerts:
        st.warning(a)

st.divider()

# Layout
left, right = st.columns([1, 1])

with left:
    st.subheader("Peso dei Gate sul totale")
    pie_df = gate_tot.sort_values(metric, ascending=False)
    fig = px.pie(pie_df, names="Gate", values=metric, hole=0.35)
    st.plotly_chart(fig, use_container_width=True)

with right:
    st.subheader("Confronto Righe vs Colli per Gate")
    bar_df = gate_tot.melt(id_vars=["Gate"], value_vars=["Righe","Colli"], var_name="Metrica", value_name="Valore")
    fig2 = px.bar(bar_df, x="Gate", y="Valore", color="Metrica", barmode="group")
    st.plotly_chart(fig2, use_container_width=True)

st.divider()

st.subheader("Dettaglio: Gate → Giri (barra stacked)")
# stacked by giri inside gate for selected metric
g = flt.groupby(["Gate","Giro"], as_index=False)[["Righe","Colli"]].sum()
g["Valore"] = g[metric]
# keep top giri per gate to avoid clutter
g_rank = g.sort_values(["Gate","Valore"], ascending=[True, False]).groupby("Gate").head(50)
fig3 = px.bar(g_rank, x="Gate", y="Valore", color="Giro", barmode="stack")
st.plotly_chart(fig3, use_container_width=True)

st.divider()

# Top giri per gate
st.subheader("Top Giri per Gate")
sel_gate = st.selectbox("Scegli Gate", options=sorted(flt["Gate"].unique().tolist()))
sub = flt[flt["Gate"] == sel_gate].groupby("Giro", as_index=False)[["Righe","Colli"]].sum()
sub["Colli_per_100_Righe"] = np.where(sub["Righe"]>0, sub["Colli"]*100.0/sub["Righe"], np.nan)
sub = sub.sort_values(metric, ascending=False).head(top_n)

colA, colB = st.columns([1,1])
with colA:
    fig4 = px.bar(sub, x="Giro", y=metric)
    st.plotly_chart(fig4, use_container_width=True)
with colB:
    st.dataframe(sub, use_container_width=True)

csv = sub.to_csv(index=False).encode("utf-8")
st.download_button("Scarica Top Giri (CSV)", data=csv, file_name=f"top_giri_{sel_gate}.csv", mime="text/csv")

st.divider()

# Trend giornaliero
st.subheader("Trend giornaliero per Gate")
trend_metric = st.radio("Metrica trend", ["Righe","Colli"], horizontal=True, key="trend_metric")
trend = flt.groupby(["Day","Gate"], as_index=False)[trend_metric].sum()

# If too many gates, let user pick a subset for readability
trend_gates = st.multiselect("Gate nel trend", options=sorted(trend["Gate"].unique().tolist()),
                             default=sorted(trend["Gate"].unique().tolist())[:min(6, trend["Gate"].nunique())])
trend = trend[trend["Gate"].isin(trend_gates)]

fig5 = px.line(trend, x="Day", y=trend_metric, color="Gate", markers=True)
st.plotly_chart(fig5, use_container_width=True)

st.caption("Nota: se l'export non contiene una colonna data, il trend usa la data inferita dal filename (es. 2026-02-18) o la data odierna.")
