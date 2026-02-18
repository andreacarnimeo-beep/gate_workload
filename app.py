
from __future__ import annotations
import json
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt

from utils.env import load_dotenv, get_env
from utils.gate_map import read_gate_map
from utils.parse_easymag import read_easymag_export_first_table, classify_export, normalize_export

st.set_page_config(page_title="Gate Workload – EasyMag", layout="wide")

load_dotenv()

DATA_LAKE = Path(get_env("DATA_LAKE", "./data/lake"))
DATA_INBOX = Path(get_env("DATA_INBOX", "./data/inbox"))
GATE_MAP_PATH = Path(get_env("GATE_MAP", "./data/GATE.xlsx"))

FACT_PATH = DATA_LAKE / "fact_gate_workload.parquet"
META_PATH = DATA_LAKE / "refresh_meta.json"

st.title("Carico di lavoro per GATE – EasyMag")

def load_fact() -> pd.DataFrame:
    if FACT_PATH.exists():
        df = pd.read_parquet(FACT_PATH)
        # Data può essere object/date
        if "Data" in df.columns:
            df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        return df
    return pd.DataFrame(columns=["Giro","Gate","Data","Righe","Colli"])

def ingest_from_uploads(gate_file, f1, f2) -> pd.DataFrame:
    gate_map = read_gate_map(gate_file)
    def parse(upload):
        df = read_easymag_export_first_table(upload)
        kind = classify_export(df)
        norm = normalize_export(df)
        return kind, norm
    k1, n1 = parse(f1)
    k2, n2 = parse(f2)
    if k1 == "unknown" or k2 == "unknown" or k1 == k2:
        # fallback: usa volume
        s1 = n1["Valore"].sum()
        s2 = n2["Valore"].sum()
        k1 = "righe" if s1 >= s2 else "colli"
        k2 = "colli" if k1 == "righe" else "righe"
    righe = (n1 if k1=="righe" else n2).rename(columns={"Valore":"Righe"})
    colli = (n1 if k1=="colli" else n2).rename(columns={"Valore":"Colli"})
    if righe["Data"].notna().any() and colli["Data"].notna().any():
        fact = pd.merge(righe[["Giro","Data","Righe"]], colli[["Giro","Data","Colli"]], on=["Giro","Data"], how="outer")
    else:
        fact = pd.merge(righe[["Giro","Righe"]], colli[["Giro","Colli"]], on=["Giro"], how="outer")
        fact["Data"] = pd.NaT
    fact = fact.merge(gate_map, on="Giro", how="left")
    fact["Gate"] = fact["Gate"].fillna("(Senza Gate)")
    for c in ["Righe","Colli"]:
        fact[c] = pd.to_numeric(fact[c], errors="coerce").fillna(0)
    fact["Data"] = pd.to_datetime(fact["Data"], errors="coerce")
    return fact

# Sidebar: data source
with st.sidebar:
    st.header("Sorgente dati")
    mode = st.radio("Modalità", ["Automatica (DATA_LAKE)", "Manuale (upload)"], index=0)
    st.caption(f"DATA_INBOX: {DATA_INBOX}")
    st.caption(f"DATA_LAKE: {DATA_LAKE}")
    if META_PATH.exists():
        meta = json.loads(META_PATH.read_text(encoding="utf-8"))
        if meta:
            st.success(f"Ultimo refresh: {meta[0].get('updated_at')}")
            st.write({"righe_file": meta[0].get("righe_file"), "colli_file": meta[0].get("colli_file")})
    if mode == "Manuale (upload)":
        gate_up = st.file_uploader("Carica GATE.xlsx", type=["xlsx"])
        f1 = st.file_uploader("Carica export EasyMag #1", type=["xlsx"])
        f2 = st.file_uploader("Carica export EasyMag #2", type=["xlsx"])
        do_load = st.button("Carica dati", type="primary")
    else:
        do_load = st.button("Ricarica da DATA_LAKE")

@st.cache_data(show_spinner=False)
def cached_fact_from_lake(mtime: float) -> pd.DataFrame:
    return load_fact()

df = pd.DataFrame()
if mode == "Automatica (DATA_LAKE)":
    mtime = FACT_PATH.stat().st_mtime if FACT_PATH.exists() else 0.0
    if do_load:
        st.cache_data.clear()
    df = cached_fact_from_lake(mtime)
else:
    if do_load and gate_up and f1 and f2:
        df = ingest_from_uploads(gate_up, f1, f2)
    else:
        df = pd.DataFrame(columns=["Giro","Gate","Data","Righe","Colli"])

if df.empty:
    st.info("Nessun dato disponibile. Se sei in modalità Automatica, esegui ingestion/run_ingestion.py oppure verifica DATA_INBOX. In modalità Manuale, carica i file.")
    st.stop()

# Ensure columns
for c in ["Gate","Giro"]:
    df[c] = df[c].astype(str)

# Filters
st.subheader("Filtri")
c1, c2, c3, c4 = st.columns([1.2,1.2,1.2,1.2])

dates_available = df["Data"].dropna()
min_date = dates_available.min().date() if not dates_available.empty else None
max_date = dates_available.max().date() if not dates_available.empty else None

with c1:
    metric = st.selectbox("Metrica principale", ["Righe", "Colli"])
with c2:
    if min_date and max_date:
        start, end = st.date_input("Periodo", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    else:
        start, end = (None, None)
with c3:
    gates = sorted(df["Gate"].unique().tolist())
    gate_sel = st.multiselect("Solo Gate selezionati", gates, default=gates)
with c4:
    # giri dipendono da gate selezionati
    df_gate = df[df["Gate"].isin(gate_sel)]
    giri = sorted(df_gate["Giro"].unique().tolist())
    giro_sel = st.multiselect("Solo Giri selezionati", giri, default=giri)

f = df.copy()
if start and end:
    f = f[(f["Data"].dt.date >= start) & (f["Data"].dt.date <= end)]
f = f[f["Gate"].isin(gate_sel) & f["Giro"].isin(giro_sel)]

# Aggregations
gate_agg = f.groupby("Gate", as_index=False)[["Righe","Colli"]].sum()
gate_agg["Totale"] = gate_agg["Righe"] + gate_agg["Colli"]
total_metric = gate_agg[metric].sum()

# Alerts
with st.sidebar:
    st.header("Alert")
    pct_thr = st.slider("Gate > X% del totale (metrica)", 5, 90, 40)
    ratio_thr = st.slider("Colli / 100 righe anomalo (>)", 10, 400, 120)

alerts = []
if total_metric > 0:
    gate_agg["pct"] = gate_agg[metric] / total_metric * 100
    heavy = gate_agg[gate_agg["pct"] > pct_thr].sort_values("pct", ascending=False)
    if not heavy.empty:
        alerts.append(("Gate pesante", heavy[["Gate","pct"]]))

gate_agg["colli_per_100_righe"] = (gate_agg["Colli"] / gate_agg["Righe"].replace({0: pd.NA}) * 100).astype("Float64")
anom = gate_agg[gate_agg["colli_per_100_righe"] > ratio_thr].sort_values("colli_per_100_righe", ascending=False)
if not anom.empty:
    alerts.append(("Colli/100 righe anomalo", anom[["Gate","colli_per_100_righe"]]))

if alerts:
    st.warning("⚠️ Alert attivi (vedi sidebar per soglie).")
    for title, dfa in alerts:
        st.sidebar.subheader(title)
        st.sidebar.dataframe(dfa, use_container_width=True)
else:
    st.sidebar.success("Nessun alert nel periodo/filtri selezionati.")

# Layout main
k1, k2, k3 = st.columns(3)
k1.metric("Tot Righe", f["Righe"].sum())
k2.metric("Tot Colli", f["Colli"].sum())
ratio = (f["Colli"].sum() / (f["Righe"].sum() if f["Righe"].sum() else 1) * 100)
k3.metric("Colli / 100 Righe", f"{ratio:.1f}")

st.divider()

left, right = st.columns([1, 1.4])

# Pie chart
with left:
    st.markdown("### Peso dei Gate sul totale")
    pie_df = gate_agg.copy()
    pie_df = pie_df[pie_df[metric] > 0]
    if pie_df.empty:
        st.info("Nessun dato per la metrica selezionata.")
    else:
        pie = alt.Chart(pie_df).mark_arc().encode(
            theta=alt.Theta(field=metric, type="quantitative"),
            color=alt.Color(field="Gate", type="nominal"),
            tooltip=["Gate", metric, alt.Tooltip("pct:Q", format=".1f")]
        ).properties(height=360)
        st.altair_chart(pie, use_container_width=True)

# Stacked bar with giro detail
with right:
    st.markdown("### Dettaglio Gate → Giri (stacked)")
    det = f.groupby(["Gate","Giro"], as_index=False)[["Righe","Colli"]].sum()
    det["Val"] = det[metric]
    bar = alt.Chart(det).mark_bar().encode(
        x=alt.X("Gate:N", sort="-y"),
        y=alt.Y("Val:Q", title=metric),
        color=alt.Color("Giro:N"),
        tooltip=["Gate","Giro", metric, "Righe", "Colli"]
    ).properties(height=360)
    st.altair_chart(bar, use_container_width=True)

st.divider()

# Compare righe vs colli (grouped bar)
st.markdown("### Confronto Righe vs Colli per Gate")
melt = gate_agg.melt(id_vars=["Gate"], value_vars=["Righe","Colli"], var_name="Metrica", value_name="Valore")
cmp = alt.Chart(melt).mark_bar().encode(
    x=alt.X("Gate:N", sort="-y"),
    xOffset="Metrica:N",
    y=alt.Y("Valore:Q"),
    color="Metrica:N",
    tooltip=["Gate","Metrica","Valore"]
).properties(height=320)
st.altair_chart(cmp, use_container_width=True)

# Top giri per gate
st.markdown("### Top Giri per Gate")
gcol1, gcol2, gcol3 = st.columns([1.2,1,1])
with gcol1:
    gate_focus = st.selectbox("Seleziona Gate", sorted(f["Gate"].unique().tolist()))
with gcol2:
    topn = st.number_input("Top N", min_value=3, max_value=50, value=10, step=1)
with gcol3:
    top_mode = st.selectbox("Vista", ["Solo metrica selezionata", "Entrambe (righe+colli)"])

df_focus = f[f["Gate"] == gate_focus].groupby("Giro", as_index=False)[["Righe","Colli"]].sum()
df_focus["Val"] = df_focus[metric]
df_focus = df_focus.sort_values("Val", ascending=False).head(int(topn))

if top_mode == "Solo metrica selezionata":
    st.dataframe(df_focus[["Giro", metric]].rename(columns={metric:"Valore"}), use_container_width=True)
else:
    df_focus["Colli/100 righe"] = (df_focus["Colli"] / df_focus["Righe"].replace({0: pd.NA}) * 100).astype("Float64")
    st.dataframe(df_focus[["Giro","Righe","Colli","Colli/100 righe"]], use_container_width=True)

# Trend giornaliero
st.markdown("### Trend giornaliero per Gate")
if f["Data"].notna().any():
    metric_trend = st.selectbox("Metrica trend", ["Righe","Colli"], index=0)
    trend_gate = st.multiselect("Gate nel trend", sorted(f["Gate"].unique().tolist()), default=sorted(f["Gate"].unique().tolist())[:3])
    tr = f[f["Gate"].isin(trend_gate)].groupby(["Data","Gate"], as_index=False)[metric_trend].sum()
    line = alt.Chart(tr).mark_line().encode(
        x=alt.X("Data:T"),
        y=alt.Y(f"{metric_trend}:Q"),
        color="Gate:N",
        tooltip=["Data:T","Gate:N", metric_trend]
    ).properties(height=320)
    st.altair_chart(line, use_container_width=True)
else:
    st.info("Nel dato corrente non è presente una colonna Data valida: trend non disponibile.")
