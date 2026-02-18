import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Gate Workload • EasyMag", layout="wide")

# -------------------------
# Normalization helpers
# -------------------------
def norm_giro(x) -> str:
    """Normalize Giro keys for robust matching (no Excel edits needed)."""
    if pd.isna(x):
        return ""
    s = str(x)
    s = s.replace("\u00A0", " ").strip()          # NBSP -> space, trim
    s = re.sub(r"\s+", " ", s)                    # collapse internal spaces
    s = s.replace("–", "-").replace("—", "-").replace("_", "-")
    s = s.upper()
    return s

def safe_to_datetime(x):
    try:
        return pd.to_datetime(x, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def find_marker_row(raw: pd.DataFrame) -> int | None:
    """Find first row containing '(*) Numero di Operazioni' (case-insensitive)."""
    pat = re.compile(r"\(\*\)\s*numero\s+di\s+operazioni", re.IGNORECASE)
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).fillna("")
        if row.str.contains(pat).any():
            return i
    return None

def find_header_row(raw: pd.DataFrame) -> int:
    """Find header row containing 'Giro' and 'Data' or 'Giro/Data'."""
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].astype(str).fillna("")
        joined = " | ".join(row.tolist()).lower()
        if "giro" in joined and ("data" in joined or "giro/data" in joined):
            return i
    # fallback: first non-empty row
    for i in range(len(raw)):
        if raw.iloc[i].notna().any():
            return i
    return 0

def parse_easymag_pivot(file_bytes: bytes) -> pd.DataFrame:
    """
    Parse EasyMag pivot-like export (dates down, giri across).
    Returns long df: Data, Giro, Valore, Giro_raw
    """
    raw = pd.read_excel(BytesIO(file_bytes), header=None, engine="openpyxl")

    marker = find_marker_row(raw)
    if marker is not None and marker > 0:
        raw = raw.iloc[:marker].copy()

    hdr = find_header_row(raw)
    header_vals = raw.iloc[hdr].tolist()
    df = raw.iloc[hdr+1:].copy()
    df.columns = header_vals

    df = df.dropna(axis=1, how="all")

    date_col = df.columns[0]
    df = df.rename(columns={date_col: "Data"})
    df["Data_str"] = df["Data"].astype(str).fillna("")

    # remove Tot: rows and empty rows
    df = df[~df["Data_str"].str.contains(r"^\s*Tot\s*:", case=False, regex=True)]
    df = df[df["Data_str"].str.strip() != ""]

    # parse dates
    df["Data"] = df["Data"].apply(safe_to_datetime)
    if df["Data"].isna().all():
        extracted = df["Data_str"].str.extract(r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", expand=False)
        df["Data"] = extracted.apply(safe_to_datetime)

    value_cols = [c for c in df.columns if c not in ["Data", "Data_str"]]

    long = df.melt(id_vars=["Data"], value_vars=value_cols, var_name="Giro", value_name="Valore")
    long["Giro_raw"] = long["Giro"]
    long["Giro"] = long["Giro"].apply(norm_giro)
    long["Valore"] = pd.to_numeric(long["Valore"], errors="coerce").fillna(0.0)

    long = long[(long["Giro"] != "") & (long["Valore"] != 0)]
    long = long.dropna(subset=["Data"])
    return long[["Data", "Giro", "Valore", "Giro_raw"]]

def infer_metric_label(filename: str, total_val: float) -> str:
    """
    Infer whether file is Righe or Colli from filename.
    If unknown, caller may still map it manually in UI.
    """
    name = (filename or "").lower()
    if any(k in name for k in ["riga", "righe", "preliev"]):
        return "Righe"
    if any(k in name for k in ["collo", "colli"]):
        return "Colli"
    # ambiguous
    return "Auto"

def build_gate_map(gate_file_bytes: bytes) -> pd.DataFrame:
    """
    Read GATE.xlsx expecting:
      - Giro in colonna B
      - Gate in colonna J
    """
    gate_raw = pd.read_excel(BytesIO(gate_file_bytes), engine="openpyxl")
    # Defensive: use positional columns if present
    if gate_raw.shape[1] >= 10:
        giro_col = gate_raw.columns[1]   # B
        gate_col = gate_raw.columns[9]   # J
    else:
        # fallback: try by name
        candidates_giro = [c for c in gate_raw.columns if str(c).strip().lower() in ["giro", "giri", "tour"]]
        candidates_gate = [c for c in gate_raw.columns if "gate" in str(c).strip().lower()]
        giro_col = candidates_giro[0] if candidates_giro else gate_raw.columns[0]
        gate_col = candidates_gate[0] if candidates_gate else gate_raw.columns[-1]

    gate_map = gate_raw[[giro_col, gate_col]].copy()
    gate_map.columns = ["Giro", "Gate"]

    gate_map["Giro_raw"] = gate_map["Giro"]
    gate_map["Giro"] = gate_map["Giro"].apply(norm_giro)
    gate_map["Gate"] = gate_map["Gate"].astype(str).replace({"nan": ""}).str.strip()

    gate_map = gate_map[gate_map["Giro"] != ""].copy()
    gate_map = gate_map.drop_duplicates(subset=["Giro"], keep="last")
    return gate_map[["Giro", "Gate", "Giro_raw"]]

def merge_with_gate(long_df: pd.DataFrame, gate_map: pd.DataFrame, metric: str) -> pd.DataFrame:
    fact = long_df.copy()
    fact["Metrica"] = metric
    fact = fact.merge(gate_map[["Giro", "Gate"]], on="Giro", how="left")
    fact["Gate"] = fact["Gate"].fillna("NON ASSEGNATO")
    return fact

# -------------------------
# UI
# -------------------------
st.title("Gate Workload • EasyMag")

st.sidebar.header("1) Caricamento file")
gate_file = st.sidebar.file_uploader("GATE.xlsx (obbligatorio)", type=["xlsx"], key="gate")

st.sidebar.caption("Gli export EasyMag sono opzionali: puoi caricare solo Righe, solo Colli, oppure entrambi.")
file1 = st.sidebar.file_uploader("Export EasyMag #1 (righe o colli) (opzionale)", type=["xlsx"], key="f1")
file2 = st.sidebar.file_uploader("Export EasyMag #2 (righe o colli) (opzionale)", type=["xlsx"], key="f2")

if gate_file is None:
    st.info("Carica prima **GATE.xlsx** per iniziare.")
    st.stop()

gate_map = build_gate_map(gate_file.getvalue())

# Parse available exports
exports = []
for f in [file1, file2]:
    if f is None:
        continue
    long_df = parse_easymag_pivot(f.getvalue())
    label = infer_metric_label(getattr(f, "name", ""), float(long_df["Valore"].sum()))
    exports.append({"file": f, "long": long_df, "label": label})

if len(exports) == 0:
    st.warning("Hai caricato solo GATE.xlsx. Carica almeno **un** export EasyMag (righe o colli) per vedere i grafici.")
    # Mostra comunque diagnostica mapping
    st.subheader("Diagnostica mapping GATE")
    st.dataframe(gate_map.head(50))
    st.stop()

# If ambiguous, let user assign metrics
st.sidebar.header("2) Assegnazione file (se serve)")
for i, ex in enumerate(exports, start=1):
    default = ex["label"]
    if default == "Auto":
        default = "Righe"
    ex["metric"] = st.sidebar.selectbox(
        f"File {i}: {getattr(ex['file'], 'name', 'export')}",
        options=["Righe", "Colli"],
        index=0 if default == "Righe" else 1,
        key=f"metric_{i}"
    )

# If user accidentally assigns both as same metric, allow but warn and keep the latest one
metrics_assigned = [ex["metric"] for ex in exports]
if len(metrics_assigned) == 2 and metrics_assigned[0] == metrics_assigned[1]:
    st.sidebar.warning("Hai assegnato entrambi i file alla stessa metrica. Verranno sommati come un unico dataset.")

# Build unified fact
facts = [merge_with_gate(ex["long"], gate_map, ex["metric"]) for ex in exports]
fact = pd.concat(facts, ignore_index=True)

# -------------------------
# Diagnostics for "NON ASSEGNATO"
# -------------------------
with st.expander("🔎 Diagnostica: perché vedo 'NON ASSEGNATO'? (click)"):
    st.write("L'app normalizza automaticamente i giri (spazi, maiuscole, numeri/testo). 'NON ASSEGNATO' resta solo se il giro non è presente nel mapping.")
    assigned = fact[fact["Gate"] != "NON ASSEGNATO"]["Valore"].sum()
    unassigned = fact[fact["Gate"] == "NON ASSEGNATO"]["Valore"].sum()
    total = fact["Valore"].sum()
    if total > 0:
        st.metric("Quota NON ASSEGNATO", f"{unassigned/total:.1%}")
    # list missing giri
    missing_giri = fact.loc[fact["Gate"] == "NON ASSEGNATO", ["Giro_raw", "Giro"]].drop_duplicates().sort_values("Giro")
    st.write("Giri presenti negli export ma senza Gate nel mapping:")
    st.dataframe(missing_giri, use_container_width=True)

# -------------------------
# Filters
# -------------------------
st.sidebar.header("3) Filtri")
min_d = fact["Data"].min().date()
max_d = fact["Data"].max().date()
start_d, end_d = st.sidebar.date_input("Periodo", value=(min_d, max_d), min_value=min_d, max_value=max_d)

if isinstance(start_d, (list, tuple)):
    # streamlit may return tuple
    start_d, end_d = start_d
start_dt = pd.to_datetime(start_d)
end_dt = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

fact_f = fact[(fact["Data"] >= start_dt) & (fact["Data"] <= end_dt)].copy()

all_gates = sorted(fact_f["Gate"].unique().tolist())
sel_gates = st.sidebar.multiselect("Solo Gate selezionati", options=all_gates, default=all_gates)

fact_f = fact_f[fact_f["Gate"].isin(sel_gates)]

all_giri = sorted(fact_f["Giro"].unique().tolist())
sel_giri = st.sidebar.multiselect("Solo Giri selezionati", options=all_giri, default=all_giri)

fact_f = fact_f[fact_f["Giro"].isin(sel_giri)]

# Settings
topn = st.sidebar.slider("Top N giri (vista dettaglio)", 5, 50, 15)
alert_gate_pct = st.sidebar.slider("Alert: Gate > X% del totale", 5, 90, 40)
alert_ratio = st.sidebar.slider("Alert: Colli/100 righe > soglia", 10, 300, 120)

# Metric selector based on available
available_metrics = sorted(fact_f["Metrica"].unique().tolist())
metric_for_pie = st.selectbox("Metrica per grafici principali", options=available_metrics, index=0)

# -------------------------
# Dashboard layout
# -------------------------
c1, c2, c3 = st.columns([1, 1, 1])

total_val = fact_f.loc[fact_f["Metrica"] == metric_for_pie, "Valore"].sum()
total_rows = fact_f.loc[fact_f["Metrica"] == "Righe", "Valore"].sum() if "Righe" in available_metrics else 0.0
total_colli = fact_f.loc[fact_f["Metrica"] == "Colli", "Valore"].sum() if "Colli" in available_metrics else 0.0

with c1:
    st.metric(f"Totale {metric_for_pie}", f"{total_val:,.0f}".replace(",", "."))
with c2:
    st.metric("Totale Righe", f"{total_rows:,.0f}".replace(",", ".")) if "Righe" in available_metrics else st.metric("Totale Righe", "—")
with c3:
    st.metric("Totale Colli", f"{total_colli:,.0f}".replace(",", ".")) if "Colli" in available_metrics else st.metric("Totale Colli", "—")

# PIE: gate share
pie_df = (fact_f[fact_f["Metrica"] == metric_for_pie]
          .groupby("Gate", as_index=False)["Valore"].sum()
          .sort_values("Valore", ascending=False))
if pie_df["Valore"].sum() > 0:
    fig_pie = px.pie(pie_df, names="Gate", values="Valore", title=f"Peso Gate sul totale ({metric_for_pie})")
    st.plotly_chart(fig_pie, use_container_width=True)
else:
    st.info("Nessun dato nel periodo/filtri selezionati.")

# Stacked bar Gate->Giro
st.subheader(f"Dettaglio Gate → Giri ({metric_for_pie})")
bar_df = fact_f[fact_f["Metrica"] == metric_for_pie].groupby(["Gate", "Giro"], as_index=False)["Valore"].sum()
if len(bar_df):
    fig_bar = px.bar(bar_df, x="Gate", y="Valore", color="Giro", title="Carico per Gate con dettaglio giri", barmode="stack")
    st.plotly_chart(fig_bar, use_container_width=True)
else:
    st.info("Nessun dato per l'istogramma.")

# Righe vs Colli dashboard
st.subheader("Confronto Righe vs Colli per Gate")
compare_df = fact_f.groupby(["Gate", "Metrica"], as_index=False)["Valore"].sum()
fig_cmp = px.bar(compare_df, x="Gate", y="Valore", color="Metrica", barmode="group", title="Righe vs Colli per Gate")
st.plotly_chart(fig_cmp, use_container_width=True)

# Ratio table if both present
if ("Righe" in available_metrics) and ("Colli" in available_metrics):
    piv = compare_df.pivot(index="Gate", columns="Metrica", values="Valore").fillna(0.0).reset_index()
    piv["Colli_per_100_righe"] = np.where(piv["Righe"] > 0, (piv["Colli"] / piv["Righe"]) * 100, np.nan)
    st.dataframe(piv.sort_values("Righe", ascending=False), use_container_width=True)

# Top giri per Gate
st.subheader("Top Giri per Gate")
gate_sel = st.selectbox("Scegli Gate", options=sorted(fact_f["Gate"].unique().tolist()))
top_df = (fact_f[fact_f["Gate"] == gate_sel]
          .groupby(["Giro", "Metrica"], as_index=False)["Valore"].sum())

# show both metrics if present, otherwise just one
if len(top_df):
    # rank by selected metric if present else first available
    rank_metric = metric_for_pie if metric_for_pie in available_metrics else available_metrics[0]
    rank = (top_df[top_df["Metrica"] == rank_metric]
            .sort_values("Valore", ascending=False)
            .head(topn)[["Giro", "Valore"]])
    st.write(f"Top {topn} giri per Gate **{gate_sel}** (ordinati per {rank_metric})")
    st.dataframe(rank, use_container_width=True)

    fig_top = px.bar(rank, x="Giro", y="Valore", title=f"Top giri ({rank_metric}) - {gate_sel}")
    st.plotly_chart(fig_top, use_container_width=True)

    # export CSV
    csv_bytes = rank.to_csv(index=False).encode("utf-8")
    st.download_button("Scarica CSV Top giri", data=csv_bytes, file_name=f"top_giri_{gate_sel}.csv", mime="text/csv")
else:
    st.info("Nessun dato per Top Giri.")

# Trend giornaliero per Gate
st.subheader("Trend giornaliero per Gate")
metric_trend = st.selectbox("Metrica per trend", options=available_metrics, key="metric_trend")
trend_df = (fact_f[fact_f["Metrica"] == metric_trend]
            .groupby(["Data", "Gate"], as_index=False)["Valore"].sum())
if len(trend_df):
    fig_trend = px.line(trend_df, x="Data", y="Valore", color="Gate", title=f"Trend giornaliero ({metric_trend})")
    st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.info("Nessun dato per trend.")

# Alerts
st.subheader("🚨 Alert")
alerts = []

# Gate share alert (on selected main metric)
tot_metric = fact_f.loc[fact_f["Metrica"] == metric_for_pie, "Valore"].sum()
if tot_metric > 0:
    share = (fact_f[fact_f["Metrica"] == metric_for_pie]
             .groupby("Gate", as_index=False)["Valore"].sum())
    share["pct"] = (share["Valore"] / tot_metric) * 100
    over = share[share["pct"] >= alert_gate_pct].sort_values("pct", ascending=False)
    for _, r in over.iterrows():
        alerts.append(f"Gate {r['Gate']} pesa {r['pct']:.1f}% del totale {metric_for_pie} (soglia {alert_gate_pct}%).")

# Colli/100 righe anomaly
if ("Righe" in available_metrics) and ("Colli" in available_metrics):
    cmp = compare_df.pivot(index="Gate", columns="Metrica", values="Valore").fillna(0.0)
    cmp["ratio"] = np.where(cmp["Righe"] > 0, (cmp["Colli"] / cmp["Righe"]) * 100, np.nan)
    anom = cmp[cmp["ratio"] >= alert_ratio].sort_values("ratio", ascending=False)
    for gate, row in anom.iterrows():
        alerts.append(f"Gate {gate}: Colli/100 righe = {row['ratio']:.1f} (soglia {alert_ratio}).")

if alerts:
    for a in alerts:
        st.warning(a)
else:
    st.success("Nessun alert nel periodo/filtri selezionati.")
