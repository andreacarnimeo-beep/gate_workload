import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from io import BytesIO

st.set_page_config(page_title="Gate Workload • EasyMag", layout="wide")

# -------------------------
# Helpers
# -------------------------
def norm_giro(x) -> str:
    if pd.isna(x):
        return ""
    try:
        if isinstance(x, (int, np.integer)):
            return str(int(x))
        if isinstance(x, (float, np.floating)) and float(x).is_integer():
            return str(int(x))
    except Exception:
        pass
    s = str(x).replace("\u00A0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("–", "-").replace("—", "-").replace("_", "-")
    s = s.upper()
    m = re.fullmatch(r"0*(\d+)\.0+", s)
    if m:
        s = m.group(1)
    return s

def safe_to_datetime(x):
    try:
        return pd.to_datetime(x, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT

def find_marker_row(raw: pd.DataFrame) -> int | None:
    pat = re.compile(r"\(\*\)\s*numero\s+di\s+operazioni", re.IGNORECASE)
    for i in range(len(raw)):
        row = raw.iloc[i].astype(str).fillna("")
        if row.str.contains(pat).any():
            return i
    return None

def find_header_row(raw: pd.DataFrame) -> int:
    for i in range(min(80, len(raw))):
        row = raw.iloc[i].astype(str).fillna("")
        joined = " | ".join(row.tolist()).lower()
        if "giro" in joined and ("data" in joined or "giro/data" in joined):
            return i
    for i in range(len(raw)):
        if raw.iloc[i].notna().any():
            return i
    return 0

def is_total_giro(colname) -> bool:
    if pd.isna(colname):
        return True
    s = str(colname).strip().lower()
    return (s in {"tot", "tot:", "totale", "totali"} or s.startswith("tot") or "totale" in s)


def sniff_report_type(file_bytes: bytes) -> str:
    """
    Auto-detect report type from the header area of the EasyMag export.
    If it contains 'preliev' near the '(*) Numero di Operazioni ...' line -> Righe.
    If it contains 'colli'/'collo' -> Colli.
    """
    try:
        raw = pd.read_excel(BytesIO(file_bytes), header=None, nrows=80, engine="openpyxl")
        blob = "\n".join(raw.astype(str).fillna("").values.flatten().tolist()).lower()
    except Exception:
        return "Auto"
    if "preliev" in blob:
        return "Righe"
    if "colli" in blob or "collo" in blob:
        return "Colli"
    return "Auto"

def parse_easymag_pivot(file_bytes: bytes) -> pd.DataFrame:
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

    # drop Tot rows in date col
    df = df[~df["Data_str"].str.contains(r"^\s*Tot\s*:", case=False, regex=True)]
    df = df[df["Data_str"].str.strip() != ""]

    df["Data"] = df["Data"].apply(safe_to_datetime)
    if df["Data"].isna().all():
        extracted = df["Data_str"].str.extract(r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", expand=False)
        df["Data"] = extracted.apply(safe_to_datetime)

    raw_value_cols = [c for c in df.columns if c not in ["Data", "Data_str"]]
    value_cols = [c for c in raw_value_cols if not is_total_giro(c)]  # IMPORTANT: remove Tot columns

    long = df.melt(id_vars=["Data"], value_vars=value_cols, var_name="Giro", value_name="Valore")
    long["Giro_raw"] = long["Giro"]
    long["Giro"] = long["Giro"].apply(norm_giro)
    long["Valore"] = pd.to_numeric(long["Valore"], errors="coerce").fillna(0.0)

    long = long[(long["Giro"] != "") & (long["Valore"] != 0)]
    long = long.dropna(subset=["Data"])
    return long[["Data", "Giro", "Valore", "Giro_raw"]]

def infer_metric_label(filename: str) -> str:
    name = (filename or "").lower()
    if any(k in name for k in ["riga", "righe", "preliev"]):
        return "Righe"
    if any(k in name for k in ["collo", "colli"]):
        return "Colli"
    return "Auto"

def build_gate_map(gate_file_bytes: bytes) -> pd.DataFrame:
    gate_raw = pd.read_excel(BytesIO(gate_file_bytes), engine="openpyxl")
    if gate_raw.shape[1] >= 10:
        giro_col = gate_raw.columns[1]   # B
        gate_col = gate_raw.columns[9]   # J
    else:
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

def merge_with_gate(long_df: pd.DataFrame, gate_map: pd.DataFrame, metric: str, source: str) -> pd.DataFrame:
    fact = long_df.copy()
    fact["Metrica"] = metric
    fact["Fonte"] = source
    fact = fact.merge(gate_map[["Giro", "Gate"]], on="Giro", how="left")
    return fact

def fmt_int(x: float) -> str:
    try:
        return f"{x:,.0f}".replace(",", ".")
    except Exception:
        return "—"

# -------------------------
# UI
# -------------------------


def parse_period_widget(min_d, max_d):
    """
    Robust date selector:
    - Supports single-day analysis (start=end).
    - Handles Streamlit returning:
        * a single date
        * a (start, end) tuple
        * (start, None) while user is still selecting
        * [date] list (some builds)
    """
    # If dataset contains a single day, show a single-date picker
    if min_d == max_d:
        d = st.sidebar.date_input("Periodo", value=min_d, min_value=min_d, max_value=max_d)
        return d, d

    period = st.sidebar.date_input("Periodo", value=(min_d, max_d), min_value=min_d, max_value=max_d)

    # Case: tuple/list
    if isinstance(period, (list, tuple)):
        if len(period) == 2:
            start, end = period[0], period[1]
            if end is None:
                end = start
            return start, end
        if len(period) == 1:
            return period[0], period[0]

    # Case: single date
    return period, period

st.title("Gate Workload • EasyMag")

st.sidebar.header("1) Caricamento file")
gate_file = st.sidebar.file_uploader("GATE.xlsx (obbligatorio)", type=["xlsx"], key="gate")
st.sidebar.caption("Gli export EasyMag sono opzionali: puoi caricare solo Righe, solo Colli, oppure entrambi.")
file1 = st.sidebar.file_uploader("Export EasyMag #1 (opzionale)", type=["xlsx"], key="f1")
file2 = st.sidebar.file_uploader("Export EasyMag #2 (opzionale)", type=["xlsx"], key="f2")

if gate_file is None:
    st.info("Carica prima **GATE.xlsx** per iniziare.")
    st.stop()

gate_map = build_gate_map(gate_file.getvalue())
known_gates = sorted([g for g in gate_map["Gate"].unique().tolist() if str(g).strip() != ""])

exports = []
for idx, f in enumerate([file1, file2], start=1):
    if f is None:
        continue
    fb = f.getvalue()
    long_df = parse_easymag_pivot(fb)
    detected = sniff_report_type(fb)
    label = infer_metric_label(getattr(f, "name", ""))
    exports.append({"file": f, "long": long_df, "label": label, "detected": detected, "idx": idx})

if len(exports) == 0:
    st.warning("Hai caricato solo GATE.xlsx. Carica almeno **un** export EasyMag per vedere i grafici.")
    st.dataframe(gate_map.head(50), use_container_width=True)
    st.stop()

st.sidebar.header("2) Assegnazione file")
for ex in exports:
    default = ex.get("detected", "Auto")
    if default == "Auto":
        default = ex.get("label", "Righe")
    if default == "Auto":
        default = "Righe"
    ex["metric"] = st.sidebar.selectbox(
        f"File {ex['idx']}: {getattr(ex['file'], 'name', 'export')}",
        options=["Righe", "Colli"],
        index=0 if default == "Righe" else 1,
        key=f"metric_{ex['idx']}"
    )

# Show summary of assignments
assign_summary = pd.DataFrame({
    "Fonte": [getattr(ex["file"], "name", f"File {ex['idx']}") for ex in exports],
    "Auto_detect": [ex.get("detected", "Auto") for ex in exports],
    "Metrica": [ex["metric"] for ex in exports],
    "Totale_valore": [float(ex["long"]["Valore"].sum()) for ex in exports],
})
st.sidebar.caption("Riepilogo file caricati:")
st.sidebar.dataframe(assign_summary, use_container_width=True, height=140)

facts = [merge_with_gate(ex["long"], gate_map, ex["metric"], getattr(ex["file"], "name", f"File {ex['idx']}")) for ex in exports]
fact = pd.concat(facts, ignore_index=True)

# -------------------------
# Enforce: no missing Gate
# -------------------------
if "gate_overrides" not in st.session_state:
    st.session_state.gate_overrides = {}

missing_giri_df = fact[fact["Gate"].isna()][["Giro", "Giro_raw"]].drop_duplicates().sort_values("Giro")

if len(missing_giri_df) > 0:
    st.error("⚠️ Ci sono giri senza Gate nel mapping. Poiché **non possono esistere Gate non assegnati**, devi assegnare ogni giro mancante a un Gate.")
    st.write("Non modifichiamo gli Excel: queste assegnazioni sono applicate **solo dentro l'app** (sessione corrente).")
    st.dataframe(missing_giri_df, use_container_width=True)

    with st.form("assign_missing_gates"):
        st.subheader("Assegna i giri mancanti")
        gate_options = known_gates + ["(Inserisci manualmente)"] if known_gates else ["(Inserisci manualmente)"]
        for _, row in missing_giri_df.iterrows():
            giro = row["Giro"]
            giro_raw = row["Giro_raw"]
            st.markdown(f"**Giro:** `{giro}` (raw: `{giro_raw}`)")
            sel = st.selectbox(f"Gate per Giro {giro}", options=gate_options, key=f"sel_gate_{giro}")
            manual = ""
            if sel == "(Inserisci manualmente)":
                manual = st.text_input(f"Inserisci Gate (testo) per Giro {giro}", key=f"manual_gate_{giro}")
            chosen = (manual or sel).strip()
            st.session_state.gate_overrides[giro] = chosen
        submitted = st.form_submit_button("✅ Salva assegnazioni e continua")

    if not submitted:
        st.stop()

    unresolved = [g for g in missing_giri_df["Giro"].tolist()
                  if st.session_state.gate_overrides.get(g, "").strip() == "" or st.session_state.gate_overrides.get(g) == "(Inserisci manualmente)"]
    if unresolved:
        st.warning("Devi assegnare un Gate valido per tutti i giri mancanti.")
        st.stop()

# Apply overrides
if st.session_state.gate_overrides:
    ov = st.session_state.gate_overrides
    fact["Gate"] = fact.apply(lambda r: ov.get(r["Giro"], r["Gate"]), axis=1)

if fact["Gate"].isna().any():
    st.error("Sono rimasti giri non assegnati. Controlla le assegnazioni.")
    st.stop()

with st.expander("🔎 Diagnostica mapping (click)"):
    if st.session_state.gate_overrides:
        st.write("Override applicati (solo app):")
        st.dataframe(pd.DataFrame({"Giro": list(st.session_state.gate_overrides.keys()),
                                   "Gate_override": list(st.session_state.gate_overrides.values())}), use_container_width=True)
    st.write("Anteprima GATE.xlsx:")
    st.dataframe(gate_map.head(50), use_container_width=True)

# -------------------------
# Filters
# -------------------------
st.sidebar.header("3) Filtri")
min_d = fact["Data"].min().date()
max_d = fact["Data"].max().date()
start_d, end_d = parse_period_widget(min_d, max_d)

start_dt = pd.to_datetime(start_d)
end_dt = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
fact_f = fact[(fact["Data"] >= start_dt) & (fact["Data"] <= end_dt)].copy()

all_gates = sorted(fact_f["Gate"].unique().tolist())
sel_gates = st.sidebar.multiselect("Solo Gate selezionati", options=all_gates, default=all_gates)
fact_f = fact_f[fact_f["Gate"].isin(sel_gates)]

all_giri = sorted(fact_f["Giro"].unique().tolist())
sel_giri = st.sidebar.multiselect("Solo Giri selezionati", options=all_giri, default=all_giri)
fact_f = fact_f[fact_f["Giro"].isin(sel_giri)]

topn = st.sidebar.slider("Top N giri (vista dettaglio)", 5, 50, 15)
alert_gate_pct = st.sidebar.slider("Alert: Gate > X% del totale", 5, 90, 40)
alert_ratio = st.sidebar.slider("Alert: Colli/100 righe > soglia", 10, 300, 120)

available_metrics = sorted(fact_f["Metrica"].unique().tolist())
metric_for_main = st.selectbox("Metrica per grafici principali", options=available_metrics, index=0)

# -------------------------
# KPIs (NO ternary expressions)
# -------------------------
c1, c2, c3 = st.columns([1, 1, 1])
total_main = fact_f.loc[fact_f["Metrica"] == metric_for_main, "Valore"].sum()

with c1:
    st.metric(f"Totale {metric_for_main}", fmt_int(total_main))

with c2:
    if "Righe" in available_metrics:
        total_rows = fact_f.loc[fact_f["Metrica"] == "Righe", "Valore"].sum()
        st.metric("Totale Righe", fmt_int(total_rows))
    else:
        st.metric("Totale Righe", "—")

with c3:
    if "Colli" in available_metrics:
        total_colli = fact_f.loc[fact_f["Metrica"] == "Colli", "Valore"].sum()
        st.metric("Totale Colli", fmt_int(total_colli))
    else:
        st.metric("Totale Colli", "—")

# -------------------------
# Charts
# -------------------------
pie_df = (fact_f[fact_f["Metrica"] == metric_for_main]
          .groupby("Gate", as_index=False)["Valore"].sum()
          .sort_values("Valore", ascending=False))
if pie_df["Valore"].sum() <= 0:
    st.info("Nessun dato nel periodo/filtri selezionati.")
    st.stop()

st.plotly_chart(px.pie(pie_df, names="Gate", values="Valore", title=f"Peso Gate sul totale ({metric_for_main})"),
                use_container_width=True)

st.subheader(f"Dettaglio Gate → Giri ({metric_for_main})")
bar_df = fact_f[fact_f["Metrica"] == metric_for_main].groupby(["Gate", "Giro"], as_index=False)["Valore"].sum()
if len(bar_df):
    st.plotly_chart(px.bar(bar_df, x="Gate", y="Valore", color="Giro", barmode="stack",
                           title="Carico per Gate con dettaglio giri"),
                    use_container_width=True)

st.subheader("Confronto Righe vs Colli per Gate")
compare_df = fact_f.groupby(["Gate", "Metrica"], as_index=False)["Valore"].sum()
st.plotly_chart(px.bar(compare_df, x="Gate", y="Valore", color="Metrica", barmode="group",
                       title="Righe vs Colli per Gate"),
                use_container_width=True)

if ("Righe" in available_metrics) and ("Colli" in available_metrics):
    piv = compare_df.pivot(index="Gate", columns="Metrica", values="Valore").fillna(0.0).reset_index()
    piv["Colli_per_100_righe"] = np.where(piv["Righe"] > 0, (piv["Colli"] / piv["Righe"]) * 100, np.nan)
    st.dataframe(piv.sort_values("Righe", ascending=False), use_container_width=True)

st.subheader("Top Giri per Gate")
gate_sel = st.selectbox("Scegli Gate", options=sorted(fact_f["Gate"].unique().tolist()))
top_df = fact_f[fact_f["Gate"] == gate_sel].groupby(["Giro", "Metrica"], as_index=False)["Valore"].sum()
if len(top_df):
    rank_metric = metric_for_main if metric_for_main in available_metrics else available_metrics[0]
    rank = (top_df[top_df["Metrica"] == rank_metric]
            .sort_values("Valore", ascending=False)
            .head(topn)[["Giro", "Valore"]])
    st.dataframe(rank, use_container_width=True)
    st.plotly_chart(px.bar(rank, x="Giro", y="Valore", title=f"Top giri ({rank_metric}) - {gate_sel}"),
                    use_container_width=True)
    st.download_button("Scarica CSV Top giri", data=rank.to_csv(index=False).encode("utf-8"),
                       file_name=f"top_giri_{gate_sel}.csv", mime="text/csv")

st.subheader("Trend giornaliero per Gate")
metric_trend = st.selectbox("Metrica per trend", options=available_metrics, key="metric_trend")
trend_df = fact_f[fact_f["Metrica"] == metric_trend].groupby(["Data", "Gate"], as_index=False)["Valore"].sum()
if len(trend_df):
    st.plotly_chart(px.line(trend_df, x="Data", y="Valore", color="Gate", title=f"Trend giornaliero ({metric_trend})"),
                    use_container_width=True)

st.subheader("🚨 Alert")
alerts = []
tot_metric = fact_f.loc[fact_f["Metrica"] == metric_for_main, "Valore"].sum()
if tot_metric > 0:
    share = fact_f[fact_f["Metrica"] == metric_for_main].groupby("Gate", as_index=False)["Valore"].sum()
    share["pct"] = (share["Valore"] / tot_metric) * 100
    over = share[share["pct"] >= alert_gate_pct].sort_values("pct", ascending=False)
    for _, r in over.iterrows():
        alerts.append(f"Gate {r['Gate']} pesa {r['pct']:.1f}% del totale {metric_for_main} (soglia {alert_gate_pct}%).")

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
