
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from io import BytesIO

st.set_page_config(page_title="Gate Workload", layout="wide")

# ---------- Helpers ----------

def norm_giro(x):
    if pd.isna(x):
        return ""
    try:
        if float(x).is_integer():
            return str(int(float(x)))
    except:
        pass
    return str(x).strip()

def safe_date(x):
    return pd.to_datetime(x, dayfirst=True, errors="coerce")

def sniff_report_type(file_bytes):
    raw = pd.read_excel(BytesIO(file_bytes), header=None, nrows=80)
    text = "\n".join(raw.astype(str).fillna("").values.flatten()).lower()
    if "preliev" in text:
        return "Righe"
    if "colli" in text or "collo" in text:
        return "Colli"
    return "Righe"

def parse_pivot(file_bytes):
    raw = pd.read_excel(BytesIO(file_bytes), header=None)
    marker = raw.astype(str).apply(lambda r: r.str.contains("Numero di Operazioni", case=False)).any(axis=1)
    if marker.any():
        raw = raw.iloc[:marker.idxmax()]

    header_row = 0
    for i in range(min(50, len(raw))):
        if raw.iloc[i].astype(str).str.contains("giro", case=False).any():
            header_row = i
            break

    df = raw.iloc[header_row+1:].copy()
    df.columns = raw.iloc[header_row]
    df = df.dropna(axis=1, how="all")

    df = df.rename(columns={df.columns[0]: "Data"})
    df["Data"] = df["Data"].apply(safe_date)
    df = df.dropna(subset=["Data"])

    value_cols = [c for c in df.columns if "tot" not in str(c).lower() and c != "Data"]
    long = df.melt(id_vars=["Data"], value_vars=value_cols, var_name="Giro", value_name="Valore")
    long["Giro"] = long["Giro"].apply(norm_giro)
    long["Valore"] = pd.to_numeric(long["Valore"], errors="coerce").fillna(0)
    long = long[long["Valore"] != 0]

    return long

def build_gate_map(file_bytes):
    g = pd.read_excel(BytesIO(file_bytes))
    giro = g.columns[1]
    gate = g.columns[9]
    m = g[[giro, gate]].copy()
    m.columns = ["Giro", "Gate"]
    m["Giro"] = m["Giro"].apply(norm_giro)
    m = m.dropna(subset=["Giro"])
    return m.drop_duplicates("Giro")

# ---------- UI ----------

st.title("Gate Workload • Stable")

gate_file = st.sidebar.file_uploader("GATE.xlsx (obbligatorio)", type=["xlsx"])
file1 = st.sidebar.file_uploader("Export EasyMag", type=["xlsx"])

if gate_file is None or file1 is None:
    st.stop()

gate_map = build_gate_map(gate_file.getvalue())

fb = file1.getvalue()
metric = sniff_report_type(fb)
data = parse_pivot(fb)
data["Metrica"] = metric

fact = data.merge(gate_map, on="Giro", how="left")

if fact["Gate"].isna().any():
    st.error("Ci sono giri senza Gate assegnato. Correggi GATE.xlsx.")
    st.stop()

min_d = fact["Data"].min().date()
max_d = fact["Data"].max().date()

if min_d == max_d:
    start = end = st.sidebar.date_input("Periodo", value=min_d)
else:
    period = st.sidebar.date_input("Periodo", value=(min_d, max_d))
    if isinstance(period, tuple):
        start, end = period
    else:
        start = end = period

fact = fact[(fact["Data"] >= pd.to_datetime(start)) &
            (fact["Data"] <= pd.to_datetime(end) + pd.Timedelta(days=1))]

st.metric(f"Totale {metric}", f"{fact['Valore'].sum():,.0f}".replace(",", "."))

pie = fact.groupby("Gate", as_index=False)["Valore"].sum()
st.plotly_chart(px.pie(pie, names="Gate", values="Valore"), use_container_width=True)

bar = fact.groupby(["Gate","Giro"], as_index=False)["Valore"].sum()
st.plotly_chart(px.bar(bar, x="Gate", y="Valore", color="Giro", barmode="stack"), use_container_width=True)
