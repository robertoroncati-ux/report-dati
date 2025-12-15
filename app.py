import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Report Vendite – Area Ponente", layout="wide")

# =====================================================
# STANDARD COLONNE
# =====================================================
STD = {
    "city": "Città",
    "agent": "Agente",
    "client": "Cliente",
    "category": "Categoria",
    "article": "Articolo",
    "f24": "Fatturato 2024",
    "f25": "Fatturato 2025",
}

RENAME_MAP = {
    "Citta": "Città",
    "CITTA": "Città",
    "AGENTE": "Agente",
    "Esercizio": "Cliente",
    "ESERCIZIO": "Cliente",
    "Fatturato2024": "Fatturato 2024",
    "Fatturato_2024": "Fatturato 2024",
    "Fatturato2025": "Fatturato 2025",
    "Fatturato_2025": "Fatturato 2025",
}

REQUIRED = list(STD.values())


# =====================================================
# UTILITY
# =====================================================
def clean_text(s):
    return s.astype(str).str.replace("\u00a0", " ", regex=False).str.strip()


@st.cache_data(show_spinner=False)
def load_excel(file):
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns=RENAME_MAP)

    for col in REQUIRED:
        if col not in df.columns:
            raise ValueError(f"Manca la colonna: {col}")

    for c in ["Agente", "Città", "Cliente", "Categoria", "Articolo"]:
        df[c] = clean_text(df[c])
        df = df[df[c] != ""]

    df["Fatturato 2024"] = pd.to_numeric(df["Fatturato 2024"], errors="coerce").fillna(0)
    df["Fatturato 2025"] = pd.to_numeric(df["Fatturato 2025"], errors="coerce").fillna(0)

    return df


def to_excel_bytes(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            sdf.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()


# =====================================================
# UI START
# =====================================================
st.title("📊 Analisi Vendite – Area Ponente")
st.caption("Versione con simulatore logistico semi-assistito")

uploaded = st.file_uploader("Carica file Excel", type=["xlsx", "xls"])
if not uploaded:
    st.stop()

df = load_excel(uploaded)

# =====================================================
# BASE CLIENTI (livello cliente)
# =====================================================
base_clienti = (
    df.groupby(["Agente", "Città", "Cliente"], as_index=False)
      .agg(fatturato_2024=("Fatturato 2024", "sum"),
           fatturato_2025=("Fatturato 2025", "sum"))
)

base_clienti = base_clienti[base_clienti["fatturato_2025"] > 0].copy()

# =====================================================
# REPORT ESISTENTI (RIDOTTI QUI PER SPAZIO)
# =====================================================
fatturato_agente = (
    base_clienti.groupby("Agente", as_index=False)
      .agg(fatturato=("fatturato_2025", "sum"),
           clienti=("Cliente", "nunique"),
           citta=("Città", "nunique"))
)

# =====================================================
# TABS
# =====================================================
tab1, tab2, tab3 = st.tabs([
    "Report sintetico",
    "Simulatore logistico",
    "Export"
])

# =====================================================
# TAB 1 – REPORT SINTETICO
# =====================================================
with tab1:
    st.subheader("Fatturato per agente (attuale)")
    st.dataframe(fatturato_agente, use_container_width=True)

# =====================================================
# TAB 2 – SIMULATORE LOGISTICO
# =====================================================
with tab2:
    st.subheader("🧠 Simulatore logistico semi-assistito")

    col1, col2, col3 = st.columns(3)
    with col1:
        ron_out = st.checkbox("Roncati non operativo", value=True)
    with col2:
        balducci_out = st.checkbox("Balducci non presente", value=True)
    with col3:
        nuovo_agente = st.checkbox("Inserimento nuovo agente", value=True)

    # STATO PRIMA
    stato_prima = (
        base_clienti.groupby("Agente", as_index=False)
          .agg(fatturato=("fatturato_2025", "sum"),
               clienti=("Cliente", "nunique"),
               citta=("Città", "nunique"))
    )

    st.markdown("### Stato PRIMA")
    st.dataframe(stato_prima, use_container_width=True)

    # MOVIMENTI
    st.markdown("### Movimenti (decisi da te)")
    movimenti = st.data_editor(
        pd.DataFrame(columns=[
            "Cliente", "Città", "Da agente", "A agente", "Fatturato", "Motivo"
        ]),
        num_rows="dynamic",
        use_container_width=True
    )

    # APPLICA SIMULAZIONE
    sim = base_clienti.copy()

    if ron_out:
        sim = sim[sim["Agente"] != "Roncati"]

    if balducci_out:
        sim = sim[sim["Agente"] != "Balducci"]

    for _, r in movimenti.dropna(subset=["Cliente"]).iterrows():
        sim.loc[
            (sim["Cliente"] == r["Cliente"]) &
            (sim["Città"] == r["Città"]) &
            (sim["Agente"] == r["Da agente"]),
            "fatturato_2025"
        ] -= r["Fatturato"]

        sim = pd.concat([sim, pd.DataFrame([{
            "Agente": r["A agente"],
            "Città": r["Città"],
            "Cliente": r["Cliente"],
            "fatturato_2024": 0,
            "fatturato_2025": r["Fatturato"]
        }])], ignore_index=True)

    # STATO DOPO
    stato_dopo = (
        sim.groupby("Agente", as_index=False)
          .agg(fatturato=("fatturato_2025", "sum"),
               clienti=("Cliente", "nunique"),
               citta=("Città", "nunique"))
    )

    confronto = stato_prima.merge(
        stato_dopo, on="Agente", how="outer", suffixes=("_prima", "_dopo")
    ).fillna(0)

    confronto["Delta fatturato"] = confronto["fatturato_dopo"] - confronto["fatturato_prima"]

    st.markdown("### Stato DOPO simulazione")
    st.dataframe(confronto, use_container_width=True)

    # CONTROLLO AREA
    tot_prima = stato_prima["fatturato"].sum()
    tot_dopo = stato_dopo["fatturato"].sum()

    if round(tot_prima, 2) != round(tot_dopo, 2):
        st.error(f"❌ Fatturato area NON coerente: {tot_prima:.0f} → {tot_dopo:.0f}")
    else:
        st.success(f"✅ Fatturato area invariato: {tot_dopo:.0f}")

    # NUOVO AGENTE
    if nuovo_agente:
        pacchetto = sim[sim["Agente"] == "Nuovo Agente"]
        st.markdown("### Pacchetto nuovo agente")
        st.dataframe(pacchetto, use_container_width=True)

# =====================================================
# TAB 3 – EXPORT
# =====================================================
with tab3:
    excel = to_excel_bytes({
        "BASE_CLIENTI": base_clienti,
        "STATO_PRIMA": stato_prima,
        "STATO_DOPO": stato_dopo,
        "MOVIMENTI": movimenti
    })

    st.download_button(
        "⬇️ Scarica scenario simulazione",
        data=excel,
        file_name="scenario_logistico_ponente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
