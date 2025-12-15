import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Report Vendite - Analisi Città/Agente", layout="wide")

# -----------------------------
# Config colonne (robusto ai nomi)
# -----------------------------
# Standard interni che useremo nell'app
STD = {
    "city": "Città",
    "agent": "Agente",
    "client": "Cliente",
    "category": "Categoria",
    "article": "Articolo",
    "f24": "Fatturato 2024",
    "f25": "Fatturato 2025",
}

# Alias possibili nel file Excel -> standard
RENAME_MAP = {
    # città/agente/cliente
    "Città": STD["city"], "Citta": STD["city"], "CITTA": STD["city"],
    "Agente": STD["agent"], "AGENTE": STD["agent"],
    "Esercizio": STD["client"], "Cliente": STD["client"], "ESERCIZIO": STD["client"],
    "Categoria": STD["category"], "CATEGORIA": STD["category"],
    "Articolo": STD["article"], "ARTICOLO": STD["article"],

    # fatturati 2024
    "Fatturato 2024": STD["f24"],
    "Fatturato2024": STD["f24"],
    "Fatturato_2024": STD["f24"],

    # fatturati 2025
    "Fatturato 2025": STD["f25"],
    "Fatturato2025": STD["f25"],
    "Fatturato_2025": STD["f25"],
}

REQUIRED = [STD["city"], STD["agent"], STD["client"], STD["category"], STD["article"], STD["f24"], STD["f25"]]

def _clean_text(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
         .str.replace("\u00a0", " ", regex=False)
         .str.strip()
    )

@st.cache_data(show_spinner=False)
def load_excel(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns=RENAME_MAP)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Mancano queste colonne nel file Excel: {missing}")

    # pulizia testi
    for c in [STD["city"], STD["agent"], STD["client"], STD["category"], STD["article"]]:
        df[c] = _clean_text(df[c]).replace({"nan": ""})

    # numerici
    for c in [STD["f24"], STD["f25"]]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    return df

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            sdf.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()

# -----------------------------
# UI
# -----------------------------
st.title("📊 Report Vendite (Analisi Città / Agente)")

uploaded = st.file_uploader("Carica il file Excel", type=["xlsx", "xls"])
if not uploaded:
    st.info("Carica un Excel per iniziare.")
    st.stop()

try:
    df = load_excel(uploaded)
except Exception as e:
    st.error(str(e))
    st.stop()

# -----------------------------
# FILTRO FONDAMENTALE: solo clienti attivi (Fatturato 2025 > 0)
# -----------------------------
df_active = df[df[STD["f25"]] > 0].copy()

if df_active.empty:
    st.warning("Dopo il filtro (Fatturato 2025 > 0) non rimane nessuna riga. Controlla il file.")
    st.stop()

st.caption(f"Righe totali file: **{len(df):,}** | Righe attive (Fatturato 2025 > 0): **{len(df_active):,}**".replace(",", "."))

# -----------------------------
# BASE PER KPI CLIENTI: aggrego a livello Agente-Città-Cliente
# così contiamo i clienti una sola volta anche se compaiono su più prodotti
# -----------------------------
ac_client = (
    df_active.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)[[STD["f24"], STD["f25"]]]
             .sum()
)

# -----------------------------
# 1) FATTURATO PER CITTA' (SINTESI + DETTAGLIO)
# -----------------------------
city_summary = (
    ac_client.groupby(STD["city"], as_index=False)
             .agg(
                 fatturato_citta_2025=(STD["f25"], "sum"),
                 n_agenti=("Agente", pd.Series.nunique),
                 n_clienti_attivi=("Cliente", pd.Series.nunique),
             )
             .sort_values("fatturato_citta_2025", ascending=False)
)

city_detail = (
    ac_client.groupby([STD["city"], STD["agent"]], as_index=False)
             .agg(
                 fatturato_agente_nella_citta_2025=(STD["f25"], "sum"),
                 n_clienti_attivi_agente_nella_citta=("Cliente", pd.Series.nunique),
             )
)

# percentuale incidenza agente sulla città
city_total_map = city_summary.set_index(STD["city"])["fatturato_citta_2025"]
city_detail["%_incidenza_su_citta"] = (
    city_detail.apply(lambda r: (r["fatturato_agente_nella_citta_2025"] / city_total_map.get(r[STD["city"]], 1)) * 100, axis=1)
)

city_detail = city_detail.sort_values([STD["city"], "fatturato_agente_nella_citta_2025"], ascending=[True, False])

# -----------------------------
# 2) FATTURATO AGENTE (fatt 2025, delta vs 2024, clienti attivi)
# -----------------------------
agent_report = (
    ac_client.groupby(STD["agent"], as_index=False)[[STD["f24"], STD["f25"]]]
             .sum()
)
agent_clients = (
    ac_client.groupby(STD["agent"], as_index=False)
             .agg(clienti_attivi_2025=(STD["client"], "nunique"))
)

agent_report = agent_report.merge(agent_clients, on=STD["agent"], how="left")
agent_report["Delta 2025 vs 2024"] = agent_report[STD["f25"]] - agent_report[STD["f24"]]
agent_report = agent_report.rename(columns={
    STD["f25"]: "Fatturato totale 2025",
    STD["f24"]: "Fatturato totale 2024",
})
agent_report = agent_report.sort_values("Fatturato totale 2025", ascending=False)

# -----------------------------
# 3) FATTURATO PER ZONA (Agente -> Città) + dettaglio per cliente
# -----------------------------
zone_agent_city = (
    ac_client.groupby([STD["agent"], STD["city"]], as_index=False)
             .agg(
                 fatturato_2025=(STD["f25"], "sum"),
                 clienti_attivi_2025=(STD["client"], "nunique"),
             )
             .sort_values([STD["agent"], "fatturato_2025"], ascending=[True, False])
)

# Dettaglio clienti per (Agente, Città)
zone_client_detail = (
    ac_client.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)
             .agg(fatturato_cliente_2025=(STD["f25"], "sum"))
             .sort_values([STD["agent"], STD["city"], "fatturato_cliente_2025"], ascending=[True, True, False])
)

# (facoltativo) Totale per agente come sintesi aggiuntiva, utile in export
zone_agent_total = (
    zone_agent_city.groupby(STD["agent"], as_index=False)
                  .agg(fatturato_totale_2025=("fatturato_2025", "sum"),
                       n_citta_attive=(STD["city"], "nunique"),
                       clienti_attivi_totali=("clienti_attivi_2025", "sum"))
                  .sort_values("fatturato_totale_2025", ascending=False)
)

# -----------------------------
# 4) FATTURATO PER CATEGORIA (Agente, Categoria, fatturato, prodotti venduti)
# "prodotti venduti" = numero articoli distinti (non quantità)
# -----------------------------
agent_category = (
    df_active.groupby([STD["agent"], STD["category"]], as_index=False)
             .agg(
                 fatturato_categoria_2025=(STD["f25"], "sum"),
                 prodotti_distinti_nella_categoria=(STD["article"], "nunique"),
             )
             .sort_values([STD["agent"], "fatturato_categoria_2025"], ascending=[True, False])
)

# -----------------------------
# UI con tabs (sintesi + dettaglio dove serve)
# -----------------------------
tab1, tab2, tab3, tab4, tab_export = st.tabs([
    "1) Fatturato per Città",
    "2) Fatturato Agente",
    "3) Fatturato per Zona",
    "4) Fatturato per Categoria",
    "Export Excel"
])

with tab1:
    st.subheader("Tabella: Fatturato per Città (Sintesi)")
    st.dataframe(city_summary, use_container_width=True, height=420)

    st.subheader("Dettaglio: Agenti nella Città")
    st.dataframe(city_detail, use_container_width=True, height=520)

with tab2:
    st.subheader("Tabella: Fatturato Agente")
    st.dataframe(agent_report, use_container_width=True, height=650)

with tab3:
    st.subheader("Tabella: Fatturato per Zona (Agente → Città)")
    st.dataframe(zone_agent_city, use_container_width=True, height=420)

    st.subheader("Dettaglio: Clienti (Agente → Città → Cliente)")
    st.dataframe(zone_client_detail, use_container_width=True, height=520)

with tab4:
    st.subheader("Tabella: Fatturato per Categoria (Agente → Categoria)")
    st.dataframe(agent_category, use_container_width=True, height=650)

with tab_export:
    st.subheader("Esporta tutte le tabelle in Excel (multi-foglio)")

    sheets = {
        "FATT_CITTA_SINTESI": city_summary,
        "FATT_CITTA_DETTAGLIO": city_detail,
        "FATT_AGENTE": agent_report,
        "ZONA_AGENTE_TOTALE": zone_agent_total,
        "ZONA_AGENTE_CITTA": zone_agent_city,
        "ZONA_CLIENTI_DETT": zone_client_detail,
        "FATT_CATEGORIA": agent_category,
    }

    excel_bytes = to_excel_bytes(sheets)

    st.download_button(
        label="⬇️ Scarica Excel report",
        data=excel_bytes,
        file_name="report_vendite_analisi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
