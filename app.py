import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Report Vendite - Analisi Città/Agente", layout="wide")

# =============================
# Colonne: standard + alias
# =============================
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
    # Anagrafiche
    "Città": STD["city"], "Citta": STD["city"], "CITTA": STD["city"],
    "Agente": STD["agent"], "AGENTE": STD["agent"],
    "Esercizio": STD["client"], "Cliente": STD["client"], "ESERCIZIO": STD["client"],
    "Categoria": STD["category"], "CATEGORIA": STD["category"],
    "Articolo": STD["article"], "ARTICOLO": STD["article"],

    # Fatturati 2024
    "Fatturato 2024": STD["f24"],
    "Fatturato2024": STD["f24"],
    "Fatturato_2024": STD["f24"],

    # Fatturati 2025
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

    # Pulizia testi
    for c in [STD["city"], STD["agent"], STD["client"], STD["category"], STD["article"]]:
        df[c] = _clean_text(df[c]).replace({"nan": ""})

    # Numerici
    for c in [STD["f24"], STD["f25"]]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    return df


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            sdf.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()


# =============================
# UI
# =============================
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

st.caption(f"Righe totali file: **{len(df):,}**".replace(",", "."))

# =========================================================
# BASE CORRETTA (qui è dove prima si “rompeva” il 2024)
# - ac_client_all: tutto il file aggregato a livello cliente
# - ac_client_active: SOLO clienti attivi (fatt 2025 > 0)
# =========================================================
ac_client_all = (
    df.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)[[STD["f24"], STD["f25"]]]
      .sum()
)

ac_client_active = ac_client_all[ac_client_all[STD["f25"]] > 0].copy()

if ac_client_active.empty:
    st.warning("Dopo il filtro (Fatturato 2025 > 0) non rimane nessuna riga. Controlla il file.")
    st.stop()

st.caption(
    f"Clienti/righe aggregati (tutto): **{len(ac_client_all):,}** | "
    f"Clienti attivi 2025 (fatt 2025 > 0): **{len(ac_client_active):,}**"
    .replace(",", ".")
)

# =============================
# 1) FATTURATO PER CITTA' (sintesi + dettaglio) - SOLO ATTIVI 2025
# =============================
city_summary = (
    ac_client_active.groupby(STD["city"], as_index=False)
      .agg(
          fatturato_citta_2025=(STD["f25"], "sum"),
          n_agenti=("Agente", pd.Series.nunique),
          n_clienti_attivi=("Cliente", pd.Series.nunique),
      )
      .sort_values("fatturato_citta_2025", ascending=False)
)

city_detail = (
    ac_client_active.groupby([STD["city"], STD["agent"]], as_index=False)
      .agg(
          fatturato_agente_nella_citta_2025=(STD["f25"], "sum"),
          n_clienti_attivi_agente_nella_citta=("Cliente", pd.Series.nunique),
      )
)

# % incidenza su città
city_total_map = city_summary.set_index(STD["city"])["fatturato_citta_2025"]
city_detail["%_incidenza_su_citta"] = (
    city_detail.apply(
        lambda r: (r["fatturato_agente_nella_citta_2025"] / city_total_map.get(r[STD["city"]], 1)) * 100,
        axis=1
    )
)
city_detail = city_detail.sort_values([STD["city"], "fatturato_agente_nella_citta_2025"], ascending=[True, False])

# =============================
# 2) FATTURATO AGENTE - Totali 2024/2025 su TUTTO + KPI clienti attivi/persi/acquisiti
# =============================
agent_totals = (
    ac_client_all.groupby(STD["agent"], as_index=False)[[STD["f24"], STD["f25"]]]
      .sum()
)

# Clienti attivi = solo fatt 2025 > 0
agent_active = (
    ac_client_active.groupby(STD["agent"], as_index=False)
      .agg(clienti_attivi_2025=(STD["client"], "nunique"))
)

# Clienti persi/acquisiti/mantenuti (basati sul confronto 2024 e 2025, su TUTTO)
tmp = ac_client_all.copy()
tmp["attivo_2024"] = tmp[STD["f24"]] > 0
tmp["attivo_2025"] = tmp[STD["f25"]] > 0

tmp["stato_cliente"] = np.select(
    [
        tmp["attivo_2024"] & tmp["attivo_2025"],
        tmp["attivo_2024"] & ~tmp["attivo_2025"],
        ~tmp["attivo_2024"] & tmp["attivo_2025"],
    ],
    ["Mantenuto", "Perso", "Acquisito"],
    default="Inattivo"
)

agent_client_flow = (
    tmp.groupby(STD["agent"], as_index=False)
      .agg(
          clienti_persi=("stato_cliente", lambda s: (s == "Perso").sum()),
          clienti_acquisiti=("stato_cliente", lambda s: (s == "Acquisito").sum()),
          clienti_mantenuti=("stato_cliente", lambda s: (s == "Mantenuto").sum()),
      )
)

agent_report = agent_totals.merge(agent_active, on=STD["agent"], how="left") \
                          .merge(agent_client_flow, on=STD["agent"], how="left")

agent_report["clienti_attivi_2025"] = agent_report["clienti_attivi_2025"].fillna(0).astype(int)
for c in ["clienti_persi", "clienti_acquisiti", "clienti_mantenuti"]:
    agent_report[c] = agent_report[c].fillna(0).astype(int)

agent_report["Delta 2025 vs 2024"] = agent_report[STD["f25"]] - agent_report[STD["f24"]]

agent_report = agent_report.rename(columns={
    STD["f25"]: "Fatturato totale 2025",
    STD["f24"]: "Fatturato totale 2024",
})
agent_report = agent_report.sort_values("Fatturato totale 2025", ascending=False)

# Dettaglio clienti persi/acquisiti per agente (utile da export)
agent_client_detail = tmp[tmp["stato_cliente"].isin(["Perso", "Acquisito", "Mantenuto"])][
    [STD["agent"], STD["city"], STD["client"], STD["f24"], STD["f25"], "stato_cliente"]
].sort_values([STD["agent"], "stato_cliente", STD["f25"]], ascending=[True, True, False])

# =============================
# 3) FATTURATO PER ZONA (Agente -> Città) + dettaglio cliente - SOLO ATTIVI 2025
# =============================
zone_agent_city = (
    ac_client_active.groupby([STD["agent"], STD["city"]], as_index=False)
      .agg(
          fatturato_2025=(STD["f25"], "sum"),
          clienti_attivi_2025=(STD["client"], "nunique"),
      )
      .sort_values([STD["agent"], "fatturato_2025"], ascending=[True, False])
)

zone_client_detail = (
    ac_client_active.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)
      .agg(fatturato_cliente_2025=(STD["f25"], "sum"))
      .sort_values([STD["agent"], STD["city"], "fatturato_cliente_2025"], ascending=[True, True, False])
)

zone_agent_total = (
    zone_agent_city.groupby(STD["agent"], as_index=False)
      .agg(
          fatturato_totale_2025=("fatturato_2025", "sum"),
          n_citta_attive=(STD["city"], "nunique"),
          clienti_attivi_totali=("clienti_attivi_2025", "sum"),
      )
      .sort_values("fatturato_totale_2025", ascending=False)
)

# =============================
# 4) FATTURATO PER CATEGORIA - SOLO ATTIVI 2025
# prodotti venduti = numero articoli distinti
# =============================
df_active_rows = df[df[STD["f25"]] > 0].copy()

agent_category = (
    df_active_rows.groupby([STD["agent"], STD["category"]], as_index=False)
      .agg(
          fatturato_categoria_2025=(STD["f25"], "sum"),
          prodotti_distinti_nella_categoria=(STD["article"], "nunique"),
      )
      .sort_values([STD["agent"], "fatturato_categoria_2025"], ascending=[True, False])
)

# =============================
# UI Tabs
# =============================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1) Fatturato per Città",
    "2) Fatturato Agente",
    "3) Fatturato per Zona",
    "4) Fatturato per Categoria",
    "Export Excel"
])

with tab1:
    st.subheader("Fatturato per Città (Sintesi) — SOLO clienti attivi 2025")
    st.dataframe(city_summary, use_container_width=True, height=420)

    st.subheader("Dettaglio: Agenti nella Città — SOLO clienti attivi 2025")
    st.dataframe(city_detail, use_container_width=True, height=520)

with tab2:
    st.subheader("Fatturato Agente — Totali 2024/2025 su TUTTO + clienti attivi/persi/acquisiti")
    st.dataframe(agent_report, use_container_width=True, height=520)

    st.subheader("Dettaglio clienti persi / acquisiti / mantenuti (per agente)")
    st.dataframe(agent_client_detail, use_container_width=True, height=520)

with tab3:
    st.subheader("Fatturato per Zona (Agente → Città) — SOLO clienti attivi 2025")
    st.dataframe(zone_agent_city, use_container_width=True, height=420)

    st.subheader("Dettaglio: Clienti (Agente → Città → Cliente) — SOLO clienti attivi 2025")
    st.dataframe(zone_client_detail, use_container_width=True, height=520)

with tab4:
    st.subheader("Fatturato per Categoria (Agente → Categoria) — SOLO clienti attivi 2025")
    st.dataframe(agent_category, use_container_width=True, height=650)

with tab5:
    st.subheader("Esporta tutte le tabelle in Excel (multi-foglio)")

    sheets = {
        "FATT_CITTA_SINTESI": city_summary,
        "FATT_CITTA_DETTAGLIO": city_detail,

        "FATT_AGENTE": agent_report,
        "DETT_CLIENTI_FLOW": agent_client_detail,

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
