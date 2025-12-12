import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import matplotlib.pyplot as plt

st.set_page_config(page_title="Report Vendite", layout="wide")

# -----------------------------
# Utils
# -----------------------------
REQUIRED_COLS = [
    "Città", "Agente", "Esercizio", "Categoria", "Articolo", "Fatturato 2024", "Fatturato 2025"
]

def _clean_text(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
         .str.replace("\u00a0", " ", regex=False)
         .str.strip()
    )

@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)

    # Normalizzo nomi colonne (tolgo spazi finali/iniziali)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Mancano colonne obbligatorie: {missing}")

    # Pulizia testi
    for c in ["Città", "Agente", "Esercizio", "Categoria", "Articolo"]:
        df[c] = _clean_text(df[c]).replace({"nan": ""})

    # Numerici
    for c in ["Fatturato 2024", "Fatturato 2025"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    # Teniamo solo righe "sensate": almeno un fatturato > 0 o almeno un campo compilato
    # (se vuoi tenerle tutte, togli questa riga)
    df = df.copy()

    return df

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            # nome foglio max 31 caratteri
            safe = name[:31]
            sdf.to_excel(writer, index=False, sheet_name=safe)
    return output.getvalue()

def euro_fmt(x: float) -> str:
    # Formattazione rapida stile IT
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def plot_bar(df: pd.DataFrame, x_col: str, y_col: str, title: str, top_n: int = 20):
    # df già ordinato o no: ordino per y desc e prendo top_n
    tmp = df[[x_col, y_col]].copy()
    tmp = tmp.sort_values(by=y_col, ascending=False).head(top_n)
    fig = plt.figure()
    plt.bar(tmp[x_col].astype(str), tmp[y_col])
    plt.xticks(rotation=45, ha="right")
    plt.title(title)
    plt.tight_layout()
    st.pyplot(fig, clear_figure=True)

# -----------------------------
# UI
# -----------------------------
st.title("Report Vendite (Excel)")

uploaded = st.file_uploader("Carica il file Excel", type=["xlsx", "xls"])

if not uploaded:
    st.info("Carica un Excel per iniziare.")
    st.stop()

try:
    df = load_excel(uploaded)
except Exception as e:
    st.error(f"Errore nel caricamento: {e}")
    st.stop()

# Rename “Esercizio” in UI (solo visuale)
df_ui = df.rename(columns={"Esercizio": "Cliente"})

# -----------------------------
# REPORT BASE (parte 1)
# -----------------------------
# Città
report_city = (
    df_ui.groupby("Città", as_index=False)[["Fatturato 2024", "Fatturato 2025"]]
        .sum()
)
report_city["Delta"] = report_city["Fatturato 2025"] - report_city["Fatturato 2024"]
report_city = report_city.sort_values("Fatturato 2025", ascending=False)

# Agente
report_agent = (
    df_ui.groupby("Agente", as_index=False)[["Fatturato 2024", "Fatturato 2025"]]
        .sum()
)
report_agent["Delta"] = report_agent["Fatturato 2025"] - report_agent["Fatturato 2024"]
report_agent = report_agent.sort_values("Fatturato 2025", ascending=False)

# -----------------------------
# NUOVI REPORT
# -----------------------------

# 1) Agente x Categoria
agent_category = (
    df_ui.groupby(["Agente", "Categoria"], as_index=False)[["Fatturato 2024", "Fatturato 2025"]]
        .sum()
)
agent_category["Delta"] = agent_category["Fatturato 2025"] - agent_category["Fatturato 2024"]
agent_category = agent_category.sort_values(["Agente", "Fatturato 2025"], ascending=[True, False])

# 2) Agente x Prodotto (Articolo)
agent_product = (
    df_ui.groupby(["Agente", "Articolo"], as_index=False)[["Fatturato 2024", "Fatturato 2025"]]
        .sum()
)
agent_product["Delta"] = agent_product["Fatturato 2025"] - agent_product["Fatturato 2024"]
agent_product = agent_product.sort_values(["Agente", "Fatturato 2025"], ascending=[True, False])

# 3) Top prodotti per fatturato per agente (top 10)
top_products_per_agent = (
    agent_product.sort_values(["Agente", "Fatturato 2025"], ascending=[True, False])
                .groupby("Agente", as_index=False)
                .head(10)
)

# 4) Confronto 2024 vs 2025 per agente + clienti persi/acquisiti
# Prima aggrego a livello Agente-Cliente
agent_client = (
    df_ui.groupby(["Agente", "Cliente"], as_index=False)[["Fatturato 2024", "Fatturato 2025"]]
        .sum()
)

agent_client["attivo_2024"] = agent_client["Fatturato 2024"] > 0
agent_client["attivo_2025"] = agent_client["Fatturato 2025"] > 0

# Classificazione clienti
agent_client["stato_cliente"] = np.select(
    [
        agent_client["attivo_2024"] & agent_client["attivo_2025"],
        agent_client["attivo_2024"] & ~agent_client["attivo_2025"],
        ~agent_client["attivo_2024"] & agent_client["attivo_2025"],
    ],
    ["Mantenuto", "Perso", "Acquisito"],
    default="Inattivo"
)

# KPI per agente
agent_kpi = (
    agent_client.groupby("Agente", as_index=False)
        .agg(
            fatt_2024=("Fatturato 2024", "sum"),
            fatt_2025=("Fatturato 2025", "sum"),
            clienti_2024=("attivo_2024", "sum"),
            clienti_2025=("attivo_2025", "sum"),
            clienti_persi=("stato_cliente", lambda s: (s == "Perso").sum()),
            clienti_acquisiti=("stato_cliente", lambda s: (s == "Acquisito").sum()),
            clienti_mantenuti=("stato_cliente", lambda s: (s == "Mantenuto").sum()),
        )
)
agent_kpi["delta"] = agent_kpi["fatt_2025"] - agent_kpi["fatt_2024"]
agent_kpi = agent_kpi.sort_values("fatt_2025", ascending=False)

# -----------------------------
# TABS UI
# -----------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Base: Città / Agente",
    "Agente x Categoria",
    "Agente x Prodotto",
    "Top prodotti per agente",
    "Confronto + Clienti persi/acquisiti"
])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Fatturato per Città")
        st.dataframe(report_city, use_container_width=True, height=520)
        plot_bar(report_city, "Città", "Fatturato 2025", "Top città per fatturato 2025", top_n=20)

    with c2:
        st.subheader("Fatturato per Agente")
        st.dataframe(report_agent, use_container_width=True, height=520)
        plot_bar(report_agent, "Agente", "Fatturato 2025", "Fatturato 2025 per agente", top_n=30)

with tab2:
    st.subheader("Agente per Categoria (somma fatturato)")
    st.dataframe(agent_category, use_container_width=True, height=650)

    # Piccolo focus: top categorie globali 2025
    st.markdown("**Top categorie 2025 (globale)**")
    top_cat = df_ui.groupby("Categoria", as_index=False)["Fatturato 2025"].sum().sort_values("Fatturato 2025", ascending=False)
    plot_bar(top_cat, "Categoria", "Fatturato 2025", "Top categorie 2025", top_n=20)

with tab3:
    st.subheader("Agente per Prodotto (Articolo)")
    st.dataframe(agent_product, use_container_width=True, height=650)

    st.markdown("**Top prodotti 2025 (globale)**")
    top_prod = df_ui.groupby("Articolo", as_index=False)["Fatturato 2025"].sum().sort_values("Fatturato 2025", ascending=False)
    plot_bar(top_prod, "Articolo", "Fatturato 2025", "Top prodotti 2025", top_n=20)

with tab4:
    st.subheader("Top 10 prodotti per agente (per fatturato 2025)")
    st.dataframe(top_products_per_agent, use_container_width=True, height=650)

with tab5:
    st.subheader("Confronto 2024 vs 2025 per agente + clienti persi/acquisiti")
    st.dataframe(agent_kpi, use_container_width=True, height=500)

    st.markdown("**Dettaglio clienti persi/acquisiti per agente (righe)**")
    # Mostro solo Persi/Acquisiti/Mantenuti (tolgo Inattivo)
    detail = agent_client[agent_client["stato_cliente"].isin(["Perso", "Acquisito", "Mantenuto"])].copy()
    detail = detail.sort_values(["Agente", "stato_cliente", "Fatturato 2025"], ascending=[True, True, False])
    st.dataframe(detail[["Agente", "Cliente", "Fatturato 2024", "Fatturato 2025", "stato_cliente"]], use_container_width=True, height=650)

# -----------------------------
# EXPORT
# -----------------------------
st.divider()
st.subheader("Esporta report in Excel")

sheets = {
    "BASE_CITTA": report_city,
    "BASE_AGENTE": report_agent,
    "AGENTE_CATEGORIA": agent_category,
    "AGENTE_PRODOTTO": agent_product,
    "TOP_PROD_PER_AGENTE": top_products_per_agent,
    "AGENTE_KPI": agent_kpi,
    "DET_CLIENTI": detail[["Agente", "Cliente", "Fatturato 2024", "Fatturato 2025", "stato_cliente"]],
}

excel_bytes = to_excel_bytes(sheets)

st.download_button(
    label="Scarica Excel (tutti i fogli)",
    data=excel_bytes,
    file_name="report_vendite.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
