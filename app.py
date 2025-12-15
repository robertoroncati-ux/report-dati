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
    "Città": STD["city"], "Citta": STD["city"], "CITTA": STD["city"],
    "Agente": STD["agent"], "AGENTE": STD["agent"],
    "Esercizio": STD["client"], "Cliente": STD["client"], "ESERCIZIO": STD["client"],
    "Categoria": STD["category"], "CATEGORIA": STD["category"],
    "Articolo": STD["article"], "ARTICOLO": STD["article"],

    "Fatturato 2024": STD["f24"],
    "Fatturato2024": STD["f24"],
    "Fatturato_2024": STD["f24"],

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

    # 🔥 FIX RIGA 0: elimino righe senza chiavi minime (Agente/Città/Cliente)
    df = df[
        (df[STD["agent"]] != "") &
        (df[STD["city"]] != "") &
        (df[STD["client"]] != "")
    ].copy()

    return df


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            sdf.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()


def add_total_row(df_in: pd.DataFrame, label_map: dict, sum_cols: list, first: bool = True) -> pd.DataFrame:
    df = df_in.copy()
    total = {c: np.nan for c in df.columns}

    for k, v in label_map.items():
        if k in total:
            total[k] = v

    for c in sum_cols:
        if c in df.columns:
            total[c] = df[c].sum()

    total_df = pd.DataFrame([total])
    if first:
        return pd.concat([total_df, df], ignore_index=True)
    return pd.concat([df, total_df], ignore_index=True)


def highlight_incidenza(val):
    try:
        return "background-color: #ffb3b3" if float(val) > 30 else ""
    except Exception:
        return ""


# =============================
# UI
# =============================
st.title("📊 Report Vendite (Analisi Città / Agente)")
st.caption("VERSIONE APP: v5 - fix righe fantasma + % incidenza categoria + export per tab")

uploaded = st.file_uploader("Carica il file Excel", type=["xlsx", "xls"])
if not uploaded:
    st.info("Carica un Excel per iniziare.")
    st.stop()

try:
    df = load_excel(uploaded)
except Exception as e:
    st.error(str(e))
    st.stop()

st.caption(f"Righe totali file (dopo pulizia chiavi): **{len(df):,}**".replace(",", "."))

# =========================================================
# Base:
# - ac_client_all: tutto il file aggregato a livello Agente-Città-Cliente
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
    f"Clienti aggregati (tutto): **{len(ac_client_all):,}** | "
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
          n_agenti=(STD["agent"], pd.Series.nunique),
          n_clienti_attivi=(STD["client"], pd.Series.nunique),
      )
      .sort_values("fatturato_citta_2025", ascending=False)
)

city_detail = (
    ac_client_active.groupby([STD["city"], STD["agent"]], as_index=False)
      .agg(
          fatturato_agente_nella_citta_2025=(STD["f25"], "sum"),
          n_clienti_attivi_agente_nella_citta=(STD["client"], pd.Series.nunique),
      )
)

city_total_map = city_summary.set_index(STD["city"])["fatturato_citta_2025"]
city_detail["%_incidenza_su_citta"] = (
    city_detail.apply(
        lambda r: (r["fatturato_agente_nella_citta_2025"] / city_total_map.get(r[STD["city"]], 1)) * 100,
        axis=1
    )
).round(2)
city_detail = city_detail.sort_values([STD["city"], "fatturato_agente_nella_citta_2025"], ascending=[True, False])

# =============================
# 2) FATTURATO AGENTE - Totali 2024/2025 su TUTTO + flusso clienti + incidenza%
# =============================
agent_totals = (
    ac_client_all.groupby(STD["agent"], as_index=False)[[STD["f24"], STD["f25"]]]
      .sum()
)

agent_active = (
    ac_client_active.groupby(STD["agent"], as_index=False)
      .agg(clienti_attivi_2025=(STD["client"], "nunique"))
)

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

agent_report = agent_report.rename(columns={
    STD["f25"]: "Fatturato totale 2025",
    STD["f24"]: "Fatturato totale 2024",
})

totale_area_2025 = agent_report["Fatturato totale 2025"].sum()
agent_report["% incidenza su totale area"] = (
    agent_report["Fatturato totale 2025"] / totale_area_2025 * 100
).round(2)

agent_report = agent_report.sort_values("Fatturato totale 2025", ascending=False)

top3 = agent_report.head(3).copy()
top3_share = (top3["Fatturato totale 2025"].sum() / totale_area_2025 * 100) if totale_area_2025 else 0
top3_table = top3[[STD["agent"], "Fatturato totale 2025", "% incidenza su totale area"]].copy()
top3_table = top3_table.rename(columns={STD["agent"]: "Agente"})

agent_client_detail = tmp[tmp["stato_cliente"].isin(["Perso", "Acquisito", "Mantenuto"])][
    [STD["agent"], STD["city"], STD["client"], STD["f24"], STD["f25"], "stato_cliente"]
].sort_values([STD["agent"], "stato_cliente", STD["f25"]], ascending=[True, True, False])

agent_view = agent_report.rename(columns={STD["agent"]: "Agente"}).copy()
agent_view = agent_view[
    [
        "Agente",
        "Fatturato totale 2024",
        "Fatturato totale 2025",
        "% incidenza su totale area",
        "clienti_attivi_2025",
        "clienti_persi",
        "clienti_acquisiti",
        "clienti_mantenuti",
    ]
]

# =============================
# 3) FATTURATO PER ZONA (Agente -> Città) + dettaglio cliente - SOLO ATTIVI 2025
# Totali nominati, UNA volta sola
# =============================
zone_agent_city = (
    ac_client_active.groupby([STD["agent"], STD["city"]], as_index=False)
      .agg(
          fatturato_2025=(STD["f25"], "sum"),
          clienti_attivi_2025=(STD["client"], "nunique"),
      )
      .sort_values([STD["agent"], "fatturato_2025"], ascending=[True, False])
)

zone_agent_city = add_total_row(
    zone_agent_city,
    label_map={STD["agent"]: "TOTALE AREA", STD["city"]: "TOTALE"},
    sum_cols=["fatturato_2025"],
    first=True
)

zone_client_detail = (
    ac_client_active.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)
      .agg(fatturato_cliente_2025=(STD["f25"], "sum"))
      .sort_values([STD["agent"], STD["city"], "fatturato_cliente_2025"], ascending=[True, True, False])
)

zone_client_detail = add_total_row(
    zone_client_detail,
    label_map={STD["agent"]: "TOTALE AREA", STD["city"]: "TOTALE", STD["client"]: "TOTALE"},
    sum_cols=["fatturato_cliente_2025"],
    first=True
)

zone_agent_total = (
    zone_agent_city[zone_agent_city[STD["agent"]] != "TOTALE AREA"]
      .groupby(STD["agent"], as_index=False)
      .agg(
          fatturato_totale_2025=("fatturato_2025", "sum"),
          n_citta_attive=(STD["city"], "nunique"),
          clienti_attivi_totali=("clienti_attivi_2025", "sum"),
      )
      .sort_values("fatturato_totale_2025", ascending=False)
)

# =============================
# 4) FATTURATO PER CATEGORIA - SOLO ATTIVI 2025
# Qui mettiamo % incidenza sul fatturato dell'agente (2025)
# =============================
df_active_rows = df[df[STD["f25"]] > 0].copy()

# Totale 2025 per agente (sul perimetro attivo 2025)
agent_total_2025_active = (
    df_active_rows.groupby(STD["agent"], as_index=False)
      .agg(fatturato_agente_2025=(STD["f25"], "sum"))
)

agent_category = (
    df_active_rows.groupby([STD["agent"], STD["category"]], as_index=False)
      .agg(fatturato_categoria_2025=(STD["f25"], "sum"))
      .merge(agent_total_2025_active, on=STD["agent"], how="left")
)

agent_category["% incidenza su fatturato agente"] = (
    agent_category["fatturato_categoria_2025"] / agent_category["fatturato_agente_2025"] * 100
).replace([np.inf, -np.inf], np.nan).fillna(0).round(2)

agent_category = agent_category.drop(columns=["fatturato_agente_2025"]) \
                               .sort_values([STD["agent"], "fatturato_categoria_2025"], ascending=[True, False])

# =============================
# Tabs + Export per singolo tab
# =============================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Fatturato per Città",
    "Fatturato Agente",
    "Fatturato per Zona",
    "Fatturato per Categoria",
    "Export Excel (tutto)"
])

# -------- TAB 1 --------
with tab1:
    st.subheader("Fatturato per Città (Sintesi) — SOLO clienti attivi 2025")
    st.dataframe(city_summary, use_container_width=True, height=420)

    st.subheader("Dettaglio: Agenti nella Città — SOLO clienti attivi 2025")
    st.dataframe(city_detail, use_container_width=True, height=520)

    city_xlsx = to_excel_bytes({
        "CITTA_SINTESI": city_summary,
        "CITTA_DETTAGLIO": city_detail
    })
    st.download_button(
        "⬇️ Scarica report Città (Excel)",
        data=city_xlsx,
        file_name="report_fatturato_citta.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------- TAB 2 --------
with tab2:
    st.subheader("Fatturato Agente — Totali 2024/2025 su TUTTO + % incidenza sul totale area (2025)")

    st.markdown(f"**Top 3 agenti = {top3_share:.2f}% del fatturato totale area (2025)**")
    st.dataframe(top3_table, use_container_width=True, height=180)

    styled = agent_view.style.applymap(highlight_incidenza, subset=["% incidenza su totale area"])
    st.dataframe(styled, use_container_width=True, height=520)

    st.subheader("Dettaglio clienti persi / acquisiti / mantenuti (per agente)")
    st.dataframe(agent_client_detail, use_container_width=True, height=520)

    agent_xlsx = to_excel_bytes({
        "AGENTI": agent_view,
        "TOP3": top3_table,
        "DETT_CLIENTI_FLOW": agent_client_detail
    })
    st.download_button(
        "⬇️ Scarica report Agenti (Excel)",
        data=agent_xlsx,
        file_name="report_fatturato_agenti.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------- TAB 3 --------
with tab3:
    st.subheader("Fatturato per Zona (Agente → Città) — SOLO clienti attivi 2025")
    st.dataframe(zone_agent_city, use_container_width=True, height=420)

    st.subheader("Dettaglio: Clienti (Agente → Città → Cliente) — SOLO clienti attivi 2025")
    st.dataframe(zone_client_detail, use_container_width=True, height=520)

    zona_xlsx = to_excel_bytes({
        "ZONA_AGENTE_CITTA": zone_agent_city,
        "ZONA_CLIENTI_DETT": zone_client_detail,
        "ZONA_SINTESI_AGENTE": zone_agent_total
    })
    st.download_button(
        "⬇️ Scarica report Zona (Excel)",
        data=zona_xlsx,
        file_name="report_fatturato_zona.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------- TAB 4 --------
with tab4:
    st.subheader("Fatturato per Categoria (Agente → Categoria) — SOLO clienti attivi 2025")
    st.dataframe(agent_category, use_container_width=True, height=650)

    cat_xlsx = to_excel_bytes({"CATEGORIA": agent_category})
    st.download_button(
        "⬇️ Scarica report Categoria (Excel)",
        data=cat_xlsx,
        file_name="report_fatturato_categoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------- TAB 5 (export totale) --------
with tab5:
    st.subheader("Esporta tutte le tabelle in Excel (multi-foglio)")

    sheets_all = {
        "FATT_CITTA_SINTESI": city_summary,
        "FATT_CITTA_DETTAGLIO": city_detail,

        "FATT_AGENTE": agent_view,
        "TOP3_AGENTI": top3_table,
        "DETT_CLIENTI_FLOW": agent_client_detail,

        "ZONA_AGENTE_TOTALE": zone_agent_total,
        "ZONA_AGENTE_CITTA": zone_agent_city,
        "ZONA_CLIENTI_DETT": zone_client_detail,

        "FATT_CATEGORIA": agent_category,
    }

    all_xlsx = to_excel_bytes(sheets_all)
    st.download_button(
        label="⬇️ Scarica Excel completo (tutti i fogli)",
        data=all_xlsx,
        file_name="report_vendite_analisi_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
