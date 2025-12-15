import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =====================================================
# CONFIG
# =====================================================
st.set_page_config(page_title="Report Vendite (Analisi Città / Agente)", layout="wide")

# =====================================================
# COLONNE STANDARD + RINOMINE (robusto)
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
def clean_text(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
        .replace({"nan": ""})
    )

@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns=RENAME_MAP)

    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"Mancano queste colonne nel file Excel: {missing}")

    # pulizia stringhe
    for c in [STD["agent"], STD["city"], STD["client"], STD["category"], STD["article"]]:
        df[c] = clean_text(df[c])

    # numerici
    df[STD["f24"]] = pd.to_numeric(df[STD["f24"]], errors="coerce").fillna(0.0)
    df[STD["f25"]] = pd.to_numeric(df[STD["f25"]], errors="coerce").fillna(0.0)

    # fix righe fantasma: servono almeno Agente/Città/Cliente
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

def guess_agent(agents: list[str], contains: str) -> str | None:
    contains = contains.lower()
    for a in agents:
        if contains in a.lower():
            return a
    return None

def highlight_incidenza(val):
    try:
        return "background-color: #ffd1d1" if float(val) >= 30 else ""
    except Exception:
        return ""

def add_total_row(df_in: pd.DataFrame, labels: dict, sum_cols: list, first=True) -> pd.DataFrame:
    df = df_in.copy()
    total = {c: np.nan for c in df.columns}
    for k, v in labels.items():
        if k in total:
            total[k] = v
    for c in sum_cols:
        if c in total:
            total[c] = df[c].sum()
    total_df = pd.DataFrame([total])
    return pd.concat([total_df, df], ignore_index=True) if first else pd.concat([df, total_df], ignore_index=True)

# =====================================================
# UI
# =====================================================
st.title("📊 Report Vendite (Analisi Città / Agente)")
st.caption("VERSIONE APP: v7 — statistiche complete + export per tab + simulatore logistico + auto-proposta per città")

uploaded = st.file_uploader("Carica il file Excel", type=["xlsx", "xls"])
if not uploaded:
    st.stop()

try:
    df = load_excel(uploaded)
except Exception as e:
    st.error(str(e))
    st.stop()

st.caption(f"Righe totali file (dopo pulizia chiavi): **{len(df):,}**".replace(",", "."))

# =====================================================
# BASE: aggregazione a livello Agente-Città-Cliente
# =====================================================
ac_client_all = (
    df.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)[[STD["f24"], STD["f25"]]]
      .sum()
)

ac_client_active = ac_client_all[ac_client_all[STD["f25"]] > 0].copy()

st.caption(
    f"Clienti aggregati (tutto): **{len(ac_client_all):,}** | "
    f"Clienti attivi 2025 (fatt 2025 > 0): **{len(ac_client_active):,}**"
    .replace(",", ".")
)

if ac_client_active.empty:
    st.warning("Dopo il filtro (Fatturato 2025 > 0) non rimane nessuna riga. Controlla il file.")
    st.stop()

# =====================================================
# REPORT 1 — FATTURATO PER CITTA' (solo attivi 2025)
# =====================================================
city_summary = (
    ac_client_active.groupby(STD["city"], as_index=False)
      .agg(
          fatturato_citta_2025=(STD["f25"], "sum"),
          clienti_attivi_2025=(STD["client"], "nunique"),
          agenti_presenti=(STD["agent"], pd.Series.nunique),
      )
      .sort_values("fatturato_citta_2025", ascending=False)
)

city_detail = (
    ac_client_active.groupby([STD["city"], STD["agent"]], as_index=False)
      .agg(
          fatturato_agente_nella_citta_2025=(STD["f25"], "sum"),
          clienti_attivi_agente_nella_citta=(STD["client"], "nunique"),
      )
)

city_tot_map = city_summary.set_index(STD["city"])["fatturato_citta_2025"]
city_detail["% incidenza su città"] = (
    city_detail.apply(
        lambda r: (r["fatturato_agente_nella_citta_2025"] / city_tot_map.get(r[STD["city"]], 1)) * 100,
        axis=1
    )
).replace([np.inf, -np.inf], np.nan).fillna(0).round(2)

city_detail = city_detail.sort_values([STD["city"], "fatturato_agente_nella_citta_2025"], ascending=[True, False])

# =====================================================
# REPORT 2 — FATTURATO AGENTE (totali 2024/2025 su TUTTO + incidenza + clienti attivi + persi/acquisiti)
# =====================================================
agent_totals = (
    ac_client_all.groupby(STD["agent"], as_index=False)
      .agg(
          fatt_2024=(STD["f24"], "sum"),
          fatt_2025=(STD["f25"], "sum"),
      )
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

agent_flow = (
    tmp.groupby(STD["agent"], as_index=False)
      .agg(
          clienti_persi=("stato_cliente", lambda s: (s == "Perso").sum()),
          clienti_acquisiti=("stato_cliente", lambda s: (s == "Acquisito").sum()),
          clienti_mantenuti=("stato_cliente", lambda s: (s == "Mantenuto").sum()),
      )
)

agent_report = agent_totals.merge(agent_active, on=STD["agent"], how="left") \
                          .merge(agent_flow, on=STD["agent"], how="left")

agent_report["clienti_attivi_2025"] = agent_report["clienti_attivi_2025"].fillna(0).astype(int)
for c in ["clienti_persi", "clienti_acquisiti", "clienti_mantenuti"]:
    agent_report[c] = agent_report[c].fillna(0).astype(int)

tot_area_2025 = agent_report["fatt_2025"].sum()
agent_report["% incidenza su totale area"] = (
    agent_report["fatt_2025"] / tot_area_2025 * 100
).replace([np.inf, -np.inf], np.nan).fillna(0).round(2)

agent_report = agent_report.sort_values("fatt_2025", ascending=False)

agent_view = agent_report.rename(columns={
    STD["agent"]: "Agente",
    "fatt_2024": "Fatturato totale 2024",
    "fatt_2025": "Fatturato totale 2025",
}).copy()

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

top3 = agent_report.head(3).copy()
top3_share = (top3["fatt_2025"].sum() / tot_area_2025 * 100) if tot_area_2025 else 0
top3_table = top3.rename(columns={STD["agent"]: "Agente"})[
    ["Agente", "fatt_2025", "% incidenza su totale area"]
].rename(columns={"fatt_2025": "Fatturato totale 2025"})

agent_client_detail = tmp[tmp["stato_cliente"].isin(["Perso", "Acquisito", "Mantenuto"])][
    [STD["agent"], STD["city"], STD["client"], STD["f24"], STD["f25"], "stato_cliente"]
].sort_values([STD["agent"], "stato_cliente", STD["f25"]], ascending=[True, True, False])

# =====================================================
# REPORT 3 — FATTURATO PER ZONA (Agente -> Città) + dettaglio cliente (solo attivi 2025)
# + Totali con etichette (niente riga 0 "misteriosa")
# =====================================================
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

zone_agent_city_tot = add_total_row(
    zone_agent_city,
    labels={STD["agent"]: "TOTALE AREA", STD["city"]: "TOTALE"},
    sum_cols=["fatturato_2025"],
    first=True
)

zone_client_detail_tot = add_total_row(
    zone_client_detail,
    labels={STD["agent"]: "TOTALE AREA", STD["city"]: "TOTALE", STD["client"]: "TOTALE"},
    sum_cols=["fatturato_cliente_2025"],
    first=True
)

tot_area_zone = zone_agent_city["fatturato_2025"].sum()

# =====================================================
# REPORT 4 — FATTURATO PER CATEGORIA (solo attivi 2025) + % incidenza su fatturato agente
# =====================================================
df_active_rows = df[df[STD["f25"]] > 0].copy()

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

# =====================================================
# TABS
# =====================================================
tab_citta, tab_agente, tab_zona, tab_categoria, tab_sim, tab_export = st.tabs([
    "Fatturato per Città",
    "Fatturato Agente",
    "Fatturato per Zona",
    "Fatturato per Categoria",
    "🧠 Simulatore logistico",
    "Export Excel (tutto)"
])

# =====================================================
# TAB: CITTA'
# =====================================================
with tab_citta:
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

# =====================================================
# TAB: AGENTE
# =====================================================
with tab_agente:
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

# =====================================================
# TAB: ZONA
# =====================================================
with tab_zona:
    st.subheader("Fatturato per Zona (Agente → Città) — SOLO clienti attivi 2025")
    st.caption(f"Totale fatturato area (attivi 2025): {tot_area_zone:,.2f}".replace(",", "."))
    st.dataframe(zone_agent_city_tot, use_container_width=True, height=450)

    st.subheader("Dettaglio: Clienti (Agente → Città → Cliente) — SOLO clienti attivi 2025")
    st.dataframe(zone_client_detail_tot, use_container_width=True, height=520)

    zona_xlsx = to_excel_bytes({
        "ZONA_AGENTE_CITTA": zone_agent_city_tot,
        "ZONA_CLIENTI_DETT": zone_client_detail_tot
    })
    st.download_button(
        "⬇️ Scarica report Zona (Excel)",
        data=zona_xlsx,
        file_name="report_fatturato_zona.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# TAB: CATEGORIA
# =====================================================
with tab_categoria:
    st.subheader("Fatturato per Categoria (Agente → Categoria) — SOLO attivi 2025 + % incidenza su fatturato agente")
    st.dataframe(agent_category, use_container_width=True, height=650)

    cat_xlsx = to_excel_bytes({"CATEGORIA": agent_category})
    st.download_button(
        "⬇️ Scarica report Categoria (Excel)",
        data=cat_xlsx,
        file_name="report_fatturato_categoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# TAB: SIMULATORE + AUTO-PROPOSTA (per città)
# =====================================================
with tab_sim:
    st.header("🧠 Simulatore logistico semi-assistito + ⚡ Auto-proposta per città")

    # base simulatore: livello cliente (attivi 2025)
    base_sim = (
        df.groupby([STD["agent"], STD["city"], STD["client"]], as_index=False)
          .agg(fatturato_2025=(STD["f25"], "sum"))
    )
    base_sim = base_sim[base_sim["fatturato_2025"] > 0].copy()

    agents_list = sorted(base_sim[STD["agent"]].unique().tolist())
    cities_list = sorted(base_sim[STD["city"]].unique().tolist())

    default_ron = guess_agent(agents_list, "roncati")
    default_bal = guess_agent(agents_list, "balducci")
    default_bit = guess_agent(agents_list, "bitto")
    default_fre = guess_agent(agents_list, "freccero")

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1.2])
    with c1:
        ron_out = st.checkbox("Roncati non operativo", value=True)
    with c2:
        bal_out = st.checkbox("Balducci non presente", value=True)
    with c3:
        include_new = st.checkbox("Inserimento nuovo agente", value=True)
    with c4:
        new_agent_name = st.text_input("Nome nuovo agente (simulazione)", value="Nuovo Agente")

    st.caption("Seleziona i nomi esatti (così l’auto-proposta non sbaglia agente).")
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        ron_agent = st.selectbox(
            "Roncati",
            options=["(nessuno)"] + agents_list,
            index=(["(nessuno)"] + agents_list).index(default_ron) if default_ron in agents_list else 0
        )
    with s2:
        bitto_agent = st.selectbox(
            "Bitto (target)",
            options=["(nessuno)"] + agents_list,
            index=(["(nessuno)"] + agents_list).index(default_bit) if default_bit in agents_list else 0
        )
    with s3:
        bal_agent = st.selectbox(
            "Balducci (uscita)",
            options=["(nessuno)"] + agents_list,
            index=(["(nessuno)"] + agents_list).index(default_bal) if default_bal in agents_list else 0
        )
    with s4:
        fre_agent = st.selectbox(
            "Freccero (preferenza logistica)",
            options=["(nessuno)"] + agents_list,
            index=(["(nessuno)"] + agents_list).index(default_fre) if default_fre in agents_list else 0
        )

    # scenario base (post esclusioni strutturali)
    sim_work = base_sim.copy()
    if ron_out and ron_agent != "(nessuno)":
        sim_work = sim_work[sim_work[STD["agent"]] != ron_agent]
    if bal_out and bal_agent != "(nessuno)":
        sim_work = sim_work[sim_work[STD["agent"]] != bal_agent]

    # Stato PRIMA
    stato_prima = (
        sim_work.groupby(STD["agent"], as_index=False)
          .agg(
              fatturato_prima=("fatturato_2025", "sum"),
              clienti_prima=(STD["client"], "nunique"),
              citta_prima=(STD["city"], "nunique"),
          )
          .sort_values("fatturato_prima", ascending=False)
    )

    st.subheader("Stato PRIMA (dopo esclusioni strutturali)")
    st.dataframe(stato_prima, use_container_width=True, height=320)

    # suggerimenti città monoreferente
    st.subheader("Suggerimenti — città con 1 solo agente (attivi 2025)")
    city_crit = (
        sim_work.groupby(STD["city"], as_index=False)
          .agg(
              fatturato=("fatturato_2025", "sum"),
              clienti=(STD["client"], "nunique"),
              agenti=(STD["agent"], "nunique"),
          )
    )
    city_crit["fatt_medio_cliente"] = (city_crit["fatturato"] / city_crit["clienti"]).replace([np.inf, -np.inf], 0).fillna(0)
    st.dataframe(
        city_crit[city_crit["agenti"] == 1].sort_values("fatturato", ascending=False).head(30),
        use_container_width=True,
        height=260
    )

    # movimenti
    if "moves" not in st.session_state:
        st.session_state["moves"] = pd.DataFrame(columns=[
            "Cliente", "Città", "Da agente", "A agente", "Fatturato", "Motivo"
        ])

    def add_move_rows(rows_df: pd.DataFrame, da: str, a: str, motivo: str):
        if rows_df.empty:
            return
        add_rows = rows_df[[STD["client"], STD["city"], "fatturato_2025"]].copy()
        add_rows = add_rows.rename(columns={
            STD["client"]: "Cliente",
            STD["city"]: "Città",
            "fatturato_2025": "Fatturato"
        })
        add_rows["Da agente"] = da
        add_rows["A agente"] = a
        add_rows["Motivo"] = motivo
        st.session_state["moves"] = pd.concat([st.session_state["moves"], add_rows[
            ["Cliente", "Città", "Da agente", "A agente", "Fatturato", "Motivo"]
        ]], ignore_index=True)

    def best_receiver_in_city(df_city: pd.DataFrame, prefer_agent: str | None = None) -> str | None:
        if df_city.empty:
            return None
        if prefer_agent and prefer_agent in df_city[STD["agent"]].unique():
            return prefer_agent
        g = df_city.groupby(STD["agent"], as_index=False)["fatturato_2025"].sum().sort_values("fatturato_2025", ascending=False)
        return g.iloc[0][STD["agent"]] if len(g) else None

    st.markdown("---")
    st.subheader("⚡ Auto-proposta (per città)")

    ap1, ap2, ap3 = st.columns([1.3, 1.2, 1.5])
    with ap1:
        ap_city_ron = st.selectbox("Città da cedere (Roncati → Bitto)", options=cities_list, index=0)
    with ap2:
        ap_package_target = st.number_input("Pacchetto minimo nuovo agente (€)", min_value=0.0, value=150000.0, step=10000.0)
    with ap3:
        ap_focus_cities = st.multiselect(
            "Città su cui costruire il pacchetto nuovo agente (se vuoto → auto)",
            options=cities_list
        )

    b_ap1, b_ap2 = st.columns([1, 1])
    with b_ap1:
        if st.button("⚡ Genera Auto-proposta"):
            # riparto pulito
            st.session_state["moves"] = st.session_state["moves"].iloc[0:0].copy()

            # 1) Roncati → Bitto nella città scelta
            if ron_agent != "(nessuno)" and bitto_agent != "(nessuno)":
                df_ron_city = base_sim[
                    (base_sim[STD["agent"]] == ron_agent) &
                    (base_sim[STD["city"]] == ap_city_ron)
                ].copy()
                add_move_rows(df_ron_city, ron_agent, bitto_agent, "Auto: Roncati→Bitto per città")

            # 2) Uscita Balducci: redistribuzione PER CITTÀ (preferenza a Freccero se già presente in città)
            if bal_agent != "(nessuno)":
                bal_rows = base_sim[base_sim[STD["agent"]] == bal_agent].copy()
                for city, part in bal_rows.groupby(STD["city"]):
                    city_pool = base_sim[(base_sim[STD["city"]] == city) & (base_sim[STD["agent"]] != bal_agent)].copy()
                    receiver = best_receiver_in_city(city_pool, prefer_agent=(fre_agent if fre_agent != "(nessuno)" else None))
                    if receiver:
                        add_move_rows(part, bal_agent, receiver, "Auto: uscita Balducci (per città)")

            # 3) Pacchetto nuovo agente: per città
            if include_new and new_agent_name:
                city_stats = (
                    base_sim.groupby(STD["city"], as_index=False)
                      .agg(fatturato=("fatturato_2025", "sum"), agenti=(STD["agent"], "nunique"))
                )
                if len(ap_focus_cities) == 0:
                    candidate_cities = city_stats[city_stats["agenti"] >= 2].sort_values("fatturato", ascending=False)[STD["city"]].head(6).tolist()
                else:
                    candidate_cities = ap_focus_cities

                pool = base_sim[base_sim[STD["city"]].isin(candidate_cities)].copy()
                pool = pool.sort_values("fatturato_2025", ascending=False)

                tot_pack = 0.0
                used = set()
                for _, r in pool.iterrows():
                    if tot_pack >= ap_package_target:
                        break
                    key = (r[STD["agent"]], r[STD["city"]], r[STD["client"]])
                    if key in used:
                        continue
                    used.add(key)
                    add_move_rows(pd.DataFrame([r]), r[STD["agent"]], new_agent_name, "Auto: pacchetto nuovo agente (per città)")
                    tot_pack += float(r["fatturato_2025"])

    with b_ap2:
        if st.button("🧹 Svuota movimenti (auto+manuale)"):
            st.session_state["moves"] = st.session_state["moves"].iloc[0:0].copy()

    # manuale
    st.markdown("---")
    st.subheader("Aggiunta movimenti (manuale)")

    g1, g2, g3 = st.columns([1.2, 1.2, 1.2])
    with g1:
        sel_city = st.selectbox("Città (manuale)", options=cities_list, key="man_city")
    sub_city = sim_work[sim_work[STD["city"]] == sel_city].copy()

    with g2:
        origin_agent = st.selectbox("Da agente (manuale)", options=sorted(sub_city[STD["agent"]].unique().tolist()), key="man_da")
    sub_origin = sub_city[sub_city[STD["agent"]] == origin_agent].copy()

    with g3:
        dest_agents = sorted(sim_work[STD["agent"]].unique().tolist())
        if include_new and new_agent_name and new_agent_name not in dest_agents:
            dest_agents = dest_agents + [new_agent_name]
        dest_agent = st.selectbox("A agente (manuale)", options=dest_agents, key="man_a")

    clients_options = sub_origin.sort_values("fatturato_2025", ascending=False)[STD["client"]].tolist()
    sel_clients = st.multiselect("Clienti da spostare (manuale)", options=clients_options, key="man_clients")
    motivo = st.selectbox("Motivo (manuale)", options=["Logistica", "Riequilibrio", "Nuovo agente"], key="man_motivo")

    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("➕ Aggiungi movimenti selezionati (manuale)"):
            if sel_clients:
                add_rows = sub_origin[sub_origin[STD["client"]].isin(sel_clients)][
                    [STD["client"], STD["city"], STD["agent"], "fatturato_2025"]
                ].copy()
                add_rows = add_rows.rename(columns={
                    STD["client"]: "Cliente",
                    STD["city"]: "Città",
                    STD["agent"]: "Da agente",
                    "fatturato_2025": "Fatturato"
                })
                add_rows["A agente"] = dest_agent
                add_rows["Motivo"] = motivo
                st.session_state["moves"] = pd.concat([st.session_state["moves"], add_rows[
                    ["Cliente", "Città", "Da agente", "A agente", "Fatturato", "Motivo"]
                ]], ignore_index=True)
    with b2:
        if st.button("🧹 Svuota movimenti (manuale)"):
            st.session_state["moves"] = st.session_state["moves"].iloc[0:0].copy()

    st.subheader("Movimenti (editabili)")
    moves_df = st.data_editor(
        st.session_state["moves"],
        num_rows="dynamic",
        use_container_width=True,
        key="moves_editor"
    )
    st.session_state["moves"] = moves_df.copy()

    # applica movimenti
    sim_after = sim_work.copy()

    moves_ok = moves_df.dropna(subset=["Cliente", "Città", "Da agente", "A agente", "Fatturato"]).copy()
    moves_ok["Fatturato"] = pd.to_numeric(moves_ok["Fatturato"], errors="coerce").fillna(0.0)
    moves_ok = moves_ok[moves_ok["Fatturato"] > 0].copy()

    for _, r in moves_ok.iterrows():
        cli = str(r["Cliente"]).strip()
        cty = str(r["Città"]).strip()
        da = str(r["Da agente"]).strip()
        a = str(r["A agente"]).strip()
        fatt = float(r["Fatturato"])

        mask = (
            (sim_after[STD["client"]] == cli) &
            (sim_after[STD["city"]] == cty) &
            (sim_after[STD["agent"]] == da)
        )

        available = sim_after.loc[mask, "fatturato_2025"].sum()
        move_amt = min(fatt, available)
        if move_amt <= 0:
            continue

        sim_after.loc[mask, "fatturato_2025"] = sim_after.loc[mask, "fatturato_2025"] - move_amt

        sim_after = pd.concat([sim_after, pd.DataFrame([{
            STD["agent"]: a,
            STD["city"]: cty,
            STD["client"]: cli,
            "fatturato_2025": move_amt
        }])], ignore_index=True)

    sim_after = sim_after[sim_after["fatturato_2025"] > 0].copy()

    stato_dopo = (
        sim_after.groupby(STD["agent"], as_index=False)
          .agg(
              fatturato_dopo=("fatturato_2025", "sum"),
              clienti_dopo=(STD["client"], "nunique"),
              citta_dopo=(STD["city"], "nunique"),
          )
          .sort_values("fatturato_dopo", ascending=False)
    )

    confronto = stato_prima.merge(stato_dopo, on=STD["agent"], how="outer").fillna(0)
    confronto["Delta fatturato"] = confronto["fatturato_dopo"] - confronto["fatturato_prima"]
    confronto = confronto.sort_values("fatturato_dopo", ascending=False)

    st.subheader("Stato DOPO")
    st.dataframe(confronto, use_container_width=True, height=380)

    tot_prima = float(stato_prima["fatturato_prima"].sum())
    tot_dopo = float(stato_dopo["fatturato_dopo"].sum())

    if round(tot_prima, 2) != round(tot_dopo, 2):
        st.error(f"⚠️ Fatturato area NON invariato: {tot_prima:,.2f} → {tot_dopo:,.2f}".replace(",", "."))
    else:
        st.success(f"✅ Fatturato area invariato: {tot_dopo:,.2f}".replace(",", "."))

    if include_new and new_agent_name:
        pac = sim_after[sim_after[STD["agent"]] == new_agent_name].copy()
        st.subheader("Pacchetto nuovo agente")
        if pac.empty:
            st.info("Ancora nessun cliente assegnato al nuovo agente.")
        else:
            pac_view = pac.groupby(STD["city"], as_index=False).agg(
                fatturato=("fatturato_2025", "sum"),
                clienti=(STD["client"], "nunique")
            ).sort_values("fatturato", ascending=False)
            st.dataframe(pac_view, use_container_width=True, height=260)

    sim_xlsx = to_excel_bytes({
        "BASE_ATTIVI_2025": sim_work.rename(columns={"fatturato_2025": "fatturato_2025_attivi"}),
        "MOVIMENTI": moves_df,
        "STATO_PRIMA": stato_prima,
        "STATO_DOPO": stato_dopo,
        "DETTAGLIO_DOPO": sim_after
    })
    st.download_button(
        "⬇️ Scarica Scenario Simulazione (Excel)",
        data=sim_xlsx,
        file_name="scenario_simulazione_logistica.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =====================================================
# TAB: EXPORT TUTTO
# =====================================================
with tab_export:
    st.subheader("Esporta tutte le tabelle in Excel (multi-foglio)")

    sheets_all = {
        "CITTA_SINTESI": city_summary,
        "CITTA_DETTAGLIO": city_detail,

        "AGENTI": agent_view,
        "TOP3_AGENTI": top3_table,
        "DETT_CLIENTI_FLOW": agent_client_detail,

        "ZONA_AGENTE_CITTA": zone_agent_city_tot,
        "ZONA_CLIENTI_DETT": zone_client_detail_tot,

        "CATEGORIA": agent_category,
    }

    all_xlsx = to_excel_bytes(sheets_all)
    st.download_button(
        label="⬇️ Scarica Excel completo (tutti i fogli)",
        data=all_xlsx,
        file_name="report_vendite_analisi_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
