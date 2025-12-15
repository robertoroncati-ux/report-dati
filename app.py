import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Report Vendite (Analisi Città / Agente)", layout="wide")

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

    # FIX righe fantasma: elimino righe senza chiavi minime (Agente/Città/Cliente)
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


def guess_agent(agents: list[str], contains: str) -> str | None:
    contains = contains.lower()
    for a in agents:
        if contains in a.lower():
            return a
    return None


# =============================
# UI
# =============================
st.title("📊 Report Vendite (Analisi Città / Agente)")
st.caption("VERSIONE APP: v6 - report + export + simulatore logistico semi-assistito")

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
# Base aggregata a livello Agente-Città-Cliente
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
# 1) FATTURATO PER CITTA' - SOLO ATTIVI 2025
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
# 4) FATTURATO PER CATEGORIA - SOLO ATTIVI 2025 + % incidenza su fatturato agente
# =============================
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

# =============================
# TAB UI
# =============================
tab_citta, tab_agente, tab_zona, tab_categoria, tab_sim, tab_export = st.tabs([
    "Fatturato per Città",
    "Fatturato Agente",
    "Fatturato per Zona",
    "Fatturato per Categoria",
    "🧠 Simulatore logistico",
    "Export Excel (tutto)"
])

# -------- TAB CITTÀ --------
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

# -------- TAB AGENTE --------
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

# -------- TAB ZONA --------
with tab_zona:
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

# -------- TAB CATEGORIA --------
with tab_categoria:
    st.subheader("Fatturato per Categoria (Agente → Categoria) — SOLO clienti attivi 2025")
    st.dataframe(agent_category, use_container_width=True, height=650)

    cat_xlsx = to_excel_bytes({"CATEGORIA": agent_category})
    st.download_button(
        "⬇️ Scarica report Categoria (Excel)",
        data=cat_xlsx,
        file_name="report_fatturato_categoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------- TAB SIMULATORE LOGISTICO --------
with tab_sim:
    st.subheader("🧠 Simulatore logistico semi-assistito (scenario)")

    # BASE simulatore: livello cliente (attivi 2025)
    base_sim = ac_client_active[[STD["agent"], STD["city"], STD["client"], STD["f25"]]].copy()
    base_sim = base_sim.rename(columns={STD["f25"]: "fatturato_2025"})

    agents_list = sorted(base_sim[STD["agent"]].unique().tolist())
    cities_list = sorted(base_sim[STD["city"]].unique().tolist())

    # default agent guesses
    default_ron = guess_agent(agents_list, "roncati")
    default_bal = guess_agent(agents_list, "balducci")

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1.2])
    with c1:
        ron_out = st.checkbox("Roncati non operativo", value=True)
    with c2:
        bal_out = st.checkbox("Balducci non presente", value=True)
    with c3:
        include_new = st.checkbox("Inserimento Nuovo Agente", value=True)
    with c4:
        new_agent_name = st.text_input("Nome nuovo agente (simulazione)", value="Nuovo Agente")

    # selezione agenti reali (non obbligatoria, ma utile se i nomi non matchano)
    st.caption("Se i nomi non coincidono (es. 'A400 - RONCATI...'), selezionali qui:")
    s1, s2 = st.columns(2)
    with s1:
        ron_agent = st.selectbox("Seleziona 'Roncati' (per esclusione)", options=["(nessuno)"] + agents_list,
                                 index=(["(nessuno)"] + agents_list).index(default_ron) if default_ron in agents_list else 0)
    with s2:
        bal_agent = st.selectbox("Seleziona 'Balducci' (per esclusione)", options=["(nessuno)"] + agents_list,
                                 index=(["(nessuno)"] + agents_list).index(default_bal) if default_bal in agents_list else 0)

    # Applica esclusioni
    sim_work = base_sim.copy()
    if ron_out and ron_agent != "(nessuno)":
        sim_work = sim_work[sim_work[STD["agent"]] != ron_agent]
    if bal_out and bal_agent != "(nessuno)":
        sim_work = sim_work[sim_work[STD["agent"]] != bal_agent]

    # Stato PRIMA (dopo eventuali esclusioni strutturali)
    stato_prima = (
        sim_work.groupby(STD["agent"], as_index=False)
          .agg(
              fatturato_prima=("fatturato_2025", "sum"),
              clienti_prima=(STD["client"], "nunique"),
              citta_prima=(STD["city"], "nunique"),
          )
          .sort_values("fatturato_prima", ascending=False)
    )

    st.markdown("### Stato PRIMA (dopo esclusioni strutturali)")
    st.dataframe(stato_prima, use_container_width=True, height=300)

    # ===== Semi-assist: città critiche =====
    st.markdown("### Suggerimenti (semi-assistito)")
    city_crit = (
        sim_work.groupby(STD["city"], as_index=False)
          .agg(
              fatturato=("fatturato_2025", "sum"),
              clienti=(STD["client"], "nunique"),
              agenti=(STD["agent"], "nunique"),
          )
    )
    # euristica: pochi agenti (1) e fatturato medio per cliente basso => potenziale
    city_crit["fatt_medio_cliente"] = (city_crit["fatturato"] / city_crit["clienti"]).replace([np.inf, -np.inf], 0).fillna(0)
    city_crit = city_crit.sort_values(["agenti", "fatturato"], ascending=[True, False])

    st.caption("Città con 1 solo agente (possibili candidate per presidio mirato / nuovo agente):")
    st.dataframe(city_crit[city_crit["agenti"] == 1].head(20), use_container_width=True, height=260)

    # ===== Movimenti in session state =====
    if "moves" not in st.session_state:
        st.session_state["moves"] = pd.DataFrame(columns=[
            "Cliente", "Città", "Da agente", "A agente", "Fatturato", "Motivo"
        ])

    st.markdown("### Aggiunta movimenti (guidata)")

    g1, g2, g3 = st.columns([1.2, 1.2, 1.2])
    with g1:
        sel_city = st.selectbox("Città", options=cities_list)
    sub_city = sim_work[sim_work[STD["city"]] == sel_city].copy()

    with g2:
        origin_agent = st.selectbox("Da agente", options=sorted(sub_city[STD["agent"]].unique().tolist()))
    sub_origin = sub_city[sub_city[STD["agent"]] == origin_agent].copy()

    with g3:
        dest_agents = sorted(sim_work[STD["agent"]].unique().tolist())
        if include_new:
            if new_agent_name not in dest_agents:
                dest_agents = dest_agents + [new_agent_name]
        dest_agent = st.selectbox("A agente", options=dest_agents)

    clients_options = sub_origin.sort_values("fatturato_2025", ascending=False)[STD["client"]].tolist()
    sel_clients = st.multiselect("Clienti da spostare (dalla città e agente selezionati)", options=clients_options)

    motivo = st.selectbox("Motivo", options=["Logistica", "Riequilibrio", "Nuovo agente"])

    b1, b2, b3 = st.columns([1, 1, 2])
    with b1:
        if st.button("➕ Aggiungi movimenti selezionati"):
            if sel_clients:
                add_rows = sub_origin[sub_origin[STD["client"]].isin(sel_clients)][[STD["client"], STD["city"], STD["agent"], "fatturato_2025"]].copy()
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
        if st.button("🧹 Svuota movimenti"):
            st.session_state["moves"] = st.session_state["moves"].iloc[0:0].copy()

    with b3:
        st.caption("Puoi anche modificare manualmente la tabella movimenti qui sotto (semi-assistito + edit libero).")

    st.markdown("### Movimenti (editabili)")
    moves_df = st.data_editor(
        st.session_state["moves"],
        num_rows="dynamic",
        use_container_width=True,
        key="moves_editor"
    )
    # sincronizza eventuali modifiche
    st.session_state["moves"] = moves_df.copy()

    # ===== Applica movimenti =====
    sim_after = sim_work.copy()

    if len(moves_df) > 0:
        # Validazioni base
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

            # quanto esiste davvero da spostare?
            available = sim_after.loc[mask, "fatturato_2025"].sum()
            move_amt = min(fatt, available)

            if move_amt <= 0:
                continue

            # sottraggo
            sim_after.loc[mask, "fatturato_2025"] = sim_after.loc[mask, "fatturato_2025"] - move_amt

            # aggiungo (nuova riga)
            sim_after = pd.concat([sim_after, pd.DataFrame([{
                STD["agent"]: a,
                STD["city"]: cty,
                STD["client"]: cli,
                "fatturato_2025": move_amt
            }])], ignore_index=True)

        # pulizia: elimino righe diventate 0 o negative
        sim_after = sim_after[sim_after["fatturato_2025"] > 0].copy()

    # ===== Stato DOPO =====
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

    st.markdown("### Stato DOPO (dopo movimenti)")
    st.dataframe(confronto, use_container_width=True, height=380)

    tot_prima = float(stato_prima["fatturato_prima"].sum())
    tot_dopo = float(stato_dopo["fatturato_dopo"].sum())

    if round(tot_prima, 2) != round(tot_dopo, 2):
        st.error(f"⚠️ Fatturato area NON invariato: {tot_prima:,.2f} → {tot_dopo:,.2f}".replace(",", "."))
    else:
        st.success(f"✅ Fatturato area invariato: {tot_dopo:,.2f}".replace(",", "."))

    # ===== Pacchetto nuovo agente =====
    if include_new:
        pac = sim_after[sim_after[STD["agent"]] == new_agent_name].copy()
        st.markdown("### Pacchetto nuovo agente")
        if pac.empty:
            st.info("Ancora nessun cliente assegnato al nuovo agente.")
        else:
            pac_view = pac.groupby([STD["city"]], as_index=False).agg(
                fatturato=("fatturato_2025", "sum"),
                clienti=(STD["client"], "nunique")
            ).sort_values("fatturato", ascending=False)
            st.dataframe(pac_view, use_container_width=True, height=260)

            st.info(
                f"Nuovo agente → Fatturato: {pac['fatturato_2025'].sum():,.2f} | "
                f"Clienti: {pac[STD['client']].nunique()} | "
                f"Città: {pac[STD['city']].nunique()}".replace(",", ".")
            )

    # Export scenario simulatore
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

# -------- TAB EXPORT TUTTO --------
with tab_export:
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
