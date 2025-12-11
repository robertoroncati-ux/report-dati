import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# ------------------ CONFIGURAZIONE BASE ------------------

st.set_page_config(page_title="Report Fatturato Agente/Città", layout="wide")
st.title("📊 Report Fatturato Agente / Città")

uploaded_file = st.file_uploader("Carica il file Excel clienti", type=["xlsx", "xls"])


# ------------------ FUNZIONI DI SUPPORTO ------------------

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> BytesIO:
    """Trasforma un DataFrame in un file Excel in memoria."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer


def full_report_excel(city_summary, city_agent, agent_city, agent_totals) -> BytesIO:
    """Crea un unico file Excel con tutti i report in fogli separati."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        city_summary.to_excel(writer, index=False, sheet_name="Fatturato_città")
        city_agent.to_excel(writer, index=False, sheet_name="Città_agente")
        agent_city.to_excel(writer, index=False, sheet_name="Agente_città_%")
        agent_totals.to_excel(writer, index=False, sheet_name="Totale_agente")
    buffer.seek(0)
    return buffer


# ------------------ ELABORAZIONE PRINCIPALE ------------------

if uploaded_file is not None:
    # Legge il file Excel
    df = pd.read_excel(uploaded_file)

    # Uniformiamo i nomi delle colonne (per essere più robusti)
    if "Citta" in df.columns and "Città" not in df.columns:
        df = df.rename(columns={"Citta": "Città"})

    if "Esercizio" in df.columns and "Cliente" not in df.columns:
        df = df.rename(columns={"Esercizio": "Cliente"})

    if "acquistato al 10/12/2025" in df.columns and "Fatturato_2025" not in df.columns:
        df = df.rename(columns={"acquistato al 10/12/2025": "Fatturato_2025"})

    # Colonne di lavoro
    city_col = "Città"
    agent_col = "Agente"
    fatt_col = "Fatturato_2025"
    client_col = "Cliente" if "Cliente" in df.columns else None

    # Anteprima
    st.subheader("Anteprima dati (prime 20 righe)")
    st.dataframe(df.head(20))

    st.markdown("---")

    # ------------------ CALCOLI DI BASE UNA SOLA VOLTA ------------------

    # Totale per agente
    agg_dict_agent = {fatt_col: "sum"}
    if city_col in df.columns:
        agg_dict_agent["Numero_città"] = (city_col, "nunique")
    if client_col:
        agg_dict_agent["Numero_clienti"] = (client_col, "nunique")

    # Pandas non ama i dict misti, quindi faccio due passaggi
    agent_totals = (
        df.groupby(agent_col)[fatt_col]
        .sum()
        .reset_index()
        .rename(columns={fatt_col: "Totale_Fatturato_2025"})
        .sort_values("Totale_Fatturato_2025", ascending=False)
    )
    if city_col in df.columns:
        agent_cities = (
            df.groupby(agent_col)[city_col].nunique().reset_index().rename(
                columns={city_col: "Numero_città"}
            )
        )
        agent_totals = agent_totals.merge(agent_cities, on=agent_col, how="left")

    if client_col:
        agent_clients = (
            df.groupby(agent_col)[client_col].nunique().reset_index().rename(
                columns={client_col: "Numero_clienti"}
            )
        )
        agent_totals = agent_totals.merge(agent_clients, on=agent_col, how="left")

    # Riassunto per città
    city_summary = (
        df.groupby(city_col)[fatt_col]
        .sum()
        .reset_index()
        .rename(columns={fatt_col: "Totale_Fatturato_2025"})
        .sort_values("Totale_Fatturato_2025", ascending=False)
    )

    if client_col:
        n_clienti = (
            df.groupby(city_col)[client_col]
            .nunique()
            .reset_index()
            .rename(columns={client_col: "Numero_clienti"})
        )
        city_summary = city_summary.merge(n_clienti, on=city_col, how="left")

    n_agenti = (
        df.groupby(city_col)[agent_col]
        .nunique()
        .reset_index()
        .rename(columns={agent_col: "Numero_agenti"})
    )
    city_summary = city_summary.merge(n_agenti, on=city_col, how="left")

    city_summary["Peso_%"] = (
        city_summary["Totale_Fatturato_2025"]
        / city_summary["Totale_Fatturato_2025"].sum()
        * 100
    )

    # Dettaglio città → agente
    city_agent = (
        df.groupby([city_col, agent_col])[fatt_col]
        .sum()
        .reset_index()
        .rename(columns={fatt_col: "Fatturato_2025"})
        .sort_values([city_col, "Fatturato_2025"], ascending=[True, False])
    )
    if client_col:
        n_clienti_ca = (
            df.groupby([city_col, agent_col])[client_col]
            .nunique()
            .reset_index()
            .rename(columns={client_col: "Numero_clienti"})
        )
        city_agent = city_agent.merge(
            n_clienti_ca, on=[city_col, agent_col], how="left"
        )

    # Vista agente → città con %
    agent_city_raw = (
        df.groupby([agent_col, city_col])[fatt_col]
        .sum()
        .reset_index()
        .rename(columns={fatt_col: "Fatturato_2025"})
    )
    if client_col:
        n_clienti_ac = (
            df.groupby([agent_col, city_col])[client_col]
            .nunique()
            .reset_index()
            .rename(columns={client_col: "Numero_clienti"})
        )
        agent_city_raw = agent_city_raw.merge(
            n_clienti_ac, on=[agent_col, city_col], how="left"
        )

    agent_city = agent_city_raw.merge(
        agent_totals[[agent_col, "Totale_Fatturato_2025"]],
        on=agent_col,
        how="left",
    )
    agent_city["Peso_%_sul_totale_agente"] = (
        agent_city["Fatturato_2025"] / agent_city["Totale_Fatturato_2025"] * 100
    )
    agent_city = agent_city.sort_values(
        [agent_col, "Fatturato_2025"], ascending=[True, False]
    )

    # ------------------ TABS ------------------

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📍 Riassunto per città",
            "🏬 Città → Agente",
            "🧑‍💼 Agente → Città (con %)",
            "📈 Totale agenti + grafico",
        ]
    )

    # -------- TAB 1: RIASSUNTO CITTÀ --------
    with tab1:
        st.markdown("### Riassunto per città")

        all_cities = sorted(city_summary[city_col].unique())
        selected_cities = st.multiselect(
            "Filtra per città", options=all_cities, default=[]
        )

        if selected_cities:
            cs_filtered = city_summary[city_summary[city_col].isin(selected_cities)]
        else:
            cs_filtered = city_summary

        st.dataframe(cs_filtered)

        excel_city = df_to_excel_bytes(cs_filtered, "Fatturato_città")
        st.download_button(
            "⬇️ Scarica riepilogo città (Excel)",
            data=excel_city,
            file_name="riepilogo_città.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------- TAB 2: CITTÀ → AGENTE --------
    with tab2:
        st.markdown("### Dettaglio città → agente")

        all_cities = sorted(city_agent[city_col].unique())
        all_agents = sorted(city_agent[agent_col].unique())

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            sel_cities = st.multiselect(
                "Filtra per città", options=all_cities, default=[]
            )
        with col_f2:
            sel_agents = st.multiselect(
                "Filtra per agente", options=all_agents, default=[]
            )

        ca_filtered = city_agent.copy()
        if sel_cities:
            ca_filtered = ca_filtered[ca_filtered[city_col].isin(sel_cities)]
        if sel_agents:
            ca_filtered = ca_filtered[ca_filtered[agent_col].isin(sel_agents)]

        st.dataframe(ca_filtered)

        excel_city_agent = df_to_excel_bytes(ca_filtered, "Città_agente")
        st.download_button(
            "⬇️ Scarica città → agente (Excel)",
            data=excel_city_agent,
            file_name="città_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------- TAB 3: AGENTE → CITTÀ CON % --------
    with tab3:
        st.markdown("### Vista agente → città con peso % sul totale agente")

        all_agents = sorted(agent_city[agent_col].unique())
        sel_agents_tab3 = st.multiselect(
            "Filtra per agente", options=all_agents, default=[]
        )

        ac_filtered = agent_city.copy()
        if sel_agents_tab3:
            ac_filtered = ac_filtered[ac_filtered[agent_col].isin(sel_agents_tab3)]

        st.dataframe(ac_filtered)

        excel_agent_city = df_to_excel_bytes(ac_filtered, "Agente_città_%")
        st.download_button(
            "⬇️ Scarica agente → città (Excel)",
            data=excel_agent_city,
            file_name="agente_città_percentuale.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------- TAB 4: TOTALE AGENTE + GRAFICO --------
    with tab4:
        st.markdown("### Totale fatturato per agente")

        sel_agents_tab4 = st.multiselect(
            "Filtra tabella per agente",
            options=sorted(agent_totals[agent_col].unique()),
            default=[],
        )

        at_filtered = agent_totals.copy()
        if sel_agents_tab4:
            at_filtered = at_filtered[at_filtered[agent_col].isin(sel_agents_tab4)]

        st.dataframe(at_filtered)

        st.markdown("---")
        st.markdown("### Dettaglio e grafico per singolo agente")

        agente_scelto = st.selectbox(
            "Seleziona agente per il grafico",
            options=sorted(agent_totals[agent_col].unique()),
        )

        df_agente = agent_city[agent_city[agent_col] == agente_scelto].copy()
        if df_agente.empty:
            st.warning("Nessun dato per questo agente.")
        else:
            # Tabella per agente
            st.markdown(f"#### Dettaglio città per agente **{agente_scelto}**")
            df_agente = df_agente.rename(
                columns={"Peso_%_sul_totale_agente": "Peso_%"}
            )
            st.dataframe(df_agente[[city_col, "Fatturato_2025", "Peso_%"]])

            # Grafico a barre orizzontali con tutte le città
            st.markdown("#### Grafico fatturato per città (tutte le città visibili)")

            df_plot = df_agente.sort_values("Fatturato_2025", ascending=True)

            fig, ax = plt.subplots(figsize=(8, 6))
            ax.barh(df_plot[city_col], df_plot["Fatturato_2025"])

            ax.set_xlabel("Fatturato 2025")
            ax.set_ylabel("Città")
            ax.set_title(f"Fatturato per città – Agente {agente_scelto}")

            for i, (val, perc) in enumerate(
                zip(df_plot["Fatturato_2025"], df_plot["Peso_%"])
            ):
                ax.text(
                    val,
                    i,
                    f"{perc:.1f}%",
                    va="center",
                    ha="left",
                    fontsize=8,
                )

            fig.tight_layout()
            st.pyplot(fig)

        st.markdown("---")
        st.markdown("### 📥 Scarica report completo (tutti i fogli)")

        excel_full = full_report_excel(
            city_summary=city_summary,
            city_agent=city_agent,
            agent_city=agent_city.rename(
                columns={"Peso_%_sul_totale_agente": "Peso_%"}
            ),
            agent_totals=agent_totals,
        )

        st.download_button(
            "⬇️ Scarica report completo (Excel)",
            data=excel_full,
            file_name="report_fatturato_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
