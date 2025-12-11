import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# ------------------ CONFIGURAZIONE ------------------

st.set_page_config(page_title="Report Fatturato Agente/Città", layout="wide")
st.title("📊 Report Fatturato Agente / Città")

uploaded_file = st.file_uploader("Carica il file Excel clienti", type=["xlsx", "xls"])


# ------------------ FUNZIONI SUPPORTO ------------------

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str) -> BytesIO:
    """Trasforma un DataFrame in un file Excel in memoria."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer


def full_report_excel(city_summary, city_agent, agent_city, agent_totals) -> BytesIO:
    """Crea un file Excel con tutti i report in fogli separati."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        city_summary.to_excel(writer, index=False, sheet_name="Riassunto_città")
        city_agent.to_excel(writer, index=False, sheet_name="Città_Agente")
        agent_city.to_excel(writer, index=False, sheet_name="Agente_Città_%")
        agent_totals.to_excel(writer, index=False, sheet_name="Totale_Agente")
    buffer.seek(0)
    return buffer


# ------------------ LOGICA PRINCIPALE ------------------

if uploaded_file is not None:
    # Legge il file Excel
    df = pd.read_excel(uploaded_file)

    # Rinomina colonne in modo robusto rispetto al file che usi
    rename_map = {}

    # Citta -> Città
    if "Citta" in df.columns and "Città" not in df.columns:
        rename_map["Citta"] = "Città"

    # Esercizio -> Cliente
    if "Esercizio" in df.columns and "Cliente" not in df.columns:
        rename_map["Esercizio"] = "Cliente"

    # acquistato al 10/12/2025 -> Fatturato2025
    if "acquistato al 10/12/2025" in df.columns and "Fatturato2025" not in df.columns:
        rename_map["acquistato al 10/12/2025"] = "Fatturato2025"

    if rename_map:
        df = df.rename(columns=rename_map)

    # Verifica colonne necessarie
    required_cols = ["Città", "Agente", "Cliente", "Fatturato2025"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Mancano queste colonne nel file Excel: {missing}")
        st.stop()

    # Anteprima
    st.subheader("Anteprima dati (prime 20 righe)")
    st.dataframe(df.head(20))

    st.markdown("---")

    # ======================
    # CALCOLI BASE
    # ======================

    # Totale per agente
    agent_totals = (
        df.groupby("Agente")
        .agg(
            Totale_Fatturato_2025=("Fatturato2025", "sum"),
            Numero_città=("Città", "nunique"),
            Numero_clienti=("Cliente", "nunique"),
        )
        .reset_index()
        .sort_values("Totale_Fatturato_2025", ascending=False)
    )

    # Riassunto per città
    city_summary = (
        df.groupby("Città")
        .agg(
            Totale_Fatturato_2025=("Fatturato2025", "sum"),
            Numero_clienti=("Cliente", "nunique"),
            Numero_agenti=("Agente", "nunique"),
        )
        .reset_index()
        .sort_values("Totale_Fatturato_2025", ascending=False)
    )
    city_summary["Peso_%"] = (
        city_summary["Totale_Fatturato_2025"]
        / city_summary["Totale_Fatturato_2025"].sum()
        * 100
    )

    # Dettaglio città → agente
    city_agent = (
        df.groupby(["Città", "Agente"])
        .agg(
            Fatturato_2025=("Fatturato2025", "sum"),
            Numero_clienti=("Cliente", "nunique"),
        )
        .reset_index()
        .sort_values(by=["Città", "Fatturato_2025"], ascending=[True, False])
    )

    # Vista agente → città con %
    agent_city_raw = (
        df.groupby(["Agente", "Città"])
        .agg(
            Fatturato_2025=("Fatturato2025", "sum"),
            Numero_clienti=("Cliente", "nunique"),
        )
        .reset_index()
    )

    agent_city = agent_city_raw.merge(
        agent_totals[["Agente", "Totale_Fatturato_2025"]],
        on="Agente",
        how="left",
    )
    agent_city["Peso_%_sul_totale_agente"] = (
        agent_city["Fatturato_2025"] / agent_city["Totale_Fatturato_2025"] * 100
    )
    agent_city = agent_city.sort_values(
        by=["Agente", "Fatturato_2025"], ascending=[True, False]
    )

    # ======================
    # TABS
    # ======================

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📍 Riassunto per città",
            "🏬 Città → Agente",
            "🧑‍💼 Agente → Città (con %)",
            "📈 Totale agenti + grafico",
        ]
    )

    # -------- TAB 1: RIASSUNTO PER CITTÀ --------
    with tab1:
        st.markdown("### Riassunto per città")
        st.dataframe(city_summary)

        excel_cs = df_to_excel_bytes(city_summary, "Riassunto_città")
        st.download_button(
            "⬇️ Scarica riassunto città (Excel)",
            data=excel_cs,
            file_name="riassunto_città.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------- TAB 2: CITTÀ → AGENTE --------
    with tab2:
        st.markdown("### Dettaglio città → agente")
        st.dataframe(city_agent)

        excel_ca = df_to_excel_bytes(city_agent, "Città_Agente")
        st.download_button(
            "⬇️ Scarica città → agente (Excel)",
            data=excel_ca,
            file_name="città_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------- TAB 3: AGENTE → CITTÀ (CON %) --------
    with tab3:
        st.markdown("### Vista agente → città (con % sul totale agente)")
        st.dataframe(agent_city)

        excel_ac = df_to_excel_bytes(agent_city, "Agente_Città_%")
        st.download_button(
            "⬇️ Scarica agente → città (Excel)",
            data=excel_ac,
            file_name="agente_città_percentuale.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -------- TAB 4: TOTALE AGENTI + GRAFICO --------
    with tab4:
        st.markdown("### Totale fatturato per agente")
        st.dataframe(agent_totals)

        excel_at = df_to_excel_bytes(agent_totals, "Totale_Agente")
        st.download_button(
            "⬇️ Scarica totale per agente (Excel)",
            data=excel_at,
            file_name="totale_per_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("---")
        st.markdown("### Dettaglio e grafico per singolo agente")

        tutti_agenti = sorted(agent_totals["Agente"].unique())
        agente_scelto = st.selectbox(
            "Seleziona agente per il grafico", options=tutti_agenti
        )

        df_agente = agent_city[agent_city["Agente"] == agente_scelto].copy()
        if df_agente.empty:
            st.warning("Nessun dato per questo agente.")
        else:
            df_agente = df_agente.rename(
                columns={"Peso_%_sul_totale_agente": "Peso_%"}
            )

            st.markdown(f"#### Dettaglio città per agente **{agente_scelto}**")
            st.dataframe(df_agente[["Città", "Fatturato_2025", "Peso_%"]])

            st.markdown("#### Grafico fatturato per città (torta)")

            fatt_per_citta = (
                df_agente.groupby("Città")["Fatturato_2025"]
                .sum()
                .reset_index()
                .sort_values("Fatturato_2025", ascending=False)
            )

            if not fatt_per_citta.empty:
                fig, ax = plt.subplots()
                wedges, texts, autotexts = ax.pie(
                    fatt_per_citta["Fatturato_2025"],
                    labels=fatt_per_citta["Città"],
                    autopct="%1.1f%%",
                    startangle=90,
                )
                ax.axis("equal")

                # Scritte piccole per evitare sovrapposizioni eccessive
                for t in texts + autotexts:
                    t.set_fontsize(6)

                st.pyplot(fig)

        st.markdown("---")
        st.markdown("### 📥 Report completo (tutti i fogli)")

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
