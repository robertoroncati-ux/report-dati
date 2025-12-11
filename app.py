import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt

st.set_page_config(page_title="Report Area Ponente", layout="wide")

st.title("📊 Report Fatturato per Città – Area Ponente")

uploaded_file = st.file_uploader("Carica il file Excel clienti", type=["xlsx", "xls"])

def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer

if uploaded_file is not None:
    # Legge il file
    df = pd.read_excel(uploaded_file)

    # Rinomina le colonne come vogliamo noi
    df = df.rename(columns={
        "Citta": "Città",
        "Agente": "Agente",
        "Esercizio": "Cliente",
        "acquistato al 10/12/2025": "Fatturato2025"
    })

    st.subheader("Anteprima dati (prime 20 righe)")
    st.dataframe(df.head(20))

    st.markdown("---")

    # Totale per agente (ci servirà in più tab)
    agent_totals = df.groupby("Agente").agg(
        Totale_Fatturato_2025=("Fatturato2025", "sum"),
        Numero_città=("Città", "nunique"),
        Numero_clienti=("Cliente", "nunique")
    ).reset_index().sort_values("Totale_Fatturato_2025", ascending=False)

    tab1, tab2, tab3, tab4 = st.tabs([
        "📍 Riassunto per città",
        "🏬 Dettaglio città → agente",
        "🧑‍💼 Vista agente → città (con %)",
        "📈 Riepilogo agente + grafico"
    ])

    # ======================
    # TAB 1 – RIASSUNTO PER CITTÀ
    # ======================
    with tab1:
        st.markdown("### Riassunto per città")

        city_summary = df.groupby("Città").agg(
            Totale_Fatturato_2025=("Fatturato2025", "sum"),
            Numero_clienti=("Cliente", "nunique"),
            Numero_agenti=("Agente", "nunique")
        ).reset_index()

        city_summary = city_summary.sort_values(
            by="Totale_Fatturato_2025", ascending=False
        )

        st.dataframe(city_summary)

        buffer1 = df_to_excel_bytes(city_summary, "Riassunto per città")
        st.download_button(
            label="⬇️ Scarica riassunto per città (Excel)",
            data=buffer1,
            file_name="riassunto_citta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ======================
    # TAB 2 – DETTAGLIO CITTÀ → AGENTE
    # ======================
    with tab2:
        st.markdown("### Dettaglio per città e agente")
        st.write("Per ogni città e agente: fatturato e numero di locali seguiti da quell’agente in quella città.")

        city_agent = df.groupby(["Città", "Agente"]).agg(
            Fatturato_2025=("Fatturato2025", "sum"),
            Numero_clienti=("Cliente", "nunique")
        ).reset_index()

        city_agent = city_agent.sort_values(
            by=["Città", "Fatturato_2025"], ascending=[True, False]
        )

        st.dataframe(city_agent)

        buffer2 = df_to_excel_bytes(city_agent, "Città-Agente")
        st.download_button(
            label="⬇️ Scarica dettaglio città → agente (Excel)",
            data=buffer2,
            file_name="dettaglio_citta_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ======================
    # TAB 3 – VISTA AGENTE → CITTÀ (CON %)
    # ======================
    with tab3:
        st.markdown("### Totale fatturato per agente e città (con peso %)")

        agent_city = df.groupby(["Agente", "Città"]).agg(
            Fatturato_2025=("Fatturato2025", "sum"),
            Numero_clienti=("Cliente", "nunique")
        ).reset_index()

        # Aggiunge il totale agente per calcolare il peso %
        agent_city = agent_city.merge(
            agent_totals[["Agente", "Totale_Fatturato_2025"]],
            on="Agente",
            how="left"
        )

        agent_city["Peso_%_sul_totale_agente"] = (
            agent_city["Fatturato_2025"] /
            agent_city["Totale_Fatturato_2025"] * 100
        )

        agent_city = agent_city.sort_values(
            by=["Agente", "Fatturato_2025"], ascending=[True, False]
        )

        st.dataframe(agent_city)

        buffer3 = df_to_excel_bytes(agent_city, "Agente-Città")
        st.download_button(
            label="⬇️ Scarica fatturato agente → città (Excel)",
            data=buffer3,
            file_name="fatturato_agente_per_citta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ======================
    # TAB 4 – RIEPILOGO AGENTE + GRAFICO A TORTA
    # ======================
    with tab4:
        st.markdown("### Riepilogo totale per agente")
        st.dataframe(agent_totals)

        buffer4 = df_to_excel_bytes(agent_totals, "Totale per agente")
        st.download_button(
            label="⬇️ Scarica totale per agente (Excel)",
            data=buffer4,
            file_name="totale_per_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.markdown("### Grafico a torta: ripartizione fatturato per città (per agente)")

        agente_scelto = st.selectbox(
            "Seleziona un agente",
            agent_totals["Agente"].tolist()
        )

        df_agente = df[df["Agente"] == agente_scelto]
        fatt_per_citta = df_agente.groupby("Città").agg(
            Fatturato_2025=("Fatturato2025", "sum")
        ).reset_index().sort_values("Fatturato_2025", ascending=False)

        if not fatt_per_citta.empty:
            # Calcola anche la % per tabella
            totale_agente = fatt_per_citta["Fatturato_2025"].sum()
            fatt_per_citta["Peso_%"] = fatt_per_citta["Fatturato_2025"] / totale_agente * 100

            cols1, cols2 = st.columns([1, 1])

            with cols1:
                st.write(f"Ripartizione fatturato 2025 per città – **{agente_scelto}**")
                st.dataframe(fatt_per_citta)

            with cols2:
                fig, ax = plt.subplots()
                ax.pie(
                    fatt_per_citta["Fatturato_2025"],
                    labels=fatt_per_citta["Città"],
                    autopct="%1.1f%%"
                )
                ax.axis("equal")
                st.pyplot(fig)
