import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# ------------------ CONFIGURAZIONE ------------------
st.set_page_config(page_title="Report Fatturato Agente/Città", layout="wide")
st.title("📊 Report Fatturato Agente / Città")

uploaded_file = st.file_uploader("Carica il file Excel clienti", type=["xlsx", "xls"])

# ------------------ FUNZIONE PER EXPORT EXCEL ------------------
def full_report_excel(city_summary, city_agent, agent_city, agent_totals):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        city_summary.to_excel(writer, index=False, sheet_name="Fatturato città")
        city_agent.to_excel(writer, index=False, sheet_name="Città -> Agente")
        agent_city.to_excel(writer, index=False, sheet_name="Agente -> Città")
        agent_totals.to_excel(writer, index=False, sheet_name="Totale agente")
    buffer.seek(0)
    return buffer

# ------------------ ELABORAZIONE ------------------
if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)

    # RINOMINA COLONNE
    df = df.rename(columns={
        "Citta": "Città",
        "Agente": "Agente",
        "Cliente": "Cliente",
        "acquistato al 10/12/2025": "Fatturato_2025"
    })

    st.subheader("Anteprima dati (prime 20 righe)")
    st.dataframe(df.head(20))

    # ----------- FATTURATO TOTALE PER CITTÀ -----------
    st.markdown("### 📍 Vista città (con %)")
    city_summary = (
        df.groupby("Città")["Fatturato_2025"]
        .sum()
        .reset_index()
        .sort_values("Fatturato_2025", ascending=False)
    )
    city_summary["Peso_%"] = (city_summary["Fatturato_2025"] /
                              city_summary["Fatturato_2025"].sum() * 100)

    st.dataframe(city_summary)

    # ----------- FATTURATO CITTÀ PER AGENTE -----------
    st.markdown("### 👥 Vista agente → città (con %)")
    city_agent = (
        df.groupby(["Città", "Agente"])["Fatturato_2025"]
        .sum()
        .reset_index()
        .sort_values(["Città", "Fatturato_2025"], ascending=[True, False])
    )

    st.dataframe(city_agent)

    # ----------- AGENTI COINVOLTI -----------
    agent_list = sorted(df["Agente"].unique())

    tab1, tab2 = st.tabs(["📌 Seleziona agente", "📥 Riepilogo agente + grafico"])

    with tab1:
        st.write("Scegli l'agente per vedere il dettaglio delle città.")

    with tab2:

        agente_scelto = st.selectbox("Seleziona agente", agent_list)

        agent_filtered = df[df["Agente"] == agente_scelto]

        if agent_filtered.empty:
            st.warning("Nessun dato per questo agente.")
        else:
            st.markdown(f"### Dettaglio città per agente **{agente_scelto}**")

            fatt_per_citta = (
                agent_filtered.groupby("Città")["Fatturato_2025"]
                .sum()
                .reset_index()
                .sort_values("Fatturato_2025", ascending=False)
            )
            fatt_per_citta["Peso_%"] = (
                fatt_per_citta["Fatturato_2025"] /
                fatt_per_citta["Fatturato_2025"].sum() * 100
            )

            st.dataframe(fatt_per_citta)

            st.markdown("### 📊 Grafico fatturato per città (tutte visibili, leggibile)")

            # ------------------ GRAFICO A BARRE ORIZZONTALI ------------------
            fig, ax = plt.subplots(figsize=(8, 8))

            fatt_sorted = fatt_per_citta.sort_values("Fatturato_2025", ascending=True)

            ax.barh(fatt_sorted["Città"], fatt_sorted["Fatturato_2025"])
            ax.set_xlabel("Fatturato 2025")
            ax.set_ylabel("Città")
            ax.set_title(f"Fatturato per città – Agente {agente_scelto}")

            # Percentuale alla fine della barra
            for i, (val, perc) in enumerate(zip(
                fatt_sorted["Fatturato_2025"], fatt_sorted["Peso_%"]
            )):
                ax.text(val, i, f"{perc:.1f}%", va="center", ha="left", fontsize=8)

            fig.tight_layout()
            st.pyplot(fig)

    # ----------- TOTALE PER AGENTE -----------
    st.markdown("### 🧮 Totale fatturato per agente")

    agent_totals = (
        df.groupby("Agente")["Fatturato_2025"]
        .sum()
        .reset_index()
        .sort_values("Fatturato_2025", ascending=False)
    )

    st.dataframe(agent_totals)

    # ----------- DOWNLOAD EXCEL COMPLETO -----------
    st.markdown("### 📥 Scarica report completo in Excel")

    excel_bytes = full_report_excel(
        city_summary=city_summary,
        city_agent=city_agent,
        agent_city=fatt_per_citta if uploaded_file is not None else pd.DataFrame(),
        agent_totals=agent_totals
    )

    st.download_button(
        "📊 Scarica Excel",
        data=excel_bytes,
        file_name="report_fatturato.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
