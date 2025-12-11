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
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer


def full_report_excel(city_summary, city_agent, agent_city, agent_totals) -> BytesIO:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        city_summary.to_excel(writer, index=False, sheet_name="Riassunto_città")
        city_agent.to_excel(writer, index=False, sheet_name="Città_agente")
        agent_city.to_excel(writer, index=False, sheet_name="Agente_città_%")
        agent_totals.to_excel(writer, index=False, sheet_name="Totale_agente")
    buffer.seek(0)
    return buffer


# ------------------ LOGICA PRINCIPALE ------------------

if uploaded_file is not None:
    # Leggo il file
    df = pd.read_excel(uploaded_file)

    # Adatto i nomi delle colonne al nostro schema
    df = df.rename(columns={
        "Citta": "Città",
        "Agente": "Agente",
        "Esercizio": "Cliente",
        "acquistato al 10/12/2025": "Fatturato2025"
    })

    st.subheader("Anteprima dati (prime 20 righe)")
    st.dataframe(df.head(20))

    st.markdown("---")

    # ---------- CALCOLI BASE UNA SOLA VOLTA ----------

    # Totale per agente
    agent_totals = (
        df.groupby("Agente")
          .agg(
              Totale_Fatturato_2025=("Fatturato2025", "sum"),
              Numero_città=("Città", "nunique"),
              Numero_clienti=("Cliente", "nunique")
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
              Numero_agenti=("Agente", "nunique")
          )
          .reset_index()
          .sort_values("Totale_Fatturato_2025", ascending=False)
    )
    city_summary["Peso_%"] = (
        city_summary["Totale_Fatturato_2025"] /
        city_summary["Totale_Fatturato_2025"].sum() * 100
    )

    # Dettaglio città → agente
    city_agent = (
        df.groupby(["Città", "Agente"])
          .agg(
              Fatturato_2025=("Fatturato2025", "sum"),
              Numero_clienti=("Cliente", "nunique")
          )
          .reset_index()
          .sort_values(["Città", "Fatturato_2025"], ascending=[True, False])
    )

    # Vista agente → città con %
    agent_city_raw = (
        df.groupby(["Agente", "Città"])
          .agg(
              Fatturato_2025=("Fatturato2025", "sum"),
              Numero_clienti=("Cliente", "nunique")
          )
          .reset_index()
    )

    agent_city = agent_city_raw.merge(
        agent_totals[["Agente", "Totale_Fatturato_2025"]],
        on="Agente",
        how="left"
    )
    agent_city["Peso_%_sul_totale_agente"] = (
        agent_city["Fatturato_2025"] /
        agent_city["Totale_Fatturato_2025"] * 100
    )
    agent_city = agent_city.sort_values(
        ["Agente", "Fatturato_2025"], ascending=[True, False]
    )

    # ------------------ TABS COME PRIMA ------------------

    tab1, tab2, tab3, tab4 = st.tabs([
        "📍 Riassunto per città",
        "🏬 Dettaglio città → agente",
        "🧑‍💼 Vista agente → città (con %)",
        "📈 Totale agenti + grafico"
    ])

    # ---------- TAB 1: RIASSUNTO PER CITTÀ ----------
    with tab1:
        st.markdown("### Riassunto per città")

        # Filtro stile Excel
        lista_citta = sorted(city_summary["Città"].unique())
        filtro_citta = st.multiselect(
            "Filtra città (lascia vuoto per tutte)",
            options=lista_citta,
            default=[]
        )

        if filtro_citta:
            cs_view = city_summary[city_summary["Città"].isin(filtro_citta)]
        else:
            cs_view = city_summary

        st.dataframe(cs_view)

        buffer1 = df_to_excel_bytes(cs_view, "Riassunto_città")
        st.download_button(
            "⬇️ Scarica riassunto per città (Excel)",
            data=buffer1,
            file_name="riassunto_città.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---------- TAB 2: DETTAGLIO CITTÀ → AGENTE ----------
    with tab2:
        st.markdown("### Dettaglio per città e agente")

        lista_citta2 = sorted(city_agent["Città"].unique())
        lista_agenti2 = sorted(city_agent["Agente"].unique())

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_citta2 = st.multiselect(
                "Filtra città", options=lista_citta2, default=[]
            )
        with col_f2:
            filtro_agenti2 = st.multiselect(
                "Filtra agenti", options=lista_agenti2, default=[]
            )

        ca_view = city_agent.copy()
        if filtro_citta2:
            ca_view = ca_view[ca_view["Città"].isin(filtro_citta2)]
        if filtro_agenti2:
            ca_view = ca_view[ca_view["Agente"].isin(filtro_agenti2)]

        st.dataframe(ca_view)

        buffer2 = df_to_excel_bytes(ca_view, "Città_agente")
        st.download_button(
            "⬇️ Scarica città → agente (Excel)",
            data=buffer2,
            file_name="città_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---------- TAB 3: VISTA AGENTE → CITTÀ (CON %) ----------
    with tab3:
        st.markdown("### Totale fatturato per agente e città (con peso %)")

        lista_agenti3 = sorted(agent_city["Agente"].unique())
        filtro_agenti3 = st.multiselect(
            "Filtra agenti", options=lista_agenti3, default=[]
        )

        ac_view = agent_city.copy()
        if filtro_agenti3:
            ac_view = ac_view[ac_view["Agente"].isin(filtro_agenti3)]

        st.dataframe(ac_view)

        buffer3 = df_to_excel_bytes(ac_view, "Agente_città_%")
        st.download_button(
            "⬇️ Scarica agente → città (Excel)",
            data=buffer3,
            file_name="agente_città_percentuale.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---------- TAB 4: TOTALE AGENTI + GRAFICO ----------
    with tab4:
        st.markdown("### Riepilogo totale per agente")
        st.dataframe(agent_totals)

        buffer4 = df_to_excel_bytes(agent_totals, "Totale_agente")
        st.download_button(
            "⬇️ Scarica totale per agente (Excel)",
            data=buffer4,
            file_name="totale_per_agente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")
        st.markdown("### Grafico fatturato per città (per singolo agente)")

        agente_scelto = st.selectbox(
            "Seleziona un agente",
            options=sorted(agent_totals["Agente"].unique())
        )

        df_agente = df[df["Agente"] == agente_scelto].copy()
        if df_agente.empty:
            st.warning("Nessun dato per questo agente.")
        else:
            fatt_per_citta = (
                df_agente.groupby("Città")["Fatturato2025"]
                .sum()
                .reset_index()
                .sort_values("Fatturato2025", ascending=False)
            )
            totale_agente = fatt_per_citta["Fatturato2025"].sum()
            fatt_per_citta["Peso_%"] = (
                fatt_per_citta["Fatturato2025"] / totale_agente * 100
            )

            col_t, col_g = st.columns([1, 1])

            with col_t:
                st.write(f"Ripartizione fatturato 2025 per città – **{agente_scelto}**")
                st.dataframe(fatt_per_citta)

            with col_g:
                # Grafico a barre orizzontali (niente tortona)
                fatt_sorted = fatt_per_citta.sort_values("Fatturato2025", ascending=True)

                fig, ax = plt.subplots(figsize=(8, 6))
                ax.barh(fatt_sorted["Città"], fatt_sorted["Fatturato2025"])
                ax.set_xlabel("Fatturato 2025")
                ax.set_ylabel("Città")
                ax.set_title(f"Fatturato per città – Agente {agente_scelto}")

                for i, (val, perc) in enumerate(
                    zip(fatt_sorted["Fatturato2025"], fatt_sorted["Peso_%"])
                ):
                    ax.text(
                        val,
                        i,
                        f"{perc:.1f}%",
                        va="center",
                        ha="left",
                        fontsize=8
                    )

                fig.tight_layout()
                st.pyplot(fig)

        st.markdown("---")
        st.markdown("### 📥 Report completo in un unico Excel")

        full_buffer = full_report_excel(
            city_summary=city_summary,
            city_agent=city_agent,
            agent_city=agent_city,
            agent_totals=agent_totals
        )
        st.download_button(
            "⬇️ Scarica report completo (tutti i fogli)",
            data=full_buffer,
            file_name="report_area_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
