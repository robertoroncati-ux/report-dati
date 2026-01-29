import io
import json
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =========================
# Config base
# =========================
st.set_page_config(
    page_title="Analisi Vendite & Ottimizzazione Area",
    layout="wide",
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_XLSX = APP_DIR / "statisticatot25.xlsx"
CACHE_DIR = APP_DIR / ".cache_app"
CACHE_DIR.mkdir(exist_ok=True)

SETTINGS_FILE = CACHE_DIR / "settings.json"
NON_MOVABLE_FILE = CACHE_DIR / "non_spostabili.json"
COLUMN_MAP_FILE = CACHE_DIR / "column_map.json"


# =========================
# Util
# =========================
def _safe_read_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _safe_write_json(path: Path, data) -> None:
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def norm_col(s: str) -> str:
    """Normalizza nomi colonne per matching robusto."""
    if s is None:
        return ""
    return (
        str(s)
        .strip()
        .lower()
        .replace("à", "a")
        .replace("è", "e")
        .replace("é", "e")
        .replace("ì", "i")
        .replace("ò", "o")
        .replace("ù", "u")
    )


def is_totali_row(row: pd.Series, key_cols: List[str]) -> bool:
    """Esclude righe 'Totali' presenti in una qualsiasi delle colonne chiave."""
    for c in key_cols:
        if c in row and isinstance(row[c], str) and row[c].strip().lower() == "totali":
            return True
    return False


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            # Excel sheet name max 31 chars
            safe_name = name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


# =========================
# Colonne attese + sinonimi
# =========================
COLUMN_SYNONYMS = {
    "agente": ["agente", "venditore", "sales", "seller", "commerciale"],
    "citta": ["citta", "città", "city", "comune", "localita", "località"],
    "cliente": ["cliente", "client", "ragione sociale", "ragionesociale", "customer"],
    "fatturato": ["fatturato", "valore", "importo", "revenue", "fatt", "vendite"],
    "categoria": ["categoria", "cat", "family", "gruppo", "category"],
    "anno": ["anno", "year"],
    # opzionale
    "zona": ["zona", "area", "region", "territorio", "zone"],
}


REQUIRED_KEYS = ["agente", "citta", "cliente", "fatturato", "categoria", "anno"]


def guess_column_map(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    normed = {norm_col(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {k: None for k in COLUMN_SYNONYMS.keys()}
    for key, syns in COLUMN_SYNONYMS.items():
        for s in syns:
            if norm_col(s) in normed:
                out[key] = normed[norm_col(s)]
                break
    return out


def validate_column_map(colmap: Dict[str, Optional[str]]) -> Tuple[bool, List[str]]:
    missing = [k for k in REQUIRED_KEYS if not colmap.get(k)]
    return (len(missing) == 0, missing)


# =========================
# Lettura dati + cache anti-freeze
# =========================
@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: Optional[bytes], file_path_str: str) -> pd.DataFrame:
    """
    Se file_bytes è valorizzato, legge da BytesIO (upload).
    Se no, legge da path (locale).
    """
    if file_bytes:
        bio = io.BytesIO(file_bytes)
        df = pd.read_excel(bio, engine="openpyxl")
    else:
        df = pd.read_excel(file_path_str, engine="openpyxl")
    return df


def clean_and_prepare(df_raw: pd.DataFrame, colmap: Dict[str, str]) -> pd.DataFrame:
    df = df_raw.copy()

    # Rinomina in chiavi standard interne
    rename = {
        colmap["agente"]: "agente",
        colmap["citta"]: "citta",
        colmap["cliente"]: "cliente",
        colmap["fatturato"]: "fatturato",
        colmap["categoria"]: "categoria",
        colmap["anno"]: "anno",
    }
    if colmap.get("zona"):
        rename[colmap["zona"]] = "zona"

    df = df.rename(columns=rename)

    # Keep solo colonne utili (se presenti)
    keep = ["agente", "citta", "cliente", "fatturato", "categoria", "anno"]
    if "zona" in df.columns:
        keep.append("zona")
    df = df[keep]

    # Pulizia stringhe
    for c in ["agente", "citta", "cliente", "categoria"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    if "zona" in df.columns:
        df["zona"] = df["zona"].astype(str).str.strip()

    # Tipi
    df["anno"] = pd.to_numeric(df["anno"], errors="coerce").fillna(0).astype(int)
    df["fatturato"] = pd.to_numeric(df["fatturato"], errors="coerce").fillna(0.0).astype(float)

    # Escludi righe "Totali"
    key_cols = ["agente", "citta", "cliente", "categoria"]
    mask_tot = df.apply(lambda r: is_totali_row(r, key_cols), axis=1)
    df = df.loc[~mask_tot].copy()

    # Righe vuote / invalid
    df = df[(df["cliente"] != "") & (df["agente"] != "")]
    return df


def filter_active_clients_2025(df: pd.DataFrame) -> pd.DataFrame:
    """
    Considerare solo clienti con fatturato 2025 > 0
    => tengo tutti i record (anche altri anni) solo per quei clienti,
       ma di default report e ottimizzazione lavorano sul 2025.
    """
    df_2025 = df[df["anno"] == 2025].copy()
    active = (
        df_2025.groupby("cliente", as_index=False)["fatturato"]
        .sum()
        .rename(columns={"fatturato": "fatt_2025"})
    )
    active = active[active["fatt_2025"] > 0]["cliente"].unique().tolist()
    return df[df["cliente"].isin(active)].copy()


# =========================
# Report
# =========================
def report_fatturato_per_citta(df_anno: pd.DataFrame) -> pd.DataFrame:
    out = (
        df_anno.groupby("citta", as_index=False)["fatturato"]
        .sum()
        .sort_values("fatturato", ascending=False)
    )
    return out


def report_fatturato_per_categoria(df_anno: pd.DataFrame) -> pd.DataFrame:
    out = (
        df_anno.groupby("categoria", as_index=False)["fatturato"]
        .sum()
        .sort_values("fatturato", ascending=False)
    )
    return out


def report_fatturato_per_agente_2025(df: pd.DataFrame) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()

    fatt = df_2025.groupby("agente", as_index=False)["fatturato"].sum()
    clienti = df_2025.groupby("agente", as_index=False)["cliente"].nunique().rename(columns={"cliente": "clienti_attivi"})
    out = fatt.merge(clienti, on="agente", how="left")

    total = out["fatturato"].sum()
    out["%_incidenza"] = (out["fatturato"] / total * 100.0) if total > 0 else 0.0

    out = out.sort_values("fatturato", ascending=False)
    return out


def report_zona_agente_citta_cliente(df_anno: pd.DataFrame) -> pd.DataFrame:
    """
    Report gerarchico: agente → città → cliente (con fatturato).
    In tabella piatta, ordinata, così è filtrabile e esportabile.
    """
    out = (
        df_anno.groupby(["agente", "citta", "cliente"], as_index=False)["fatturato"]
        .sum()
        .sort_values(["agente", "citta", "fatturato"], ascending=[True, True, False])
    )
    return out


# =========================
# Ottimizzazione area (euristica)
# =========================
@dataclass
class AgentLoad:
    agente: str
    clienti: int
    citta_distinte: int
    fatturato: float
    load_score: float


def compute_agent_loads(df_2025: pd.DataFrame, dispersion_weight: float) -> pd.DataFrame:
    base = df_2025.groupby("agente").agg(
        clienti=("cliente", "nunique"),
        citta_distinte=("citta", "nunique"),
        fatturato=("fatturato", "sum"),
    ).reset_index()
    base["load_score"] = base["clienti"] + dispersion_weight * base["citta_distinte"]
    return base.sort_values("load_score", ascending=False)


def build_client_table_2025(df: pd.DataFrame) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()
    cols = ["agente", "citta", "cliente", "categoria"]
    if "zona" in df_2025.columns:
        cols.append("zona")

    out = (
        df_2025.groupby(cols, as_index=False)["fatturato"]
        .sum()
        .rename(columns={"fatturato": "fatt_2025"})
    )
    out = out.sort_values("fatt_2025", ascending=True)  # piccoli prima
    return out


def simulate_reassignment(
    df_clients: pd.DataFrame,
    dispersion_weight: float,
    target_max_clienti: int,
    max_fatt_loss_pct: float,
    non_spostabili: List[str],
    prefer_same_city: bool = True,
    max_moves: int = 10_000,
) -> Dict[str, pd.DataFrame]:
    """
    Euristica:
    - identifica agenti sovraccarichi (clienti > target_max_clienti OR load_score alto)
    - sposta prima clienti piccoli (fatt_2025 basso)
    - regola assegnazione: stessa città -> agente meno carico
    - vincolo: non far perdere troppo fatturato agli agenti (max_fatt_loss_pct)
    """
    # Stato iniziale
    dfc = df_clients.copy()
    dfc["movable"] = ~dfc["cliente"].isin(set(non_spostabili))

    # Per rapidità: indicizzazioni
    # Fatturato per agente (iniziale)
    fatt_init = dfc.groupby("agente")["fatt_2025"].sum().to_dict()
    clienti_init = dfc.groupby("agente")["cliente"].nunique().to_dict()

    # Carichi iniziali
    def loads_snapshot(d: pd.DataFrame) -> pd.DataFrame:
        return compute_agent_loads(
            d.rename(columns={"fatt_2025": "fatturato"})[["agente", "citta", "cliente", "fatt_2025"]]
             .rename(columns={"fatt_2025": "fatturato"})
             .assign(anno=2025)  # dummy
             .rename(columns={"fatturato": "fatturato"})[["agente", "citta", "cliente", "fatturato", "anno"]],
            dispersion_weight=dispersion_weight
        )

    # Funzione alternativa più diretta (senza ricostruire) per prestazioni
    def loads_fast(d: pd.DataFrame) -> pd.DataFrame:
        base = d.groupby("agente").agg(
            clienti=("cliente", "nunique"),
            citta_distinte=("citta", "nunique"),
            fatturato=("fatt_2025", "sum"),
        ).reset_index()
        base["load_score"] = base["clienti"] + dispersion_weight * base["citta_distinte"]
        return base

    loads = loads_fast(dfc)

    # Definizione sovraccarico: clienti sopra soglia
    # (se vuoi, qui puoi evolvere con percentili su load_score)
    overloaded = loads[loads["clienti"] > target_max_clienti].copy()
    overloaded = overloaded.sort_values("load_score", ascending=False)

    # Se nessuno è sovraccarico, esci
    if overloaded.empty:
        return {
            "moves": pd.DataFrame(columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025"]),
            "before": loads.sort_values("load_score", ascending=False),
            "after": loads.sort_values("load_score", ascending=False),
            "note": pd.DataFrame([{"msg": "Nessun agente risulta sovraccarico con i parametri attuali."}]),
        }

    # Stato dinamico
    current_agent_of = dfc.set_index("cliente")["agente"].to_dict()
    current_city_of = dfc.set_index("cliente")["citta"].to_dict()
    current_fatt_of = dfc.set_index("cliente")["fatt_2025"].to_dict()

    # Tabelle dinamiche per agenti
    agent_clients = dfc.groupby("agente")["cliente"].apply(lambda x: set(x.tolist())).to_dict()
    agent_cities = dfc.groupby("agente")["citta"].apply(lambda x: set(x.tolist())).to_dict()
    agent_fatt = dfc.groupby("agente")["fatt_2025"].sum().to_dict()

    def agent_load_score(a: str) -> float:
        c = len(agent_clients.get(a, set()))
        z = len(agent_cities.get(a, set()))
        return c + dispersion_weight * z

    # Per scegliere candidato in città
    # pre-calcolo: per città, lista agenti che già hanno clienti lì
    def rebuild_city_agents() -> Dict[str, set]:
        city_agents: Dict[str, set] = {}
        for a, cities in agent_cities.items():
            for city in cities:
                city_agents.setdefault(city, set()).add(a)
        return city_agents

    city_agents = rebuild_city_agents()

    # Vincolo fatturato: non perdere oltre X% rispetto al fatturato iniziale (2025)
    min_fatt_allowed = {a: fatt_init.get(a, 0.0) * (1.0 - max_fatt_loss_pct) for a in fatt_init.keys()}

    # Lista clienti spostabili ordinata per piccolo fatturato
    # (ma andiamo agente per agente sovraccarico)
    moves = []
    moves_count = 0

    # Utility: scegli miglior destinatario
    all_agents = sorted(agent_clients.keys())

    def pick_target_agent(city: str, donor: str) -> Optional[str]:
        # candidati: stessa città se preferito
        candidates = []
        if prefer_same_city:
            for a in city_agents.get(city, set()):
                if a != donor:
                    candidates.append(a)

        # se vuoto, prova tutti
        if not candidates:
            candidates = [a for a in all_agents if a != donor]

        # ordina per load_score crescente (meno carico)
        candidates = sorted(candidates, key=lambda a: agent_load_score(a))

        # regole hard: non superare target_max_clienti (tolleranza piccola)
        for a in candidates:
            if len(agent_clients.get(a, set())) <= target_max_clienti:
                return a

        # se tutti sopra, almeno prendi il meno carico
        return candidates[0] if candidates else None

    # Processa agenti sovraccarichi
    overloaded_agents = overloaded["agente"].tolist()

    for donor in overloaded_agents:
        # clienti del donor ordinati per fatturato crescente e spostabili
        donor_clients = [c for c in agent_clients.get(donor, set()) if c not in non_spostabili]
        donor_clients = sorted(donor_clients, key=lambda c: current_fatt_of.get(c, 0.0))

        for client in donor_clients:
            if moves_count >= max_moves:
                break

            # se donor già sotto soglia clienti, stop
            if len(agent_clients.get(donor, set())) <= target_max_clienti:
                break

            city = current_city_of.get(client, "")
            fatt = current_fatt_of.get(client, 0.0)

            # vincolo: non far scendere donor sotto la soglia fatturato minima
            if (agent_fatt.get(donor, 0.0) - fatt) < min_fatt_allowed.get(donor, 0.0):
                continue

            target = pick_target_agent(city=city, donor=donor)
            if not target:
                continue

            # Esegui spostamento
            # rimuovi da donor
            agent_clients[donor].discard(client)
            agent_fatt[donor] = agent_fatt.get(donor, 0.0) - fatt
            # città: potrebbe diventare vuota, va ricalcolata bene (per semplicità la ricalcoliamo smart)
            # aggiungi a target
            agent_clients.setdefault(target, set()).add(client)
            agent_fatt[target] = agent_fatt.get(target, 0.0) + fatt

            # aggiorna mapping cliente
            current_agent_of[client] = target

            # ricostruzione city sets per donor/target (corretta ma rapida: aggiorno via ricalcolo solo per quei 2)
            # donor cities
            donor_cities = set(current_city_of[c] for c in agent_clients.get(donor, set()))
            agent_cities[donor] = donor_cities
            # target cities
            target_cities = set(current_city_of[c] for c in agent_clients.get(target, set()))
            agent_cities[target] = target_cities

            # aggiorna city_agents (ricostruzione leggera e sicura)
            city_agents = rebuild_city_agents()

            moves.append(
                {
                    "cliente": client,
                    "citta": city,
                    "da_agente": donor,
                    "a_agente": target,
                    "fatt_2025": fatt,
                }
            )
            moves_count += 1

        if moves_count >= max_moves:
            break

    moves_df = pd.DataFrame(moves)

    # BEFORE
    before = loads_fast(dfc).sort_values("load_score", ascending=False)

    # AFTER: costruisci df_after da agent_clients
    rows = []
    for a, clients_set in agent_clients.items():
        for c in clients_set:
            rows.append(
                {
                    "agente": a,
                    "citta": current_city_of.get(c, ""),
                    "cliente": c,
                    "fatt_2025": current_fatt_of.get(c, 0.0),
                }
            )
    after_clients = pd.DataFrame(rows)
    after = loads_fast(after_clients).sort_values("load_score", ascending=False)

    return {
        "moves": moves_df.sort_values("fatt_2025", ascending=True),
        "before": before,
        "after": after,
        "after_clients": after_clients.sort_values(["agente", "citta", "fatt_2025"], ascending=[True, True, False]),
        "note": pd.DataFrame(
            [{
                "msg": f"Spostamenti effettuati: {len(moves_df)}. Parametri: max_clienti={target_max_clienti}, "
                       f"peso_dispersione={dispersion_weight}, max_loss_fatt={max_fatt_loss_pct:.0%}."
            }]
        ),
    }


# =========================
# UI
# =========================
st.title("📊 Analisi vendite + 🧭 Ottimizzazione area (Streamlit)")

with st.sidebar:
    st.header("Dati")
    mode = st.radio("Sorgente file", ["Locale (statisticatot25.xlsx)", "Upload Excel"], index=0)

    uploaded_bytes = None
    file_path = str(DEFAULT_XLSX)

    if mode == "Upload Excel":
        up = st.file_uploader("Carica statisticatot25.xlsx", type=["xlsx"])
        if up is not None:
            uploaded_bytes = up.getvalue()
            file_path = up.name
    else:
        if not DEFAULT_XLSX.exists():
            st.warning("Non trovo statisticatot25.xlsx nella stessa cartella di app.py. Puoi usare Upload.")
        file_path = str(DEFAULT_XLSX)

    st.divider()
    st.header("Impostazioni colonne")

# Carico settings
saved_colmap = _safe_read_json(COLUMN_MAP_FILE, default={})
saved_non_movable = _safe_read_json(NON_MOVABLE_FILE, default=[])

# Leggi file (con cache)
df_raw = None
try:
    if mode == "Upload Excel" and uploaded_bytes is None:
        df_raw = None
    else:
        df_raw = load_excel_cached(uploaded_bytes, file_path)
except Exception as e:
    st.error(f"Errore lettura Excel: {e}")

if df_raw is None:
    st.info("Carica un file Excel (upload) oppure metti statisticatot25.xlsx accanto a app.py.")
    st.stop()

# Mappa colonne (guess + fallback a salvato)
guess = guess_column_map(df_raw)
colmap = {**guess, **saved_colmap}  # salvato vince

# Sidebar mapping interattivo
with st.sidebar:
    df_cols = list(df_raw.columns)

    def pick_col(label: str, keyname: str, optional: bool = False) -> Optional[str]:
        current = colmap.get(keyname)
        options = ["(non impostata)"] + df_cols if optional else df_cols
        idx = 0
        if current in df_cols:
            idx = options.index(current) if current in options else 0
        sel = st.selectbox(label, options=options, index=idx, key=f"map_{keyname}")
        if optional and sel == "(non impostata)":
            return None
        return sel

    colmap_ui = {}
    colmap_ui["agente"] = pick_col("Colonna Agente", "agente")
    colmap_ui["citta"] = pick_col("Colonna Città", "citta")
    colmap_ui["cliente"] = pick_col("Colonna Cliente", "cliente")
    colmap_ui["fatturato"] = pick_col("Colonna Fatturato", "fatturato")
    colmap_ui["categoria"] = pick_col("Colonna Categoria", "categoria")
    colmap_ui["anno"] = pick_col("Colonna Anno", "anno")
    colmap_ui["zona"] = pick_col("Colonna Zona (opzionale)", "zona", optional=True)

    ok, missing = validate_column_map(colmap_ui)
    if not ok:
        st.error(f"Colonne mancanti: {', '.join(missing)}")
        st.stop()

    if st.button("💾 Salva mapping colonne"):
        _safe_write_json(COLUMN_MAP_FILE, colmap_ui)
        st.success("Salvato!")

# Preparazione dati
df = clean_and_prepare(df_raw, colmap_ui)
df = filter_active_clients_2025(df)

# Filtri generali
with st.sidebar:
    st.divider()
    st.header("Filtri (report)")

    anni = sorted([a for a in df["anno"].unique().tolist() if a != 0])
    default_year = 2025 if 2025 in anni else (anni[-1] if anni else 2025)
    year_sel = st.selectbox("Anno", options=anni if anni else [2025], index=(anni.index(default_year) if anni and default_year in anni else 0))

    agents_all = sorted(df["agente"].unique().tolist())
    cities_all = sorted(df["citta"].unique().tolist())
    cats_all = sorted(df["categoria"].unique().tolist())

    filt_agents = st.multiselect("Agenti", options=agents_all, default=[])
    filt_cities = st.multiselect("Città", options=cities_all, default=[])
    filt_cats = st.multiselect("Categorie", options=cats_all, default=[])

def apply_filters(d: pd.DataFrame) -> pd.DataFrame:
    out = d.copy()
    out = out[out["anno"] == int(year_sel)]
    if filt_agents:
        out = out[out["agente"].isin(filt_agents)]
    if filt_cities:
        out = out[out["citta"].isin(filt_cities)]
    if filt_cats:
        out = out[out["categoria"].isin(filt_cats)]
    return out

df_year = apply_filters(df)

# Tabs
tab_report, tab_opt, tab_data = st.tabs(["📈 Report", "🧭 Ottimizzazione area", "🧾 Anteprima dati"])

# -------------------------
# Report
# -------------------------
with tab_report:
    st.subheader(f"Report anno {year_sel} (solo clienti con fatturato 2025 > 0)")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### Fatturato per città")
        rep_city = report_fatturato_per_citta(df_year)
        st.dataframe(rep_city, use_container_width=True, hide_index=True)
        if not rep_city.empty:
            st.bar_chart(rep_city.set_index("citta")["fatturato"])

    with c2:
        st.markdown("### Fatturato per categoria")
        rep_cat = report_fatturato_per_categoria(df_year)
        st.dataframe(rep_cat, use_container_width=True, hide_index=True)
        if not rep_cat.empty:
            st.bar_chart(rep_cat.set_index("categoria")["fatturato"])

    st.divider()

    st.markdown("### Fatturato per agente (solo 2025 + clienti attivi + % incidenza)")
    rep_agent_2025 = report_fatturato_per_agente_2025(df)
    st.dataframe(
        rep_agent_2025.style.format({"fatturato": "{:,.2f}", "%_incidenza": "{:.2f}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    st.markdown("### Fatturato per zona (agente → città → cliente)")
    rep_zone = report_zona_agente_citta_cliente(df_year)
    st.dataframe(rep_zone, use_container_width=True, hide_index=True)

    # Export report
    st.markdown("### Export report (Excel)")
    report_bytes = to_excel_bytes(
        {
            f"fatturato_citta_{year_sel}": rep_city,
            f"fatturato_categoria_{year_sel}": rep_cat,
            "fatturato_agente_2025": rep_agent_2025,
            f"agente_citta_cliente_{year_sel}": rep_zone,
        }
    )
    st.download_button(
        "⬇️ Scarica Report Excel",
        data=report_bytes,
        file_name=f"report_vendite_{year_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# -------------------------
# Ottimizzazione area
# -------------------------
with tab_opt:
    st.subheader("Ottimizzazione area (euristica): sovraccarico = clienti + dispersione città")

    df_clients = build_client_table_2025(df)

    left, right = st.columns([1, 1])

    with left:
        st.markdown("#### Parametri")
        target_max_clienti = st.number_input("Target max clienti per agente", min_value=50, max_value=300, value=140, step=5)
        dispersion_weight = st.number_input("Peso dispersione (numero città)", min_value=0.0, max_value=10.0, value=1.0, step=0.1)
        max_fatt_loss_pct = st.slider("Max perdita fatturato per agente (donatore)", min_value=0, max_value=50, value=15, step=1) / 100.0
        prefer_same_city = st.checkbox("Preferisci assegnazione nella stessa città", value=True)

        st.markdown("#### Clienti NON spostabili (gestione in app)")
        all_clients = sorted(df_clients["cliente"].unique().tolist())
        default_nonmov = [c for c in saved_non_movable if c in all_clients]

        non_movable_sel = st.multiselect(
            "Seleziona clienti non spostabili",
            options=all_clients,
            default=default_nonmov,
        )

        cbtn1, cbtn2 = st.columns(2)
        with cbtn1:
            if st.button("💾 Salva lista NON spostabili"):
                _safe_write_json(NON_MOVABLE_FILE, non_movable_sel)
                st.success("Salvata!")
        with cbtn2:
            if st.button("🧹 Svuota lista"):
                non_movable_sel = []
                _safe_write_json(NON_MOVABLE_FILE, non_movable_sel)
                st.success("Svuotata!")

    with right:
        st.markdown("#### Situazione iniziale (2025)")
        loads0 = compute_agent_loads(
            df[df["anno"] == 2025][["agente", "citta", "cliente", "fatturato", "anno"]].copy(),
            dispersion_weight=dispersion_weight,
        )
        st.dataframe(loads0.style.format({"fatturato": "{:,.2f}", "load_score": "{:,.2f}"}), use_container_width=True, hide_index=True)

    st.divider()

    run = st.button("🚀 Esegui simulazione ottimizzazione (prima/dopo)")

    if run:
        with st.spinner("Simulazione in corso..."):
            sim = simulate_reassignment(
                df_clients=df_clients,
                dispersion_weight=dispersion_weight,
                target_max_clienti=int(target_max_clienti),
                max_fatt_loss_pct=float(max_fatt_loss_pct),
                non_spostabili=non_movable_sel,
                prefer_same_city=prefer_same_city,
            )

        st.success("Simulazione completata.")

        st.info(sim["note"].iloc[0]["msg"] if not sim["note"].empty else "OK")

        st.markdown("### Spostamenti effettuati (clienti piccoli prima)")
        st.dataframe(sim["moves"].style.format({"fatt_2025": "{:,.2f}"}), use_container_width=True, hide_index=True)

        cA, cB = st.columns(2)
        with cA:
            st.markdown("### Prima")
            st.dataframe(sim["before"].style.format({"fatturato": "{:,.2f}", "load_score": "{:,.2f}"}), use_container_width=True, hide_index=True)
        with cB:
            st.markdown("### Dopo")
            st.dataframe(sim["after"].style.format({"fatturato": "{:,.2f}", "load_score": "{:,.2f}"}), use_container_width=True, hide_index=True)

        # Grafico prima/dopo load_score per agente (top 20)
        st.markdown("### Confronto load_score (Top 20 per 'Prima')")
        before_plot = sim["before"].sort_values("load_score", ascending=False).head(20).set_index("agente")[["load_score"]]
        after_plot = sim["after"].set_index("agente")[["load_score"]].reindex(before_plot.index).fillna(0)

        chart_df = pd.DataFrame({
            "Prima": before_plot["load_score"],
            "Dopo": after_plot["load_score"],
        })
        st.bar_chart(chart_df)

        # Export Excel simulazione
        st.markdown("### Export simulazione (Excel)")
        export_sheets = {
            "spostamenti": sim["moves"],
            "prima_carichi": sim["before"],
            "dopo_carichi": sim["after"],
        }
        if "after_clients" in sim:
            export_sheets["assegnazioni_dopo"] = sim["after_clients"]

        sim_bytes = to_excel_bytes(export_sheets)
        st.download_button(
            "⬇️ Scarica Simulazione Excel (prima/dopo)",
            data=sim_bytes,
            file_name=f"simulazione_ottimizzazione_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    st.markdown("#### Tabella clienti 2025 (base ottimizzazione)")
    st.dataframe(df_clients.sort_values("fatt_2025", ascending=False), use_container_width=True, hide_index=True)

# -------------------------
# Anteprima dati
# -------------------------
with tab_data:
    st.subheader("Anteprima dati puliti (dopo esclusione Totali + solo clienti 2025>0)")
    st.write(f"Righe: {len(df):,}")
    st.dataframe(df.head(200), use_container_width=True, hide_index=True)

