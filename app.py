import io
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="Analisi Vendite & Ottimizzazione Area",
    layout="wide",
)

APP_DIR = Path(__file__).resolve().parent
DEFAULT_XLSX = APP_DIR / "statisticatot25.xlsx"

CACHE_DIR = APP_DIR / ".cache_app"
CACHE_DIR.mkdir(exist_ok=True)

COLUMN_MAP_FILE = CACHE_DIR / "column_map.json"
NON_MOVABLE_FILE = CACHE_DIR / "non_spostabili.json"


# ============================================================
# UTIL - JSON SAFE
# ============================================================
def safe_read_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def safe_write_json(path: Path, data) -> None:
    try:
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


# ============================================================
# UTIL - EXCEL EXPORT
# ============================================================
def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            df = df.copy()
            safe_name = name[:31] if name else "Sheet1"
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


# ============================================================
# COLONNE - MAPPING ROBUSTO
# ============================================================
def norm_col(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    repl = (
        ("à", "a"),
        ("è", "e"),
        ("é", "e"),
        ("ì", "i"),
        ("ò", "o"),
        ("ù", "u"),
        ("’", "'"),
    )
    for a, b in repl:
        s = s.replace(a, b)
    return s


COLUMN_SYNONYMS = {
    "agente": ["agente", "venditore", "sales", "seller", "commerciale"],
    "citta": ["citta", "città", "city", "comune", "localita", "località"],
    "cliente": ["cliente", "client", "ragione sociale", "ragionesociale", "customer"],
    "fatturato": ["fatturato", "valore", "importo", "revenue", "fatt", "vendite"],
    "categoria": ["categoria", "cat", "family", "gruppo", "category"],
    "anno": ["anno", "year"],
    "zona": ["zona", "area", "region", "territorio", "zone"],  # opzionale
}

REQUIRED_KEYS = ["agente", "citta", "cliente", "fatturato", "categoria", "anno"]


def guess_column_map(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    # mappa: nome_normalizzato -> nome_originale
    normed = {norm_col(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {k: None for k in COLUMN_SYNONYMS.keys()}
    for key, syns in COLUMN_SYNONYMS.items():
        for s in syns:
            if norm_col(s) in normed:
                out[key] = normed[norm_col(s)]
                break
    return out


def validate_map(colmap: Dict[str, Optional[str]]) -> Tuple[bool, List[str], bool]:
    missing = [k for k in REQUIRED_KEYS if not colmap.get(k)]
    vals = [colmap[k] for k in REQUIRED_KEYS if colmap.get(k)]
    has_duplicates = len(set(vals)) != len(vals)
    return (len(missing) == 0, missing, has_duplicates)


# ============================================================
# LETTURA EXCEL - CACHE
# ============================================================
@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: Optional[bytes], file_path: str) -> pd.DataFrame:
    if file_bytes:
        bio = io.BytesIO(file_bytes)
        df = pd.read_excel(bio, engine="openpyxl")
    else:
        df = pd.read_excel(file_path, engine="openpyxl")
    return df


# ============================================================
# CLEANING DATA
# ============================================================
def drop_totali_rows(df: pd.DataFrame, key_cols: List[str]) -> pd.DataFrame:
    """
    Esclude righe in cui una qualsiasi colonna chiave contiene 'Totali' (case-insensitive).
    Implementazione vettoriale (molto più veloce di apply).
    """
    masks = []
    for c in key_cols:
        if c in df.columns:
            s = df[c].astype(str).str.strip().str.lower()
            masks.append(s.eq("totali"))
    if not masks:
        return df
    mask_tot = masks[0]
    for m in masks[1:]:
        mask_tot = mask_tot | m
    return df.loc[~mask_tot].copy()


def clean_and_prepare(df_raw: pd.DataFrame, colmap: Dict[str, Optional[str]]) -> pd.DataFrame:
    df = df_raw.copy()

    # Rinomina colonne selezionate in nomi standard interni
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

    # Keep colonne utili (solo se presenti)
    keep = ["agente", "citta", "cliente", "fatturato", "categoria", "anno"]
    if "zona" in df.columns:
        keep.append("zona")

    missing_after_rename = [c for c in keep if c not in df.columns]
    if missing_after_rename:
        # Non deve succedere se mapping ok, ma se succede: messaggio chiaro
        raise KeyError(f"Mancano colonne dopo mapping: {missing_after_rename}")

    df = df[keep].copy()

    # Pulizia stringhe
    for c in ["agente", "citta", "cliente", "categoria"]:
        df[c] = df[c].astype(str).str.strip()

    if "zona" in df.columns:
        df["zona"] = df["zona"].astype(str).str.strip()

    # Tipi
    df["anno"] = pd.to_numeric(df["anno"], errors="coerce").fillna(0).astype(int)
    df["fatturato"] = pd.to_numeric(df["fatturato"], errors="coerce").fillna(0.0).astype(float)

    # Escludi righe "Totali"
    df = drop_totali_rows(df, key_cols=["agente", "citta", "cliente", "categoria"])

    # Drop righe senza dati chiave
    df = df[(df["agente"] != "") & (df["cliente"] != "")]
    return df


def filter_active_clients_2025(df: pd.DataFrame) -> pd.DataFrame:
    """
    Considerare solo clienti con fatturato 2025 > 0
    """
    df_2025 = df[df["anno"] == 2025].copy()
    if df_2025.empty:
        return df.iloc[0:0].copy()  # nessun 2025 -> dataset vuoto
    active = df_2025.groupby("cliente", as_index=False)["fatturato"].sum()
    active = active[active["fatturato"] > 0]["cliente"].unique().tolist()
    return df[df["cliente"].isin(active)].copy()


# ============================================================
# REPORT
# ============================================================
def report_fatturato_per_citta(df_anno: pd.DataFrame) -> pd.DataFrame:
    return (
        df_anno.groupby("citta", as_index=False)["fatturato"]
        .sum()
        .sort_values("fatturato", ascending=False)
    )


def report_fatturato_per_categoria(df_anno: pd.DataFrame) -> pd.DataFrame:
    return (
        df_anno.groupby("categoria", as_index=False)["fatturato"]
        .sum()
        .sort_values("fatturato", ascending=False)
    )


def report_fatturato_per_agente_2025(df: pd.DataFrame) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()
    if df_2025.empty:
        return pd.DataFrame(columns=["agente", "fatturato", "clienti_attivi", "%_incidenza"])

    fatt = df_2025.groupby("agente", as_index=False)["fatturato"].sum()
    clienti = (
        df_2025.groupby("agente", as_index=False)["cliente"]
        .nunique()
        .rename(columns={"cliente": "clienti_attivi"})
    )
    out = fatt.merge(clienti, on="agente", how="left")

    total = out["fatturato"].sum()
    out["%_incidenza"] = (out["fatturato"] / total * 100.0) if total > 0 else 0.0
    return out.sort_values("fatturato", ascending=False)


def report_agente_citta_cliente(df_anno: pd.DataFrame) -> pd.DataFrame:
    return (
        df_anno.groupby(["agente", "citta", "cliente"], as_index=False)["fatturato"]
        .sum()
        .sort_values(["agente", "citta", "fatturato"], ascending=[True, True, False])
    )


# ============================================================
# OTTIMIZZAZIONE AREA (euristica semplice e stabile)
# ============================================================
def compute_agent_loads_2025(df_2025: pd.DataFrame, dispersion_weight: float) -> pd.DataFrame:
    if df_2025.empty:
        return pd.DataFrame(columns=["agente", "clienti", "citta_distinte", "fatturato", "load_score"])

    base = df_2025.groupby("agente").agg(
        clienti=("cliente", "nunique"),
        citta_distinte=("citta", "nunique"),
        fatturato=("fatturato", "sum"),
    ).reset_index()
    base["load_score"] = base["clienti"] + dispersion_weight * base["citta_distinte"]
    return base.sort_values("load_score", ascending=False)


def build_client_table_2025(df: pd.DataFrame) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()
    if df_2025.empty:
        return pd.DataFrame(columns=["agente", "citta", "cliente", "categoria", "fatt_2025"])

    cols = ["agente", "citta", "cliente", "categoria"]
    if "zona" in df_2025.columns:
        cols.append("zona")

    out = (
        df_2025.groupby(cols, as_index=False)["fatturato"]
        .sum()
        .rename(columns={"fatturato": "fatt_2025"})
        .sort_values("fatt_2025", ascending=True)  # piccoli prima
    )
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
    - sovraccarico: clienti > target_max_clienti
    - sposta prima clienti piccoli
    - assegna preferibilmente stessa città all’agente meno carico
    - vincolo: il donatore non perde oltre max_fatt_loss_pct del suo fatturato 2025
    """
    if df_clients.empty:
        return {
            "moves": pd.DataFrame(columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025"]),
            "before": pd.DataFrame(),
            "after": pd.DataFrame(),
            "note": pd.DataFrame([{"msg": "Nessun dato 2025 disponibile."}]),
        }

    dfc = df_clients.copy()
    dfc["movable"] = ~dfc["cliente"].isin(set(non_spostabili))

    # snapshot iniziale
    before_clients = dfc[["agente", "citta", "cliente", "fatt_2025"]].copy()
    before_loads = compute_agent_loads_2025(
        before_clients.rename(columns={"fatt_2025": "fatturato"}).assign(anno=2025),
        dispersion_weight=dispersion_weight,
    )

    # agent state
    agent_clients = dfc.groupby("agente")["cliente"].apply(lambda x: set(x.tolist())).to_dict()
    agent_cities = dfc.groupby("agente")["citta"].apply(lambda x: set(x.tolist())).to_dict()
    agent_fatt_init = dfc.groupby("agente")["fatt_2025"].sum().to_dict()
    agent_fatt = dict(agent_fatt_init)

    # min fatt allowed for donor
    min_fatt_allowed = {a: agent_fatt_init.get(a, 0.0) * (1.0 - max_fatt_loss_pct) for a in agent_fatt_init.keys()}

    # current lookups by client
    current_city_of = dfc.set_index("cliente")["citta"].to_dict()
    current_fatt_of = dfc.set_index("cliente")["fatt_2025"].to_dict()
    current_agent_of = dfc.set_index("cliente")["agente"].to_dict()

    all_agents = sorted(agent_clients.keys())

    def agent_load_score(a: str) -> float:
        return len(agent_clients.get(a, set())) + dispersion_weight * len(agent_cities.get(a, set()))

    def rebuild_city_agents() -> Dict[str, set]:
        city_agents: Dict[str, set] = {}
        for a, cities in agent_cities.items():
            for city in cities:
                city_agents.setdefault(city, set()).add(a)
        return city_agents

    city_agents = rebuild_city_agents()

    # chi è sovraccarico?
    overloaded = [a for a in all_agents if len(agent_clients.get(a, set())) > target_max_clienti]
    overloaded = sorted(overloaded, key=lambda a: agent_load_score(a), reverse=True)

    if not overloaded:
        return {
            "moves": pd.DataFrame(columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025"]),
            "before": before_loads,
            "after": before_loads,
            "after_clients": before_clients.sort_values(["agente", "citta", "fatt_2025"], ascending=[True, True, False]),
            "note": pd.DataFrame([{"msg": "Nessun agente sovraccarico con i parametri attuali."}]),
        }

    def pick_target(city: str, donor: str) -> Optional[str]:
        candidates = []

        if prefer_same_city:
            candidates = [a for a in city_agents.get(city, set()) if a != donor]

        if not candidates:
            candidates = [a for a in all_agents if a != donor]

        candidates = sorted(candidates, key=lambda a: agent_load_score(a))

        # prova a non superare troppo la soglia
        for a in candidates:
            if len(agent_clients.get(a, set())) <= target_max_clienti:
                return a

        return candidates[0] if candidates else None

    moves = []
    moves_count = 0

    for donor in overloaded:
        donor_clients = [c for c in agent_clients.get(donor, set()) if c not in non_spostabili]
        donor_clients = sorted(donor_clients, key=lambda c: current_fatt_of.get(c, 0.0))  # piccoli prima

        for client in donor_clients:
            if moves_count >= max_moves:
                break

            # se donor è rientrato
            if len(agent_clients.get(donor, set())) <= target_max_clienti:
                break

            fatt = float(current_fatt_of.get(client, 0.0))
            city = str(current_city_of.get(client, "") or "")

            # vincolo fatturato donatore
            if (agent_fatt.get(donor, 0.0) - fatt) < min_fatt_allowed.get(donor, 0.0):
                continue

            target = pick_target(city, donor)
            if not target:
                continue

            # move
            agent_clients[donor].discard(client)
            agent_fatt[donor] = agent_fatt.get(donor, 0.0) - fatt

            agent_clients.setdefault(target, set()).add(client)
            agent_fatt[target] = agent_fatt.get(target, 0.0) + fatt

            current_agent_of[client] = target

            # aggiorna cities per donor e target (ricalcolo localizzato)
            agent_cities[donor] = set(current_city_of[c] for c in agent_clients.get(donor, set()))
            agent_cities[target] = set(current_city_of[c] for c in agent_clients.get(target, set()))
            city_agents = rebuild_city_agents()

            moves.append(
                {"cliente": client, "citta": city, "da_agente": donor, "a_agente": target, "fatt_2025": fatt}
            )
            moves_count += 1

        if moves_count >= max_moves:
            break

    moves_df = pd.DataFrame(moves).sort_values("fatt_2025", ascending=True) if moves else pd.DataFrame(
        columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025"]
    )

    # build after clients
    rows = []
    for a, clset in agent_clients.items():
        for c in clset:
            rows.append(
                {
                    "agente": a,
                    "citta": current_city_of.get(c, ""),
                    "cliente": c,
                    "fatt_2025": float(current_fatt_of.get(c, 0.0)),
                }
            )
    after_clients = pd.DataFrame(rows)
    after_loads = compute_agent_loads_2025(
        after_clients.rename(columns={"fatt_2025": "fatturato"}).assign(anno=2025),
        dispersion_weight=dispersion_weight,
    )

    note_msg = (
        f"Spostamenti: {len(moves_df)} | max_clienti={target_max_clienti} | "
        f"peso_dispersione={dispersion_weight} | max_loss_fatt={max_fatt_loss_pct:.0%}"
    )

    return {
        "moves": moves_df,
        "before": before_loads,
        "after": after_loads,
        "after_clients": after_clients.sort_values(["agente", "citta", "fatt_2025"], ascending=[True, True, False]),
        "note": pd.DataFrame([{"msg": note_msg}]),
    }


# ============================================================
# UI - HEADER
# ============================================================
st.title("📊 Analisi vendite + 🧭 Ottimizzazione area")

with st.sidebar:
    st.header("Dati")

    # IMPORTANT: su Streamlit Cloud è più sicuro default su UPLOAD
    mode = st.radio("Sorgente file", ["Upload Excel", "Locale (statisticatot25.xlsx)"], index=0)

    uploaded_bytes = None
    file_path = str(DEFAULT_XLSX)

    if mode == "Upload Excel":
        up = st.file_uploader("Carica statisticatot25.xlsx", type=["xlsx"])
        if up is None:
            st.info("Carica un file Excel per iniziare.")
            st.stop()
        uploaded_bytes = up.getvalue()
    else:
        if not DEFAULT_XLSX.exists():
            st.warning("Non trovo statisticatot25.xlsx accanto a app.py. Usa Upload Excel.")
            st.stop()

    st.divider()
    st.header("Mapping colonne")


# ============================================================
# LOAD RAW
# ============================================================
try:
    df_raw = load_excel_cached(uploaded_bytes, file_path)
except Exception as e:
    st.error(f"Errore lettura Excel: {e}")
    st.stop()

if df_raw is None or df_raw.empty:
    st.error("Excel vuoto o non leggibile.")
    st.stop()

df_cols = list(df_raw.columns)

# guess + saved
guess = guess_column_map(df_raw)
saved_map = safe_read_json(COLUMN_MAP_FILE, default={})
colmap_pref = {**guess, **saved_map}

saved_non_movable = safe_read_json(NON_MOVABLE_FILE, default=[])


# ============================================================
# SIDEBAR - PICKERS (ANTI-ERRORI)
# ============================================================
with st.sidebar:
    st.caption("Seleziona le colonne corrette. Le obbligatorie devono essere tutte diverse.")

    def pick_col(label: str, keyname: str, optional: bool = False) -> Optional[str]:
        current = colmap_pref.get(keyname)

        if optional:
            options = ["(non impostata)"] + df_cols
            idx = options.index(current) if current in options else 0
            sel = st.selectbox(label, options=options, index=idx, key=f"map_{keyname}")
            return None if sel == "(non impostata)" else sel

        # obbligatoria
        placeholder = "(seleziona...)"
        options = [placeholder] + df_cols
        idx = options.index(current) if current in options else 0
        sel = st.selectbox(label, options=options, index=idx, key=f"map_{keyname}")
        return None if sel == placeholder else sel

    colmap_ui: Dict[str, Optional[str]] = {}
    colmap_ui["agente"] = pick_col("Colonna Agente", "agente")
    colmap_ui["citta"] = pick_col("Colonna Città", "citta")
    colmap_ui["cliente"] = pick_col("Colonna Cliente", "cliente")
    colmap_ui["fatturato"] = pick_col("Colonna Fatturato", "fatturato")
    colmap_ui["categoria"] = pick_col("Colonna Categoria", "categoria")
    colmap_ui["anno"] = pick_col("Colonna Anno", "anno")
    colmap_ui["zona"] = pick_col("Colonna Zona (opzionale)", "zona", optional=True)

    ok, missing, has_dup = validate_map(colmap_ui)
    if not ok:
        st.error(f"Mancano colonne: {', '.join(missing)}")
        st.stop()
    if has_dup:
        st.error("Hai assegnato la stessa colonna a più campi. Seleziona colonne diverse.")
        st.stop()

    if st.button("💾 Salva mapping colonne"):
        safe_write_json(COLUMN_MAP_FILE, colmap_ui)
        st.success("Mapping salvato!")


# ============================================================
# PREP DATA
# ============================================================
try:
    df = clean_and_prepare(df_raw, colmap_ui)
except KeyError as e:
    st.error(f"Errore mapping colonne: {e}")
    st.stop()
except Exception as e:
    st.error(f"Errore pulizia dati: {e}")
    st.stop()

df = filter_active_clients_2025(df)

if df.empty:
    st.warning("Nessun cliente con fatturato 2025 > 0 (oppure anno 2025 assente).")
    st.stop()


# ============================================================
# FILTRI REPORT
# ============================================================
with st.sidebar:
    st.divider()
    st.header("Filtri report")

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
    out = d[d["anno"] == int(year_sel)].copy()
    if filt_agents:
        out = out[out["agente"].isin(filt_agents)]
    if filt_cities:
        out = out[out["citta"].isin(filt_cities)]
    if filt_cats:
        out = out[out["categoria"].isin(filt_cats)]
    return out


df_year = apply_filters(df)


# ============================================================
# TABS
# ============================================================
tab_report, tab_opt, tab_data = st.tabs(["📈 Report", "🧭 Ottimizzazione area", "🧾 Anteprima dati"])


# =========================
# TAB REPORT
# =========================
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
    rep_zone = report_agente_citta_cliente(df_year)
    st.dataframe(rep_zone, use_container_width=True, hide_index=True)

    st.divider()
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


# =========================
# TAB OTTIMIZZAZIONE
# =========================
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
                safe_write_json(NON_MOVABLE_FILE, non_movable_sel)
                st.success("Salvata!")
        with cbtn2:
            if st.button("🧹 Svuota lista"):
                non_movable_sel = []
                safe_write_json(NON_MOVABLE_FILE, non_movable_sel)
                st.success("Svuotata!")

    with right:
        st.markdown("#### Situazione iniziale (2025)")
        df_2025_for_loads = df[df["anno"] == 2025].copy()
        loads0 = compute_agent_loads_2025(df_2025_for_loads, dispersion_weight=dispersion_weight)
        st.dataframe(loads0.style.format({"fatturato": "{:,.2f}", "load_score": "{:,.2f}"}), use_container_width=True, hide_index=True)

    st.divider()

    run = st.button("🚀 Esegui simulazione ottimizzazione (prima/dopo)")

    if run:
        with st.spinner("Simulazione in corso..."):
            sim = simulate_reassignment(
                df_clients=df_clients,
                dispersion_weight=float(dispersion_weight),
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

        st.markdown("### Confronto load_score (Top 20 per 'Prima')")
        if not sim["before"].empty:
            before_plot = sim["before"].sort_values("load_score", ascending=False).head(20).set_index("agente")["load_score"]
            after_plot = sim["after"].set_index("agente")["load_score"].reindex(before_plot.index).fillna(0)
            chart_df = pd.DataFrame({"Prima": before_plot, "Dopo": after_plot})
            st.bar_chart(chart_df)

        st.markdown("### Export simulazione (Excel)")
        sim_bytes = to_excel_bytes(
            {
                "spostamenti": sim["moves"],
                "prima_carichi": sim["before"],
                "dopo_carichi": sim["after"],
                "assegnazioni_dopo": sim.get("after_clients", pd.DataFrame()),
            }
        )

        st.download_button(
            "⬇️ Scarica Simulazione Excel (prima/dopo)",
            data=sim_bytes,
            file_name=f"simulazione_ottimizzazione_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    st.markdown("#### Tabella clienti 2025 (base ottimizzazione)")
    st.dataframe(df_clients.sort_values("fatt_2025", ascending=False), use_container_width=True, hide_index=True)


# =========================
# TAB DATA
# =========================
with tab_data:
    st.subheader("Anteprima dati puliti (Totali esclusi + solo clienti 2025>0)")
    st.write(f"Righe: {len(df):,}")
    st.dataframe(df.head(300), use_container_width=True, hide_index=True)
