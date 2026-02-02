import io
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(page_title="Analisi Vendite & Ottimizzazione Area", layout="wide")

APP_DIR = Path(__file__).resolve().parent
DEFAULT_VENDITE_XLSX = APP_DIR / "statisticatot25.xlsx"
DEFAULT_PROVV_FILE = APP_DIR / "Progetto_ponente_26.xls"  # opzionale se lo tieni accanto

CACHE_DIR = APP_DIR / ".cache_app"
CACHE_DIR.mkdir(exist_ok=True)

COLUMN_MAP_VENDITE_FILE = CACHE_DIR / "column_map_vendite.json"
COLUMN_MAP_PROVV_FILE = CACHE_DIR / "column_map_provv.json"
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
            safe_name = (name or "Sheet1")[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return output.getvalue()


# ============================================================
# COLONNE - MAPPING ROBUSTO
# ============================================================
def norm_col(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    repl = (("à", "a"), ("è", "e"), ("é", "e"), ("ì", "i"), ("ò", "o"), ("ù", "u"), ("’", "'"))
    for a, b in repl:
        s = s.replace(a, b)
    return s


VENDITE_SYNONYMS = {
    "agente": ["agente", "venditore", "sales", "seller", "commerciale", "codice agente"],
    "citta": ["citta", "città", "city", "comune", "localita", "località"],
    "cliente": ["cliente", "client", "esercizio", "ragione sociale", "ragionesociale", "customer", "nominativo cliente"],
    "categoria": ["categoria", "cat", "family", "gruppo", "category"],
    "articolo": ["articolo", "prodotto", "item", "sku"],
    "fatturato": ["fatturato", "valore", "importo", "revenue", "vendite", "acquistato"],
    "anno": ["anno", "year"],
    "zona": ["zona", "area", "region", "territorio", "zone"],
}

PROVV_SYNONYMS = {
    "cod_agente": ["codice agente", "cod agente", "cod_agente", "id agente"],
    "agente": ["agente", "venditore", "sales", "seller", "commerciale"],
    "cod_cliente": ["cod cliente", "codice cliente", "cod_cliente", "id cliente", "cliente codice", "cod cliente "],
    "cliente_nome": ["nominativo cliente", "ragione sociale", "cliente", "esercizio", "customer"],
    "citta": ["citta", "città", "city", "comune", "localita", "località"],
    "prov": ["pr", "provincia", "prov"],
    "fatt_provv": ["fatturato per calcolo provvigioni", "fatturato provvigioni", "imponibile provvigioni", "base provvigioni"],
    "provvigione": ["importo provvigione", "provvigione", "commissione", "commissions"],
    "indirizzo": ["indirizzo", "address"],
    "cap": ["cap", "zip"],
}

REQUIRED_VENDITE = ["agente", "citta", "cliente", "categoria", "fatturato"]
REQUIRED_PROVV = ["cod_cliente", "provvigione"]


def guess_map(df: pd.DataFrame, synonyms: Dict[str, List[str]]) -> Dict[str, Optional[str]]:
    normed = {norm_col(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {k: None for k in synonyms.keys()}
    for key, syns in synonyms.items():
        for s in syns:
            if norm_col(s) in normed:
                out[key] = normed[norm_col(s)]
                break
    return out


def validate_map(colmap: Dict[str, Optional[str]], required_keys: List[str]) -> Tuple[bool, List[str], bool]:
    missing = [k for k in required_keys if not colmap.get(k)]
    vals = [colmap[k] for k in required_keys if colmap.get(k)]
    has_duplicates = len(set(vals)) != len(vals)
    return (len(missing) == 0, missing, has_duplicates)


def infer_year_from_header(header: str) -> Optional[int]:
    if not header:
        return None
    m = re.search(r"(20\d{2})", str(header))
    return int(m.group(1)) if m else None


def extract_client_code(s: str) -> str:
    """
    Estrae il codice cliente dal campo "cliente" delle vendite.
    Esempi:
      "C70987003 - NOME" -> "C70987003"
      "C70987003" -> "C70987003"
      "70987003 - NOME" -> "70987003"
    """
    if s is None:
        return ""
    t = str(s).strip()
    if not t:
        return ""
    # prende tutto prima di " - " se presente
    left = t.split(" - ")[0].strip()
    # oppure prima di "-" senza spazi
    left = left.split("-")[0].strip()
    return left


# ============================================================
# LETTURA FILE - CACHE
# ============================================================
@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: Optional[bytes], file_path: str) -> pd.DataFrame:
    if file_bytes:
        bio = io.BytesIO(file_bytes)
        return pd.read_excel(bio)
    return pd.read_excel(file_path)


# ============================================================
# CLEANING VENDITE
# ============================================================
def drop_totali_rows(df: pd.DataFrame, key_cols: List[str]) -> pd.DataFrame:
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


def clean_and_prepare_vendite(df_raw: pd.DataFrame, colmap: Dict[str, Optional[str]], fixed_year: Optional[int]) -> pd.DataFrame:
    df = df_raw.copy()

    rename = {
        colmap["agente"]: "agente",
        colmap["citta"]: "citta",
        colmap["cliente"]: "cliente",
        colmap["categoria"]: "categoria",
        colmap["fatturato"]: "fatturato",
    }
    if colmap.get("anno"):
        rename[colmap["anno"]] = "anno"
    if colmap.get("zona"):
        rename[colmap["zona"]] = "zona"
    if colmap.get("articolo"):
        rename[colmap["articolo"]] = "articolo"

    df = df.rename(columns=rename)

    keep = ["agente", "citta", "cliente", "categoria", "fatturato"]
    if "anno" in df.columns:
        keep.append("anno")
    if "zona" in df.columns:
        keep.append("zona")
    if "articolo" in df.columns:
        keep.append("articolo")

    missing = [c for c in keep if c not in df.columns]
    if missing:
        raise KeyError(f"Mancano colonne dopo mapping vendite: {missing}")

    df = df[keep].copy()

    for c in ["agente", "citta", "cliente", "categoria"]:
        df[c] = df[c].astype(str).str.strip()

    if "zona" in df.columns:
        df["zona"] = df["zona"].astype(str).str.strip()
    if "articolo" in df.columns:
        df["articolo"] = df["articolo"].astype(str).str.strip()

    df["fatturato"] = pd.to_numeric(df["fatturato"], errors="coerce").fillna(0.0).astype(float)

    if "anno" in df.columns:
        df["anno"] = pd.to_numeric(df["anno"], errors="coerce").fillna(0).astype(int)
    else:
        df["anno"] = int(fixed_year or 2025)

    key_cols = ["agente", "citta", "cliente", "categoria"]
    if "articolo" in df.columns:
        key_cols.append("articolo")

    df = drop_totali_rows(df, key_cols=key_cols)
    df = df[(df["agente"] != "") & (df["cliente"] != "")]
    return df


def filter_active_clients_2025(df: pd.DataFrame) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()
    if df_2025.empty:
        return df.iloc[0:0].copy()
    active = df_2025.groupby("cliente", as_index=False)["fatturato"].sum()
    active = active[active["fatturato"] > 0]["cliente"].unique().tolist()
    return df[df["cliente"].isin(active)].copy()


# ============================================================
# CLEANING PROVVIGIONI
# ============================================================
def clean_and_prepare_provv(df_raw: pd.DataFrame, colmap: Dict[str, Optional[str]]) -> pd.DataFrame:
    df = df_raw.copy()

    rename = {}
    for k, v in colmap.items():
        if v:
            rename[v] = k
    df = df.rename(columns=rename)

    # richieste minime
    missing = [k for k in REQUIRED_PROVV if k not in df.columns]
    if missing:
        raise KeyError(f"Mancano colonne dopo mapping provvigioni: {missing}")

    # normalizza campi base
    df["cod_cliente"] = df["cod_cliente"].astype(str).str.strip()
    df["provvigione"] = pd.to_numeric(df["provvigione"], errors="coerce").fillna(0.0).astype(float)

    if "fatt_provv" in df.columns:
        df["fatt_provv"] = pd.to_numeric(df["fatt_provv"], errors="coerce").fillna(0.0).astype(float)

    for c in ["agente", "cliente_nome", "citta", "prov", "indirizzo", "cap", "cod_agente"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    # elimina righe senza cod_cliente
    df = df[df["cod_cliente"].astype(str).str.strip() != ""].copy()
    return df


def aggregate_provv_by_cliente(df_provv: pd.DataFrame) -> pd.DataFrame:
    """
    Locale per locale = per cliente:
      provv_locale = somma(importo provvigione) per cod_cliente
      fatt_provv_locale = somma(fatturato per calcolo provvigioni) per cod_cliente (se presente)
    """
    if df_provv.empty:
        return pd.DataFrame(columns=["cod_cliente", "provv_locale", "fatt_provv_locale"])

    agg_dict = {"provvigione": "sum"}
    if "fatt_provv" in df_provv.columns:
        agg_dict["fatt_provv"] = "sum"

    out = df_provv.groupby("cod_cliente", as_index=False).agg(agg_dict)
    out = out.rename(columns={"provvigione": "provv_locale"})
    if "fatt_provv" in out.columns:
        out = out.rename(columns={"fatt_provv": "fatt_provv_locale"})
    else:
        out["fatt_provv_locale"] = 0.0
    return out


# ============================================================
# REPORT
# ============================================================
def report_city_summary(df_anno: pd.DataFrame, include_agents: bool, include_provv: bool) -> pd.DataFrame:
    agg = {
        "fatturato": ("fatturato", "sum"),
        "n_clienti": ("cliente", "nunique"),
    }
    if include_provv and "provv_locale" in df_anno.columns:
        agg["provvigioni"] = ("provv_locale", "sum")

    if include_agents:
        rep = (
            df_anno.groupby("citta")
            .agg(
                **agg,
                agenti=("agente", lambda s: ", ".join(sorted(set(map(str, s.dropna()))))),
            )
            .reset_index()
            .sort_values("fatturato", ascending=False)
        )
    else:
        rep = (
            df_anno.groupby("citta")
            .agg(**agg)
            .reset_index()
            .sort_values("fatturato", ascending=False)
        )

    return rep


def report_fatturato_per_categoria(df_anno: pd.DataFrame) -> pd.DataFrame:
    return df_anno.groupby("categoria", as_index=False)["fatturato"].sum().sort_values("fatturato", ascending=False)


def report_fatturato_per_agente_2025(df: pd.DataFrame, include_provv: bool) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()
    if df_2025.empty:
        cols = ["agente", "fatturato", "clienti_attivi", "%_incidenza"]
        if include_provv:
            cols.append("provvigioni")
        return pd.DataFrame(columns=cols)

    fatt = df_2025.groupby("agente", as_index=False)["fatturato"].sum()
    clienti = df_2025.groupby("agente", as_index=False)["cliente"].nunique().rename(columns={"cliente": "clienti_attivi"})
    out = fatt.merge(clienti, on="agente", how="left")

    total = out["fatturato"].sum()
    out["%_incidenza"] = (out["fatturato"] / total * 100.0) if total > 0 else 0.0

    if include_provv and "provv_locale" in df_2025.columns:
        provv = df_2025.groupby("agente", as_index=False)["provv_locale"].sum().rename(columns={"provv_locale": "provvigioni"})
        out = out.merge(provv, on="agente", how="left")
        out["provvigioni"] = out["provvigioni"].fillna(0.0)

    return out.sort_values("fatturato", ascending=False)


def report_agente_citta_cliente(df_anno: pd.DataFrame, include_provv: bool) -> pd.DataFrame:
    gcols = ["agente", "citta", "cliente"]
    agg = {"fatturato": "sum"}
    if include_provv and "provv_locale" in df_anno.columns:
        agg["provv_locale"] = "sum"

    out = df_anno.groupby(gcols, as_index=False).agg(agg)
    out = out.sort_values(["agente", "citta", "fatturato"], ascending=[True, True, False])
    return out


# ============================================================
# OTTIMIZZAZIONE
# ============================================================
def build_client_table_2025(df: pd.DataFrame) -> pd.DataFrame:
    df_2025 = df[df["anno"] == 2025].copy()
    if df_2025.empty:
        return pd.DataFrame(columns=["agente", "citta", "cliente", "fatt_2025", "client_code", "provv_locale", "fatt_provv_locale"])

    cols = ["agente", "citta", "cliente"]
    if "zona" in df_2025.columns:
        cols.append("zona")

    out = (
        df_2025.groupby(cols, as_index=False)["fatturato"]
        .sum()
        .rename(columns={"fatturato": "fatt_2025"})
    )

    out["client_code"] = out["cliente"].apply(extract_client_code)

    # se il df ha già provv_locale a riga vendite (dopo merge), sommiamo per cliente
    if "provv_locale" in df_2025.columns:
        provv = (
            df_2025.groupby(["cliente"], as_index=False)["provv_locale"].sum()
            .rename(columns={"provv_locale": "provv_locale_sum"})
        )
        out = out.merge(provv, on="cliente", how="left")
        out["provv_locale"] = out["provv_locale_sum"].fillna(0.0)
        out = out.drop(columns=["provv_locale_sum"])
    else:
        out["provv_locale"] = 0.0

    if "fatt_provv_locale" in df_2025.columns:
        fp = (
            df_2025.groupby(["cliente"], as_index=False)["fatt_provv_locale"].sum()
            .rename(columns={"fatt_provv_locale": "fatt_provv_locale_sum"})
        )
        out = out.merge(fp, on="cliente", how="left")
        out["fatt_provv_locale"] = out["fatt_provv_locale_sum"].fillna(0.0)
        out = out.drop(columns=["fatt_provv_locale_sum"])
    else:
        out["fatt_provv_locale"] = 0.0

    return out.sort_values("fatt_2025", ascending=True)


def apply_manual_overrides(df_clients: pd.DataFrame, overrides: Dict[str, str]) -> pd.DataFrame:
    if df_clients.empty or not overrides:
        return df_clients.copy()
    ov = pd.DataFrame({"cliente": list(overrides.keys()), "agente_new": list(overrides.values())})
    out = df_clients.merge(ov, on="cliente", how="left")
    out["agente"] = out["agente_new"].fillna(out["agente"])
    return out.drop(columns=["agente_new"])


def compute_agent_loads_from_clients(df_clients_2025: pd.DataFrame, dispersion_weight: float, all_agents: List[str]) -> pd.DataFrame:
    if df_clients_2025.empty:
        base = pd.DataFrame(columns=["agente", "clienti", "citta_distinte", "fatturato", "provvigioni", "load_score"])
    else:
        base = (
            df_clients_2025.groupby("agente")
            .agg(
                clienti=("cliente", "nunique"),
                citta_distinte=("citta", "nunique"),
                fatturato=("fatt_2025", "sum"),
                provvigioni=("provv_locale", "sum"),
            )
            .reset_index()
        )
        base["load_score"] = base["clienti"] + dispersion_weight * base["citta_distinte"]
        base = base.sort_values("load_score", ascending=False)

    present = set(base["agente"].tolist()) if not base.empty else set()
    missing = [a for a in all_agents if a not in present]
    if missing:
        add = pd.DataFrame(
            {
                "agente": missing,
                "clienti": [0] * len(missing),
                "citta_distinte": [0] * len(missing),
                "fatturato": [0.0] * len(missing),
                "provvigioni": [0.0] * len(missing),
            }
        )
        add["load_score"] = add["clienti"] + dispersion_weight * add["citta_distinte"]
        base = pd.concat([base, add], ignore_index=True).sort_values("load_score", ascending=False)

    return base


def simulate_reassignment(
    df_clients: pd.DataFrame,
    dispersion_weight: float,
    target_max_clienti: int,
    max_fatt_loss_pct: float,
    non_spostabili: List[str],
    extra_agents: List[str],
    prefer_same_city: bool = True,
    max_moves: int = 10_000,
) -> Dict[str, pd.DataFrame]:
    if df_clients.empty:
        return {
            "moves": pd.DataFrame(columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025", "provv_locale"]),
            "before": pd.DataFrame(),
            "after": pd.DataFrame(),
            "after_clients": pd.DataFrame(),
            "note": pd.DataFrame([{"msg": "Nessun dato 2025 disponibile."}]),
        }

    dfc = df_clients.copy()

    agent_clients = dfc.groupby("agente")["cliente"].apply(lambda x: set(x.tolist())).to_dict()
    agent_cities = dfc.groupby("agente")["citta"].apply(lambda x: set(x.tolist())).to_dict()
    agent_fatt_init = dfc.groupby("agente")["fatt_2025"].sum().to_dict()
    agent_fatt = dict(agent_fatt_init)

    # include agenti papabili (a 0)
    for a in extra_agents:
        a = str(a).strip()
        if not a:
            continue
        agent_clients.setdefault(a, set())
        agent_cities.setdefault(a, set())
        agent_fatt_init.setdefault(a, 0.0)
        agent_fatt.setdefault(a, 0.0)

    all_agents = sorted(agent_clients.keys())
    before_loads = compute_agent_loads_from_clients(dfc, dispersion_weight, all_agents)

    current_city_of = dfc.set_index("cliente")["citta"].to_dict()
    current_fatt_of = dfc.set_index("cliente")["fatt_2025"].to_dict()
    current_provv_of = dfc.set_index("cliente")["provv_locale"].to_dict()

    min_fatt_allowed = {a: agent_fatt_init.get(a, 0.0) * (1.0 - max_fatt_loss_pct) for a in all_agents}

    def agent_load_score(a: str) -> float:
        return len(agent_clients.get(a, set())) + dispersion_weight * len(agent_cities.get(a, set()))

    def rebuild_city_agents() -> Dict[str, set]:
        city_agents: Dict[str, set] = {}
        for a, cities in agent_cities.items():
            for city in cities:
                if city:
                    city_agents.setdefault(city, set()).add(a)
        return city_agents

    city_agents = rebuild_city_agents()

    overloaded = [a for a in all_agents if len(agent_clients.get(a, set())) > target_max_clienti]
    overloaded = sorted(overloaded, key=lambda a: agent_load_score(a), reverse=True)

    if not overloaded:
        return {
            "moves": pd.DataFrame(columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025", "provv_locale"]),
            "before": before_loads,
            "after": before_loads,
            "after_clients": dfc.sort_values(["agente", "citta", "fatt_2025"], ascending=[True, True, False]),
            "note": pd.DataFrame([{"msg": "Nessun agente sovraccarico con i parametri attuali."}]),
        }

    def pick_target(city: str, donor: str) -> Optional[str]:
        candidates = []
        if prefer_same_city and city:
            candidates = [a for a in city_agents.get(city, set()) if a != donor]
        if not candidates:
            candidates = [a for a in all_agents if a != donor]
        candidates = sorted(candidates, key=lambda a: agent_load_score(a))
        # preferisci chi è sotto target
        for a in candidates:
            if len(agent_clients.get(a, set())) <= target_max_clienti:
                return a
        return candidates[0] if candidates else None

    moves = []
    moves_count = 0

    for donor in overloaded:
        donor_clients = [c for c in agent_clients.get(donor, set()) if c not in non_spostabili]
        donor_clients = sorted(donor_clients, key=lambda c: float(current_fatt_of.get(c, 0.0)))  # piccoli prima

        for client in donor_clients:
            if moves_count >= max_moves:
                break
            if len(agent_clients.get(donor, set())) <= target_max_clienti:
                break

            fatt = float(current_fatt_of.get(client, 0.0))
            provv = float(current_provv_of.get(client, 0.0))
            city = str(current_city_of.get(client, "") or "")

            # vincolo perdita fatturato
            if (agent_fatt.get(donor, 0.0) - fatt) < min_fatt_allowed.get(donor, 0.0):
                continue

            target = pick_target(city, donor)
            if not target:
                continue

            # sposta
            agent_clients[donor].discard(client)
            agent_fatt[donor] = agent_fatt.get(donor, 0.0) - fatt

            agent_clients.setdefault(target, set()).add(client)
            agent_fatt[target] = agent_fatt.get(target, 0.0) + fatt

            # aggiorna città
            agent_cities[donor] = set(current_city_of.get(c, "") for c in agent_clients.get(donor, set()))
            agent_cities[target] = set(current_city_of.get(c, "") for c in agent_clients.get(target, set()))
            city_agents = rebuild_city_agents()

            moves.append({"cliente": client, "citta": city, "da_agente": donor, "a_agente": target, "fatt_2025": fatt, "provv_locale": provv})
            moves_count += 1

        if moves_count >= max_moves:
            break

    moves_df = pd.DataFrame(moves) if moves else pd.DataFrame(columns=["cliente", "citta", "da_agente", "a_agente", "fatt_2025", "provv_locale"])

    rows = []
    for a, clset in agent_clients.items():
        for c in clset:
            rows.append(
                {
                    "agente": a,
                    "citta": current_city_of.get(c, ""),
                    "cliente": c,
                    "fatt_2025": float(current_fatt_of.get(c, 0.0)),
                    "provv_locale": float(current_provv_of.get(c, 0.0)),
                }
            )
    after_clients = pd.DataFrame(rows)
    after_loads = compute_agent_loads_from_clients(after_clients, dispersion_weight, all_agents)

    note_msg = (
        f"Spostamenti: {len(moves_df)} | max_clienti={target_max_clienti} | "
        f"peso_dispersione={dispersion_weight} | max_loss_fatt={max_fatt_loss_pct:.0%}"
    )

    return {
        "moves": moves_df.sort_values("fatt_2025", ascending=True) if not moves_df.empty else moves_df,
        "before": before_loads,
        "after": after_loads,
        "after_clients": after_clients.sort_values(["agente", "citta", "fatt_2025"], ascending=[True, True, False]) if not after_clients.empty else after_clients,
        "note": pd.DataFrame([{"msg": note_msg}]),
    }


# ============================================================
# SESSION STATE
# ============================================================
st.title("📊 Analisi vendite + 🧭 Ottimizzazione area")

if "extra_agents" not in st.session_state:
    st.session_state.extra_agents = []  # non persistente (come richiesto)
if "manual_overrides" not in st.session_state:
    st.session_state.manual_overrides = {}  # {cliente: nuovo_agente}
if "manual_moves" not in st.session_state:
    st.session_state.manual_moves = []  # lista dict


# ============================================================
# SIDEBAR - FILES
# ============================================================
with st.sidebar:
    st.header("Dati vendite")
    mode_v = st.radio("Sorgente vendite", ["Upload Excel", "Locale (statisticatot25.xlsx)"], index=0)

    vend_bytes = None
    vend_path = str(DEFAULT_VENDITE_XLSX)

    if mode_v == "Upload Excel":
        up = st.file_uploader("Carica statisticatot25.xlsx", type=["xlsx"], key="vend_up")
        if up is None:
            st.info("Carica il file vendite per iniziare.")
            st.stop()
        vend_bytes = up.getvalue()
    else:
        if not DEFAULT_VENDITE_XLSX.exists():
            st.warning("Non trovo statisticatot25.xlsx accanto a app.py. Usa Upload.")
            st.stop()

    st.divider()
    st.header("Dati provvigioni (opzionale)")
    mode_p = st.radio("Sorgente provvigioni", ["Nessun file", "Upload provvigioni (.xls/.xlsx)", "Locale (Progetto_ponente_26.xls)"], index=0)

    provv_bytes = None
    provv_path = str(DEFAULT_PROVV_FILE)
    provv_enabled = mode_p != "Nessun file"

    if mode_p == "Upload provvigioni (.xls/.xlsx)":
        up2 = st.file_uploader("Carica file provvigioni", type=["xls", "xlsx"], key="provv_up")
        if up2 is not None:
            provv_bytes = up2.getvalue()
        else:
            st.warning("Se vuoi usare le provvigioni, carica il file.")
            provv_enabled = False
    elif mode_p == "Locale (Progetto_ponente_26.xls)":
        if not DEFAULT_PROVV_FILE.exists():
            st.warning("Non trovo Progetto_ponente_26.xls accanto a app.py. Usa Upload.")
            provv_enabled = False

    st.divider()
    st.header("Mapping colonne vendite")


# ============================================================
# LOAD VENDITE RAW
# ============================================================
try:
    df_raw_v = load_excel_cached(vend_bytes, vend_path)
except Exception as e:
    st.error(f"Errore lettura vendite: {e}")
    st.stop()

if df_raw_v is None or df_raw_v.empty:
    st.error("File vendite vuoto o non leggibile.")
    st.stop()

vend_cols = list(df_raw_v.columns)

guess_v = guess_map(df_raw_v, VENDITE_SYNONYMS)
saved_v = safe_read_json(COLUMN_MAP_VENDITE_FILE, default={})
colmap_v_pref = {**guess_v, **saved_v}

fixed_year = None

with st.sidebar:
    st.caption("Seleziona le colonne corrette (vendite). Le obbligatorie devono essere diverse.")

    def pick_col(label: str, keyname: str, options: List[str], current: Optional[str], optional: bool = False) -> Optional[str]:
        if optional:
            opts = ["(non impostata)"] + options
            idx = opts.index(current) if current in opts else 0
            sel = st.selectbox(label, options=opts, index=idx, key=f"map_v_{keyname}")
            return None if sel == "(non impostata)" else sel

        placeholder = "(seleziona...)"
        opts = [placeholder] + options
        idx = opts.index(current) if current in opts else 0
        sel = st.selectbox(label, options=opts, index=idx, key=f"map_v_{keyname}")
        return None if sel == placeholder else sel

    colmap_v: Dict[str, Optional[str]] = {}
    colmap_v["agente"] = pick_col("Colonna Agente", "agente", vend_cols, colmap_v_pref.get("agente"))
    colmap_v["citta"] = pick_col("Colonna Città", "citta", vend_cols, colmap_v_pref.get("citta"))
    colmap_v["cliente"] = pick_col("Colonna Cliente", "cliente", vend_cols, colmap_v_pref.get("cliente"))
    colmap_v["categoria"] = pick_col("Colonna Categoria", "categoria", vend_cols, colmap_v_pref.get("categoria"))
    colmap_v["fatturato"] = pick_col("Colonna Valore (es. acquistato 2025)", "fatturato", vend_cols, colmap_v_pref.get("fatturato"))
    colmap_v["anno"] = pick_col("Colonna Anno (opzionale)", "anno", vend_cols, colmap_v_pref.get("anno"), optional=True)
    colmap_v["articolo"] = pick_col("Colonna Articolo (opzionale)", "articolo", vend_cols, colmap_v_pref.get("articolo"), optional=True)
    colmap_v["zona"] = pick_col("Colonna Zona (opzionale)", "zona", vend_cols, colmap_v_pref.get("zona"), optional=True)

    ok, missing, has_dup = validate_map(colmap_v, REQUIRED_VENDITE)
    if not ok:
        st.error(f"Mancano colonne vendite: {', '.join(missing)}")
        st.stop()
    if has_dup:
        st.error("Hai assegnato la stessa colonna vendite a più campi. Scegli colonne diverse.")
        st.stop()

    if not colmap_v.get("anno"):
        fixed_year = infer_year_from_header(colmap_v.get("fatturato") or "")
        if fixed_year:
            st.info(f"Anno dedotto dal nome colonna valore: {fixed_year}")
        else:
            fixed_year = st.number_input("Anno fisso per la colonna valore", min_value=2000, max_value=2100, value=2025, step=1)

    if st.button("💾 Salva mapping vendite"):
        safe_write_json(COLUMN_MAP_VENDITE_FILE, colmap_v)
        st.success("Mapping vendite salvato!")


# ============================================================
# PREP VENDITE
# ============================================================
try:
    df_v = clean_and_prepare_vendite(df_raw_v, colmap_v, fixed_year=fixed_year)
except Exception as e:
    st.error(f"Errore preparazione vendite: {e}")
    st.stop()

df_v = filter_active_clients_2025(df_v)
if df_v.empty:
    st.warning("Nessun cliente con fatturato 2025 > 0 (o anno 2025 assente).")
    st.stop()


# ============================================================
# LOAD + PREP PROVV (OPZIONALE)
# ============================================================
df_provv_raw = None
df_provv = None
provv_agg = None
unmatched_codes_df = pd.DataFrame(columns=["client_code", "cliente", "citta", "agente"])

if provv_enabled:
    try:
        df_provv_raw = load_excel_cached(provv_bytes, provv_path)
    except Exception as e:
        st.warning(
            f"Non riesco a leggere il file provvigioni ({e}). "
            "Se è .xls su Streamlit Cloud, aggiungi xlrd==2.0.1 in requirements, oppure converti in .xlsx."
        )
        provv_enabled = False

if provv_enabled and df_provv_raw is not None and not df_provv_raw.empty:
    provv_cols = list(df_provv_raw.columns)
    guess_p = guess_map(df_provv_raw, PROVV_SYNONYMS)
    saved_p = safe_read_json(COLUMN_MAP_PROVV_FILE, default={})
    colmap_p_pref = {**guess_p, **saved_p}

    with st.sidebar:
        st.divider()
        st.header("Mapping colonne provvigioni")
        st.caption("Minimo richiesto: cod cliente + importo provvigione")

        def pick_col_p(label: str, keyname: str, optional: bool = False) -> Optional[str]:
            current = colmap_p_pref.get(keyname)
            if optional:
                opts = ["(non impostata)"] + provv_cols
                idx = opts.index(current) if current in opts else 0
                sel = st.selectbox(label, options=opts, index=idx, key=f"map_p_{keyname}")
                return None if sel == "(non impostata)" else sel
            placeholder = "(seleziona...)"
            opts = [placeholder] + provv_cols
            idx = opts.index(current) if current in opts else 0
            sel = st.selectbox(label, options=opts, index=idx, key=f"map_p_{keyname}")
            return None if sel == placeholder else sel

        colmap_p: Dict[str, Optional[str]] = {}
        colmap_p["cod_cliente"] = pick_col_p("Colonna Cod Cliente", "cod_cliente")
        colmap_p["provvigione"] = pick_col_p("Colonna Importo Provvigione", "provvigione")
        colmap_p["fatt_provv"] = pick_col_p("Colonna Fatturato per calcolo provvigioni (opz.)", "fatt_provv", optional=True)
        colmap_p["agente"] = pick_col_p("Colonna Agente (opz.)", "agente", optional=True)
        colmap_p["cod_agente"] = pick_col_p("Colonna Cod Agente (opz.)", "cod_agente", optional=True)
        colmap_p["cliente_nome"] = pick_col_p("Colonna Nominativo cliente (opz.)", "cliente_nome", optional=True)
        colmap_p["citta"] = pick_col_p("Colonna Città (opz.)", "citta", optional=True)
        colmap_p["prov"] = pick_col_p("Colonna Provincia (opz.)", "prov", optional=True)
        colmap_p["indirizzo"] = pick_col_p("Colonna Indirizzo (opz.)", "indirizzo", optional=True)
        colmap_p["cap"] = pick_col_p("Colonna CAP (opz.)", "cap", optional=True)

        okp, missp, dupp = validate_map(colmap_p, REQUIRED_PROVV)
        if not okp:
            st.error(f"Mancano colonne provvigioni: {', '.join(missp)}")
            provv_enabled = False
        if dupp:
            st.error("Hai assegnato la stessa colonna provvigioni a più campi (nelle obbligatorie).")
            provv_enabled = False

        if st.button("💾 Salva mapping provvigioni"):
            safe_write_json(COLUMN_MAP_PROVV_FILE, colmap_p)
            st.success("Mapping provvigioni salvato!")

    if provv_enabled:
        try:
            df_provv = clean_and_prepare_provv(df_provv_raw, colmap_p)
            provv_agg = aggregate_provv_by_cliente(df_provv)
        except Exception as e:
            st.warning(f"Problema preparazione provvigioni: {e}")
            provv_enabled = False

# ============================================================
# MERGE PROVV -> VENDITE (solo se disponibile)
# ============================================================
include_provv = provv_enabled and (provv_agg is not None) and (not provv_agg.empty)

if include_provv:
    # aggiungo cod cliente alle vendite
    df_v = df_v.copy()
    df_v["client_code"] = df_v["cliente"].apply(extract_client_code)

    df_v = df_v.merge(provv_agg, left_on="client_code", right_on="cod_cliente", how="left")
    df_v["provv_locale"] = df_v["provv_locale"].fillna(0.0)
    df_v["fatt_provv_locale"] = df_v["fatt_provv_locale"].fillna(0.0)
    df_v = df_v.drop(columns=["cod_cliente"], errors="ignore")

    # unmatched per debug/controllo
    unmatched = df_v[(df_v["client_code"].astype(str).str.strip() != "") & (df_v["provv_locale"] == 0.0)].copy()
    # (attenzione: può essere 0 anche se match reale ma provv 0; qui è solo segnalazione)
    unmatched_codes_df = unmatched[["client_code", "cliente", "citta", "agente"]].drop_duplicates().head(200)


# ============================================================
# SIDEBAR - FILTRI REPORT
# ============================================================
with st.sidebar:
    st.divider()
    st.header("Filtri report")

    anni = sorted([a for a in df_v["anno"].unique().tolist() if a != 0])
    default_year = 2025 if 2025 in anni else (anni[-1] if anni else 2025)
    year_sel = st.selectbox("Anno", options=anni if anni else [2025], index=(anni.index(default_year) if anni and default_year in anni else 0))

    agents_all_base = sorted(df_v["agente"].unique().tolist())
    cities_all = sorted(df_v["citta"].unique().tolist())
    cats_all = sorted(df_v["categoria"].unique().tolist())

    filt_agents = st.multiselect("Agenti", options=agents_all_base, default=[])
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


df_year = apply_filters(df_v)


# ============================================================
# TABS
# ============================================================
tab_report, tab_opt, tab_data = st.tabs(["📈 Report", "🧭 Ottimizzazione area", "🧾 Anteprima dati"])


# =========================
# TAB REPORT
# =========================
with tab_report:
    st.subheader(f"Report anno {year_sel} (solo clienti con totale 2025 > 0)")

    st.markdown("### Città: fatturato + numero clienti")
    show_agents_in_city = st.checkbox("Mostra anche agenti per città (opzionale)", value=False)
    show_provv_in_city = st.checkbox("Mostra provvigioni per città (opzionale)", value=include_provv, disabled=not include_provv)

    rep_city = report_city_summary(df_year, include_agents=show_agents_in_city, include_provv=show_provv_in_city)
    fmt = {"fatturato": "{:,.2f}"}
    if show_provv_in_city and "provvigioni" in rep_city.columns:
        fmt["provvigioni"] = "{:,.2f}"

    st.dataframe(rep_city.style.format(fmt), use_container_width=True, hide_index=True)

    if not rep_city.empty:
        st.bar_chart(rep_city.set_index("citta")["fatturato"])

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Fatturato per categoria")
        rep_cat = report_fatturato_per_categoria(df_year)
        st.dataframe(rep_cat, use_container_width=True, hide_index=True)
        if not rep_cat.empty:
            st.bar_chart(rep_cat.set_index("categoria")["fatturato"])

    with c2:
        st.markdown("### Fatturato per agente (solo 2025 + clienti attivi + % incidenza)")
        rep_agent_2025 = report_fatturato_per_agente_2025(df_v, include_provv=include_provv)
        fmt2 = {"fatturato": "{:,.2f}", "%_incidenza": "{:.2f}"}
        if include_provv and "provvigioni" in rep_agent_2025.columns:
            fmt2["provvigioni"] = "{:,.2f}"
        st.dataframe(rep_agent_2025.style.format(fmt2), use_container_width=True, hide_index=True)

    st.divider()

    st.markdown("### Fatturato per zona (agente → città → cliente)")
    rep_zone = report_agente_citta_cliente(df_year, include_provv=include_provv)
    fmt3 = {"fatturato": "{:,.2f}"}
    if include_provv and "provv_locale" in rep_zone.columns:
        fmt3["provv_locale"] = "{:,.2f}"
    st.dataframe(rep_zone.style.format(fmt3), use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("### Export report (Excel)")

    sheets = {
        f"citta_{year_sel}": rep_city,
        f"fatturato_categoria_{year_sel}": rep_cat,
        "fatturato_agente_2025": rep_agent_2025,
        f"agente_citta_cliente_{year_sel}": rep_zone,
    }
    if include_provv:
        sheets["unmatched_codes_preview"] = unmatched_codes_df

    report_bytes = to_excel_bytes(sheets)

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
    st.subheader("Ottimizzazione area: automatica + manuale + agenti papabili (solo in app)")

    df_clients_base = build_client_table_2025(df_v)
    df_clients_current = apply_manual_overrides(df_clients_base, st.session_state.manual_overrides)

    agents_from_file = sorted(df_clients_base["agente"].unique().tolist()) if not df_clients_base.empty else []
    agents_from_overrides = sorted(list(set(st.session_state.manual_overrides.values()))) if st.session_state.manual_overrides else []
    all_agents = sorted(list(set(agents_from_file + st.session_state.extra_agents + agents_from_overrides)))

    # -------------------------
    # AGENTI PAPABILI
    # -------------------------
    st.markdown("### 👤 Agenti papabili (solo in app)")
    colA, colB, colC = st.columns([1.2, 0.8, 1.5])
    with colA:
        new_agent_name = st.text_input("Nome nuovo agente", value="", placeholder="Es. Nuovo Agente 1")
    with colB:
        if st.button("➕ Aggiungi"):
            name = (new_agent_name or "").strip()
            if not name:
                st.warning("Inserisci un nome agente.")
            else:
                low_all = {a.strip().lower() for a in all_agents}
                if name.lower() in low_all:
                    st.info("Agente già presente (nel file o già aggiunto).")
                else:
                    st.session_state.extra_agents.append(name)
                    st.success(f"Aggiunto: {name}")
                    st.rerun()
    with colC:
        if st.session_state.extra_agents:
            to_remove = st.selectbox("Rimuovi agente papabile", options=["(nessuno)"] + st.session_state.extra_agents)
            if st.button("🗑️ Rimuovi"):
                if to_remove != "(nessuno)":
                    st.session_state.extra_agents = [a for a in st.session_state.extra_agents if a != to_remove]
                    if st.session_state.manual_overrides:
                        st.session_state.manual_overrides = {k: v for k, v in st.session_state.manual_overrides.items() if v != to_remove}
                    st.success(f"Rimosso: {to_remove}")
                    st.rerun()
        else:
            st.caption("Nessun agente papabile aggiunto.")

    st.divider()

    sub_auto, sub_manual = st.tabs(["🤖 Automatica", "🖐️ Manuale"])

    # ---------- AUTOMATICA ----------
    with sub_auto:
        left, right = st.columns([1, 1])

        with left:
            st.markdown("#### Parametri automatica")
            target_max_clienti = st.number_input("Target max clienti per agente", min_value=50, max_value=300, value=140, step=5)
            dispersion_weight = st.number_input("Peso dispersione (numero città)", min_value=0.0, max_value=10.0, value=1.0, step=0.1)
            max_fatt_loss_pct = st.slider("Max perdita fatturato per agente (donatore)", min_value=0, max_value=50, value=15, step=1) / 100.0
            prefer_same_city = st.checkbox("Preferisci assegnazione nella stessa città", value=True)

            st.markdown("#### Clienti NON spostabili")
            all_clients = sorted(df_clients_base["cliente"].unique().tolist()) if not df_clients_base.empty else []
            saved_nonmov = safe_read_json(NON_MOVABLE_FILE, default=[])
            default_nonmov = [c for c in saved_nonmov if c in all_clients]
            non_movable_sel = st.multiselect("Lista non spostabili (persistente)", options=all_clients, default=default_nonmov)

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

            st.divider()
            lock_manual = st.checkbox("Non toccare i clienti spostati a mano (consigliato)", value=True)

        with right:
            st.markdown("#### Carichi attuali (prima) - includendo provvigioni")
            loads0 = compute_agent_loads_from_clients(df_clients_current, dispersion_weight, all_agents)
            st.dataframe(
                loads0.style.format({"fatturato": "{:,.2f}", "provvigioni": "{:,.2f}", "load_score": "{:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()
        run = st.button("🚀 Esegui simulazione automatica (prima/dopo)")

        if run:
            non_spost = list(set(non_movable_sel))
            if lock_manual and st.session_state.manual_overrides:
                non_spost = list(set(non_spost + list(st.session_state.manual_overrides.keys())))

            with st.spinner("Simulazione in corso..."):
                sim = simulate_reassignment(
                    df_clients=df_clients_current,
                    dispersion_weight=float(dispersion_weight),
                    target_max_clienti=int(target_max_clienti),
                    max_fatt_loss_pct=float(max_fatt_loss_pct),
                    non_spostabili=non_spost,
                    extra_agents=st.session_state.extra_agents,
                    prefer_same_city=prefer_same_city,
                )

            st.success("Simulazione completata.")
            st.info(sim["note"].iloc[0]["msg"] if not sim["note"].empty else "OK")

            st.markdown("### Spostamenti suggeriti (clienti piccoli prima)")
            st.dataframe(sim["moves"].style.format({"fatt_2025": "{:,.2f}", "provv_locale": "{:,.2f}"}), use_container_width=True, hide_index=True)

            cA, cB = st.columns(2)
            with cA:
                st.markdown("### Prima")
                st.dataframe(sim["before"].style.format({"fatturato": "{:,.2f}", "provvigioni": "{:,.2f}", "load_score": "{:,.2f}"}), use_container_width=True, hide_index=True)
            with cB:
                st.markdown("### Dopo")
                st.dataframe(sim["after"].style.format({"fatturato": "{:,.2f}", "provvigioni": "{:,.2f}", "load_score": "{:,.2f}"}), use_container_width=True, hide_index=True)

            st.markdown("### Export simulazione automatica (Excel)")
            sim_bytes = to_excel_bytes(
                {
                    "spostamenti_auto": sim["moves"],
                    "carichi_prima": sim["before"],
                    "carichi_dopo": sim["after"],
                    "assegnazioni_dopo": sim.get("after_clients", pd.DataFrame()),
                }
            )
            st.download_button(
                "⬇️ Scarica Simulazione Excel (automatica)",
                data=sim_bytes,
                file_name=f"simulazione_automatica_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()
        st.markdown("#### Base clienti 2025 (dopo eventuali spostamenti manuali)")
        st.dataframe(
            df_clients_current.sort_values("fatt_2025", ascending=False),
            use_container_width=True,
            hide_index=True,
        )

    # ---------- MANUALE ----------
    with sub_manual:
        st.subheader("Spostamenti manuali (città → agenti → locali → sposta)")

        if df_clients_current.empty:
            st.info("Nessun dato clienti 2025 disponibile.")
            st.stop()

        cP1, cP2 = st.columns(2)
        with cP1:
            dispersion_weight_m = st.number_input("Peso dispersione (manuale)", min_value=0.0, max_value=10.0, value=1.0, step=0.1)
        with cP2:
            max_fatt_loss_pct_m = st.slider("Max perdita fatturato donatore (manuale)", min_value=0, max_value=50, value=15, step=1) / 100.0

        ignore_fatt_vincolo = st.checkbox("Ignora vincolo fatturato donatore (manuale)", value=False)
        ignore_dispersion = st.checkbox("Ignora dispersione (manuale) (usa solo clienti)", value=False)
        ignore_non_spostabili = st.checkbox("Ignora lista NON spostabili (manuale)", value=False)

        non_spostabili = safe_read_json(NON_MOVABLE_FILE, default=[])

        left, right = st.columns([1, 1])

        with left:
            city_list = sorted(df_clients_current["citta"].dropna().unique().tolist())
            city_sel = st.selectbox("Seleziona Città", options=city_list)

            df_city = df_clients_current[df_clients_current["citta"] == city_sel].copy()
            src_agent_options = sorted(df_city["agente"].dropna().unique().tolist())
            if not src_agent_options:
                st.warning("In questa città non ci sono clienti assegnati.")
                st.stop()

            src_agent = st.selectbox("Agente (sorgente)", options=src_agent_options)

            df_src = df_city[df_city["agente"] == src_agent].copy().sort_values("fatt_2025", ascending=True)

            search = st.text_input("Cerca locale (opzionale)", value="")
            if search.strip():
                df_src_view = df_src[df_src["cliente"].str.contains(search, case=False, na=False)].copy()
            else:
                df_src_view = df_src

            st.caption("Locali dell’agente in questa città (2025) - con provvigioni")
            st.dataframe(
                df_src_view[["cliente", "fatt_2025", "provv_locale"]].style.format({"fatt_2025": "{:,.2f}", "provv_locale": "{:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )

            clients_src = df_src_view["cliente"].tolist()
            sel_clients = st.multiselect("Seleziona locale/i da spostare", options=clients_src)

        with right:
            only_same_city = st.checkbox("Destinazione solo agenti della stessa città", value=True)

            agents_in_city = sorted(df_city["agente"].dropna().unique().tolist())

            if st.session_state.extra_agents:
                include_papabili = st.checkbox("Includi agenti papabili", value=True)
                if include_papabili:
                    for a in st.session_state.extra_agents:
                        if a not in agents_in_city:
                            agents_in_city.append(a)
                    agents_in_city = sorted(list(set(agents_in_city)))

            if only_same_city:
                tgt_options = [a for a in agents_in_city if a != src_agent]
            else:
                tgt_options = [a for a in all_agents if a != src_agent]

            tgt_options = sorted(list(dict.fromkeys(tgt_options)))
            tgt_agent = st.selectbox("Agente (destinazione)", options=tgt_options if tgt_options else ["(nessuno)"])

            st.markdown("### Carichi attuali (prima) - includendo provvigioni")
            dw = 0.0 if ignore_dispersion else dispersion_weight_m
            loads_before = compute_agent_loads_from_clients(df_clients_current, dw, all_agents)
            st.dataframe(
                loads_before.style.format({"fatturato": "{:,.2f}", "provvigioni": "{:,.2f}", "load_score": "{:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        move_all = st.button("🚚 Sposta TUTTI i locali dell’agente sorgente (rispetta ricerca)", disabled=(tgt_agent == "(nessuno)"))
        if move_all:
            sel_clients = df_src_view["cliente"].tolist()

        do_move = st.button("➡️ Sposta selezionati", disabled=(not sel_clients or tgt_agent == "(nessuno)"))

        if do_move:
            fatt_by_agent = df_clients_current.groupby("agente")["fatt_2025"].sum().to_dict()
            provv_by_agent = df_clients_current.groupby("agente")["provv_locale"].sum().to_dict()

            min_allowed = {a: fatt_by_agent.get(a, 0.0) * (1.0 - max_fatt_loss_pct_m) for a in all_agents}

            fatt_lookup = df_clients_current.set_index("cliente")["fatt_2025"].to_dict()
            provv_lookup = df_clients_current.set_index("cliente")["provv_locale"].to_dict()
            city_lookup = df_clients_current.set_index("cliente")["citta"].to_dict()
            agent_lookup = df_clients_current.set_index("cliente")["agente"].to_dict()

            blocked = []
            moved = 0

            for cl in sel_clients:
                if (not ignore_non_spostabili) and (cl in non_spostabili):
                    blocked.append((cl, "non spostabile"))
                    continue

                current_agent = agent_lookup.get(cl, src_agent)
                fatt = float(fatt_lookup.get(cl, 0.0))
                provv = float(provv_lookup.get(cl, 0.0))

                if not ignore_fatt_vincolo:
                    if (fatt_by_agent.get(current_agent, 0.0) - fatt) < min_allowed.get(current_agent, 0.0):
                        blocked.append((cl, "vincolo fatturato donatore"))
                        continue

                # snapshot prima/dopo per i due agenti coinvolti
                donor_f_before = fatt_by_agent.get(current_agent, 0.0)
                donor_p_before = provv_by_agent.get(current_agent, 0.0)
                dest_f_before = fatt_by_agent.get(tgt_agent, 0.0)
                dest_p_before = provv_by_agent.get(tgt_agent, 0.0)

                donor_f_after = donor_f_before - fatt
                donor_p_after = donor_p_before - provv
                dest_f_after = dest_f_before + fatt
                dest_p_after = dest_p_before + provv

                st.session_state.manual_overrides[cl] = tgt_agent
                st.session_state.manual_moves.append(
                    {
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "citta": city_lookup.get(cl, city_sel),
                        "cliente": cl,
                        "da_agente": current_agent,
                        "a_agente": tgt_agent,
                        "fatt_2025": fatt,
                        "provv_locale": provv,
                        "donatore_fatt_prima": donor_f_before,
                        "donatore_fatt_dopo": donor_f_after,
                        "donatore_provv_prima": donor_p_before,
                        "donatore_provv_dopo": donor_p_after,
                        "dest_fatt_prima": dest_f_before,
                        "dest_fatt_dopo": dest_f_after,
                        "dest_provv_prima": dest_p_before,
                        "dest_provv_dopo": dest_p_after,
                    }
                )

                # aggiorna contatori locali per vincoli successivi nello stesso click
                fatt_by_agent[current_agent] = donor_f_after
                provv_by_agent[current_agent] = donor_p_after
                fatt_by_agent[tgt_agent] = dest_f_after
                provv_by_agent[tgt_agent] = dest_p_after

                moved += 1

            if moved:
                st.success(f"Spostati {moved} locale/i.")
            if blocked:
                st.warning("Alcuni locali non sono stati spostati:")
                st.dataframe(pd.DataFrame(blocked, columns=["cliente", "motivo"]), use_container_width=True, hide_index=True)

            st.rerun()

        st.divider()

        moves_df = pd.DataFrame(st.session_state.manual_moves)
        df_after_manual = apply_manual_overrides(df_clients_base, st.session_state.manual_overrides)

        agents_from_overrides_now = sorted(list(set(st.session_state.manual_overrides.values()))) if st.session_state.manual_overrides else []
        all_agents_now = sorted(list(set(agents_from_file + st.session_state.extra_agents + agents_from_overrides_now)))

        dw2 = 0.0 if ignore_dispersion else dispersion_weight_m
        loads_after = compute_agent_loads_from_clients(df_after_manual, dw2, all_agents_now)

        cA, cB = st.columns([1, 1])
        with cA:
            st.markdown("### Movimenti manuali registrati (con provvigioni prima/dopo)")
            if moves_df.empty:
                st.info("Nessun movimento manuale ancora.")
            else:
                st.dataframe(
                    moves_df.sort_values("timestamp", ascending=False).style.format(
                        {
                            "fatt_2025": "{:,.2f}",
                            "provv_locale": "{:,.2f}",
                            "donatore_provv_prima": "{:,.2f}",
                            "donatore_provv_dopo": "{:,.2f}",
                            "dest_provv_prima": "{:,.2f}",
                            "dest_provv_dopo": "{:,.2f}",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

        with cB:
            st.markdown("### Carichi dopo movimenti manuali (includendo provvigioni)")
            st.dataframe(
                loads_after.style.format({"fatturato": "{:,.2f}", "provvigioni": "{:,.2f}", "load_score": "{:,.2f}"}),
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("### Export manuale (Excel) - unico file")
        export_bytes = to_excel_bytes(
            {
                "movimenti_manuali": moves_df if not moves_df.empty else pd.DataFrame(columns=[
                    "timestamp","citta","cliente","da_agente","a_agente","fatt_2025","provv_locale",
                    "donatore_fatt_prima","donatore_fatt_dopo","donatore_provv_prima","donatore_provv_dopo",
                    "dest_fatt_prima","dest_fatt_dopo","dest_provv_prima","dest_provv_dopo"
                ]),
                "clienti_dopo_manuale": df_after_manual.sort_values(["agente", "citta", "fatt_2025"], ascending=[True, True, False]) if not df_after_manual.empty else df_after_manual,
                "carichi_prima": loads_before,
                "carichi_dopo": loads_after,
                "agenti_papabili": pd.DataFrame({"agente_papabile": st.session_state.extra_agents}),
            }
        )

        st.download_button(
            "⬇️ Scarica Excel (manuale)",
            data=export_bytes,
            file_name=f"spostamenti_manuali_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.divider()
        if st.button("🧨 Reset movimenti manuali (torna al file originale)"):
            st.session_state.manual_overrides = {}
            st.session_state.manual_moves = []
            st.success("Reset fatto.")
            st.rerun()


# =========================
# TAB DATA
# =========================
with tab_data:
    st.subheader("Anteprima dati puliti (Totali esclusi + solo clienti 2025>0)")
    st.write(f"Righe vendite: {len(df_v):,} | Provvigioni attive: {'Sì' if include_provv else 'No'}")
    st.dataframe(df_v.head(300), use_container_width=True, hide_index=True)

    if include_provv:
        st.markdown("### Preview codici senza match provvigioni (campione)")
        st.dataframe(unmatched_codes_df, use_container_width=True, hide_index=True)
