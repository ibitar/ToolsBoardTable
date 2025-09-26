# app.py
# Tableau de bord OCS/OSS – Streamlit
# Dépendances : pip install streamlit pandas openpyxl matplotlib

import re
import unicodedata
from pathlib import Path
from typing import Dict, Tuple, Optional, List
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ==========
# 1) Config
# ==========
DEFAULT_PATH = r"C:\Users\i.bitar\OneDrive - EGIS Group\Documents\Professionnel\EPR2\DEVELOPEMENT-OCS-OSS\OCS-OSS\2. Feuilles Excel\OCS_OSS_2025_TABLEAU_DE_BORD.xlsm"
SHEET = "Liste_OCS_OSS"

# Motifs robustes pour détecter les entêtes (insensibles casse/typos)
PATTERNS = {
    "logiciel": r"\blogiciel\b",
    "date_maj_info": r"derni[eè]re\s+m(?:ise)?\s*à\s*jour",
    "date_ajout": r"date\s+de\s+rajout.*excel",
    "date_debut": r"date\s+de\s+d[ée]but.*qualifi",
    "date_fin": r"date\s+de\s+fin.*qualifi",
}

# Abréviations FR des mois (indexés 1..12)
MOIS_ABBR_FR = ["", "Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Aoû", "Sep", "Oct", "Nov", "Déc"]

DOC_STATUS_COLUMNS = [
    "etat de la notice de description",
    "etat de la notice de validation",
    "etat de la notice de vérification",
    "etat du pv recettage",
]
REVISION_STALENESS_DAYS = 30
STATUS_EMPTY_COLOR = "#ffcccc"
REVISION_STALE_COLOR = "#ffe0b2"

# ==========
# 2) Utilitaires
# ==========
def _detect_header_row(df_nohdr: pd.DataFrame) -> int:
    """Cherche la ligne d'entêtes en repérant la colonne 'Logiciel' (ou équivalent)."""
    for i in range(min(30, len(df_nohdr))):
        row = df_nohdr.iloc[i].astype(str).str.strip().str.lower()
        if row.str.contains(PATTERNS["logiciel"], regex=True, na=False).any():
            return i
    return 0

def _normalize_headers(headers) -> pd.Index:
    """Nettoie les entêtes: trim + espaces multiples -> simple + lower."""
    h = (
        pd.Series(headers)
        .astype(str).str.replace(r"\s+", " ", regex=True)
        .str.strip().str.lower()
    )
    return pd.Index(h)

def _find_columns(cols: pd.Index) -> Dict[str, str]:
    """Associe les noms effectifs de colonnes aux clés du PATTERNS."""
    found: Dict[str, str] = {}
    for key, pat in PATTERNS.items():
        mask = cols.to_series().str.contains(pat, regex=True, na=False)
        if mask.any():
            found[key] = cols[mask][0]
    requis = {"logiciel", "date_debut", "date_fin", "date_ajout"}
    if not requis.issubset(found.keys()):
        manquantes = sorted(list(requis - set(found.keys())))
        raise ValueError(f"Colonnes non trouvées: {manquantes}\nColonnes disponibles: {list(cols)}")
    return found

@st.cache_data(show_spinner=False)
def lire_table(path_or_buffer, sheet: str = SHEET) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Lit l’onglet, détecte entêtes, convertit dates, renvoie (df, mapping_colonnes)."""
    raw = pd.read_excel(path_or_buffer, sheet_name=sheet, header=None, engine="openpyxl")
    hdr_row = _detect_header_row(raw)
    headers = _normalize_headers(raw.iloc[hdr_row])
    df = raw.iloc[hdr_row + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    mapping = _find_columns(df.columns)
    # conversions dates
    for k in ("date_ajout", "date_debut", "date_fin", "date_maj_info"):
        if k in mapping:
            df[mapping[k]] = pd.to_datetime(df[mapping[k]], errors="coerce")
    # nettoyer libellés
    df[mapping["logiciel"]] = df[mapping["logiciel"]].astype(str).str.strip()
    return df, mapping

def _by_year(s: pd.Series, year: int) -> pd.Series:
    return (s.notna()) & (s.dt.year == year)

def kpis_globaux(df: pd.DataFrame, m: Dict[str, str], year: int) -> Dict[str, int]:
    return {
        "Ajoutés au fichier": int(_by_year(df[m["date_ajout"]], year).sum()),
        "Démarrés": int(_by_year(df[m["date_debut"]], year).sum()),
        "Finis": int(_by_year(df[m["date_fin"]], year).sum()),
        "En cours (démarrés non finis)": int((_by_year(df[m["date_debut"]], year) & (~df[m["date_fin"]].notna())).sum()),
    }

def stats_par_logiciel(df: pd.DataFrame, m: Dict[str, str], year: int) -> pd.DataFrame:
    g = df.groupby(m["logiciel"], dropna=False)
    out = pd.DataFrame({
        "Démarrés": g[m["date_debut"]].apply(lambda s: _by_year(s, year).sum()),
        "Ajoutés au fichier": g[m["date_ajout"]].apply(lambda s: _by_year(s, year).sum()),
        "Finis": g[m["date_fin"]].apply(lambda s: _by_year(s, year).sum()),
    })
    out["En cours (démarrés non finis)"] = g.apply(lambda x: (_by_year(x[m["date_debut"]], year) & (~x[m["date_fin"]].notna())).sum())
    out = (
        out.sort_values(["Démarrés","Finis","Ajoutés au fichier"], ascending=False)
           .reset_index()
           .rename(columns={m["logiciel"]:"Logiciel"})
    )
    return out

def stats_mensuelles(df: pd.DataFrame, m: Dict[str, str], year: int, champ_key: str) -> pd.DataFrame:
    """Renvoie colonnes: Mois (1..12), Libellé (FR), Valeur."""
    s = df[m[champ_key]]
    mois = s.loc[_by_year(s, year)].dt.month
    counts = mois.value_counts().sort_index()
    idx = pd.Index(range(1,13), name="Mois")
    counts = counts.reindex(idx, fill_value=0)
    valeur_col = {"date_ajout":"Ajouts", "date_debut":"Démarrages", "date_fin":"Fins"}[champ_key]
    out = counts.reset_index(name=valeur_col)
    out["Libellé"] = out["Mois"].apply(lambda i: MOIS_ABBR_FR[i])
    return out  # colonnes: Mois, valeur_col, Libellé


def doc_status_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Synthèse des statuts documentaires (OK / En attente / NaN)."""
    lignes: List[Dict[str, int]] = []
    for col in DOC_STATUS_COLUMNS:
        if col not in df.columns:
            continue
        serie = df[col]
        normalisee = (
            serie.dropna()
            .astype(str)
            .str.strip()
            .str.lower()
        )
        lignes.append(
            {
                "Matrice": col,
                "OK": int((normalisee == "ok").sum()),
                "En attente": int((normalisee == "en attente").sum()),
                "NaN": int(serie.isna().sum()),
            }
        )
    return pd.DataFrame(lignes)


def _style_doc_matrix(df: pd.DataFrame, doc_cols: List[str], revision_cols: List[str], delai_jours: int):
    if not doc_cols and not revision_cols:
        return df

    matrice = df[doc_cols + revision_cols].copy()

    for col in revision_cols:
        matrice[col] = pd.to_datetime(matrice[col], errors="coerce")

    styler = matrice.style

    if doc_cols:
        styler = styler.applymap(
            lambda val: f"background-color: {STATUS_EMPTY_COLOR}" if pd.isna(val) or str(val).strip() == "" else "",
            subset=doc_cols,
        )

    if revision_cols:
        seuil = pd.Timestamp.now() - pd.Timedelta(days=delai_jours)
        styler = styler.format(
            {
                col: (lambda v: v.strftime("%d/%m/%Y") if pd.notna(v) else "")
                for col in revision_cols
            }
        )

        def _couleur_revision(val):
            if pd.isna(val):
                return f"background-color: {STATUS_EMPTY_COLOR}"
            if val < seuil:
                return f"background-color: {REVISION_STALE_COLOR}"
            return ""

        styler = styler.applymap(_couleur_revision, subset=revision_cols)

    return styler


def _normalize_text(text: str) -> str:
    """Supprime les accents et met en minuscules pour une comparaison robuste."""
    if text is None:
        return ""
    normalized = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).lower().strip()


def stats_referent(df: pd.DataFrame) -> pd.DataFrame:
    """Compte les dossiers par référent et sépare En cours / À faire."""
    if df.empty:
        return pd.DataFrame(columns=["Référent", "À faire", "En cours", "Dossiers ouverts"])

    referent_col = None
    for col in df.columns:
        if "referent" in _normalize_text(col):
            referent_col = col
            break
    if referent_col is None:
        raise ValueError("Aucune colonne référent détectée dans les données.")

    status_col = None
    for col in df.columns:
        norm_col = _normalize_text(col)
        if "action" in norm_col and "faire" in norm_col:
            status_col = col
            break
    if status_col is None:
        for col in df.columns:
            if "avancement" in _normalize_text(col):
                status_col = col
                break

    referents = df[referent_col].fillna("Non renseigné").astype(str).str.strip()
    referents = referents.replace("", "Non renseigné")

    if status_col is not None:
        status_values = df[status_col].fillna("").apply(_normalize_text)
    else:
        status_values = pd.Series(["en cours"] * len(df), index=df.index)

    statut = pd.Series("En cours", index=df.index)
    statut.loc[status_values.str.contains("a faire", regex=False)] = "À faire"
    statut.loc[status_values.str.contains("en cours", regex=False)] = "En cours"

    stats = (
        pd.crosstab(referents, statut)
        .reindex(columns=["À faire", "En cours"], fill_value=0)
        .rename_axis("Référent")
        .reset_index()
    )
    stats["À faire"] = stats.get("À faire", 0)
    stats["En cours"] = stats.get("En cours", 0)
    stats["Dossiers ouverts"] = stats["À faire"] + stats["En cours"]
    stats = stats.sort_values("Dossiers ouverts", ascending=False).reset_index(drop=True)
    return stats

def bar_plot(df: pd.DataFrame, x: str, y: str, title: str, ylim: Optional[Tuple[float, float]]=None):
    """Trace un histogramme (Matplotlib) avec abscisses textuelles et y commun optionnel."""
    fig, ax = plt.subplots()
    ax.bar(df[x], df[y])
    ax.set_xlabel("Mois")
    ax.set_ylabel(y)
    ax.set_title(title)
    if ylim is not None:
        ax.set_ylim(ylim)
    ax.grid(True, axis="y")
    # Rotation nulle (éventuellement 45 si labels longs)
    for tick in ax.get_xticklabels():
        tick.set_rotation(0)
    st.pyplot(fig, clear_figure=True)

# ==========
# 3) UI
# ==========
st.set_page_config(page_title="OCS/OSS – Tableau de bord", layout="wide")
st.title("Tableau de bord OCS/OSS")

with st.sidebar:
    st.header("Source de données")
    choix = st.radio("Choisir la source", ["Chemin par défaut", "Uploader un fichier"])
    if choix == "Chemin par défaut":
        path = st.text_input("Chemin XLSM", value=DEFAULT_PATH)
        file_ok = Path(path).exists()
        if not file_ok:
            st.warning("Le fichier par défaut est introuvable. Vérifie le chemin ou uploade un fichier.")
    else:
        up = st.file_uploader("Uploader le fichier Excel/XLSM", type=["xlsx","xlsm"])
        path = up

    btn_load = st.button("Charger / Recharger les données", type="primary")

# Chargement
if btn_load or "df" not in st.session_state:
    if isinstance(path, (str, Path)) and path and Path(path).exists():
        df, mapping = lire_table(str(path), SHEET)
        st.session_state["df"] = df
        st.session_state["mapping"] = mapping
    elif path is not None:  # uploaded buffer
        df, mapping = lire_table(path, SHEET)
        st.session_state["df"] = df
        st.session_state["mapping"] = mapping

if "df" not in st.session_state:
    st.stop()

df = st.session_state["df"]
m = st.session_state["mapping"]
st.success("Données chargées ✅")

# Choix de l'année (auto-remplie depuis les colonnes dates)
years = set()
for key in ("date_ajout","date_debut","date_fin"):
    if key in m:
        years |= set(df[m[key]].dropna().dt.year.unique())
years = sorted([int(y) for y in years if pd.notna(y)])
default_year = 2025 if 2025 in years else (years[-1] if years else 2025)

col_y1, col_y2 = st.columns([1,3])
with col_y1:
    year = st.number_input("Année d'analyse", min_value=1900, max_value=2100, value=int(default_year), step=1)
with col_y2:
    st.caption(f"Colonnes détectées: {m}")

# KPIs
k = kpis_globaux(df, m, year)
c1,c2,c3,c4 = st.columns(4)
c1.metric("Ajoutés au fichier", k["Ajoutés au fichier"])
c2.metric("Démarrés", k["Démarrés"])
c3.metric("Finis", k["Finis"])
c4.metric("En cours (démarrés non finis)", k["En cours (démarrés non finis)"])

# Répartition par logiciel / référent
tab_logiciels, tab_referents = st.tabs(["Logiciels", "Référents"])

with tab_logiciels:
    st.subheader("Répartition par logiciel")
    tab_logiciel = stats_par_logiciel(df, m, year)
    st.dataframe(tab_logiciel, use_container_width=True)
    st.download_button(
        "⬇️ Export CSV – Répartition par logiciel",
        tab_logiciel.to_csv(index=False).encode("utf-8"),
        file_name=f"repartition_logiciel_{year}.csv",
    )

with tab_referents:
    st.subheader("Charge par référent")
    try:
        tab_ref = stats_referent(df)
    except ValueError as exc:
        st.info(str(exc))
        tab_ref = pd.DataFrame(columns=["Référent", "À faire", "En cours", "Dossiers ouverts"])

    if tab_ref.empty:
        st.info("Aucune donnée référent disponible.")
    else:
        column_config = {
            "Référent": st.column_config.TextColumn(
                "Référent",
                help="Consultez la colonne 'Commentaire/action' pour analyser les référents surchargés.",
            ),
            "À faire": st.column_config.NumberColumn("À faire"),
            "En cours": st.column_config.NumberColumn("En cours"),
            "Dossiers ouverts": st.column_config.NumberColumn(
                "Dossiers ouverts",
                help="Les référents avec beaucoup de dossiers ouverts doivent être suivis via 'Commentaire/action'.",
            ),
        }
        st.dataframe(tab_ref, use_container_width=True, column_config=column_config)
        chart_data = tab_ref.set_index("Référent")[["Dossiers ouverts"]]
        st.bar_chart(chart_data, height=400, use_container_width=True, horizontal=True)

st.divider()

# Suivi documentaire
doc_cols = [col for col in DOC_STATUS_COLUMNS if col in df.columns]
revision_cols = [
    col
    for col in df.columns
    if col.startswith("derniere revision") or col.startswith("dernière revision")
]

with st.expander("Suivi documentaire"):
    if not doc_cols and not revision_cols:
        st.info("Aucune colonne documentaire détectée dans les données chargées.")
    else:
        synthese = doc_status_summary(df)
        if not synthese.empty:
            st.markdown("**Synthèse des statuts**")
            st.dataframe(synthese, use_container_width=True)

        if doc_cols:
            st.markdown("**Liens manquants**")
            cols_missing = st.columns(len(doc_cols))
            for col_container, col_name in zip(cols_missing, doc_cols):
                col_container.metric(col_name.title(), int(df[col_name].isna().sum()))

        affichage_cols = doc_cols + [c for c in revision_cols if c not in doc_cols]
        if affichage_cols:
            st.markdown("**Matrice documentaire**")
            column_config = {}
            for col_name in doc_cols:
                column_config[col_name] = st.column_config.TextColumn(col_name.title())
            for col_name in revision_cols:
                column_config[col_name] = st.column_config.DateColumn(col_name.title(), format="DD/MM/YYYY")

            matrice_style = _style_doc_matrix(df, doc_cols, revision_cols, REVISION_STALENESS_DAYS)
            st.dataframe(
                matrice_style,
                use_container_width=True,
                column_config=column_config,
            )

# Séries mensuelles
ajout_m = stats_mensuelles(df, m, year, "date_ajout")      # Mois, Ajouts, Libellé
debut_m = stats_mensuelles(df, m, year, "date_debut")      # Mois, Démarrages, Libellé
fin_m   = stats_mensuelles(df, m, year, "date_fin")        # Mois, Fins, Libellé

# Échelle verticale uniforme (max global + marge 10 %)
max_val = max(ajout_m["Ajouts"].max(), debut_m["Démarrages"].max(), fin_m["Fins"].max())
ylim = (0, max(1, int(max_val * 1.1)))

col_a, col_b, col_c = st.columns(3)

with col_a:
    st.markdown("**Ajouts par mois**")
    st.dataframe(ajout_m[["Mois","Libellé","Ajouts"]], use_container_width=True)
    bar_plot(ajout_m.rename(columns={"Libellé":"Mois_txt"}), "Mois_txt", "Ajouts", f"Ajouts {year}", ylim=ylim)

with col_b:
    st.markdown("**Démarrages par mois**")
    st.dataframe(debut_m[["Mois","Libellé","Démarrages"]], use_container_width=True)
    bar_plot(debut_m.rename(columns={"Libellé":"Mois_txt"}), "Mois_txt", "Démarrages", f"Démarrages {year}", ylim=ylim)

with col_c:
    st.markdown("**Fins par mois**")
    st.dataframe(fin_m[["Mois","Libellé","Fins"]], use_container_width=True)
    bar_plot(fin_m.rename(columns={"Libellé":"Mois_txt"}), "Mois_txt", "Fins", f"Fins {year}", ylim=ylim)

# Données brutes
with st.expander("Données brutes (nettoyées)"):
    st.dataframe(df, use_container_width=True)
    st.download_button("⬇️ Export CSV – Données brutes",
                       df.to_csv(index=False).encode("utf-8"),
                       file_name="donnees_brutes.csv")

st.caption("Astuce : si la détection d’entêtes échoue, vérifie que la colonne 'Logiciel' existe bien dans les premières lignes du tableau.")
st.caption("Développé par I.Bitar – EGIS")