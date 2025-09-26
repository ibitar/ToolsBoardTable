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


def _style_doc_matrix(df: pd.DataFrame, display_cols: List[str], doc_cols: List[str]):
    """Prépare un Styler pour la matrice documentaire en surlignant les statuts manquants."""
    if not display_cols:
        return df

    matrice = df[display_cols].copy()
    styler = matrice.style

    doc_cols_present = [col for col in doc_cols if col in matrice.columns]
    if doc_cols_present:
        styler = styler.applymap(
            lambda val: f"background-color: {STATUS_EMPTY_COLOR}" if pd.isna(val) or str(val).strip() == "" else "",
            subset=doc_cols_present,
        )

    return styler


def _normalize_text(text: str) -> str:
    """Supprime les accents et met en minuscules pour une comparaison robuste."""
    if text is None:
        return ""
    normalized = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).lower().strip()


def find_column_by_keywords(df: pd.DataFrame, keywords: List[str]) -> Optional[str]:
    """Retourne la première colonne dont le libellé contient tous les mots-clés normalisés."""
    normalized_keywords = [_normalize_text(keyword) for keyword in keywords]
    for col in df.columns:
        norm = _normalize_text(col)
        if all(keyword in norm for keyword in normalized_keywords):
            return col
    return None


def _format_unique_values(series: pd.Series, max_items: int = 5) -> str:
    """Met en forme les valeurs uniques d'une série pour affichage synthétique."""
    if series is None:
        return ""
    values = (
        series.dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
    )
    if len(values) == 0:
        return "Non renseigné"
    values_list = list(values[:max_items])
    if len(values) > max_items:
        values_list.append("…")
    return ", ".join(values_list)


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

def bar_plot(
    df: pd.DataFrame,
    x: str,
    y: str,
    title: str,
    ylim: Optional[Tuple[float, float]] = None,
    total: Optional[int] = None,
):
    """Trace un histogramme (Matplotlib) avec abscisses textuelles et y commun optionnel."""
    fig, ax = plt.subplots()
    ax.bar(df[x], df[y])
    ax.set_xlabel("Mois")
    ax.set_ylabel(y)
    if total is not None:
        ax.set_title(f"{title}\nTotal annuel : {total}")
    else:
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

# Colonnes clés utilisées dans plusieurs sections de l'application
logiciel_col = m.get("logiciel")
doc_cols = [col for col in DOC_STATUS_COLUMNS if col in df.columns]
version_col = find_column_by_keywords(df, ["version", "logiciel"])

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

# Répartition par logiciel / référent / rapport détaillé
tab_logiciels, tab_referents, tab_rapport = st.tabs(["Logiciels", "Référents", "Rapport logiciel"])

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

with tab_rapport:
    st.subheader("Rapport détaillé par logiciel")
    if logiciel_col is None:
        st.info("Aucune colonne 'Logiciel' détectée dans les données.")
    else:
        options = (
            df[logiciel_col]
            .dropna()
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .unique()
        )
        options = sorted(options)
        if not options:
            st.info("Aucun logiciel disponible pour le rapport.")
        else:
            selected_logiciel = st.selectbox("Choisir un logiciel", options)
            selection = df[df[logiciel_col].astype(str).str.strip() == selected_logiciel]

            if selection.empty:
                st.info("Aucune ligne correspondante trouvée pour ce logiciel.")
            else:
                st.markdown(f"### {selected_logiciel}")

                info_mapping = [
                    ("Version", version_col),
                    ("Dernière version du logiciel", find_column_by_keywords(df, ["derniere", "version", "logiciel"])),
                    ("Logiciel rattaché à", find_column_by_keywords(df, ["rattache", "logiciel"])),
                    ("Éditeur", find_column_by_keywords(df, ["editeur"])),
                    ("Référent", find_column_by_keywords(df, ["referent"])),
                    ("Ressources", find_column_by_keywords(df, ["ressources"])),
                    ("Métier", find_column_by_keywords(df, ["metier"])),
                    ("Type", find_column_by_keywords(df, ["type"])),
                    ("Avancement", find_column_by_keywords(df, ["avancement"])),
                    ("Jalon", find_column_by_keywords(df, ["jalon"])),
                    ("État du PV de recettage", find_column_by_keywords(df, ["etat", "pv", "recettage"])),
                    ("Date du dernier recettage", find_column_by_keywords(df, ["date", "dernier", "recettage"])),
                    ("Machines recettées", find_column_by_keywords(df, ["machine", "recette"])),
                    ("Confidentialité", find_column_by_keywords(df, ["confidentialite"])),
                ]

                info_rows = []
                for label, col_name in info_mapping:
                    if col_name and col_name in selection.columns:
                        formatted = _format_unique_values(selection[col_name])
                        if formatted and formatted != "Non renseigné":
                            info_rows.append((label, formatted))

                if info_rows:
                    st.table(pd.DataFrame(info_rows, columns=["Champ", "Valeur"]))

                if doc_cols:
                    st.markdown("**Statuts documentaires**")
                    doc_summary_rows = []
                    for col_name in doc_cols:
                        formatted = _format_unique_values(selection[col_name])
                        doc_summary_rows.append((col_name.title(), formatted))
                    st.table(pd.DataFrame(doc_summary_rows, columns=["Document", "État"]))

                commentaire_col = find_column_by_keywords(df, ["commentaire"])
                action_col = find_column_by_keywords(df, ["action", "faire"])
                if commentaire_col or action_col:
                    st.markdown("**Suivi des actions**")
                    columns_to_keep = [col for col in [commentaire_col, action_col] if col]
                    actions_df = selection[columns_to_keep].copy()
                    rename_map = {}
                    if commentaire_col:
                        rename_map[commentaire_col] = "Commentaire / action"
                    if action_col:
                        rename_map[action_col] = "Action à faire"
                    actions_df = actions_df.rename(columns=rename_map)
                    st.dataframe(actions_df, use_container_width=True, hide_index=True)

st.divider()

# Suivi documentaire
base_doc_cols = [col for col in [logiciel_col, version_col] if col]

with st.expander("Suivi documentaire"):
    if not doc_cols and not base_doc_cols:
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

        affichage_cols = base_doc_cols + [c for c in doc_cols if c not in base_doc_cols]
        if affichage_cols:
            st.markdown("**Matrice documentaire**")
            column_config = {}
            for col_name in base_doc_cols:
                column_config[col_name] = st.column_config.TextColumn(col_name.title())
            for col_name in doc_cols:
                column_config[col_name] = st.column_config.TextColumn(col_name.title())
            matrice_style = _style_doc_matrix(df, affichage_cols, doc_cols)
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
    bar_plot(
        ajout_m.rename(columns={"Libellé": "Mois_txt"}),
        "Mois_txt",
        "Ajouts",
        f"Ajouts {year}",
        ylim=ylim,
        total=int(ajout_m["Ajouts"].sum()),
    )

with col_b:
    st.markdown("**Démarrages par mois**")
    st.dataframe(debut_m[["Mois","Libellé","Démarrages"]], use_container_width=True)
    bar_plot(
        debut_m.rename(columns={"Libellé": "Mois_txt"}),
        "Mois_txt",
        "Démarrages",
        f"Démarrages {year}",
        ylim=ylim,
        total=int(debut_m["Démarrages"].sum()),
    )

with col_c:
    st.markdown("**Fins par mois**")
    st.dataframe(fin_m[["Mois","Libellé","Fins"]], use_container_width=True)
    bar_plot(
        fin_m.rename(columns={"Libellé": "Mois_txt"}),
        "Mois_txt",
        "Fins",
        f"Fins {year}",
        ylim=ylim,
        total=int(fin_m["Fins"].sum()),
    )

# Données brutes
with st.expander("Données brutes (nettoyées)"):
    st.dataframe(df, use_container_width=True)
    st.download_button("⬇️ Export CSV – Données brutes",
                       df.to_csv(index=False).encode("utf-8"),
                       file_name="donnees_brutes.csv")

st.caption("Astuce : si la détection d’entêtes échoue, vérifie que la colonne 'Logiciel' existe bien dans les premières lignes du tableau.")
st.caption("Développé par I.Bitar – EGIS")