import streamlit as st
import pandas as pd
import io
import plotly.express as px
from datetime import datetime
import re

st.set_page_config(page_title="Analyse Coût Global", layout="wide")

st.title("📊 Analyse des coûts globaux des fiches individuelles")

st.write(
    "1️⃣ Importez le fichier de fiches individuelles (obligatoire)\n"
    "2️⃣ (Optionnel) importez un fichier Salarié / Sous_groupe\n"
    "3️⃣ Cliquez sur le bouton pour générer le tableau par salarié\n"
    "4️⃣ Choisissez le type de graphique et cliquez pour le générer."
)

# =========================================================
#  FONCTIONS UTILITAIRES
# =========================================================

def clean_salarie(name: str) -> str:
    """Supprime '- Total' en fin de nom et nettoie."""
    if not isinstance(name, str):
        name = str(name)
    name = re.sub(r"\s*-\s*Total\s*$", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s+Total\s*$", "", name, flags=re.IGNORECASE)
    return name.strip()


def construire_tables(uploaded_file):
    """Lit le fichier Excel principal et renvoie (long_df, wide_df)."""
    xls = pd.ExcelFile(uploaded_file)
    rows = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)

        # Feuilles trop petites → on ignore
        if df.shape[0] < 3 or df.shape[1] < 3:
            continue

        # Nom du salarié (en-tête de la 1ère colonne)
        col0 = str(df.columns[0])
        salarie = col0
        if "Fiche individuelle" in col0:
            try:
                salarie = col0.split("Fiche individuelle -", 1)[1].split("- De", 1)[0].strip()
            except Exception:
                pass

        salarie = clean_salarie(salarie)

        # Ligne "Coût global"
        mask_cg = df.iloc[:, 1] == "Coût global"
        if not mask_cg.any():
            continue
        idx_cg = df.index[mask_cg][0]

        # Ligne "Libellé" (les mois)
        mask_lib = df.iloc[:, 1] == "Libellé"
        if not mask_lib.any():
            continue
        idx_lib = df.index[mask_lib][0]

        mois = df.iloc[idx_lib, 2:-1]
        couts = df.iloc[idx_cg, 2:-1]

        for m, c in zip(mois, couts):
            if pd.isna(m) or pd.isna(c):
                continue
            rows.append(
                {
                    "Salarie": salarie,
                    "Mois": str(m),
                    "Cout_global": float(c),
                }
            )

    long_df = pd.DataFrame(rows)

    if long_df.empty:
        return long_df, pd.DataFrame()

    # Tableau large : 1 ligne / salarié, 1 colonne / mois (on tri les colonnes après)
    wide_df = long_df.pivot_table(
        index="Salarie",
        columns="Mois",
        values="Cout_global",
        aggfunc="sum"
    ).reset_index()

    return long_df, wide_df


def enrichir_mois(long_df: pd.DataFrame) -> pd.DataFrame:
    """Ajoute une vraie Date pour trier/filtrer chronologiquement."""
    mois_map = {
        "Janvier": 1, "Février": 2, "Fevrier": 2, "Mars": 3, "Avril": 4,
        "Mai": 5, "Juin": 6, "Juillet": 7, "Août": 8, "Aout": 8,
        "Septembre": 9, "Octobre": 10, "Novembre": 11,
        "Décembre": 12, "Decembre": 12,
    }

    def parse_date(m):
        parts = str(m).split()
        if len(parts) >= 2:
            nom = parts[0]
            annee = parts[-1]
            try:
                mois_num = mois_map.get(nom, 1)
                annee_num = int(annee)
                return datetime(annee_num, mois_num, 1)
            except Exception:
                return None
        return None

    df = long_df.copy()
    df["Date"] = df["Mois"].apply(parse_date)
    df = df.dropna(subset=["Date"])
    df = df.sort_values("Date")
    return df


def reorder_wide_columns_chrono(wide_df: pd.DataFrame) -> pd.DataFrame:
    """Réordonne les colonnes de wide_df par ordre chronologique des mois."""
    mois_map = {
        "Janvier": 1, "Février": 2, "Fevrier": 2, "Mars": 3, "Avril": 4,
        "Mai": 5, "Juin": 6, "Juillet": 7, "Août": 8, "Aout": 8,
        "Septembre": 9, "Octobre": 10, "Novembre": 11,
        "Décembre": 12, "Decembre": 12,
    }

    def parse_date_from_col(col):
        parts = str(col).split()
        if len(parts) >= 2:
            nom = parts[0]
            annee = parts[-1]
            try:
                mois_num = mois_map.get(nom, 0)
                annee_num = int(annee)
                return datetime(annee_num, mois_num, 1)
            except Exception:
                return None
        return None

    base_cols = ["Salarie"]
    if "Sous_groupe" in wide_df.columns:
        base_cols.append("Sous_groupe")

    mois_cols = [c for c in wide_df.columns if c not in base_cols]

    mois_cols_sorted = sorted(
        mois_cols,
        key=lambda c: (parse_date_from_col(c) is None, parse_date_from_col(c) or datetime.max)
    )

    return wide_df[base_cols + mois_cols_sorted]


def appliquer_mapping_sous_groupes(
    long_df: pd.DataFrame,
    wide_df: pd.DataFrame,
    mapping_df: pd.DataFrame | None
):
    """Ajoute une colonne 'Sous_groupe' (ou 'non renseigné') à chaque salarié."""
    sal_df = pd.DataFrame({"Salarie": sorted(long_df["Salarie"].unique())})
    sal_df["Sous_groupe"] = "non renseigné"

    if mapping_df is not None and not mapping_df.empty:
        # devine les colonnes salarié / sous_groupe
        sal_col = None
        sg_col = None
        for c in mapping_df.columns:
            cl = c.lower()
            if sal_col is None and "alar" in cl:  # salarie / salarié
                sal_col = c
            if sg_col is None and ("sous" in cl or "groupe" in cl):
                sg_col = c

        if sal_col and sg_col:
            tmp = mapping_df[[sal_col, sg_col]].dropna()
            tmp = tmp.rename(columns={sal_col: "Salarie", sg_col: "Sous_groupe"})
            tmp["Salarie"] = tmp["Salarie"].apply(clean_salarie)
            sal_df = sal_df.merge(tmp, on="Salarie", how="left", suffixes=("", "_map"))
            sal_df["Sous_groupe"] = sal_df["Sous_groupe_map"].fillna(sal_df["Sous_groupe"])
            sal_df = sal_df[["Salarie", "Sous_groupe"]]

    long_df = long_df.merge(sal_df, on="Salarie", how="left")
    wide_df = wide_df.merge(sal_df, on="Salarie", how="left")

    return long_df, wide_df, sal_df


# =========================================================
#  INTERFACE
# =========================================================

# 1) Import principal
uploaded_file = st.file_uploader("📂 1️⃣ Importer le fichier Excel principal", type=["xlsx"])

# 2) Import mapping sous-groupes (optionnel)
mapping_file = st.file_uploader(
    "📂 2️⃣ (Optionnel) Importer le fichier Salarié / Sous_groupe",
    type=["xlsx", "csv"],
    key="mapping_file"
)

# Bouton pour préparer les données + afficher le tableau
st.markdown("---")
if st.button("▶️ Générer / actualiser le tableau salariés"):
    if uploaded_file is None:
        st.error("Veuillez d'abord importer le fichier principal.")
    else:
        # Lecture principale
        long_df, wide_df = construire_tables(uploaded_file)
        if long_df.empty or wide_df.empty:
            st.error("⚠️ Aucun coût global détecté. Vérifiez la présence de la ligne 'Coût global'.")
        else:
            # Enrichir avec la date
            long_df = enrichir_mois(long_df)

            # Lecture mapping si présent
            mapping_df = None
            if mapping_file is not None:
                try:
                    if mapping_file.name.lower().endswith(".csv"):
                        mapping_df = pd.read_csv(mapping_file, sep=None, engine="python")
                    else:
                        mapping_df = pd.read_excel(mapping_file)
                except Exception as e:
                    st.error(f"Impossible de lire le fichier de sous-groupes : {e}")
                    mapping_df = None

            # Application mapping + sous_groupe
            long_df, wide_df, sal_df = appliquer_mapping_sous_groupes(long_df, wide_df, mapping_df)

            # Réordonner les colonnes du tableau par ordre chronologique
            wide_df = reorder_wide_columns_chrono(wide_df)

            # Stocker en session pour réutilisation (graphique)
            st.session_state["long_df"] = long_df
            st.session_state["wide_df"] = wide_df
            st.session_state["sal_df"] = sal_df
            st.session_state["dates_uniques"] = sorted(long_df["Date"].unique())

            st.success("Données générées ✔️")

            # Affichage du tableau
            st.subheader("📄 Tableau par salarié (colonnes triées chronologiquement)")
            st.dataframe(wide_df, use_container_width=True)

# Si les données sont prêtes, on propose la partie graphique
if "long_df" in st.session_state:
    st.markdown("---")
    st.subheader("📊 3️⃣ Génération de graphiques")

    long_df = st.session_state["long_df"]
    wide_df = st.session_state["wide_df"]
    sal_df = st.session_state["sal_df"]
    unique_dates = st.session_state["dates_uniques"]

    # Choix du type de graphique
    type_graph = st.radio(
        "Type de graphique :",
        options=["Par salarié", "Par sous-groupe"],
        horizontal=True
    )

    # Sélection de la période
    min_idx, max_idx = 0, len(unique_dates) - 1
    idx_start, idx_end = st.slider(
        "Sélectionnez la période à afficher :",
        min_value=min_idx,
        max_value=max_idx,
        value=(min_idx, max_idx),
    )
    start_date = unique_dates[idx_start]
    end_date = unique_dates[idx_end]

    # Contrôles selon le type
    if type_graph == "Par salarié":
        liste_salaries = sorted(sal_df["Salarie"].unique().tolist())
        sel_salaries = st.multiselect(
            "Salariés à afficher :",
            options=liste_salaries,
            default=liste_salaries
        )
    else:
        liste_sg = sorted(sal_df["Sous_groupe"].unique().tolist())
        sel_sg = st.multiselect(
            "Sous-groupes à afficher :",
            options=liste_sg,
            default=liste_sg
        )

    # Bouton pour générer le graphique
    if st.button("📈 Générer le graphique"):
        data_plot = long_df.copy()
        data_plot = data_plot[(data_plot["Date"] >= start_date) & (data_plot["Date"] <= end_date)]

        if type_graph == "Par salarié":
            data_plot = data_plot[data_plot["Salarie"].isin(sel_salaries)]
            if data_plot.empty:
                st.info("Aucune donnée pour les salariés / période sélectionnés.")
            else:
                fig = px.line(
                    data_plot,
                    x="Date",
                    y="Cout_global",
                    color="Salarie",
                    markers=True,
                    hover_data=["Salarie", "Mois", "Cout_global", "Sous_groupe"]
                )
                fig.update_layout(
                    xaxis_title="Mois",
                    yaxis_title="Coût global (€)",
                    title="Évolution du coût global mensuel par salarié",
                    xaxis_tickformat="%m/%Y",
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.25,
                        xanchor="center",
                        x=0.5
                    ),
                    margin=dict(l=40, r=40, t=60, b=120),
                )
                st.plotly_chart(fig, use_container_width=True)

        else:  # Par sous-groupe
            data_plot = data_plot[data_plot["Sous_groupe"].isin(sel_sg)]
            if data_plot.empty:
                st.info("Aucune donnée pour les sous-groupes / période sélectionnés.")
            else:
                agg = (
                    data_plot
                    .groupby(["Date", "Mois", "Sous_groupe"], as_index=False)["Cout_global"]
                    .sum()
                )
                fig = px.line(
                    agg,
                    x="Date",
                    y="Cout_global",
                    color="Sous_groupe",
                    markers=True,
                    hover_data=["Sous_groupe", "Mois", "Cout_global"]
                )
                fig.update_layout(
                    xaxis_title="Mois",
                    yaxis_title="Coût global (€)",
                    title="Évolution du coût global mensuel par sous-groupe",
                    xaxis_tickformat="%m/%Y",
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.25,
                        xanchor="center",
                        x=0.5
                    ),
                    margin=dict(l=40, r=40, t=60, b=120),
                )
                st.plotly_chart(fig, use_container_width=True)

    # Export Excel du tableau trié
    st.subheader("💾 Export du tableau par salarié")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wide_df.to_excel(writer, index=False, sheet_name="Récap")
    st.download_button(
        "📥 Télécharger le tableau (colonnes triées)",
        buffer.getvalue(),
        "recap_cout_global_par_salarie_par_mois.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
