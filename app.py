import streamlit as st
import pandas as pd
import io
import plotly.express as px
from datetime import datetime
import re

st.set_page_config(page_title="Analyse Coût Global", layout="wide")

st.title("📊 Analyse des coûts globaux des fiches individuelles")

st.write(
    "Importez un fichier Excel contenant les fiches individuelles. "
    "L'application génère un récap (1 salarié par ligne, 1 colonne par mois), "
    "permet un classement par sous-groupe (administratif / entretien / soin / non renseigné) "
    "et affiche un graphique comparatif interactif avec Plotly."
)


# -----------------------------------------------------------
#  FONCTIONS UTILITAIRES
# -----------------------------------------------------------
def clean_salarie(name: str) -> str:
    """Nettoie le nom du salarié (supprime ' - Total', espaces, etc.)."""
    if not isinstance(name, str):
        name = str(name)
    # supprime " - Total" ou "Total" en fin de chaîne, insensible à la casse
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

        # 1) Nom du salarié à partir du titre de la première colonne
        col0 = str(df.columns[0])
        salarie = col0
        if "Fiche individuelle" in col0:
            try:
                salarie = col0.split("Fiche individuelle -", 1)[1].split("- De", 1)[0].strip()
            except Exception:
                pass

        salarie = clean_salarie(salarie)

        # 2) Ligne "Coût global"
        mask_cg = df.iloc[:, 1] == "Coût global"
        if not mask_cg.any():
            continue
        idx_cg = df.index[mask_cg][0]

        # 3) Ligne "Libellé" (ligne des mois)
        mask_lib = df.iloc[:, 1] == "Libellé"
        if not mask_lib.any():
            continue
        idx_lib = df.index[mask_lib][0]

        # 4) Extraction mois + coûts (colonnes 2 à avant-dernière, on enlève "Total")
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

    # Tableau large : 1 ligne par salarié, 1 colonne par mois
    wide_df = (
        long_df.pivot_table(
            index="Salarie",
            columns="Mois",
            values="Cout_global",
            aggfunc="sum",
        )
        .reset_index()
    )

    return long_df, wide_df


def enrichir_mois(long_df: pd.DataFrame) -> pd.DataFrame:
    """Ajoute une vraie Date pour trier/filtrer chronologiquement."""
    mois_map = {
        "Janvier": 1,
        "Février": 2,
        "Fevrier": 2,
        "Mars": 3,
        "Avril": 4,
        "Mai": 5,
        "Juin": 6,
        "Juillet": 7,
        "Août": 8,
        "Aout": 8,
        "Septembre": 9,
        "Octobre": 10,
        "Novembre": 11,
        "Décembre": 12,
        "Decembre": 12,
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


def appliquer_mapping_sous_groupes(
    long_df: pd.DataFrame, wide_df: pd.DataFrame, mapping_df: pd.DataFrame | None
):
    """
    Ajoute une colonne 'Sous_groupe' aux salariés.
    - Par défaut : 'non renseigné'
    - Si un mapping est fourni : on remplace par le sous-groupe indiqué.
    """
    sal_df = pd.DataFrame({"Salarie": sorted(long_df["Salarie"].unique())})
    sal_df["Sous_groupe"] = "non renseigné"

    if mapping_df is not None and not mapping_df.empty:
        # Essaie d'identifier automatiquement les colonnes salarié / sous-groupe
        sal_col = None
        sg_col = None
        for c in mapping_df.columns:
            cl = c.lower()
            if sal_col is None and "alar" in cl:  # salarie, salarié…
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

    # Merge sur les tables principales
    long_df = long_df.merge(sal_df, on="Salarie", how="left")
    wide_df = wide_df.merge(sal_df, on="Salarie", how="left")

    return long_df, wide_df, sal_df


# -----------------------------------------------------------
#  APPLICATION
# -----------------------------------------------------------
uploaded_file = st.file_uploader("📂 Importer le fichier Excel principal", type=["xlsx"])

if uploaded_file is not None:
    st.success("Fichier principal importé ✔️")

    long_df, wide_df = construire_tables(uploaded_file)

    if long_df.empty or wide_df.empty:
        st.error("⚠️ Aucun coût global détecté. Vérifiez la présence de la ligne 'Coût global'.")
        st.stop()

    # Ajout colonne Date
    long_df = enrichir_mois(long_df)

    # -------------------------------------------------------
    # 2ᵉ BOUTON : IMPORT DE LA TABLE DE SOUS-GROUPES
    # -------------------------------------------------------
    st.subheader("🏷️ Sous-groupes métiers (optionnel)")

    mapping_df = None
    if st.checkbox("Importer un fichier de correspondance Salarié / Sous_groupe ?"):
        mapping_file = st.file_uploader(
            "Fichier de sous-groupes (Excel ou CSV avec colonnes Salarié / Sous_groupe)",
            type=["xlsx", "csv"],
            key="mapping_uploader",
        )
        if mapping_file is not None:
            try:
                if mapping_file.name.lower().endswith(".csv"):
                    mapping_df = pd.read_csv(mapping_file, sep=None, engine="python")
                else:
                    mapping_df = pd.read_excel(mapping_file)
                st.success("Fichier de sous-groupes importé ✔️")
            except Exception as e:
                st.error(f"Impossible de lire le fichier de sous-groupes : {e}")

    # Appliquer la correspondance (ou 'non renseigné' par défaut)
    long_df, wide_df, sal_df = appliquer_mapping_sous_groupes(long_df, wide_df, mapping_df)

    # -------------------------------------------------------
    # SLIDER DE PÉRIODE
    # -------------------------------------------------------
    unique_dates = sorted(long_df["Date"].unique())
    if not unique_dates:
        st.error("Impossible de déterminer les dates (colonne 'Mois').")
        st.stop()

    st.subheader("📆 Période analysée")

    min_idx, max_idx = 0, len(unique_dates) - 1
    idx_start, idx_end = st.slider(
        "Sélectionnez la période à afficher :",
        min_value=min_idx,
        max_value=max_idx,
        value=(min_idx, max_idx),
    )
    start_date = unique_dates[idx_start]
    end_date = unique_dates[idx_end]

    # -------------------------------------------------------
    # SÉLECTION DES SOUS-GROUPES
    # -------------------------------------------------------
    st.subheader("🧩 Filtre par sous-groupe")

    sous_groupes = sorted(sal_df["Sous_groupe"].unique().tolist())
    selected_groups = st.multiselect(
        "Sous-groupes à afficher sur le graphique :",
        options=sous_groupes,
        default=sous_groupes,  # tous par défaut, y compris 'non renseigné'
    )

    # -------------------------------------------------------
    # TABLE SALARIÉS / SOUS-GROUPE (pour contrôle visuel)
    # -------------------------------------------------------
    st.markdown("**Table des salariés et de leur sous-groupe :**")
    st.dataframe(sal_df, use_container_width=True)

    # -------------------------------------------------------
    # GRAPHIQUE PLOTLY AGRÉGÉ PAR SOUS-GROUPE
    # -------------------------------------------------------
    st.subheader("📈 Coût global par sous-groupe (agrégé)")

    data_plot = long_df.copy()
    data_plot = data_plot[
        (data_plot["Date"] >= start_date)
        & (data_plot["Date"] <= end_date)
        & (data_plot["Sous_groupe"].isin(selected_groups))
    ]

    if data_plot.empty:
        st.info("Aucune donnée dans la période / les sous-groupes sélectionnés.")
    else:
        # Agrégation par Date & Sous_groupe
        agg_df = (
            data_plot.groupby(["Date", "Mois", "Sous_groupe"], as_index=False)["Cout_global"]
            .sum()
        )

        fig = px.line(
            agg_df,
            x="Date",
            y="Cout_global",
            color="Sous_groupe",
            markers=True,
            hover_data=["Sous_groupe", "Mois", "Cout_global"],
        )

        fig.update_layout(
            xaxis_title="Mois",
            yaxis_title="Coût global (€)",
            title="Évolution du coût global mensuel par sous-groupe",
            xaxis_tickformat="%m/%Y",
            legend_title_text="Sous-groupe",
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.25,  # légende sous le graphique
                xanchor="center",
                x=0.5,
            ),
            margin=dict(l=40, r=40, t=60, b=120),
        )

        st.plotly_chart(fig, use_container_width=True)

    # -------------------------------------------------------
    # EXPORT EXCEL
    # -------------------------------------------------------
    st.subheader("💾 Export Excel complet (par salarié / mois)")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wide_df.to_excel(writer, index=False, sheet_name="Récap")

    st.download_button(
        label="📥 Télécharger le récap (tous les salariés)",
        data=buffer.getvalue(),
        file_name="recap_cout_global_par_salarie_par_mois.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Veuillez importer un fichier Excel principal (.xlsx) pour commencer.")
