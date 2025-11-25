import streamlit as st
import pandas as pd
import io
import plotly.express as px

st.set_page_config(page_title="Analyse Coût Global", layout="wide")

st.title("📊 Analyse des coûts globaux des fiches individuelles")

st.write(
    "Importez un fichier Excel contenant les fiches individuelles. "
    "L'application génère un récap (1 salarié par ligne, 1 colonne par mois) "
    "et affiche un graphique comparatif interactif avec Plotly."
)

uploaded_file = st.file_uploader("📂 Importer le fichier Excel", type=["xlsx"])


# -----------------------------------------------------------
#  FONCTIONS
# -----------------------------------------------------------
def construire_tables(uploaded_file):
    """Lit le fichier Excel et renvoie (long_df, wide_df)."""
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
            rows.append({
                "Salarie": salarie,
                "Mois": str(m),
                "Cout_global": float(c)
            })

    long_df = pd.DataFrame(rows)

    if long_df.empty:
        return long_df, pd.DataFrame()

    # Tableau large : 1 ligne par salarié, 1 colonne par mois
    wide_df = long_df.pivot_table(
        index="Salarie",
        columns="Mois",
        values="Cout_global",
        aggfunc="sum"
    ).reset_index()

    return long_df, wide_df


def ajouter_ordre_mois(df):
    """Ajoute une colonne 'ordre_mois' pour trier correctement les mois (Janvier 2024, etc.)."""
    mois_map = {
        "Janvier": 1, "Février": 2, "Fevrier": 2, "Mars": 3, "Avril": 4,
        "Mai": 5, "Juin": 6, "Juillet": 7, "Août": 8, "Aout": 8,
        "Septembre": 9, "Octobre": 10, "Novembre": 11, "Décembre": 12, "Decembre": 12
    }

    def parse(m):
        parts = str(m).split()
        if len(parts) >= 2:
            nom = parts[0]
            annee = parts[-1]
            try:
                return int(annee) * 100 + mois_map.get(nom, 0)
            except Exception:
                return 0
        return 0

    df = df.copy()
    df["ordre_mois"] = df["Mois"].apply(parse)
    return df


# -----------------------------------------------------------
#  APPLICATION
# -----------------------------------------------------------
if uploaded_file is not None:
    st.success("Fichier importé ✔️")

    long_df, wide_df = construire_tables(uploaded_file)

    if long_df.empty or wide_df.empty:
        st.error("⚠️ Aucun coût global détecté. Vérifiez la présence de la ligne 'Coût global'.")
        st.stop()

    # --- Sélection des salariés ---
    st.subheader("👤 Sélection des salariés")

    liste_salaries = sorted(wide_df["Salarie"].unique().tolist())

    selection = st.multiselect(
        "Salariés à comparer sur le graphique :",
        options=liste_salaries,
        default=liste_salaries  # TOUS sélectionnés par défaut
    )

    # Tableau récap filtré
    if selection:
        wide_sel = wide_df[wide_df["Salarie"].isin(selection)]
    else:
        wide_sel = wide_df.iloc[0:0]

    st.subheader("📄 Tableau récapitulatif (coût global)")
    st.dataframe(wide_sel, use_container_width=True)

    # --- Graphique Plotly ---
    st.subheader("📈 Coût global comparé (graphique interactif Plotly)")

    if selection:
        data_plot = long_df[long_df["Salarie"].isin(selection)]
        data_plot = ajouter_ordre_mois(data_plot)
        data_plot = data_plot.sort_values("ordre_mois")

        fig = px.line(
            data_plot,
            x="Mois",
            y="Cout_global",
            color="Salarie",
            markers=True,
            hover_data=["Salarie", "Mois", "Cout_global"]
        )

        fig.update_layout(
            xaxis_title="Mois",
            yaxis_title="Coût global (€)",
            title="Évolution du coût global mensuel par salarié",
            xaxis_tickangle=-45,
            legend_title_text="Salarié",
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.3,        # en dessous du graph
                xanchor="center",
                x=0.5
            ),
            margin=dict(l=40, r=40, t=60, b=120),
        )

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sélectionnez au moins un salarié pour afficher le graphique.")

    # --- Export Excel ---
    st.subheader("💾 Export Excel complet")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wide_df.to_excel(writer, index=False, sheet_name="Récap")

    st.download_button(
        label="📥 Télécharger le récap (tous les salariés)",
        data=buffer.getvalue(),
        file_name="recap_cout_global_par_salarie_par_mois.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Veuillez importer un fichier Excel (.xlsx) pour commencer.")
