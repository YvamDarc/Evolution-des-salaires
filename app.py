import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt

st.set_page_config(page_title="Analyse Coût Global", layout="wide")

st.title("📊 Analyse des coûts globaux des fiches individuelles")

st.write(
    "Importez un fichier Excel contenant les fiches individuelles. "
    "L'application génère un récap (1 salarié par ligne, 1 colonne par mois) "
    "et affiche un graphique comparatif du coût global pour les salariés sélectionnés."
)

uploaded_file = st.file_uploader("📂 Importer le fichier Excel", type=["xlsx"])


def construire_tables(uploaded_file):
    """Lit le fichier Excel et renvoie (long_df, wide_df)."""
    xls = pd.ExcelFile(uploaded_file)
    enregistrements = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        # On saute les feuilles trop petites / vides
        if df.shape[0] < 3 or df.shape[1] < 3:
            continue

        # 1) Récupérer le nom du salarié à partir du titre de la première colonne
        col0 = str(df.columns[0])
        salarie = col0
        if "Fiche individuelle" in col0:
            try:
                part = col0.split("Fiche individuelle -", 1)[1]
                part = part.split("- De", 1)[0]
                salarie = part.strip()
            except Exception:
                pass

        # 2) Ligne "Coût global"
        mask_cout = df.iloc[:, 1] == "Coût global"
        if not mask_cout.any():
            continue
        idx_cout = df.index[mask_cout][0]

        # 3) Ligne "Libellé" (en-têtes de colonnes de mois)
        mask_header = df.iloc[:, 1] == "Libellé"
        if not mask_header.any():
            continue
        idx_header = df.index[mask_header][0]

        # 4) Extraction des mois + coûts globaux (colonnes 2 à l'avant-dernière)
        mois_labels = df.iloc[idx_header, 2:-1]
        cout_values = df.iloc[idx_cout, 2:-1]

        for mois, cout in zip(mois_labels, cout_values):
            if pd.isna(mois) or pd.isna(cout):
                continue
            enregistrements.append({
                "Salarie": salarie,
                "Mois": str(mois),
                "Cout_global": float(cout)
            })

    long_df = pd.DataFrame(enregistrements)

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


def ordonner_mois(df):
    """Ajoute une colonne d'ordre temporel à partir de 'Mois' (ex: 'Janvier 2024') et trie."""
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

    def parse_mois(m):
        parts = str(m).split()
        if len(parts) >= 2:
            nom = parts[0]
            annee = parts[-1]
            try:
                mois_num = mois_map.get(nom, 0)
                annee_num = int(annee)
            except Exception:
                mois_num, annee_num = 0, 0
        else:
            mois_num, annee_num = 0, 0
        return annee_num * 100 + mois_num

    out = df.copy()
    out["ordre_mois"] = out["Mois"].apply(parse_mois)
    out = out.sort_values("ordre_mois")
    return out


if uploaded_file is not None:
    st.success("Fichier importé ✔️")

    long_df, wide_df = construire_tables(uploaded_file)

    if long_df.empty or wide_df.empty:
        st.error("⚠️ Aucun coût global détecté dans ce fichier. Vérifiez la présence de la ligne 'Coût global'.")
    else:
        # --- Sélection des salariés ---
        st.subheader("👤 Sélection des salariés")

        liste_salaries = sorted(wide_df["Salarie"].unique().tolist())
        selection = st.multiselect(
            "Sélectionnez les salariés à comparer sur le graphique :",
            options=liste_salaries,
            default=liste_salaries  # TOUS sélectionnés par défaut
        )

        # Filtrer le tableau large pour l’affichage
        if selection:
            wide_sel = wide_df[wide_df["Salarie"].isin(selection)]
        else:
            wide_sel = wide_df.iloc[0:0]

        st.subheader("📄 Tableau récapitulatif (coût global)")
        st.dataframe(wide_sel, use_container_width=True)

        # --- Graphique unique comparatif ---
        st.subheader("📈 Coût global comparé (tous les salariés sélectionnés)")

        if selection:
            fig, ax = plt.subplots()

            # On construit un df filtré et trié par mois
            for salarie in selection:
                data_sal = long_df[long_df["Salarie"] == salarie]
                if data_sal.empty:
                    continue
                data_sal = ordonner_mois(data_sal)
                ax.plot(
                    data_sal["Mois"],
                    data_sal["Cout_global"],
                    marker="o",
                    label=salarie
                )

            ax.set_xlabel("Mois")
            ax.set_ylabel("Coût global")
            ax.set_title("Évolution du coût global par mois")
            plt.xticks(rotation=45, ha="right")
            ax.legend()
            plt.tight_layout()

            st.pyplot(fig)
        else:
            st.info("Sélectionnez au moins un salarié pour afficher le graphique.")

        # --- Export Excel du tableau large complet ---
        st.subheader("💾 Export")

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            wide_df.to_excel(writer, index=False, sheet_name="Récap")

        st.download_button(
            label="📥 Télécharger le récap complet (tous les salariés)",
            data=buffer.getvalue(),
            file_name="recap_cout_global_par_salarie_par_mois.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Veuillez importer un fichier Excel (.xlsx) pour commencer.")
    
