import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Analyse Cost Global", layout="wide")

st.title("📊 Analyse des coûts globaux des fiches individuelles")

st.write(
    "Importez un fichier Excel contenant les fiches individuelles "
    "et l'application générera un tableau récapitulatif : 1 salarié par ligne × 1 colonne par mois."
)

uploaded_file = st.file_uploader("📂 Importer le fichier Excel", type=["xlsx"])

if uploaded_file is not None:
    st.success("Fichier importé ✔️")

    # Lecture des feuilles
    xls = pd.ExcelFile(uploaded_file)
    all_rows = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        if df.shape[0] < 3 or df.shape[1] < 3:
            continue

        # 1) Extraction du nom du salarié
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
        row_cout = df.index[mask_cout][0]

        # 3) Ligne "Libellé"
        mask_header = df.iloc[:, 1] == "Libellé"
        if not mask_header.any():
            continue
        row_header = df.index[mask_header][0]

        # 4) Extraction mois + coûts
        mois_labels = df.iloc[row_header, 2:-1]
        cout_values = df.iloc[row_cout, 2:-1]

        for mois, cout in zip(mois_labels, cout_values):
            if pd.isna(mois) or pd.isna(cout):
                continue

            all_rows.append({
                "Salarie": salarie,
                "Mois": str(mois),
                "Cout_global": float(cout)
            })

    # DataFrame long
    long_df = pd.DataFrame(all_rows)

    if long_df.empty:
        st.error("⚠️ Aucun coût global détecté dans ce fichier.")
    else:
        # Pivot large
        wide_df = long_df.pivot_table(
            index="Salarie",
            columns="Mois",
            values="Cout_global",
            aggfunc="sum"
        ).reset_index()

        st.subheader("📄 Tableau récapitulatif")
        st.dataframe(wide_df, use_container_width=True)

        # Export Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wide_df.to_excel(writer, index=False, sheet_name="Récap")

        st.download_button(
            label="📥 Télécharger le fichier récapitulatif",
            data=output.getvalue(),
            file_name="recap_cout_global.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

