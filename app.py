import streamlit as st
import pandas as pd
import tempfile
from io import BytesIO

# -----------------------------------------------------
#              LOGIQUE DE TRAITEMENT EXCEL
# -----------------------------------------------------
def traiter_fichier(file):
    START_COL = "Resource"
    id_col = "External ID"

    df = pd.read_excel(file)

    # Rendre colonnes uniques
    cols, counts = [], {}
    for c in df.columns:
        if c not in counts:
            counts[c] = 1
            cols.append(c)
        else:
            counts[c] += 1
            cols.append(f"{c}_{counts[c]}")
    df.columns = cols

    # Formatage dates → dd/mm/YYYY
    def format_value(v):
        if pd.isna(v):
            return v
        if isinstance(v, pd.Timestamp):
            return v.strftime("%d/%m/%Y")
        try:
            return pd.to_datetime(v).strftime("%d/%m/%Y")
        except:
            return v

    df = df.applymap(format_value)

    # Séparation colonnes
    start_index = df.columns.tolist().index(START_COL)
    fixed_cols = df.columns.tolist()[:start_index]
    expand_cols = df.columns.tolist()[start_index:]

    # Base par External ID
    base = df.groupby(id_col, as_index=False).first()[fixed_cols]

    # Fusion ressources → horizontal
    def expand(group):
        rows = group[expand_cols]
        if len(rows) == 1:
            return rows.iloc[0]
        values = []
        for _, row in rows.iterrows():
            values.extend(row.tolist())
        return pd.Series(values)

    expanded = df.groupby(id_col, group_keys=False).apply(expand)

    # Renommer si plusieurs blocs
    if expanded.shape[1] > len(expand_cols):
        new_cols = []
        for i in range(expanded.shape[1]):
            base_col = expand_cols[i % len(expand_cols)]
            index = (i // len(expand_cols)) + 1
            new_cols.append(f"{base_col}_{index}")
        expanded.columns = new_cols
    else:
        expanded.columns = expand_cols

    expanded = expanded.reset_index()
    result = base.merge(expanded, on=id_col, how="left")

    return result


# -----------------------------------------------------
#                    UI STREAMLIT
# -----------------------------------------------------
st.set_page_config(
    page_title="Fusion Excel Tool",
    layout="centered",
    page_icon="📄"
)

st.title("📊 Fusion Ressources Excel")
st.subheader("Convertir et fusionner les ressources automatiquement")

uploaded = st.file_uploader("📂 Uploader ton fichier Excel INPUT.xlsx", type=['xlsx'])

if uploaded:
    if st.button("🚀 Lancer le traitement"):
        with st.spinner("Traitement en cours..."):
            df = traiter_fichier(uploaded)

            # Stocker fichier output
            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            st.success("✔ Traitement terminé, ton fichier est prêt !")

            st.download_button(
                "⬇ Télécharger OUTPUT.xlsx",
                data=output,
                file_name="OUTPUT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(df.head(20))
else:
    st.info("👆 Import un fichier pour commencer")
