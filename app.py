import streamlit as st
import pandas as pd
from io import BytesIO

# -----------------------------------------------------
#              LOGIQUE DE TRAITEMENT EXCEL
# -----------------------------------------------------
def traiter_fichier(file):
    START_COL = "Resource"
    id_col = "External ID"

    # Lecture Excel
    df = pd.read_excel(file)

    # Vérifier colonnes obligatoires
    if START_COL not in df.columns:
        raise Exception(f"Colonne '{START_COL}' introuvable dans le fichier.")
    if id_col not in df.columns:
        raise Exception(f"Colonne '{id_col}' introuvable dans le fichier.")

    # --- Rendre les colonnes uniques (en cas de doublons) ---
    cols, counts = [], {}
    for c in df.columns:
        if c not in counts:
            counts[c] = 1
            cols.append(c)
        else:
            counts[c] += 1
            cols.append(f"{c}_{counts[c]}")
    df = df.copy()
    df.columns = cols

    # --- Formater seulement les vraies dates (Timestamp) ---
    def format_value(v):
        if pd.isna(v):
            return v
        if isinstance(v, pd.Timestamp):
            # On convertit juste les datetime en string dd/mm/YYYY
            return v.strftime("%d/%m/%Y")
        return v  # on ne touche pas aux nombres / strings

    df = df.applymap(format_value)

    # --- Séparer colonnes fixes / colonnes à étendre ---
    start_index = df.columns.tolist().index(START_COL)
    fixed_cols = df.columns[:start_index]
    expand_cols = df.columns[start_index:]

    # Forcer External ID en string pour éviter les soucis au merge
    df[id_col] = df[id_col].astype(str)

    # Base = 1 ligne par External ID, avec les colonnes fixes
    base = df.groupby(id_col, as_index=False).first()[fixed_cols]

    # --- Fusion horizontale des ressources ---
    def expand(group):
        rows = group[expand_cols]
        values = []
        for _, r in rows.iterrows():
            values.extend(r.tolist())
        # Une seule Series par External ID : [Res1, Work1, Units1, ..., ResN, WorkN, UnitsN...]
        return pd.Series(values)

    expanded = df.groupby(id_col).apply(expand)

    # ⚠️ Avec les versions récentes de pandas, groupby+apply renvoie une Series à multi-index.
    # On la remet en DataFrame "classique" avec unstack().
    if isinstance(expanded, pd.Series):
        expanded = expanded.unstack()

    expanded = expanded.reset_index()

    # --- Renommer les colonnes générées : Resource_1, Resource Estimated Work_1, etc. ---
    new_cols = [id_col]
    for i in range(1, len(expanded.columns)):
        idx = (i - 1) % len(expand_cols)          # position dans le bloc (Resource, Work, Units, ...)
        rep = (i - 1) // len(expand_cols) + 1     # n° du bloc (1, 2, 3, ...)
        new_cols.append(f"{expand_cols[idx]}_{rep}")

    expanded.columns = new_cols

    # --- Fusion finale : colonnes fixes + colonnes ressources étendues ---
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
st.subheader("Fusionner les ressources par External ID (1 ligne par tâche/projet)")

uploaded = st.file_uploader("📂 Uploader ton fichier Excel (export des tâches)", type=['xlsx'])

if uploaded:
    if st.button("🚀 Lancer le traitement"):
        with st.spinner("Traitement en cours..."):
            try:
                df_result = traiter_fichier(uploaded)

                # Export dans un fichier Excel en mémoire
                output = BytesIO()
                df_result.to_excel(output, index=False)
                output.seek(0)

                st.success("✔ Fichier généré avec succès")

                # Statistiques
                nb_rows = len(df_result)
                nb_resource_cols = sum(
                    1 for col in df_result.columns
                    if col.startswith("Resource") and col != "Resource"
                )

                st.info(f"📊 **{nb_rows}** lignes (External ID uniques) | **{nb_resource_cols}** colonnes de ressources")

                # Bouton de téléchargement
                st.download_button(
                    "⬇ Télécharger OUTPUT.xlsx",
                    data=output,
                    file_name="OUTPUT.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Aperçu
                st.write("📄 Aperçu du résultat :")
                st.dataframe(df_result.head(20))

            except Exception as e:
                st.error(f"❌ Erreur lors du traitement : {str(e)}")
else:
    st.info("👆 Import un fichier pour commencer.")
