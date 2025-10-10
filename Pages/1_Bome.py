import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import timedelta

st.set_page_config(page_title="Forum Orientation - Affectation automatique", layout="wide")

# --------------------------------------------
# 📌 PAGE D'ACCUEIL — explication pédagogique
# --------------------------------------------
st.title("🎓 Forum des métiers - Générateur de plannings")

st.markdown("""
Bienvenue sur l'application de gestion automatique du Forum des Métiers.

Elle permet de :
- ✅ Répartir automatiquement les élèves selon leurs vœux,
- 📅 Générer un planning personnalisé pour chaque élève,
- 👩‍🏫 Éditer les plannings par professionnel ou intervenant,
- 🗂️ Exporter tous les documents au format Excel ou Word.

### 📁 Fichiers requis :

1. `01_voeux_eleves.xlsx`  : Fichier des vœux des élèves avec les colonnes : Nom, Prénom, Classe, Vœu 1 à Vœu 6
2. `02_tables_poles_metiers.xlsx` : Liste des métiers avec intervenants, horaires, capacités
3. `03_groupes_horaires.xlsx` : Liste des groupes (classes) avec horaires début/fin

✨ Appuie ensuite sur le bouton pour générer automatiquement tous les plannings.
""")

# Téléversement des fichiers
st.header("📤 Téléverse tes fichiers")

f_voeux = st.file_uploader("1. Vœux des élèves (01_voeux_eleves.xlsx)", type="xlsx")
f_tables = st.file_uploader("2. Tables métiers (02_tables_poles_metiers.xlsx)", type="xlsx")
f_groupes = st.file_uploader("3. Groupes horaires (03_groupes_horaires.xlsx)", type="xlsx")

if f_voeux and f_tables and f_groupes:
    st.success("✅ Tous les fichiers ont été chargés !")

    # Lecture des fichiers
    df_voeux = pd.read_excel(f_voeux)
    df_tables = pd.read_excel(f_tables)
    df_groupes = pd.read_excel(f_groupes)

    # Nettoyage des noms de colonnes
    df_tables.columns = df_tables.columns.str.strip()
    df_groupes.columns = df_groupes.columns.str.strip()

    # Génération bouton
    if st.button("🚀 Générer les plannings maintenant"):
        st.switch_page("2_Affectation_Eleves.py")