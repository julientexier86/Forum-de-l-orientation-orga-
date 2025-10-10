import streamlit as st
import pandas as pd

st.set_page_config(page_title="Planning Intervenants", layout="wide")
st.title("🗂️ Planning par Intervenant")

# Chargement des données mises en cache dans 2_Affectation_Eleves
@st.cache_data

def load_data():
    try:
        df_affectations = pd.read_csv("data/df_affectations.csv")
        df_tables = pd.read_csv("data/df_tables.csv")
        return df_affectations, df_tables
    except Exception as e:
        st.error(f"Erreur lors du chargement des données : {e}")
        return None, None

# Chargement des fichiers nécessaires
df_affectations, df_tables = load_data()

if df_affectations is None or df_tables is None:
    st.warning("Aucune donnée disponible. Veuillez passer par l'étape d'affectation des élèves.")
    st.stop()

# Association métiers ↔ intervenants
metier_to_intervenant = dict(zip(df_tables["Metier"], df_tables["Nom Intervenant"]))
df_affectations["Intervenant"] = df_affectations["Métier"].map(metier_to_intervenant)

# Sélection d’un intervenant
intervenants = sorted(df_affectations["Intervenant"].dropna().unique())
choix = st.selectbox("👤 Sélectionne un intervenant :", intervenants)

# Filtrage
df_int = df_affectations[df_affectations["Intervenant"] == choix]
df_int = df_int.sort_values(by="Heure")

# Affichage paginé par heure
st.subheader(f"🕒 Planning pour {choix}")

for heure in df_int["Heure"].unique():
    st.markdown(f"### ⏰ {heure}")
    df_heure = df_int[df_int["Heure"] == heure][["Nom", "Prénom", "Classe", "Métier"]]
    df_heure = df_heure.rename(columns={"Nom": "Nom élève", "Prénom": "Prénom", "Classe": "Classe", "Métier": "Métier présenté"})
    st.table(df_heure.reset_index(drop=True))