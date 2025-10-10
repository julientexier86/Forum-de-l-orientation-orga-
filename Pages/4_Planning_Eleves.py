import streamlit as st
import pandas as pd

st.set_page_config(page_title="Planning Élèves", layout="wide")
st.title("📅 Planning par Élève")

# Chargement des données depuis session_state ou fichiers CSV
def load_data():
    if "df_affect" in st.session_state and "df_tables" in st.session_state:
        return st.session_state["df_affect"], st.session_state["df_tables"]
    try:
        df_affectations = pd.read_csv("data/df_affectations.csv")
        df_tables = pd.read_csv("data/df_tables.csv")
        return df_affectations, df_tables
    except Exception as e:
        st.error(f"Erreur lors du chargement des données : {e}")
        return None, None

df_affectations, df_tables = load_data()

if df_affectations is None or df_tables is None:
    st.warning("Aucune donnée disponible. Veuillez passer par l'étape d'affectation des élèves.")
    st.stop()

# Création d’un identifiant élève pour sélection
df_affectations["Élève"] = df_affectations["Nom"] + " " + df_affectations["Prénom"] + " (" + df_affectations["Classe"] + ")"

# Liste triée des élèves
liste_eleves = sorted(df_affectations["Élève"].unique())
choix = st.selectbox("🎓 Sélectionne un élève :", liste_eleves)

# Filtrage des lignes correspondantes
df_eleve = df_affectations[df_affectations["Élève"] == choix].sort_values(by="Heure")

# Ajout des intervenants
metier_to_intervenant = dict(zip(df_tables["Metier"], df_tables["Nom Intervenant"]))
df_eleve["Intervenant"] = df_eleve["Métier"].map(metier_to_intervenant)

# Affichage par créneau horaire
st.subheader(f"🕒 Planning pour {choix}")

for heure in df_eleve["Heure"].unique():
    st.markdown(f"### ⏰ {heure}")
    df_heure = df_eleve[df_eleve["Heure"] == heure][["Métier", "Intervenant"]]
    df_heure = df_heure.rename(columns={
        "Métier": "Métier rencontré",
        "Intervenant": "Intervenant"
    })
    st.table(df_heure.reset_index(drop=True))