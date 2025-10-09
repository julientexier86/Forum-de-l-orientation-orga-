# 📄 2_Affectation_Eleves.py

import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

# 🔐 Titre
st.title("📚 Forum Orientation - Traitement des affectations")
st.markdown("""
Cette page vous permet de :
- Charger les fichiers `voeux`, `tables`, `groupes`
- Générer les plannings des élèves (Excel, Word, PDF)
- Produire les plannings par intervenant

Toutes les affectations respectent les vœux des élèves et les créneaux disponibles.
""")

# ✨ Upload des fichiers
col1, col2, col3 = st.columns(3)

with col1:
    voeux_file = st.file_uploader("1. Vœux des élèves", type="xlsx")
with col2:
    tables_file = st.file_uploader("2. Tables / Métiers", type="xlsx")
with col3:
    groupes_file = st.file_uploader("3. Groupes horaires", type="xlsx")

if voeux_file and tables_file and groupes_file:

    # 📂 Lecture des fichiers
    df_voeux = pd.read_excel(voeux_file)
    df_tables = pd.read_excel(tables_file)
    df_groupes = pd.read_excel(groupes_file)

    # Nettoyage
    df_tables.columns = df_tables.columns.str.strip()
    df_groupes.columns = df_groupes.columns.str.strip()

    # ⏰ Créneaux horaires par classe
    groupes_horaires = {}
    for _, row in df_groupes.iterrows():
        groupe = row["Groupe"]
        debut = pd.to_datetime(row["Horaire début"], format="%H:%M")
        fin = pd.to_datetime(row["Horaire fin"], format="%H:%M")

        creneaux = []
        current = debut
        while current + timedelta(minutes=15) <= fin:
            creneaux.append(current.strftime("%H:%M"))
            current += timedelta(minutes=15)

        groupes_horaires[groupe] = creneaux

    # 🗓️ Agenda des tables
    agenda_tables = {}
    for _, row in df_tables.iterrows():
        capacite = row["Capacite par creneau"]
        debut = pd.to_datetime(row["Heure debut"], format="%H:%M")
        fin = pd.to_datetime(row["Heure fin"], format="%H:%M")
        current = debut
        while current + timedelta(minutes=15) <= fin:
            heure = current.strftime("%H:%M")
            agenda_tables[(row["Metier"], heure)] = capacite
            current += timedelta(minutes=15)

    # 📊 Affectations
    affectations = []
    for _, eleve in df_voeux.iterrows():
        nom = eleve["Nom"]
        prenom = eleve["Prénom"]
        classe = eleve["Classe"]
        creneaux = groupes_horaires.get(classe, [])
        voeux = [eleve[f"Vœu {i}"] for i in range(1, 7)]
        dejas = set()
        affecte_creneaux = set()

        for vœu in voeux:
            for heure in creneaux:
                key = (vœu, heure)
                if (key in agenda_tables
                    and agenda_tables[key] > 0
                    and vœu not in dejas
                    and heure not in affecte_creneaux):

                    affectations.append({
                        "Nom": nom,
                        "Prénom": prenom,
                        "Classe": classe,
                        "Heure": heure,
                        "Métier": vœu,
                        "Vœu n°": voeux.index(vœu) + 1
                    })
                    agenda_tables[key] -= 1
                    dejas.add(vœu)
                    affecte_creneaux.add(heure)
                    break
            if len(affecte_creneaux) >= 4:
                break

    df_affect = pd.DataFrame(affectations)

    st.success(f"{len(df_affect)} affectations réalisées")

    # 👤 Affichage preview
    st.subheader("🎓 Aperçu des affectations")
    st.dataframe(df_affect.head(15))

    # 💾 Export Excel
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output

    st.download_button("🔗 Télécharger les affectations (.xlsx)", data=to_excel(df_affect), file_name="affectations_forum.xlsx")

else:
    st.warning("Merci d'importer les 3 fichiers .xlsx pour démarrer.")
