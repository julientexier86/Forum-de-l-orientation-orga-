import streamlit as st

st.markdown("""# 🧭 Forum des Métiers — Application de Gestion et de Diffusion des Plannings

Bienvenue dans l'application dédiée à la gestion du **Forum d'Orientation** de votre établissement.  
Cette plateforme permet de traiter automatiquement les vœux des élèves, d'optimiser les affectations par créneaux et métiers, et de générer des plannings clairs pour chaque acteur impliqué.

---

## 🛠️ Fonctionnalités disponibles

### 🟡 1. Import des données & affectations
> Traitement des vœux des élèves et génération automatique des affectations

- Téléversez trois fichiers :
  - `01_voeux_eleves.xlsx` → contenant les **vœux (1 à 6)** de chaque élève
  - `02_tables_poles_metiers.xlsx` → liste des **métiers / intervenants** avec horaires et capacités
  - `03_groupes_horaires.xlsx` → les **créneaux horaires** disponibles par classe

- L’application :
  - **diagnostique** les données (cohérences, métiers inconnus, classes manquantes…)
  - **génère les affectations** selon les vœux et les créneaux disponibles
  - propose un export Excel des affectations individuelles

---

### 🟢 2. Planning Élèves
> Génère automatiquement des plannings individuels élèves

- Deux formats d’export :
  - `planning_par_eleve.xlsx` → pour usage interne (inclut le rang du vœu)
  - `planning_publipostage.xlsx` → prêt pour l’envoi aux familles

- Chaque élève a jusqu’à **4 passages** de 15 minutes dans des pôles métiers différents.

---

### 🔵 3. Planning Intervenants
> Permet à chaque professionnel de visualiser les élèves qu’il rencontrera par créneau

- Génération :
  - Un fichier Excel : `planning_par_professionnel_onglets.xlsx` (1 onglet/intervenant)
  - Un document Word : `planning_par_intervenant.docx` (1 page/intervenant)

- Les créneaux sont regroupés **visuellement** pour une lecture rapide.

---

## 📂 Structure attendue des fichiers d’entrée

### 🧑‍🎓 `01_voeux_eleves.xlsx`
| Nom | Prénom | Classe | Vœu 1 | Vœu 2 | ... | Vœu 6 |
|-----|--------|--------|-------|-------|-----|-------|

### 🧰 `02_tables_poles_metiers.xlsx`
| Metier | Nom Intervenant | Heure debut | Heure fin | Capacite par creneau |
|--------|------------------|--------------|------------|------------------------|

### ⏰ `03_groupes_horaires.xlsx`
| Groupe | Horaire début | Horaire fin |
|--------|----------------|-------------|

---

## 🧩 Informations techniques

- Créneaux découpés **par tranches de 15 minutes**
- Gestion jusqu’à **6 vœux par élève**
- Chaque élève peut être affecté à **maximum 4 pôles**
- Attribution prioritaire aux **vœux les plus élevés** disponibles
- Gestion fine des **capacités et disponibilités horaires** de chaque intervenant

---

## 🙋‍♂️ Support

Pour toute question, vous pouvez contacter l’administrateur de l’application ou le porteur du projet au sein de l’établissement.  
Cette application est hébergée localement ou sur une plateforme web selon les options de déploiement choisies.""")