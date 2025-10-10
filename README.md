# 🎓 Forum de l'Orientation – Générateur de Plannings

https://forumorientation.streamlit.app

Cette application Streamlit permet de gérer efficacement le **Forum des Métiers** en automatisant :

- ✅ Le traitement des vœux des élèves
- 🧠 L'affectation intelligente sur les créneaux disponibles
- 📅 La génération de plannings personnalisés (élèves & intervenants)
- 📤 L’export en **Excel** et **Word** pour diffusion facile

---

## 🚀 Fonctionnalités

| Module | Description |
|--------|-------------|
| `0_Home.py` | Page d'accueil explicative |
| `1_Depot.py` | Dépôt des fichiers nécessaires |
| `2_Affectation_Eleves.py` | Traitement automatique des affectations |
| `3_Planning_Intervenants.py` | Planning par intervenant avec export Word |
| `4_Planning_Eleves.py` | Planning par élève avec export Word |

---

## 📁 Fichiers requis

Avant de démarrer, prépare les fichiers suivants (au format `.xlsx`) :

1. **`01_voeux_eleves.xlsx`**  
   Contient les vœux des élèves (colonnes : Nom, Prénom, Classe, Vœu 1 → Vœu 6)

2. **`02_tables_poles_metiers.xlsx`**  
   Liste des métiers, intervenants, horaires, capacités (colonnes : Métier, Intervenant, Heure début, Heure fin, Capacité par créneau)

3. **`03_groupes_horaires.xlsx`**  
   Créneaux disponibles par classe (colonnes : Groupe, Horaire début, Horaire fin)

---

## 🛠️ Lancer en local

### 1. Clone du dépôt

```bash
git clone https://github.com/julientexier86/Forum-de-l-orientation-orga-.git
cd Forum-de-l-orientation-orga-
