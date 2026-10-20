# Data Analyst Suite : analyse du trafic portuaire (ACP • IA • AFC • Cyber)

Application de bureau (Python / Tkinter) qui propose un pipeline d'analyse de données complet, appliqué à un jeu de données de **trafic portuaire** (navires, tonnage, durée d'escale, consommation de carburant, émissions de CO₂...).

Projet académique **Data Analyst 2025/2026**, réalisé par **Sara Naim**.

## Modules

| Module | Contenu |
|---|---|
| **ACP** | Statistiques, corrélations (heatmap), inerties, plan factoriel, cercle de corrélation, qualité de représentation, contributions |
| **IA** | Clustering **K-Means** (k réglable de 3 à 7), pourcentages par cluster, classification **Random Forest** (métriques, prédiction d'un nouvel individu) |
| **AFC** | Tableau de contingence et fréquences, test du χ² de dépendance, distances du χ², plan factoriel |
| **Cyber** | Détection d'anomalies avec **Isolation Forest** et **Local Outlier Factor (LOF)** sur un second jeu de données |

L'interface comprend des graphiques zoomables (barre d'outils Matplotlib), des tableaux avec lignes alternées et un menu en cartes.

## Technologies

Python, Tkinter, pandas, NumPy, scikit-learn, SciPy, Matplotlib, seaborn.

## Données

Dans le dossier `data/` :

- `DATA36_Trafic_portuaire_100_individus.xlsx` : 100 navires, 8 variables (tonnage, longueur, vitesse d'approche, durée d'escale, nombre de manœuvres, capacité en EVP, consommation de carburant, émissions de CO₂). Chargé par défaut.
- `DATA36_trafic_portuaire_ServerNav_100_lignes.xlsx` : jeu de données pour le module **Cyber** (à importer depuis l'application).

## Installation

Prérequis : [Python 3.10+](https://www.python.org/downloads/) (Tkinter est inclus dans l'installateur Windows).

```bash
git clone https://github.com/sara-nm22/data-analyst-suite-port-traffic.git
cd data-analyst-suite-port-traffic

python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Lancement

```bash
python data_analyst_suite.py
```

Pour le module Cyber, importez `data/DATA36_trafic_portuaire_ServerNav_100_lignes.xlsx` depuis l'application.

## Créer un exécutable Windows (optionnel)

```bash
pip install pyinstaller
pyinstaller DataAnalystSuite.spec
```

L'exécutable est généré dans `dist/`.

## Documentation

- Rapport complet : [`docs/Data_Analyst_Suite_Report.docx`](docs/Data_Analyst_Suite_Report.docx)
- Présentation : [`docs/Port_Traffic_Analytics_Presentation.pptx`](docs/Port_Traffic_Analytics_Presentation.pptx)

## Structure

```
data-analyst-suite-port-traffic/
├── data_analyst_suite.py   # Application
├── DataAnalystSuite.spec   # Configuration PyInstaller
├── requirements.txt
├── data/                   # Jeux de données Excel
└── docs/                   # Rapport et présentation
```

## Auteur

Sara Naim
