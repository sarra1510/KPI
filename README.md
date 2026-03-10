# Sprint KPI Calculator

Calcule automatiquement les KPIs sprint à partir d'un fichier Excel extrait de Jira.

## Prérequis

- Python 3.8+
- pip

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python sprint_kpi_calculator.py <fichier_excel.xlsx>
```

**Exemple :**

```bash
python sprint_kpi_calculator.py sprint_data.xlsx
```

Le script vous demandera ensuite de saisir la **capacité équipe en heures** (calculée manuellement chaque sprint en fonction de la disponibilité des ressources).

---

## Format du fichier Excel attendu

Le fichier `.xlsx` doit contenir **3 feuilles** :

### Feuille 1 — `Start`

Extract Jira au **démarrage** du sprint.

Colonnes importantes :
- `Key` ou `Issue Key` — identifiant unique du ticket (colonne de jointure)
- `Summary` — résumé du ticket
- `Status` — statut du ticket
- `Assignee` — responsable
- `Original Estimate` — estimation initiale en secondes

### Feuille 2 — `End Sprint`

Extract Jira à la **fin** du sprint.

Colonnes importantes :
- `Key` ou `Issue Key` — identifiant unique (colonne de jointure)
- `Summary` — résumé du ticket
- `Status` — statut du ticket
- `Assignee` — responsable
- `Resolved` — date de résolution (ou `Resolution Date`)
- `Original Estimate` (ou `Σ Original Estimate`) — estimation initiale
- `Issue Type` — type de ticket
- `Priority` — priorité

### Feuille 3 — `Worklogs`

Worklog de l'équipe sur la période du sprint.

Colonnes importantes :
- `Issue Key` (ou `Key`) — identifiant du ticket
- `Hours` (ou `Time Spent` / `Heures`) — heures loggées

---

## KPIs calculés

| KPI | Formule |
|-----|---------|
| **Capacity Utilization (%)** | `Σ Hours (worklog) / Capacité équipe × 100` |
| **Throughput** | `COUNT tickets où Resolved is not null` (feuille End Sprint) |
| **Unplanned Tickets** | Tickets présents dans `End Sprint` mais absents de `Start` |
| **WIP End Sprint** | Tickets dans `End Sprint` dont le statut n'est PAS dans : `Closed`, `Customer Pending`, `Released`, `Canceled`, `Done` |
| **Support Load (%)** | `Unplanned / Throughput × 100` |
| **Tickets sans estimation** | Tickets où `Original Estimate` est null ou 0 |
| **Tickets sans tempo** | Tickets de `End Sprint` absents du Worklog |

---

## Fichier de sortie

Le script génère automatiquement un fichier `*_KPI_Report.xlsx` avec les onglets suivants :

| Onglet | Contenu |
|--------|---------|
| `KPI Summary` | Tableau récapitulatif de tous les KPIs |
| `Unplanned Tickets` | Détail des tickets non planifiés (Key, Summary, Status, Assignee, Issue Type, Priority) |
| `WIP End Sprint` | Détail des tickets encore en cours (Key, Summary, Status, Assignee, Issue Type, Priority) |
| `Sans Estimation` | Tickets sans estimation initiale (Key, Summary, Assignee, Status) |
| `Sans Tempo` | Tickets sans entrée de worklog (Key, Summary, Assignee, Status) |

---

## Structure du projet

```
KPI/
├── sprint_kpi_calculator.py   # Script principal
├── requirements.txt           # Dépendances Python
└── README.md                  # Documentation
```