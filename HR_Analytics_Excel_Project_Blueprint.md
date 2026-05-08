# 📊 HR Analytics — Reporting Pro 100% Excel
## Blueprint Projet Portfolio

---

## 1. CONTEXTE & POSITIONNEMENT

### Problématique métier
> Une grande entreprise constate un turnover croissant et souhaite comprendre les facteurs qui poussent ses collaborateurs à partir. La direction RH a besoin d'un outil de pilotage accessible — sans licence Power BI — pour monitorer l'attrition, la satisfaction et la performance en temps réel.

### Proposition de valeur du projet
- **Message clé** : "Un reporting RH stratégique, 100% Excel, sans licence supplémentaire"
- **Cible** : Grandes entreprises dont les équipes RH n'ont pas accès à Power BI/Tableau
- **Différenciateur portfolio** : Peu de candidats montrent Power Query + Power Pivot + DAX dans Excel — c'est un marqueur de niveau avancé

### Dataset
- **Source** : IBM HR Analytics Employee Attrition & Performance (Kaggle)
- **Volume** : 1 470 employés × 35 colonnes
- **Lien** : https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset

---

## 2. COLONNES DU DATASET & CATÉGORISATION

### Démographie
| Colonne | Type | Description |
|---------|------|-------------|
| Age | Int | Âge de l'employé |
| Gender | Cat | Homme / Femme |
| MaritalStatus | Cat | Célibataire / Marié(e) / Divorcé(e) |
| Education | Ord (1-5) | 1=Below College → 5=Doctor |
| EducationField | Cat | Domaine d'études |

### Poste & Organisation
| Colonne | Type | Description |
|---------|------|-------------|
| Department | Cat | Sales / R&D / HR |
| JobRole | Cat | 9 rôles différents |
| JobLevel | Ord (1-5) | Niveau hiérarchique |
| BusinessTravel | Cat | Fréquence des déplacements |

### Rémunération
| Colonne | Type | Description |
|---------|------|-------------|
| MonthlyIncome | Int | Revenu mensuel |
| DailyRate | Int | Taux journalier |
| HourlyRate | Int | Taux horaire |
| MonthlyRate | Int | Taux mensuel |
| PercentSalaryHike | Int | % d'augmentation |
| StockOptionLevel | Ord (0-3) | Niveau de stock-options |

### Performance & Satisfaction
| Colonne | Type | Description |
|---------|------|-------------|
| PerformanceRating | Ord (1-4) | Note de performance |
| JobSatisfaction | Ord (1-4) | Satisfaction poste |
| EnvironmentSatisfaction | Ord (1-4) | Satisfaction environnement |
| RelationshipSatisfaction | Ord (1-4) | Satisfaction relations |
| WorkLifeBalance | Ord (1-4) | Équilibre vie pro/perso |
| JobInvolvement | Ord (1-4) | Implication au travail |

### Ancienneté & Parcours
| Colonne | Type | Description |
|---------|------|-------------|
| YearsAtCompany | Int | Années dans l'entreprise |
| YearsInCurrentRole | Int | Années dans le rôle actuel |
| YearsSinceLastPromotion | Int | Années depuis dernière promotion |
| YearsWithCurrManager | Int | Années avec le manager actuel |
| TotalWorkingYears | Int | Expérience totale |
| NumCompaniesWorked | Int | Nb d'entreprises précédentes |
| TrainingTimesLastYear | Int | Formations suivies l'an dernier |

### Variable cible
| Colonne | Type | Description |
|---------|------|-------------|
| Attrition | Cat | Yes / No — L'employé a-t-il quitté ? |
| OverTime | Cat | Yes / No — Heures supplémentaires |
| DistanceFromHome | Int | Distance domicile-travail (km) |

### Colonnes à supprimer (bruit)
- `EmployeeCount` (toujours = 1)
- `EmployeeNumber` (simple ID)
- `Over18` (toujours = Y)
- `StandardHours` (toujours = 80)

---

## 3. MODÈLE DIMENSIONNEL (Power Pivot)

### Architecture en étoile

```
         ┌──────────────┐
         │  dim_Employee │
         │──────────────│
         │ EmployeeID   │
         │ Age          │
         │ Gender       │
         │ MaritalStatus│
         │ Education    │
         │ EducationField│
         │ DistanceFromHome│
         │ NumCompanies │
         │ TotalWorkYrs │
         └──────┬───────┘
                │
┌───────────┐   │   ┌──────────────┐
│ dim_Job    │   │   │dim_Satisfaction│
│───────────│   │   │──────────────│
│ JobRole    │   │   │ JobSatisfaction│
│ JobLevel   │   │   │ EnvSatisfaction│
│ Department │   │   │ RelSatisfaction│
│ BusinessTravel│ │   │ WorkLifeBalance│
│ OverTime   │   │   │ JobInvolvement │
└─────┬─────┘   │   └──────┬───────┘
      │         │          │
      └────┐    │    ┌─────┘
           ▼    ▼    ▼
      ┌─────────────────────┐
      │    fact_Employee     │
      │─────────────────────│
      │ EmployeeID (FK)     │
      │ Attrition           │
      │ MonthlyIncome       │
      │ DailyRate           │
      │ HourlyRate          │
      │ PercentSalaryHike   │
      │ PerformanceRating   │
      │ StockOptionLevel    │
      │ TrainingTimesLastYr │
      │ YearsAtCompany      │
      │ YearsInCurrentRole  │
      │ YrsSinceLastPromo   │
      │ YearsWithCurrMgr    │
      └─────────────────────┘
```

### Note importante
Le dataset IBM est une snapshot (pas de dimension temps avec des dates). On va **créer une dim_AncienneteRange** (tranches d'ancienneté : 0-2 ans, 3-5 ans, 6-10 ans, 11-20 ans, 20+ ans) et une **dim_AgeRange** (18-25, 26-35, 36-45, 46-55, 55+) dans Power Query pour compenser l'absence de dates et permettre des analyses temporelles simulées.

---

## 4. PIPELINE ETL (Power Query)

### Étape 1 — Ingestion
- Importer le CSV dans Excel via `Données > Obtenir des données > À partir d'un fichier CSV`
- Renommer la requête : `raw_hr_data`

### Étape 2 — Nettoyage
| Action | Détail |
|--------|--------|
| Supprimer colonnes inutiles | EmployeeCount, EmployeeNumber, Over18, StandardHours |
| Vérifier types | Int pour les numériques, Texte pour les catégorielles |
| Vérifier les nulls | Normalement 0 nulls dans ce dataset, mais vérifier |
| Renommer colonnes | Passer en français si tu veux un rendu francophone (ex: MonthlyIncome → RevenuMensuel) |

### Étape 3 — Enrichissement (colonnes calculées)
| Nouvelle colonne | Formule Power Query |
|------------------|---------------------|
| TrancheAge | `if [Age] <= 25 then "18-25" else if [Age] <= 35 then "26-35" else if [Age] <= 45 then "36-45" else if [Age] <= 55 then "46-55" else "55+"` |
| TrancheAnciennete | `if [YearsAtCompany] <= 2 then "0-2 ans" else if [YearsAtCompany] <= 5 then "3-5 ans" else if [YearsAtCompany] <= 10 then "6-10 ans" else if [YearsAtCompany] <= 20 then "11-20 ans" else "20+ ans"` |
| TrancheRevenu | `if [MonthlyIncome] <= 3000 then "Bas" else if [MonthlyIncome] <= 7000 then "Moyen" else if [MonthlyIncome] <= 12000 then "Élevé" else "Très élevé"` |
| SatisfactionGlobale | `([JobSatisfaction] + [EnvironmentSatisfaction] + [RelationshipSatisfaction] + [WorkLifeBalance]) / 4` |
| NiveauSatisfaction | Basé sur SatisfactionGlobale : "Critique" / "Faible" / "Correcte" / "Excellente" |
| RisqueAttrition | Flag combinant : OverTime=Yes + JobSatisfaction<=2 + YearsSinceLastPromotion>=5 |

### Étape 4 — Création des tables dimensions
Créer des requêtes séparées (références ou doublons filtrés) pour chaque dimension :
- `dim_Employee` : colonnes démographiques uniques
- `dim_Job` : colonnes poste/organisation uniques  
- `dim_Satisfaction` : scores de satisfaction
- `fact_Employee` : table centrale avec FK + métriques

### Étape 5 — Chargement
- Charger chaque requête dans le **modèle de données uniquement** (ne pas charger dans des feuilles)
- Établir les relations dans Power Pivot

---

## 5. MESURES DAX

### KPIs principaux

```dax
// --- EFFECTIFS ---
Effectif Total = COUNTROWS(fact_Employee)

Nb Départs = CALCULATE(COUNTROWS(fact_Employee), fact_Employee[Attrition] = "Yes")

Nb Actifs = CALCULATE(COUNTROWS(fact_Employee), fact_Employee[Attrition] = "No")

// --- TAUX D'ATTRITION ---
Taux Attrition = 
    DIVIDE([Nb Départs], [Effectif Total], 0)

// --- RÉMUNÉRATION ---
Revenu Moyen = AVERAGE(fact_Employee[MonthlyIncome])

Revenu Médian = MEDIAN(fact_Employee[MonthlyIncome])

Masse Salariale = SUM(fact_Employee[MonthlyIncome])

Écart Revenu Moyen Actifs vs Partis = 
    CALCULATE(AVERAGE(fact_Employee[MonthlyIncome]), fact_Employee[Attrition] = "No")
    - CALCULATE(AVERAGE(fact_Employee[MonthlyIncome]), fact_Employee[Attrition] = "Yes")

// --- SATISFACTION ---
Score Satisfaction Moyen = AVERAGE(fact_Employee[SatisfactionGlobale])

Satisfaction Actifs = 
    CALCULATE(AVERAGE(fact_Employee[SatisfactionGlobale]), fact_Employee[Attrition] = "No")

Satisfaction Partis = 
    CALCULATE(AVERAGE(fact_Employee[SatisfactionGlobale]), fact_Employee[Attrition] = "Yes")

// --- PERFORMANCE ---
Note Performance Moyenne = AVERAGE(fact_Employee[PerformanceRating])

% Hauts Performeurs = 
    DIVIDE(
        CALCULATE(COUNTROWS(fact_Employee), fact_Employee[PerformanceRating] >= 3),
        [Effectif Total],
        0
    )

// --- ANCIENNETÉ ---
Ancienneté Moyenne = AVERAGE(fact_Employee[YearsAtCompany])

Délai Moyen Depuis Promo = AVERAGE(fact_Employee[YearsSinceLastPromotion])

// --- RISQUE ---
Nb Employés à Risque = 
    CALCULATE(
        COUNTROWS(fact_Employee),
        fact_Employee[RisqueAttrition] = "Oui"
    )

% Population à Risque = DIVIDE([Nb Employés à Risque], [Nb Actifs], 0)
```

### Mesures avancées (pour impressionner)

```dax
// Taux d'attrition par département (pour matrice)
Attrition par Dept = 
    CALCULATE(
        [Taux Attrition],
        ALLEXCEPT(dim_Job, dim_Job[Department])
    )

// Classement des départements par attrition
Rang Attrition Dept = 
    RANKX(
        ALL(dim_Job[Department]),
        [Taux Attrition],
        ,
        DESC
    )

// Écart satisfaction vs moyenne globale
Delta Satisfaction = 
    [Score Satisfaction Moyen] 
    - CALCULATE([Score Satisfaction Moyen], ALL(fact_Employee))
```

---

## 6. STRUCTURE DU CLASSEUR EXCEL

### Onglets (7 au total)

| # | Onglet | Rôle | Contenu |
|---|--------|------|---------|
| 1 | **ACCUEIL** | Navigation | Logo projet, titre, boutons de navigation vers chaque page, date de mise à jour, légende |
| 2 | **VUE EXECUTIVE** | Dashboard stratégique | 6 KPI cards en haut (Effectif, Taux Attrition, Revenu Moyen, Satisfaction, Performance, % Risque) + 2-3 graphiques clés |
| 3 | **ATTRITION** | Deep-dive attrition | Attrition par département, par tranche d'âge, par ancienneté, par niveau de satisfaction, par overtime — graphiques + tableau matriciel |
| 4 | **RÉMUNÉRATION** | Analyse salariale | Distribution des revenus, comparaison actifs vs partis, revenu par JobRole, par JobLevel, écarts H/F |
| 5 | **SATISFACTION** | Bien-être employés | Radar chart des 4 dimensions satisfaction, heatmap satisfaction × département, lien satisfaction-attrition |
| 6 | **PROFIL RISQUE** | Alerte & prévention | Liste des employés à risque (filtrée), profil-type du collaborateur qui part, facteurs de risque hiérarchisés |
| 7 | **MÉTHODOLOGIE** | README technique | Source des données, description du pipeline Power Query, modèle dimensionnel, glossaire KPI, parti pris du projet |

---

## 7. DESIGN & UX DU DASHBOARD

### Palette de couleurs recommandée
| Usage | Couleur | Hex |
|-------|---------|-----|
| Fond dashboard | Gris très clair | `#F5F5F5` |
| Titres & texte principal | Bleu nuit | `#1B2A4A` |
| Accent principal (KPI positifs) | Bleu corporate | `#2E75B6` |
| Accent secondaire (alertes) | Orange vif | `#ED7D31` |
| Danger / attrition | Rouge corail | `#E74C3C` |
| Succès / rétention | Vert émeraude | `#27AE60` |
| Fond cards KPI | Blanc | `#FFFFFF` |

### Règles de design
- **Police unique** : Calibri ou Segoe UI (natif Windows, propre)
- **KPI Cards** : Cellules fusionnées avec bordures fines, valeur en gras taille 20+, libellé en dessous taille 9
- **Graphiques** : Pas de quadrillage lourd, légende discrète, titres courts
- **Slicers** : Même style visuel, alignés horizontalement en haut de chaque page
- **Navigation** : Formes cliquables avec hyperliens vers chaque onglet
- **Pas de barres de défilement visibles** : Tout doit tenir dans un seul écran par onglet

---

## 8. PLAN D'EXÉCUTION

### Phase 1 — Setup & Exploration (Jour 1)
- [ ] Télécharger le dataset IBM HR depuis Kaggle
- [ ] Exploration rapide dans Excel (filtres, stats rapides)
- [ ] Identifier les distributions, valeurs aberrantes
- [ ] Documenter les observations initiales

### Phase 2 — ETL Power Query (Jour 2)
- [ ] Importer le CSV via Power Query
- [ ] Nettoyer (supprimer colonnes inutiles, typer correctement)
- [ ] Créer les colonnes enrichies (tranches, scores, flags)
- [ ] Créer les tables dimensions séparées
- [ ] Charger dans le modèle de données

### Phase 3 — Modèle Power Pivot + DAX (Jour 3)
- [ ] Configurer les relations entre tables dans Power Pivot
- [ ] Créer toutes les mesures DAX
- [ ] Tester chaque mesure individuellement
- [ ] Vérifier la cohérence (ex: Nb Départs + Nb Actifs = Effectif Total)

### Phase 4 — Construction des dashboards (Jours 4-5)
- [ ] Construire la page d'accueil avec navigation
- [ ] Construire la vue exécutive (KPI + graphiques principaux)
- [ ] Construire les 4 pages de deep-dive
- [ ] Ajouter les slicers synchronisés
- [ ] Appliquer le design (couleurs, polices, alignement)

### Phase 5 — Finitions & Documentation (Jour 6)
- [ ] Page méthodologie / README
- [ ] Revue complète UX (tout tient en un écran ?)
- [ ] Tests : tous les slicers fonctionnent ? Les KPI sont cohérents ?
- [ ] Export final + screenshots pour le portfolio GitHub

---

## 9. LIVRABLES FINAUX

| Livrable | Format | Destination |
|----------|--------|-------------|
| Classeur Excel complet | `.xlsx` | GitHub + portfolio |
| Screenshots des dashboards | `.png` | README GitHub |
| README.md du projet | `.md` | GitHub |
| (Optionnel) Vidéo walkthrough | `.mp4` | LinkedIn / YouTube |

### Structure GitHub recommandée
```
hr-analytics-excel/
├── README.md
├── data/
│   └── ibm_hr_analytics.csv
├── dashboard/
│   └── HR_Analytics_Dashboard.xlsx
├── screenshots/
│   ├── 01_accueil.png
│   ├── 02_vue_executive.png
│   ├── 03_attrition.png
│   ├── 04_remuneration.png
│   ├── 05_satisfaction.png
│   └── 06_profil_risque.png
└── docs/
    └── methodology.md
```

---

## 10. CE QUE CE PROJET DÉMONTRE AUX RECRUTEURS

| Compétence | Comment elle est démontrée |
|------------|---------------------------|
| **ETL / Data Engineering** | Pipeline Power Query complet (ingestion → nettoyage → enrichissement → modèle) |
| **Modélisation dimensionnelle** | Schéma en étoile fait/dimensions dans Power Pivot |
| **DAX / Mesures calculées** | KPIs complexes, CALCULATE, RANKX, DIVIDE, ALL/ALLEXCEPT |
| **Data Visualization** | Dashboard multi-pages, design soigné, storytelling visuel |
| **Compréhension métier RH** | Attrition, satisfaction, rémunération, gestion des risques |
| **Autonomie technique** | Tout fait dans un seul outil, sans dépendances |
| **Documentation** | README, méthodologie, glossaire — montre la rigueur |
