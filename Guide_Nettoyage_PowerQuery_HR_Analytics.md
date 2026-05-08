# 🧹 Guide de nettoyage Power Query — HR Analytics
## Instructions étape par étape

---

## RÉSUMÉ DE L'AUDIT

| Métrique | Valeur |
|----------|--------|
| Lignes | 311 employés |
| Colonnes (brutes) | 36 |
| Colonnes à garder | 22 (après suppression des ID redondants) |
| Colonnes à créer | 7 (enrichissement) |
| Nulls | DateofTermination : 207 (normal = employés actifs) / ManagerID : 8 |
| Problèmes trouvés | 12 (détaillés ci-dessous) |

---

## ÉTAPE 0 — IMPORTER LE FICHIER

> ⚠️ **IMPORTANT** : Ton fichier `.xls` est en réalité un CSV encodé en UTF-8-BOM. Excel le gère nativement mais il faut bien l'importer via Power Query.

1. Ouvre un **nouveau classeur Excel vide**
2. Va dans **Données → Obtenir des données → À partir d'un fichier → À partir d'un classeur**
   - Si Excel ne le reconnaît pas bien, utilise plutôt **À partir d'un fichier texte/CSV**
3. Sélectionne `HRDataset_v14.xls`
4. Dans l'aperçu, vérifie que :
   - Le délimiteur est bien **Virgule**
   - L'encodage est **UTF-8**
   - La première ligne est utilisée comme en-têtes
5. Clique sur **Transformer les données** (pas "Charger" — on veut l'éditeur Power Query)

Tu devrais voir 311 lignes × 36 colonnes. Si c'est le cas, on continue.

---

## ÉTAPE 1 — SUPPRIMER LES COLONNES REDONDANTES

### Pourquoi
Le dataset contient 9 colonnes d'ID numériques qui dupliquent des colonnes texte déjà présentes. Certaines sont même **incohérentes** (DeptID, EmpStatusID, PerfScoreID ne mappent pas proprement leurs libellés). On les supprime.

### Action dans Power Query
1. Sélectionne les colonnes suivantes (Ctrl+Clic) :
   - `MarriedID`
   - `MaritalStatusID`
   - `GenderID`
   - `EmpStatusID`
   - `DeptID`
   - `PerfScoreID`
   - `PositionID`
   - `Termd` (doublon de EmploymentStatus)
   - `FromDiversityJobFairID` (on recréera un flag propre)
2. Clic droit → **Supprimer les colonnes**

**Résultat attendu** : 311 lignes × 27 colonnes

---

## ÉTAPE 2 — NETTOYER LES ESPACES PARASITES (TRIM)

### Problèmes détectés

| Colonne | Problème | Exemple |
|---------|----------|---------|
| Employee_Name | 70 valeurs avec espaces à droite | `"Ait Sidi, Karthikeyan   "` |
| Sex | TOUS les "M" ont un espace | `"M "` au lieu de `"M"` |
| Department | "Production" a ~15 espaces | `"Production       "` |
| Position | 1 valeur "Data Analyst " avec espace | `"Data Analyst "` au lieu de `"Data Analyst"` |

### Action dans Power Query
1. Sélectionne les 4 colonnes : `Employee_Name`, `Sex`, `Department`, `Position`
   (Ctrl+Clic sur les en-têtes)
2. Menu **Transformer → Format → Découper le texte (Trim)**

### Vérification
Après le Trim, vérifie dans la barre de filtre :
- `Sex` ne doit plus avoir que **F** et **M** (pas "M ")
- `Department` doit avoir **6 valeurs propres** : Admin Offices, Executive Office, IT/IS, Production, Sales, Software Engineering
- `Position` doit avoir **31 valeurs** (Data Analyst ne doit apparaître qu'une seule fois)

---

## ÉTAPE 3 — CORRIGER LA CASSE INCOHÉRENTE

### Problème détecté

| Colonne | Problème |
|---------|----------|
| HispanicLatino | 4 valeurs : "No" (282), "Yes" (27), "no" (1), "yes" (1) |

### Action dans Power Query
1. Sélectionne la colonne `HispanicLatino`
2. Menu **Transformer → Format → Mettre en majuscule chaque mot** (Capitalize Each Word)

### Vérification
`HispanicLatino` ne doit plus avoir que **Yes** (28) et **No** (283).

---

## ÉTAPE 4 — CONVERTIR LES DATES

### Problèmes détectés
Les 4 colonnes de dates sont actuellement en **texte** avec des formats variés.

| Colonne | Format | Exemple | Nulls |
|---------|--------|---------|-------|
| DOB | MM/DD/YY (année sur 2 chiffres, toutes en 19xx) | `"07/10/83"` → 10 juillet 1983 | 0 |
| DateofHire | M/D/YYYY | `"7/5/2011"` | 0 |
| DateofTermination | M/D/YYYY | `"6/16/2016"` | 207 (employés actifs) |
| LastPerformanceReview_Date | M/D/YYYY | `"1/17/2019"` | 0 |

### Action pour DOB (cas spécial — année sur 2 chiffres)
> ⚠️ Les DOB vont de `51` (1951) à `92` (1992). Aucun employé n'est né après 2000.
> Power Query risque d'interpréter `51` comme `2051` au lieu de `1951`. On va le traiter manuellement.

1. Sélectionne la colonne `DOB`
2. Va dans **Ajouter une colonne → Colonne personnalisée**
3. Nom : `DateNaissance`
4. Formule :
```
Date.From(
    DateTime.FromText(
        Text.BeforeDelimiter([DOB], "/", 1) & "/" &
        Text.AfterDelimiter([DOB], "/", 1) & "/19" &
        Text.End([DOB], 2)
    )
)
```
> **Explication** : on reconstruit la date en forçant le siècle "19" devant l'année à 2 chiffres.

> **Alternative plus simple** si la formule ci-dessus pose problème :
> - Ajouter une colonne personnalisée nommée `DOB_Annee4` avec la formule :
>   `Text.BeforeDelimiter([DOB], "/") & "/" & Text.BetweenDelimiters([DOB], "/", "/") & "/19" & Text.End([DOB], 2)`
> - Puis changer le type de cette colonne en **Date** (clic droit sur l'en-tête → Type → Date)

5. Supprime l'ancienne colonne `DOB`
6. Renomme `DateNaissance` en `DOB` si tu veux garder le nom original

### Action pour les 3 autres dates
1. Sélectionne `DateofHire`, puis Ctrl+Clic sur `DateofTermination` et `LastPerformanceReview_Date`
2. Clic droit → **Modifier le type → Date**
3. Si Power Query demande confirmation de remplacement, accepte

### Vérification
- `DOB` : toutes les années doivent être entre **1951 et 1992**
- `DateofHire` : entre **2006 et 2018** environ
- `DateofTermination` : 207 valeurs null (normal) + dates entre **2009 et 2018**
- `LastPerformanceReview_Date` : entre **2012 et 2019**

---

## ÉTAPE 5 — TRAITER LES VALEURS NULLES

### Inventaire des nulls

| Colonne | Nb nulls | Traitement |
|---------|----------|------------|
| DateofTermination | 207 | ✅ **Normal** — ce sont les employés actifs. NE PAS remplacer. |
| ManagerID | 8 | ⚠️ Les 8 employés sans manager sont tous des Production Technicians. Remplacer par **0** (= "Non assigné"). |

### Action pour ManagerID
1. Sélectionne la colonne `ManagerID`
2. Menu **Transformer → Remplacer les valeurs**
   - Valeur à rechercher : `null`
   - Remplacer par : `0`
3. Change le type en **Nombre entier** (clic droit → Type → Nombre entier)

---

## ÉTAPE 6 — RENOMMER LES COLONNES

### Pour plus de clarté dans le modèle final, renomme ces colonnes :

| Ancien nom | Nouveau nom | Raison |
|------------|-------------|--------|
| Employee_Name | NomComplet | Cohérence FR |
| EmpID | EmployeID | Clarté |
| DOB | DateNaissance | Cohérence FR |
| Sex | Genre | Cohérence FR |
| MaritalDesc | StatutMarital | Cohérence FR |
| CitizenDesc | Citoyennete | Cohérence FR |
| HispanicLatino | OrigineHispanique | Cohérence FR |
| RaceDesc | Ethnie | Cohérence FR |
| DateofHire | DateEmbauche | Cohérence FR |
| DateofTermination | DateDepart | Cohérence FR |
| TermReason | RaisonDepart | Cohérence FR |
| EmploymentStatus | Statut | Cohérence FR |
| Department | Departement | Cohérence FR |
| ManagerName | NomManager | Cohérence FR |
| ManagerID | IDManager | Cohérence FR |
| RecruitmentSource | SourceRecrutement | Cohérence FR |
| PerformanceScore | ScorePerformance | Cohérence FR |
| EngagementSurvey | Engagement | Cohérence FR |
| EmpSatisfaction | Satisfaction | Cohérence FR |
| SpecialProjectsCount | NbProjetsSpeciaux | Cohérence FR |
| LastPerformanceReview_Date | DateDerniereEval | Cohérence FR |
| DaysLateLast30 | JoursRetard30j | Cohérence FR |
| Salary | Salaire | Cohérence FR |
| Position | Poste | Cohérence FR |

> 💡 **Astuce Power Query** : double-clic sur l'en-tête de chaque colonne pour la renommer directement.

> ⚠️ **Optionnel** : si tu préfères garder les noms anglais pour un portfolio international, c'est tout à fait valable. Le plus important est la cohérence — soit tout en FR, soit tout en EN.

---

## ÉTAPE 7 — CRÉER LES COLONNES ENRICHIES

### 7a — Tranche d'âge (à partir de DateNaissance)
1. **Ajouter une colonne → Colonne personnalisée**
2. Nom : `Age`
3. Formule :
```
Duration.TotalDays(DateTime.LocalNow() - [DateNaissance]) / 365.25
```
4. Change le type en **Nombre entier**

5. **Ajouter une colonne → Colonne conditionnelle**
6. Nom : `TrancheAge`
7. Conditions :
   - Si `Age` est inférieur à 30 → `"< 30 ans"`
   - Si `Age` est inférieur à 40 → `"30-39 ans"`
   - Si `Age` est inférieur à 50 → `"40-49 ans"`
   - Si `Age` est inférieur à 60 → `"50-59 ans"`
   - Sinon → `"60+ ans"`

### 7b — Tranche d'ancienneté
1. **Ajouter une colonne → Colonne personnalisée**
2. Nom : `Anciennete`
3. Formule (utilise DateDepart si l'employé est parti, sinon aujourd'hui) :
```
let
    dateFin = if [DateDepart] = null then DateTime.LocalNow() else [DateDepart]
in
    Number.RoundDown(Duration.TotalDays(dateFin - [DateEmbauche]) / 365.25)
```
4. Change le type en **Nombre entier**

5. **Ajouter une colonne → Colonne conditionnelle**
6. Nom : `TrancheAnciennete`
7. Conditions :
   - Si `Anciennete` est inférieur ou égal à 2 → `"0-2 ans"`
   - Si `Anciennete` est inférieur ou égal à 5 → `"3-5 ans"`
   - Si `Anciennete` est inférieur ou égal à 10 → `"6-10 ans"`
   - Sinon → `"10+ ans"`

### 7c — Tranche de salaire
1. **Ajouter une colonne → Colonne conditionnelle**
2. Nom : `TrancheSalaire`
3. Conditions :
   - Si `Salaire` est inférieur ou égal à 55000 → `"< 55K"`
   - Si `Salaire` est inférieur ou égal à 70000 → `"55K-70K"`
   - Si `Salaire` est inférieur ou égal à 100000 → `"70K-100K"`
   - Sinon → `"> 100K"`

### 7d — Indicateur de départ (flag clair)
1. **Ajouter une colonne → Colonne conditionnelle**
2. Nom : `AQuitte`
3. Conditions :
   - Si `Statut` est égal à `Active` → `"Non"`
   - Sinon → `"Oui"`

### 7e — Type de départ
1. **Ajouter une colonne → Colonne conditionnelle**
2. Nom : `TypeDepart`
3. Conditions :
   - Si `Statut` est égal à `Active` → `"Actif"`
   - Si `Statut` est égal à `Voluntarily Terminated` → `"Volontaire"`
   - Si `Statut` est égal à `Terminated for Cause` → `"Licenciement"`

### 7f — Catégorie de raison de départ
1. **Ajouter une colonne → Colonne personnalisée**
2. Nom : `CategorieDepart`
3. Formule :
```
if [RaisonDepart] = "N/A-StillEmployed" then "Actif"
else if [RaisonDepart] = "career change" or [RaisonDepart] = "Another position" or [RaisonDepart] = "more money" then "Opportunité externe"
else if [RaisonDepart] = "unhappy" or [RaisonDepart] = "hours" then "Insatisfaction"
else if [RaisonDepart] = "attendance" or [RaisonDepart] = "performance" or [RaisonDepart] = "no-call, no-show" or [RaisonDepart] = "gross misconduct" then "Cause disciplinaire"
else if [RaisonDepart] = "retiring" then "Retraite"
else if [RaisonDepart] = "return to school" or [RaisonDepart] = "military" then "Raison personnelle"
else if [RaisonDepart] = "relocation out of area" or [RaisonDepart] = "medical issues" or [RaisonDepart] = "maternity leave - did not return" then "Contrainte de vie"
else "Autre"
```

### 7g — Niveau de performance simplifié
1. **Ajouter une colonne → Colonne conditionnelle**
2. Nom : `NiveauPerformance`
3. Conditions :
   - Si `ScorePerformance` est égal à `Exceeds` → `"Haute"`
   - Si `ScorePerformance` est égal à `Fully Meets` → `"Satisfaisante"`
   - Si `ScorePerformance` est égal à `Needs Improvement` → `"À améliorer"`
   - Si `ScorePerformance` est égal à `PIP` → `"Critique"`

---

## ÉTAPE 8 — VÉRIFICATIONS FINALES

Avant de charger, vérifie :

### Checklist
- [ ] **27 colonnes originales nettoyées + 7 enrichies = 34 colonnes** (l'Age est aussi une nouvelle colonne, donc potentiellement 35 avec Age + Anciennete comme colonnes intermédiaires)
- [ ] **311 lignes** (pas de perte de données)
- [ ] Aucun espace parasite dans Sex, Department, Position
- [ ] HispanicLatino : seulement "Yes" et "No"
- [ ] DOB : toutes les dates entre 1951 et 1992
- [ ] DateEmbauche : toutes les dates entre 2006 et 2018
- [ ] ManagerID : aucun null (les 8 remplacés par 0)
- [ ] Les 7 colonnes enrichies n'ont aucune erreur

### Contrôle de cohérence
- Nombre d'employés actifs (AQuitte = "Non") = **207**
- Nombre de départs (AQuitte = "Oui") = **104** (88 volontaires + 16 licenciements)
- Taux d'attrition brut = 104 / 311 = **33.4%**

---

## ÉTAPE 9 — CHARGER DANS LE MODÈLE DE DONNÉES

### Option recommandée
1. Dans Power Query : **Accueil → Fermer et charger → Fermer et charger dans...**
2. Sélectionne : **Uniquement créer la connexion**
3. Coche : **Ajouter ces données au modèle de données**
4. Valide

> Cela charge les données dans **Power Pivot** sans créer de feuille de calcul. C'est la bonne pratique pour un modèle analytique — les données vivent dans le modèle, pas dans des onglets.

---

## RÉCAPITULATIF DES 12 PROBLÈMES DÉTECTÉS ET CORRIGÉS

| # | Problème | Colonne(s) | Sévérité | Correction |
|---|----------|-----------|----------|------------|
| 1 | 9 colonnes ID redondantes et parfois incohérentes | MarriedID, MaritalStatusID, GenderID, etc. | 🔴 Haute | Suppression |
| 2 | Espaces à droite sur les noms | Employee_Name | 🟡 Moyenne | Trim |
| 3 | Espace sur genre masculin ("M ") | Sex | 🔴 Haute | Trim |
| 4 | ~15 espaces sur "Production" | Department | 🔴 Haute | Trim |
| 5 | Espace sur 1 "Data Analyst " | Position | 🟡 Moyenne | Trim |
| 6 | Casse incohérente "no"/"No", "yes"/"Yes" | HispanicLatino | 🟡 Moyenne | Capitalize |
| 7 | Dates de naissance en année sur 2 chiffres | DOB | 🔴 Haute | Reconstruction avec "19" |
| 8 | 3 colonnes dates en format texte | DateofHire, etc. | 🔴 Haute | Conversion type Date |
| 9 | 207 nulls dans DateofTermination | DateofTermination | 🟢 Attendu | Aucune (normal) |
| 10 | 8 nulls dans ManagerID | ManagerID | 🟡 Moyenne | Remplacement par 0 |
| 11 | Fichier .xls est en fait un CSV UTF-8-BOM | Fichier entier | 🟢 Info | Import correct |
| 12 | Raisons de départ non catégorisées (18 valeurs) | TermReason | 🟡 Moyenne | Colonne CategorieDepart |

---

## APRÈS LE NETTOYAGE — PROCHAINES ÉTAPES

Une fois ce nettoyage terminé et les données chargées dans le modèle, les étapes suivantes seront :

1. **Créer les tables dimensions** dans Power Pivot (dim_Employe, dim_Poste, dim_Departement, dim_Manager)
2. **Établir les relations** entre les tables
3. **Écrire les mesures DAX** (taux d'attrition, salaire moyen, satisfaction par département, etc.)
4. **Construire les dashboards** onglet par onglet

> 💡 Dis-moi quand tu as terminé le nettoyage et je te guide pour la suite.
