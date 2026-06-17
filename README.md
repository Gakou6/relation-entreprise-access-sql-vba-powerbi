# Base de données Relation Entreprise — Access, SQL, VBA et Power BI

## Objectif du projet

Ce projet a pour objectif de concevoir et d’exploiter une base de données relationnelle autour de la relation entreprise de l’IUT de Roubaix.

Le travail porte principalement sur les contrats d’apprentissage, les entreprises partenaires, les formations, les OPCO, les conventions collectives, les résiliations et les versements de taxe d’apprentissage.

## Contexte

Ce projet a été réalisé dans le cadre de la SAÉ 2.01 : Conception et implémentation d’une base de données.

Les données de départ étaient fournies sous forme de fichiers Excel anonymisés. L’objectif était de transformer ces fichiers plats en une base relationnelle structurée, exploitable dans Microsoft Access, puis d’en tirer des indicateurs à l’aide de requêtes SQL, d’états, de formulaires et d’un tableau de bord Power BI.

## Outils utilisés

- Microsoft Access
- SQL Access
- VBA
- Power BI
- Excel
- Modélisation de base de données
- MCD / MLD

## Étapes du projet

1. Analyse des fichiers sources
2. Identification des redondances et incohérences
3. Création du dictionnaire des données
4. Définition des dépendances fonctionnelles
5. Réalisation du MCD
6. Passage au MLD
7. Création des tables dans Access avec SQL
8. Alimentation automatique des tables avec VBA
9. Création de requêtes SQL d’analyse
10. Création d’états et de formulaires Access
11. Réalisation d’un tableau de bord Power BI

## Structure du dépôt

- `sql/` : requêtes SQL Access utilisées pour analyser les données.
- `vba/` : scripts VBA utilisés pour automatiser la création et l’alimentation des tables.
- `docs/` : documents de conception du projet, MCD, MLD et rapport.
- `visuals/` : captures des formulaires Access, états Access et tableau de bord Power BI.
- `data/` : description des données utilisées.

## Analyses réalisées

Plusieurs requêtes ont été réalisées afin de produire des indicateurs utiles au pilotage de la relation entreprise.

Exemples d’analyses :

- taux de rétention des contrats par OPCO ;
- durée moyenne des contrats avant rupture ;
- taux d’affectation des taxes par formation ;
- total des versements par année fiscale ;
- top 10 des entreprises par montant de taxe versée ;
- nombre de contrats et masse salariale par entreprise ;
- évolution annuelle du nombre de contrats ;
- taux de résiliation par formation ;
- analyse géographique des entreprises partenaires.

## États et formulaires Access

La base Access contient également plusieurs états et formulaires permettant de consulter les données de manière plus claire.

Exemples d’éléments réalisés :

- bilan des versements par campagne ;
- liste des versements ;
- fiche contrat d’apprentissage ;
- tableau de bord des formations ;
- top 10 des entreprises versantes ;
- formulaire de gestion des contrats ;
- formulaire de gestion des versements ;
- menu général de navigation.

## Tableau de bord Power BI

Un tableau de bord Power BI a été réalisé à partir de la base Access.

Il permet notamment de suivre :

- le nombre total de contrats ;
- le taux de rupture ;
- le nombre d’entreprises partenaires ;
- la masse salariale ;
- les principaux employeurs ;
- l’évolution des contrats dans le temps ;
- la répartition géographique des entreprises ;
- la répartition par type de formation.

## Compétences mises en avant

- Modélisation de base de données
- Conception MCD / MLD
- Création de tables SQL
- Requêtes SQL Access
- Automatisation avec VBA
- Alimentation de tables depuis Excel
- Création d’états et formulaires Access
- Reporting avec Power BI
- Analyse et restitution d’indicateurs

## Remarque sur les données

Les fichiers sources complets ne sont pas inclus dans ce dépôt afin d’éviter de publier des données sensibles ou trop volumineuses.

Le dépôt présente principalement la structure du projet, les scripts SQL, les procédures VBA, les éléments de documentation et les captures des résultats obtenus.

## Résultat

Le projet aboutit à une base de données Access structurée, alimentée automatiquement et exploitable à travers des requêtes SQL, des états, des formulaires et un tableau de bord Power BI.

L’ensemble permet de mieux suivre les contrats d’apprentissage, les entreprises partenaires et les versements de taxe d’apprentissage liés à la relation entreprise.
