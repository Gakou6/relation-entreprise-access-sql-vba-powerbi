-- ============================================================
-- SAÉ 2.01 – Relation Entreprise – IUT de Roubaix
-- AMENTI Delali Merveille – BUT Sciences des Données
-- Fichier : requetes_relation_entreprise.sql
-- Toutes les requêtes SQL du rapport (Access SQL)
-- ============================================================


-- ============================================================
-- A. REQUÊTES ANALYTIQUES
-- ============================================================


-- A.1 – Analyse croisée : Taux de rétention des contrats par OPCO
-- Analyse quels OPCO s'en sortent le mieux en comparant le volume
-- de contrats signés par rapport au volume de résiliations.
-- ============================================================
SELECT
    O.Nom_OPCO,
    COUNT(C.Id_Contrat)                                          AS Nb_Contrats,
    SUM(IIF(C.Code_Resiliation IS NOT NULL, 1, 0))               AS Nb_Resiliations,
    FORMAT(
        (COUNT(C.Id_Contrat) - SUM(IIF(C.Code_Resiliation IS NOT NULL, 1, 0)))
        / COUNT(C.Id_Contrat) * 100,
    '0.0') & ' %'                                                AS Taux_Retention
FROM ((OPCO AS O
    INNER JOIN ENTREPRISE AS E   ON O.Code_OPCO = E.Code_OPCO)
    INNER JOIN ETABLISSEMENT AS ET ON E.SIREN    = ET.SIREN)
    INNER JOIN CONTRAT AS C        ON ET.SIRET   = C.SIRET
GROUP BY O.Nom_OPCO
ORDER BY COUNT(C.Id_Contrat) DESC;


-- A.2 – Analyse sectorielle : Durée moyenne des contrats avant rupture
-- Calcule combien de temps (en mois) s'écoule avant une résiliation,
-- selon la ville de l'entreprise.
-- ============================================================
SELECT
    E.Ville                                              AS Secteur_Geographique,
    COUNT(C.Id_Contrat)                                  AS Nb_Ruptures,
    AVG(DateDiff('m', C.Date_Debut, C.Date_Fin))         AS Duree_Moy_Avant_Rupture
FROM (CONTRAT AS C
    INNER JOIN ETABLISSEMENT AS ET ON C.SIRET   = ET.SIRET)
    INNER JOIN ENTREPRISE    AS E  ON ET.SIREN  = E.SIREN
WHERE C.Code_Resiliation IS NOT NULL
  AND C.Date_Fin         IS NOT NULL
GROUP BY E.Ville
ORDER BY AVG(DateDiff('m', C.Date_Debut, C.Date_Fin)) ASC;


-- A.3 – Analyse de l'offre de formation : Taux d'affectation des taxes par formation
-- Croise la table VERSEMENT_TAXE pour analyser comment la taxe est
-- fléchée et distribuée selon les formations.
-- ============================================================
SELECT
    Nom_Formation,
    SUM(Montant)                                                         AS Montant_total,
    ROUND(SUM(Montant) * 100 / (SELECT SUM(montant) FROM VERSEMENT_TAXE), 2) AS Taux_affectation
FROM FORMATION
    INNER JOIN VERSEMENT_TAXE ON FORMATION.Id_formation = VERSEMENT_TAXE.Id_formation
GROUP BY Nom_Formation
ORDER BY SUM(Montant) DESC;


-- A.4 – Total des versements par année fiscale et par fléchage
-- Synthèse du nombre et du montant des versements groupés
-- par année et par libellé de fléchage.
-- ============================================================
SELECT
    VERSEMENT_TAXE.Libelle_Flechage,
    COUNT(*)        AS Nb_versements,
    SUM(Montant)    AS total_versements
FROM VERSEMENT_TAXE
GROUP BY annee, Libelle_flechage
ORDER BY annee, Libelle_flechage;


-- ============================================================
-- B. AGRÉGATIONS PAR SIREN (PERFORMANCE DES ENTREPRISES)
-- ============================================================


-- B.1 – Top 10 des entreprises (SIREN) par montant total de taxe versée
-- Identifie les plus gros contributeurs financiers de l'IUT.
-- ============================================================
SELECT TOP 10
    raison_sociale,
    SUM(VERSEMENT_TAXE.montant) AS total_taxe
FROM VERSEMENT_TAXE
    INNER JOIN CALCUL_SIREN ON VERSEMENT_TAXE.SIRET = CALCUL_SIREN.SIRET
GROUP BY raison_sociale
ORDER BY SUM(VERSEMENT_TAXE.montant) DESC;


-- B.2 – Nombre de contrats et masse salariale brute par SIREN
-- Mesure le volume d'apprentis et l'investissement dans leurs salaires
-- par entreprise.
-- ============================================================
SELECT
    entreprise.raison_sociale,
    COUNT(CONTRAT.Id_Contrat)         AS nb_contrats,
    SUM(CONTRAT.Salaire_brut_mensuel) AS masse_salariale
FROM entreprise
    INNER JOIN CONTRAT ON entreprise.siret = CONTRAT.SIRET
GROUP BY entreprise.raison_sociale
ORDER BY SUM(CONTRAT.Salaire_brut_mensuel) DESC;


-- ============================================================
-- C. ANALYSES MULTI-ANNUELLES (ÉVOLUTION TEMPORELLE)
-- ============================================================


-- C.1 – Évolution annuelle du nombre de contrats et du taux de rupture
-- Indicateur clé pour suivre la santé des formations
-- et le taux de résiliation d'une année sur l'autre.
-- ============================================================
SELECT
    YEAR(CONTRAT.date_debut)                                                    AS Annee,
    COUNT(CONTRAT.Id_Contrat)                                                   AS NbContrats,
    COUNT(RESILIATION.Code_Resiliation)                                         AS NbResiliations,
    ROUND(
        (COUNT(RESILIATION.Code_Resiliation) / COUNT(CONTRAT.Id_Contrat)) * 100,
    2)                                                                          AS TauxRupture
FROM CONTRAT
    LEFT JOIN RESILIATION ON CONTRAT.Id_Contrat = RESILIATION.Id_Contrat
GROUP BY YEAR(CONTRAT.date_debut)
ORDER BY YEAR(CONTRAT.date_debut)
HAVING ROUND(
        (COUNT(RESILIATION.Code_Resiliation) / COUNT(CONTRAT.Id_Contrat)) * 100,
    2)  ;


-- C.2 – Suivi des objectifs de collecte de taxe par campagne
-- Compare le montant collecté par rapport aux objectifs fixés
-- dans la table CAMPAGNE_TAXE.
-- ============================================================
SELECT
    ct.annee,
    ct.objectif_montant,
    Nz(SUM(v.Montant), 0)                                                      AS total_collecte,
    IIf(Nz(SUM(v.Montant), 0) >= ct.objectif_montant, "Atteint", "Non atteint") AS objectif_atteint
FROM CAMPAGNE_TAXE AS ct
    LEFT JOIN VERSEMENT_TAXE AS v ON ct.annee = YEAR(v.Date_versement)
GROUP BY ct.annee, ct.objectif_montant
ORDER BY ct.annee;


-- ============================================================
-- D. INDICATEURS STRATÉGIQUES (PILOTAGE ET PRISE DE DÉCISION)
-- ============================================================


-- D.1 – Volume de contrats et masse salariale brute annuelle
-- Pilotage de l'activité globale par année.
-- ============================================================
SELECT
    YEAR(CONTRAT.Date_Debut)                   AS annee,
    COUNT(CONTRAT.Id_Contrat)                  AS nb_contrat,
    SUM(CONTRAT.Salaire_brut_mensuel * 12)     AS masse_salariale_brute_annuelle,
    AVG(CONTRAT.Salaire_brut_mensuel)          AS salaire_moyen_mensuel
FROM CONTRAT
GROUP BY YEAR(CONTRAT.Date_Debut)
ORDER BY YEAR(CONTRAT.Date_Debut);


-- D.2 – Taux de résiliation des contrats par formation
-- Indicateur stratégique pour mesurer le ciblage
-- et le suivi des formations.
-- ============================================================
SELECT
    f.nom_Formation,
    COUNT(c.Id_Contrat)          AS NbContrats,
    COUNT(r.Code_Resiliation)    AS NbResiliations,
    ROUND(
        IIf(
            COUNT(c.Id_Contrat) = 0,
            0,
            (COUNT(r.Code_Resiliation) / COUNT(c.Id_Contrat)) * 100
        ),
    2)                           AS TauxResiliation
FROM (FORMATION AS f
    LEFT JOIN CONTRAT    AS c ON f.Id_Formation = c.Id_Formation)
    LEFT JOIN RESILIATION AS r ON c.Id_Contrat  = r.Id_Contrat
GROUP BY f.nom_Formation;


-- D.3 – Coût moyen et reste à charge estimé par type de formation
-- Permet de savoir quelles formations sont les mieux financées
-- par les OPCO.
-- ============================================================
SELECT
    FORMATION.Type_Formation,
    COUNT(CONTRAT.Id_Contrat)                                         AS nb_contrats,
    AVG(CONTRAT.[prise_en charge_totale])                             AS cout_moyen,
    AVG(CONTRAT.[prise_en charge_totale] - CONTRAT.[prise_en charge_annuelle]) AS reste_a_charge_moyen
FROM FORMATION
    INNER JOIN CONTRAT ON FORMATION.Id_Formation = CONTRAT.Id_formation
GROUP BY FORMATION.Type_Formation, CONTRAT.[prise_en charge_annuelle]
ORDER BY AVG(CONTRAT.[prise_en charge_totale]);


-- D.4 – Top 5 des motifs de résiliation les plus fréquents
-- Identifie les causes d'échec des contrats d'apprentissage.
-- ============================================================
SELECT TOP 5
    MOTIF_RESILIATION.Libelle_Resiliation,
    COUNT(*) AS Nb_Occurrences
FROM MOTIF_RESILIATION
    INNER JOIN RESILIATION ON MOTIF_RESILIATION.Code_Resiliation = RESILIATION.Code_Resiliation
GROUP BY MOTIF_RESILIATION.Libelle_Resiliation
ORDER BY COUNT(*) DESC;


-- D.5 – Analyse géographique (Ville) des entreprises partenaires
-- Cartographie la provenance des entreprises qui recrutent
-- ou versent la taxe.
-- ============================================================
SELECT
    ville,
    COUNT(*) AS nb_entreprise
FROM ENTREPRISE
GROUP BY ville
ORDER BY COUNT(*) DESC;


-- ============================================================
-- E. REQUÊTES PARAMÉTRÉES
-- (saisie interactive au moment de l'exécution)
-- ============================================================


-- E.1 – Lister les contrats d'une entreprise spécifique
-- L'utilisateur saisit le nom (ou une partie) de la raison sociale.
-- ============================================================
SELECT
    CONTRAT.Id_Contrat,
    ENTREPRISE.raison_sociale,
    CONTRAT.Titre_mission,
    CONTRAT.Date_Debut,
    CONTRAT.Date_Fin
FROM ENTREPRISE
    INNER JOIN CONTRAT ON ENTREPRISE.SIRET = CONTRAT.Siret
WHERE ENTREPRISE.raison_sociale LIKE "*" & [Saisir le nom de l'entreprise :] & "*";


-- E.2 – Fiche d'un apprenant par nom
-- Affiche tous les contrats, formations et entreprises associés
-- à l'apprenant saisi.
-- ============================================================
SELECT
    A.Nom,
    A.Prenom,
    F.Nom_Formation,
    E.raison_sociale,
    C.Date_Debut,
    C.Date_Fin,
    C.Code_Resiliation
FROM (((CONTRAT AS C
    INNER JOIN APPRENANT  AS A  ON C.INE         = A.INE)
    INNER JOIN PARCOURS   AS P  ON C.Id_Parcours = P.Id_Parcours)
    INNER JOIN FORMATION  AS F  ON P.Code_RNCP   = F.Code_RNCP)
    INNER JOIN ETABLISSEMENT AS ET ON C.SIRET    = ET.SIRET
    INNER JOIN ENTREPRISE    AS E  ON ET.SIREN   = E.SIREN
WHERE A.Nom LIKE [Entrez le nom de l'apprenant] & "*"
ORDER BY C.Date_Debut DESC;


-- E.3 – Versements d'une entreprise sur une année fiscale donnée
-- L'utilisateur saisit le SIREN et l'année fiscale.
-- ============================================================
SELECT
    E.raison_sociale,
    V.Date_versement,
    V.Reference,
    V.Montant,
    V.Libelle_Flechage
FROM ENTREPRISE AS E
    INNER JOIN VERSEMENT_TAXE AS V ON E.SIREN = V.SIREN
WHERE E.SIREN = [Entrez le SIREN de l'entreprise]
  AND V.annee = [Entrez l'année fiscale (ex. 2024)]
ORDER BY V.Date_versement;


-- E.4 – Contrats actifs par formation et par année de parcours
-- Affiche les contrats en cours (sans résiliation) selon
-- la formation et l'année de début saisies.
-- ============================================================
SELECT
    F.Nom_Formation,
    P.Annee_Debut,
    A.Nom,
    A.Prenom,
    E.raison_sociale,
    C.Date_Debut,
    C.Date_Fin,
    C.Salaire_brut_mensuel
FROM (((CONTRAT AS C
    INNER JOIN APPRENANT  AS A  ON C.INE         = A.INE)
    INNER JOIN PARCOURS   AS P  ON C.Id_Parcours = P.Id_Parcours)
    INNER JOIN FORMATION  AS F  ON P.Code_RNCP   = F.Code_RNCP)
    INNER JOIN ETABLISSEMENT AS ET ON C.SIRET    = ET.SIRET
    INNER JOIN ENTREPRISE    AS E  ON ET.SIREN   = E.SIREN
WHERE F.Nom_Formation LIKE [Entrez le nom de la formation] & "*"
  AND P.Annee_Debut   = [Entrez l'année de début de parcours]
  AND C.Code_Resiliation IS NULL
ORDER BY A.Nom;


-- E.5 – Bilan des versements d'un OPCO entre deux années
-- Synthèse du volume financier fléché vers l'IUT par un opérateur
-- sur une plage d'années saisie.
-- ============================================================
SELECT
    O.Nom_OPCO,
    V.annee,
    COUNT(V.Id_Versement) AS Nb_Versements,
    SUM(V.Montant)        AS Total_Verse
FROM OPCO AS O
    INNER JOIN ENTREPRISE    AS E  ON O.Code_OPCO = E.Code_OPCO
    INNER JOIN VERSEMENT_TAXE AS V ON E.SIREN     = V.SIREN
WHERE O.Nom_OPCO LIKE [Entrez le nom de l'OPCO] & "*"
  AND V.annee BETWEEN [Année de début] AND [Année de fin]
GROUP BY O.Nom_OPCO, V.annee
ORDER BY V.annee;











































