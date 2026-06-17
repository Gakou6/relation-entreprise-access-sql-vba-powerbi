Option Compare Database
Option Explicit

Public Sub Creation()

Dim chainesql As String
DoCmd.SetWarnings False

'================ SITE ================
chainesql = "CREATE TABLE SITE (Id_site COUNTER, nomSite VARCHAR(50) NOT NULL, adresse VARCHAR(120), ville VARCHAR(50), PRIMARY KEY(Id_site));"
DoCmd.RunSQL chainesql

'================ DEPARTEMENT ================
chainesql = "CREATE TABLE DEPARTEMENT (Id_Departement COUNTER, nomDept VARCHAR(50) NOT NULL, Id_site INT NOT NULL, PRIMARY KEY(Id_Departement), FOREIGN KEY(Id_site) REFERENCES SITE(Id_site));"
DoCmd.RunSQL chainesql

'================ FORMATION ================
chainesql = "CREATE TABLE FORMATION (Id_Formation COUNTER, nomFormation VARCHAR(50), TypeFormation VARCHAR(50), niveau VARCHAR(50), Id_Departement INT NOT NULL, PRIMARY KEY(Id_Formation), FOREIGN KEY(Id_Departement) REFERENCES DEPARTEMENT(Id_Departement));"
DoCmd.RunSQL chainesql

'================ ENTREPRISE ================
chainesql = "CREATE TABLE ENTREPRISE (SIRET VARCHAR(14), raison_social VARCHAR(120), adresse VARCHAR(120), ville VARCHAR(50), codePostal VARCHAR(5), secteur_activite VARCHAR(50), email VARCHAR(50), telephone VARCHAR(13), PRIMARY KEY(SIRET));"
DoCmd.RunSQL chainesql

'================ CAMPAGNE_TAXE ================
chainesql = "CREATE TABLE CAMPAGNE_TAXE (annee INT, date_debut DATE, date_fin DATE, objectif_montant VARCHAR(50), PRIMARY KEY(annee));"
DoCmd.RunSQL chainesql

'================ OPCO ================
chainesql = "CREATE TABLE OPCO (code_opco VARCHAR(50), nom_opco VARCHAR(50), secteur VARCHAR(50), PRIMARY KEY(code_opco));"
DoCmd.RunSQL chainesql

'================ CONVENTION_COLLECTIVE ================
chainesql = "CREATE TABLE CONVENTION_COLLECTIVE (IDCC INT, libelle VARCHAR(50), Description VARCHAR(150), PRIMARY KEY(IDCC));"
DoCmd.RunSQL chainesql

'================ VERSEMENT_TAXE ================
chainesql = "CREATE TABLE VERSEMENT_TAXE (Id_VersementTaxe COUNTER,Id_Formation INT, date_versement DATE, montant DOUBLE, mode_paiement VARCHAR(50), reference INT, montant_affecter DOUBLE, SIRET VARCHAR(14), annee INT, PRIMARY KEY(Id_VersementTaxe), FOREIGN KEY(SIRET) REFERENCES ENTREPRISE(SIRET), FOREIGN KEY(annee) REFERENCES CAMPAGNE_TAXE(annee), FOREIGN KEY(Id_Formation) REFERENCES FORMATION(Id_Formation));"
DoCmd.RunSQL chainesql

'================ CPNE ================
chainesql = "CREATE TABLE CPNE (code_cpne INT, nom_cpne VARCHAR(50), PRIMARY KEY(code_cpne));"
DoCmd.RunSQL chainesql

'================ CONTRAT_APPRENTISSAGE ================
chainesql = "CREATE TABLE CONTRAT_APPRENTISSAGE (Id_Contrat COUNTER, date_debut DATE, date_fin DATE, type_contrat VARCHAR(50), tempsTravailHebdomadaire INT, salaire_brut DOUBLE, date_signature DATE, priseChargeAnnuelle DOUBLE, priseChargeTotale DOUBLE, Id_Formation INT, SIRET VARCHAR(14),code_cpne INT, code_opco VARCHAR(50), IDCC INT, PRIMARY KEY(Id_Contrat),FOREIGN KEY(code_cpne) REFERENCES CPNE(code_cpne), FOREIGN KEY(Id_Formation) REFERENCES FORMATION(Id_Formation), FOREIGN KEY(SIRET) REFERENCES ENTREPRISE(SIRET), FOREIGN KEY(code_opco) REFERENCES OPCO(code_opco), FOREIGN KEY(IDCC) REFERENCES CONVENTION_COLLECTIVE(IDCC));"
DoCmd.RunSQL chainesql

'================ RESILIATION ================
chainesql = "CREATE TABLE RESILIATION (Id_Resiliation COUNTER, motif VARCHAR(50), date_resiliation DATE, Id_Contrat INT, PRIMARY KEY(Id_Resiliation), FOREIGN KEY(Id_Contrat) REFERENCES CONTRAT_APPRENTISSAGE(Id_Contrat));"
DoCmd.RunSQL chainesql



MsgBox "Les tables ont bien été créées !"

End Sub


Private Sub DropTables_Click()

Dim chainesql As String

DoCmd.SetWarnings False
On Error Resume Next

chainesql = "DROP TABLE RESILIATION;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE CONTRAT_APPRENTISSAGE;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE VERSEMENT_TAXE;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE FORMATION;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE DEPARTEMENT;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE ENTREPRISE;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE OPCO;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE CONVENTION_COLLECTIVE;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE CAMPAGNE_TAXE;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE SITE;"
DoCmd.RunSQL chainesql

chainesql = "DROP TABLE CPNE;"
DoCmd.RunSQL chainesql

DoCmd.SetWarnings True

MsgBox "Les tables sont bien retirées !"

End Sub


Option Compare Database
Option Explicit

Public Sub Implementation()

    Dim CONTRAT     As String
    Dim chainesql   As String

    CONTRAT = "C:\Users\menti\OneDrive\DEVOIR SQL\Desktop\PREMIERE ANNEE'\Semestre_2\Sae\BDD\Data\DONNEES ANONYMISES\CONTRAT\CONTRAT 2022-2025.xlsx"

    'Désactivation des messages d'avertissement Access
    DoCmd.SetWarnings False

    ' --- Import des fichiers Excel ---
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "CONTRAT 2022-2025", CONTRAT, True

    ' =====================================================
    ' INSERT ENTREPRISE
    ' =====================================================
    chainesql = "INSERT INTO ENTREPRISE " & _
                "(SIRET, raison_social, adresse, ville, codePostal) " & _
                "SELECT DISTINCT [Établissement d'exécution : SIRET], [Employeur signataire  : nom], [Lieu d'exécution du contrat : nom de rue],[Lieu d'exécution du contrat : ville], [Lieu d'exécution du contrat : code postal] " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Établissement d'exécution : SIRET] IS NOT NULL AND [Établissement d'exécution : SIRET] NOT IN(SELECT SIRET FROM ENTREPRISE)"
    DoCmd.RunSQL chainesql
    MsgBox "ENTREPRISE implémentée"

    ' =====================================================
    ' INSERT CONVENTION_COLLECTIVE
    ' =====================================================
    chainesql = "INSERT INTO CONVENTION_COLLECTIVE (IDCC, libelle) " & _
                "SELECT DISTINCT CInt([Idcc : code]), [Idcc : nom] " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Idcc : code] IS NOT NULL AND IsNumeric([Idcc : code]) = True AND CInt([Idcc : code]) NOT IN(SELECT IDCC FROM CONVENTION_COLLECTIVE)"
    DoCmd.RunSQL chainesql
    MsgBox "CONVENTION_COLLECTIVE implémentée"

    ' =====================================================
    ' INSERT OPCO
    ' =====================================================
    chainesql = "INSERT INTO OPCO (code_opco, nom_opco) " & _
                "SELECT DISTINCT [Opco : code], [Opco : nom] " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Opco : code] IS NOT NULL AND [Opco : code] NOT IN(SELECT code_opco FROM OPCO)"
    DoCmd.RunSQL chainesql
    MsgBox "OPCO implémentée"

    ' =====================================================
    ' INSERT CPNE
    ' =====================================================
    chainesql = "INSERT INTO CPNE (code_cpne, nom_cpne) " & _
                "SELECT DISTINCT CInt([Cpne : code]), [Cpne : nom] " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Cpne : code] IS NOT NULL AND IsNumeric([Cpne : code]) = True AND CInt([Cpne : code]) NOT IN (SELECT code_cpne FROM CPNE)"
    DoCmd.RunSQL chainesql
    MsgBox "CPNE implémentée"

    ' =====================================================
    ' INSERT RESILIATION
    ' =====================================================
    chainesql = "INSERT INTO RESILIATION (motif) " & _
                "SELECT DISTINCT [Résiliation : nom] " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Résiliation : code] IS NOT NULL AND [Résiliation : nom] NOT IN (SELECT motif FROM RESILIATION WHERE motif IS NOT NULL)"
    DoCmd.RunSQL chainesql
    MsgBox "RESILIATION implémentée"

    ' =====================================================
    ' INSERT CONTRAT_APPRENTISSAGE
    ' =====================================================
    chainesql = "INSERT INTO CONTRAT_APPRENTISSAGE " & _
                "(date_debut, date_fin, type_contrat, tempsTravailHebdomadaire, salaire_brut, " & _
                "SIRET, Id_Formation, IDCC, code_opco) " & _
                "SELECT [Début du contrat], [Fin du contrat], IIf([Contrat précédent]='Non','Nouveau','Succession'), [Temps de travail hebdomaire (minutes)], [Salaire brut mensuel à l'embauche versé à l'apprentie], CStr([Établissement d'exécution : SIRET]), " & _
                "(SELECT Id_Formation FROM FORMATION WHERE nomFormation = [CONTRAT 2022-2025].[Formation : nom]), " & _
                "IIf(IsNumeric([Idcc : code]), CInt([Idcc : code]), Null), [Opco : code] " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Établissement d'exécution : SIRET] IS NOT NULL"
    DoCmd.RunSQL chainesql
    MsgBox "CONTRAT_APPRENTISSAGE implémentée"

    ' =====================================================
    ' INSERT CAMPAGNE_TAXE
    ' =====================================================
    chainesql = "INSERT INTO CAMPAGNE_TAXE " & _
                "(annee, date_debut, date_fin) " & _
                "SELECT Year([Début du contrat]), Min([Début du contrat]),Max([Fin du contrat]) " & _
                "FROM [CONTRAT 2022-2025] " & _
                "WHERE [Début du contrat] IS NOT NULL AND [Fin du contrat] IS NOT NULL AND Year([Début du contrat]) NOT IN (SELECT annee FROM CAMPAGNE_TAXE) " & _
                "GROUP BY Year([Début du contrat])"
    DoCmd.RunSQL chainesql
    MsgBox "CAMPAGNE_TAXE implémentée"

    ' --- Suppression des tables temporaires ---
    DoCmd.DeleteObject acTable, "CONTRAT 2022-2025"


    ' Réactivation des messages d'avertissement
    DoCmd.SetWarnings True

    MsgBox "Implémentation complète terminée avec succès !", vbInformation

End Sub

' =====================================================
' VIDER TOUTES LES TABLES
' =====================================================
Public Sub vide_tables()
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM CONTRAT_APPRENTISSAGE"
    DoCmd.RunSQL "DELETE FROM RESILIATION"
    DoCmd.RunSQL "DELETE FROM VERSEMENT_TAXE"
    DoCmd.RunSQL "DELETE FROM CAMPAGNE_TAXE"
    DoCmd.RunSQL "DELETE FROM ENTREPRISE"
    DoCmd.RunSQL "DELETE FROM CONVENTION_COLLECTIVE"
    DoCmd.RunSQL "DELETE FROM OPCO"
    DoCmd.RunSQL "DELETE FROM CPNE"
    DoCmd.SetWarnings True
    MsgBox "Toutes les tables ont été vidées !", vbInformation
End Sub


