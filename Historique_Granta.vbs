' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

Const FICHIER       = "Historique_Granta.vbs"
Const DESCRIPTION   = "Retourne l'historique des entrées/sortie dans Granta pour un utilisateur."
Const VERSION       = "3.1"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - On demande la date de début de l'historique.                       
' - On demande le nom de l'utilisateur (nom et prénom)                 
' - La requête est dans la constante REQUETE_SQL
' - Les informations de connexion sont dans la constante INFOS_CONNEXION
' - À enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la déclaration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo Historique_Granta.vbs > Historique_Granta.log


' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+
' 1) Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
' 2) Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
' 3) Const LISTE_DOSSIERS = "Z:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"

Const REQUETE_SQL     = "SELECT TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, (cast(TIMELOG32.LOGDATE as datetime)-2) as LOGDATE, TIMELOG32.LOGTIME, TIMELOG32.CARDNO, USER32.ID FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno');"
' Const REQUETE_SQL     = "SELET TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, (cast(TIMELOG32.LOGDATE as datetime)-2) as LOGDATE, TIMELOG32.LOGTIME, TIMELOG32.CARDNO, USER32.ID FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno');"
Const INFOS_CONNEXION = "DSN=granta;UID=sa;PWD=;Database=granta"
' Const INFOS_CONNEXION = "DSN=grant;UID=sa;PWD=;Database=granta"
Const DSN             = "granta"
Const UID             = "sa"
Const PWD             = ""
Const DATABASE        = "granta"

'+----------------------------------------------------------------------------+
'|                             PROGRAMME PRINCIPAL                            |
'+----------------------------------------------------------------------------+

Init
Main  ' ou Call Main()
Terminate



'+----------------------------------------------------------------------------+
'|                             PROCÉDURES/FONCTIONS                           |
'+----------------------------------------------------------------------------+

Sub Main()
  ' ------------------------------------------------------------
  ' -                        Variables                         -
  ' ------------------------------------------------------------
  LanceRequeteSQL INFOS_CONNEXION, REQUETE_SQL  

End Sub


'+----------------------------------------------------------------------------+
'| Nom          : LanceRequeteSQL                                             |
'| Description  : affiche les résultats de la requête strRequeteSQL sur le    |
'|                serveur défini par strConnexion.                            |
'| strConnexion : Les infos de connexion (ex: "DSN=d;UID=u;PWD=p;Database=db")|
'| strRequeteSQL : La requête SQL.                                            |
'+----------------------------------------------------------------------------+

Sub LanceRequeteSQL(strConnexion, strRequeteSQL)

  On Error Resume Next

  Dim Connection
  Dim Recordset
  Dim strLine
  
  'create an instance of the ADO connection and recordset objects
  Set Connection = WScript.CreateObject("ADODB.Connection")
  Set Recordset = WScript.CreateObject("ADODB.Recordset")
  
  'open the connection to the database
  Connection.Open strConnexion
  If Err.Number <> 0 Then
    WScript.Echo "Erreur lors de l'appel de la fonction Open (Numéro: " & Err.Number & ", Description: " &  Err.Description & ", Infos de connexion : " & strConnexion & ")"
    Err.Clear
  Else
    ' WScript.Echo("Connexion """ & strConnexion & """ : OK.")

    'Open the recordset object executing the SQL statement and return records
    Recordset.Open strRequeteSQL,Connection
    If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction Open (Numéro: " & Err.Number & ", Description: " &  Err.Description & ", Infos de connexion : " & strConnexion & ", Requête SQL : " & strRequeteSQL & ")"
      Err.Clear
    Else
      ' WScript.Echo("Recordset """ & strConnexion & ", Requête SQL : " & strRequeteSQL & """ : OK.")

      'first of all determine whether there are any records
      If Recordset.EOF Then
        ' On écrit dans le fichier
          WScript.Echo("No records returned.")
        'Response.Write("No records returned.")
      Else

        WScript.Echo("")
        
        Dim elt, strLigne
        strLigne = ""
        For each elt in Recordset.Fields
          strLigne = strLigne & elt.name & ";"
        Next
        WScript.Echo(strLigne)
        strLigne = ""
        ' WScript.Echo("DESCRIPTN;FORENAME;CARDHOLDER;LOGDATE;LOGTIME;CARDNO;ID")
        
        'if there are records then loop through the fields
        Do While NOT Recordset.Eof
          ' Recordset.MoveFirst '     => Revient au premier enregistrement
          ' Recordset.MoveNext        => Passe à l'enregistrement suivant
          ' Recordset.Fields.Count    => Nombre de champs
          ' Recordset.Fields(0)       => Contenu de l'élément 0 du recordset (ou Recordset.Fields(0).value)
          ' Recordset("DESCRIPTN")    => Contenu de l'élément correspondant à la colonne passée en paramètre (mais issu de la ligne du SELECT : ici DESCRIPTN).
          '                              ATTENTION !!! S'il n'y a pas de nom dans le select, il faut mettre "" pour retrouver l'élément ou utiliser les index.
          ' Recordset.Fields(0).name  => Nom de l'élément. Permet de connaître les noms par lesquels récupérer les valeurs.
          '                              Exemple : WScript.Echo(Recordset.Fields(0).name & ";" & Recordset.Fields(1).name & ";" & Recordset.Fields(2).name & ";" & Recordset.Fields(3).name & ";" & Recordset.Fields(4).name & ";" & Recordset.Fields(5).name & ";" & Recordset.Fields(6).name )
          
    
          ' DESCRIPTN;FORENAME;CARDHOLDER;;LOGTIME;CARDNO;ID        => si on ne met pas "as LOGDATE" dans le SELECT
          ' DESCRIPTN;FORENAME;CARDHOLDER;LOGDATE;LOGTIME;CARDNO;ID => si on met "as LOGDATE" dans le SELECT
    
          For each elt in Recordset.Fields
            strLigne = strLigne & elt.value & ";"
          Next
          WScript.Echo(strLigne)
          ' Récupération à partir de l'index : ' WScript.Echo(Recordset.Fields(0) & ";" & Recordset.Fields(1) & ";" & Recordset.Fields(2) & ";" & Recordset.Fields(3) & ";" & Recordset.Fields(4) & ";" & Recordset.Fields(5) & ";" & Recordset.Fields(6) )
          ' Récupération à partir du nom : ' WScript.Echo(Recordset("DESCRIPTN") & ";" & Recordset("FORENAME") & ";" & Recordset("CARDHOLDER") & ";" & Recordset("LOGDATE") & ";" & Recordset("LOGTIME") & ";" & Recordset("CARDNO") )' & ";" & Recordset("ID") )
          strLigne = ""
    
          Recordset.MoveNext
        Loop
  
      End If 'Fin test Recordset
    
    End If ' Fin test requête

  End If ' Fin test connexion
  
  
  
  'close the connection and recordset objects to free up resources
  Recordset.Close
  Set Recordset=nothing
  Connection.Close
  Set Connection=nothing

End Sub


'+----------------------------------------------------------------------------+
'| Nom         : Init                                                         |
'| Description : Affiche les informations sur le script.                      |
'|               Nom du script, version, auteur et date de création           |
'|               cf constantes : FICHIER, VERSION, AUTEUR et DATE_CREATION    |
'+----------------------------------------------------------------------------+

Sub Init()

  Banniere(FICHIER & " - " & VERSION & " - " & AUTEUR & " - " & DATE_CREATION)

End Sub


'+----------------------------------------------------------------------------+
'| Nom         : Terminate                                                    |
'| Description : Affiche la fin du script avec le nom complet.                |
'+----------------------------------------------------------------------------+

Sub Terminate()

  ' Banniere(" Fin du script   (" & WScript.ScriptFullName & ").")
  Banniere("")

End Sub


'+----------------------------------------------------------------------------+
'| Nom         : Banniere                                                     |
'| Description : Ecrit un message encadré entre 2 lignes.                     |
'| strMessage  : Le message à écrire.                                         |
'+----------------------------------------------------------------------------+

Sub Banniere(strMessage)

  Dim strTrace

  strTrace = vbCRLF
  strTrace = strTrace & "------------------------------------------------------------------------" & vbCRLF
  If strMessage <> "" Then
    strTrace = strTrace & "  " & strMessage & vbCRLF
  End If
  strTrace = strTrace & "------------------------------------------------------------------------" & vbCRLF
  strTrace = strTrace & vbCRLF

  WScript.Echo strTrace

End Sub

'+----------------------------------------------------------------------------+
'|                              FIN DU SCRIPT                                 |
'+----------------------------------------------------------------------------+

'+----------------------------------------------------------------------------+
'|                                   TESTS                                    |
'+----------------------------------------------------------------------------+
'|                                                                            |
'| 1) Tout est correct.                                                       |
'| 2) Les informations de connexion sont fausses (grant).                     |
'| 3) Le requête est mal formée (SELET).                                      |
'|                                                                            |
'+----------------------------------------------------------------------------+

' 1) Const REQUETE_SQL     = "SELECT TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, (cast(TIMELOG32.LOGDATE as datetime)-2) as LOGDATE, TIMELOG32.LOGTIME, TIMELOG32.CARDNO, USER32.ID FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno');"
' 1) Const INFOS_CONNEXION = "DSN=granta;UID=sa;PWD=;Database=granta"
' 2) Const REQUETE_SQL     = "SELECT TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, (cast(TIMELOG32.LOGDATE as datetime)-2) as LOGDATE, TIMELOG32.LOGTIME, TIMELOG32.CARDNO, USER32.ID FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno');"
' 2) Const INFOS_CONNEXION = "DSN=grant;UID=sa;PWD=;Database=granta"
' 3) Const REQUETE_SQL     = "SELET TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, (cast(TIMELOG32.LOGDATE as datetime)-2) as LOGDATE, TIMELOG32.LOGTIME, TIMELOG32.CARDNO, USER32.ID FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno');"
' 3) Const INFOS_CONNEXION = "DSN=granta;UID=sa;PWD=;Database=granta"
