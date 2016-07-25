'+----------------------------------------------------------------------------+
'| Fichier     : Historique_Granta.vbs                                        |
'+----------------------------------------------------------------------------+
'| Version     : 1.0                                                          |
'+----------------------------------------------------------------------------+
'| Description :                                                              |
'|                                                                            |
'| Retourne l'historique des entrées/sortie dans Granta pour un utilisateur.  |
'|                                                                            |
'|    - On demande la date de début de l'historique.                          |
'|    - On demande le nom de l'utilisateur (nom et prénom)                    |
'|    - On récupère l'historique dans un fichier (historique_Granta.txt)      |
'|    - Une fenêtre s'affiche à la fin du script pour dire si la sauvegarde   |
'|      a réussi ou non.                                                      |
'+----------------------------------------------------------------------------+


' Force la déclaration des variables : on est obligé de faire : `Dim Variable`
Option Explicit

' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
'On Error Resume Next

Const TITRE_FENETRE = "Création des dossiers sur D:"

Init
Main
Terminate


'------------------------------------------------------------------------------
'                            PROGRAMME PRINCIPAL
'------------------------------------------------------------------------------


Sub Main()

  ' ------------------------------------------------------------
  ' -                        Variables                         -
  ' ------------------------------------------------------------

  Dim Connection
  Dim Recordset
  Dim SQL
  Dim strLine, strCheminFichierHistorique
  
  
  Const NOM_FICHIER = "historique_Granta.txt"
  
  ' Chemin complet du fichier des adresses mac
  strCheminFichierHistorique = CheminDossierParent(Wscript.ScriptFullName) & NOM_FICHIER
  
  
  'declare the SQL statement that will query the database
  'SQL = "SELECT InitialesPersonnel,NomPersonnel,PrenomPersonnel FROM dbo.ENVPersonnel WHERE NOMPERSONNEL='DUGUE'"
  SQL = "SELECT TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, cast(TIMELOG32.LOGDATE as datetime)-2, TIMELOG32.LOGTIME, TIMELOG32.CARDNO, USER32.ID FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno') ORDER BY USER32.NAME;"
  ' SQL = "SELECT TIMELOG32.DESCRIPTN, TIMELOG32.FORENAME, TIMELOG32.CARDHOLDER, cast(TIMELOG32.LOGDATE as datetime)-2, TIMELOG32.LOGTIME, TIMELOG32.CARDNO FROM Granta.dbo.TIMELOG32 TIMELOG32, Granta.dbo.USER32 USER32 WHERE TIMELOG32.CARDHOLDER = USER32.NAME AND TIMELOG32.FORENAME = USER32.FIRSTNAME AND (TIMELOG32.LOGDATE>42488 AND USER32.NAME = 'BOISSONNET' AND USER32.FIRSTNAME = 'Bruno') ORDER BY USER32.NAME;"


  'create an instance of the ADO connection and recordset objects
  Set Connection = WScript.CreateObject("ADODB.Connection")
  Set Recordset = WScript.CreateObject("ADODB.Recordset")
  
  'open the connection to the database
  Connection.Open "DSN=granta;UID=sa;PWD=;Database=granta"
  
  'Open the recordset object executing the SQL statement and return records
  Recordset.Open SQL,Connection
  
  'first of all determine whether there are any records
  If Recordset.EOF Then
  	' On écrit dans le fichier
      call EcritDansFichier(strCheminFichierHistorique, "No records returned.")
  	'Response.Write("No records returned.")
  Else
  call EcritDansFichier(strCheminFichierHistorique, "Description;FORENAME;CARDHOLDER;LOGDATE;LOGTIME;CARDNO;ID")
  'if there are records then loop through the fields
  	Do While NOT Recordset.Eof
  		' Recordset.MoveFirst
      ' WScript.Echo("Nombre de champs : " & Recordset.Fields.Count)
      ' Exit Do
      ' WScript.Echo(Recordset.Fields(0) & " " & Recordset.Fields(1) & " " & Recordset.Fields(2) & " " & Recordset.Fields(3) & " " & Recordset.Fields(4) & " " & Recordset.Fields(5) & " " & Recordset.Fields(6))
      ' Exit Do
      call EcritDansFichier(strCheminFichierHistorique, Recordset.Fields(0) & ";" & Recordset.Fields(1) & ";" & Recordset.Fields(2) & ";" & Recordset.Fields(3) & ";" & Recordset.Fields(4) & ";" & Recordset.Fields(5) & ";" & Recordset.Fields(6) )
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.DESCRIPTN") & ";" & Recordset("TIMELOG32.FORENAME") & ";" & Recordset("TIMELOG32.CARDHOLDER") & ";" & Recordset("TIMELOG32.LOGDATE") & ";" & Recordset("TIMELOG32.LOGTIME") & ";" & Recordset("TIMELOG32.CARDNO") )' & ";" & Recordset("USER32.ID") )
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.DESCRIPTN"))
        ' WSCript.Echo(Recordset("Granta.dbo.TIMELOG32.DESCRIPTN"))
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.FORENAME"))
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.CARDHOLDER"))
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.LOGDATE"))
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.LOGTIME"))
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("TIMELOG32.CARDNO"))
  		' call EcritDansFichier(strCheminFichierHistorique, Recordset("USER32.ID"))
  		call EcritDansFichier(strCheminFichierHistorique, "")
  		Recordset.MoveNext
  	Loop
  End If
  
  'close the connection and recordset objects to free up resources
  Recordset.Close
  Set Recordset=nothing
  Connection.Close
  Set Connection=nothing

End Sub


'------------------------------------------------------------------------------
' Nom         : Init
' Description : Ecrit un repère de début du script dans le fichier de trace.
'------------------------------------------------------------------------------

Sub Init()

  'Dim fichierTrace
  '
  'fichierTrace  = CheminDossierParent(WScript.ScriptFullName) & FICHIER_TRACE
  '
  'call Tracer(fichierTrace, "")
  'call Tracer(fichierTrace, "------------------------------------------------------------------------")
  'call Tracer(fichierTrace, " Début du script   (" & WScript.ScriptFullName & ").")
  'call Tracer(fichierTrace, "------------------------------------------------------------------------")
  'call Tracer(fichierTrace, "")

  Dim result
  result = MsgBox ("Début du script", vbOK+vbExclamation, TITRE_FENETRE)

End Sub


'------------------------------------------------------------------------------
' Nom         : Terminate
' Description : Ecrit un repère de fin de script dans le fichier de trace.
'------------------------------------------------------------------------------

Sub Terminate()

  'Dim fichierTrace, objShell
  '
  'fichierTrace  = CheminDossierParent(WScript.ScriptFullName) & FICHIER_TRACE
  '
  'call Tracer(fichierTrace, "")
  'call Tracer(fichierTrace, "------------------------------------------------------------------------")
  'call Tracer(fichierTrace, " Fin du script   (" & WScript.ScriptFullName & ").")
  'call Tracer(fichierTrace, "------------------------------------------------------------------------")
  'call Tracer(fichierTrace, "")
  '
  'If erreurTrouvee Then
  '  WScript.echo "Script terminé avec des erreurs !"
  '  Set objShell = CreateObject("Wscript.Shell")
  '  objShell.Run "notepad.exe " & fichierTrace
  '  set objShell = Nothing
  'Else
  '  WScript.echo "Script terminé avec succès !"
  'end if

  Dim result
  result = MsgBox ("Fin du script", vbOK+vbExclamation, TITRE_FENETRE)

End Sub

'------------------------------------------------------------------------------
' Nom              : CheminDossierParent.
' strCheminComplet : chemin complet du fichier.
' retour           : Le chemin du dossier parent terminé par un "\".
'------------------------------------------------------------------------------
Public Function CheminDossierParent(strCheminComplet)
    On Error Resume Next
    Dim objFSO, strCheminDossierParent, fin

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strCheminDossierParent = objFSO.GetParentFolderName(strCheminComplet)
    ' Pas besoin de vérification d'erreur car GetParentFolderName ne travaille
    ' pas sur des fichiers mais sur une chaîne de caractère.

    Set objFSO = Nothing
    ' On ajoute une barre oblique inversée au cas où il n'y en aurait pas
    fin = Right(strCheminDossierParent, 1)
    if fin = "\" Then
        CheminDossierParent = strCheminDossierParent
    Else
        CheminDossierParent = strCheminDossierParent  & "\"
    End If
End Function



'------------------------------------------------------------------------------
' Nom                          : EcritDansFichier.
' strCheminCompletFichierTrace : Chemin complet du fichier.
' strTrace                     : Ce qu'il faut écrire dans le fichier.
'------------------------------------------------------------------------------
Public Sub EcritDansFichier(strCheminCompletFichierTrace, strTrace)
    On Error Resume Next
    Dim objFSO, objFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strCheminCompletFichierTrace, 8, True, -1) ' 8 = ForAppending, True pour créer le fichier s'il n'existe pas, -1 pour écrire au format Unicode

    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de l'appel de la fonction OpenTextFile." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
        Err.Clear
    Else
        ' On écrit dans le fichier
        objFile.WriteLine strTrace

        If Err.Number <> 0 Then
            WScript.Echo "Erreur lors de l'appel de la fonction WriteLine." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
            Err.Clear
        End If

        ' On ferme le fichier
        objFile.Close
        Set objFile = Nothing
    End If

    Set objFSO = Nothing

End Sub
