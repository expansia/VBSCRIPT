'+----------------------------------------------------------------------------+
'| Fichier     : Ouvrir_Dossier.vbs                                           |
'+----------------------------------------------------------------------------+
'| Version     : 3.1                                                          |
'+----------------------------------------------------------------------------+
'| Description :                                                              |
'|                                                                            |
'| Script qui crée les dossiers nécessaires sur D.                            |
'|                                                                            |
'|    - Il faut renseigner les variables strDossier*                          |
'|    - Dans ce dossier le fichier EXPANSIA.erv doit avoir la date de la      |
'|      veille.                                                               |
'|    - Un fichier trace (constante FICHIER_TRACE) situé dans le dossier du   |
'|      script permet de vérifier ce qu'il s'est passé.                       |
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

	Dim tableauDesNomsDeDossier(3), i
	
	tableauDesNomsDeDossier(0) = "D:\INFORMATIQUE\"
	tableauDesNomsDeDossier(1) = "D:\MesDocuments\"
	tableauDesNomsDeDossier(2) = "D:\modele\"
	tableauDesNomsDeDossier(3) = "D:\PERSONNEL\"
	
	For i = 0 to 3
		
		CreerDossier( tableauDesNomsDeDossier(i) )
	
	Next

End Sub




'------------------------------------------------------------------------------
'                                PROCEDURES
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Nom            : CreerDossier
' Description    : Crée un dossier à partir du chemin complet strNomDossier
' strNomDossier  : Chemin complet du dossier
'------------------------------------------------------------------------------


Sub CreerDossier(strNomDossier)

	Dim oFSO
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If Not oFSO.FolderExists(strNomDossier) Then
	  oFSO.CreateFolder strNomDossier
	Else
		result = MsgBox ("Le dossier """ & strNomDossier & """ existe déjà.", _
		vbOK+vbExclamation, strTitre)
	End If
	
	set oFSO = Nothing

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