'+----------------------------------------------------------------------------+
'| Fichier     : set_env.vbs                                                  |
'+----------------------------------------------------------------------------+
'| Version     : 2.0                                                          |
'+----------------------------------------------------------------------------+
'| Description :                                                              |
'|                                                                            |
'| Script qui change les variables d'environnement TEMP et TMP.               |
'|                                                                            |
'|    - SYSTEM TEMP = C:\Temp                                                 |
'|    - SYSTEM TMP  = C:\Temp                                                 |
'|    - USER TEMP   = C:\Temp                                                 |
'|    - USER TEMP   = C:\Temp                                                 |
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


Sub Main

call changeVariableEnv("SYSTEM", "TMP", "C:\Temp")
call changeVariableEnv("SYSTEM", "TEMP", "C:\Temp")
call changeVariableEnv("USER", "TMP", "C:\Temp")
call changeVariableEnv("USER", "TEMP", "C:\Temp")

End Sub


Sub changeVariableEnv(strCategorie, strNom, strValeur)

	Dim wshShell, wshEnv

	Set wshShell = CreateObject( "WScript.Shell" )
	Set wshEnv = wshShell.Environment( strCategorie )
	' Display the current value
	WScript.Echo strCategorie & ":" & strNom & " = " & wshEnv( strNom )
	
	' Set the environment variable
	wshEnv( strNom ) = strValeur
	' Display the result
	WScript.Echo strCategorie & ":" & strNom & " = " & wshEnv( strNom )
	
	Set wshEnv    = Nothing
	Set wshShell  = Nothing

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