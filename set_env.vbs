' Force la d�claration des variables : on est oblig� de faire `Dim Variable`
Option Explicit

Const FICHIER       = "set_env.vbs"
Const DESCRIPTION   = "Change les variables d'environnement TEMP et TMP pour l'installation d'un ordinateur EXPANSIA."
Const VERSION       = "3.0"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - SYSTEM TEMP = C:\Temp
' - SYSTEM TMP  = C:\Temp
' - USER TEMP   = C:\Temp
' - USER TEMP   = C:\Temp 
' - � enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la d�claration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo set_env.vbs > set_env.log


' Emp�che les erreurs de s'afficher (� supprimer lors du d�bogage)
' Doit �tre ajout� dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+

'+----------------------------------------------------------------------------+
'|                             PROGRAMME PRINCIPAL                            |
'+----------------------------------------------------------------------------+

Init
Main  ' ou Call Main()
Terminate



'+----------------------------------------------------------------------------+
'|                             PROC�DURES/FONCTIONS                           |
'+----------------------------------------------------------------------------+

Sub Main()

call changeVariableEnv("SYSTEM", "TMP", "C:\Temp")
call changeVariableEnv("SYSTEM", "TEMP", "C:\Temp")
call changeVariableEnv("USER", "TMP", "C:\Temp")
call changeVariableEnv("USER", "TEMP", "C:\Temp")

End Sub


'+----------------------------------------------------------------------------+
'| Nom          : changeVariableEnv                                           |
'| Description  : Modifie la variable d'environement.                         |
'| strCategorie : Cat�gorie de la variable (system, user, etc...).            |
'| strNom       : Nom de la variable.                                         |
'| strValeur    : Nouvelle valeur de la variable.                             |
'+----------------------------------------------------------------------------+

Sub changeVariableEnv(strCategorie, strNom, strValeur)

	Dim wshShell, wshEnv

	Set wshShell = CreateObject( "WScript.Shell" )
	Set wshEnv = wshShell.Environment( strCategorie )
	' Display the current value
	WScript.Echo "[Avant] - " & strCategorie & ":" & strNom & " = " & wshEnv( strNom )
	
	' Set the environment variable
	wshEnv( strNom ) = strValeur
	' Display the result
	WScript.Echo "[Apr�s] - " & strCategorie & ":" & strNom & " = " & wshEnv( strNom )
	
	Set wshEnv    = Nothing
	Set wshShell  = Nothing

End Sub



'+----------------------------------------------------------------------------+
'| Nom         : Init                                                         |
'| Description : Affiche les informations sur le script.                      |
'|               Nom du script, version, auteur et date de cr�ation           |
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
'| Description : Ecrit un message encadr� entre 2 lignes.                     |
'| strMessage  : Le message � �crire.                                         |
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
'| 1) Il n'y a pas vraiment de tests � faire. Il faut juste bien v�rifier :   |
'|      - la cat�gorie                                                        |
'|      - le nom de la variable d'environement                                |
'|      - la valeur de la variable d'environement                             |
'|                                                                            |
'+----------------------------------------------------------------------------+
