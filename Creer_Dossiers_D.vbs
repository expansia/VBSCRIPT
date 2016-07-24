' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

Const FICHIER       = "Creer_Dossiers_D.vbs"
Const DESCRIPTION   = "Crée sur D: les dossiers nécessaires à l'installation d'un ordinateur EXPANSIA."
Const VERSION       = "3.2"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - Les noms des dossiers sont dans la constante LISTE_DOSSIERS, séparés par une virgule 
' - À enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la déclaration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo Liste_Montage_Disques_Reseau.vbs > Liste_Montage_Disques_Reseau.log


' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+
' Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
Const LISTE_DOSSIERS = "D:\INFORMATIQUE\,D:\MesDocuments\,D:\modele\,D:\PERSONNEL\,D:\Data\"

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

	' Dim tableauDesNomsDeDossier(3), i
  Dim tableauDesNomsDeDossier, nomDossier

  tableauDesNomsDeDossier = Split(LISTE_DOSSIERS, ",")
	
  For Each nomDossier in tableauDesNomsDeDossier
		
    CreerDossier( nomDossier )
	
	Next

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
'| Nom         : CreerDossier                                                 |
'| Description : Crée un dossier à partir du chemin complet strNomDossier.    |
'| strNomDossier  : Chemin complet du dossier.                                |
'+----------------------------------------------------------------------------+

Sub CreerDossier(strNomDossier)

	Dim oFSO
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If Not oFSO.FolderExists(strNomDossier) Then
	  oFSO.CreateFolder strNomDossier
    WScript.Echo("Création du dossier """ & strNomDossier & """ : OK.")
	Else
		WScript.Echo("Le dossier """ & strNomDossier & """ existe déjà.")
	End If
	
	set oFSO = Nothing

End Sub


'+----------------------------------------------------------------------------+
'|                              FIN DU SCRIPT                                 |
'+----------------------------------------------------------------------------+
