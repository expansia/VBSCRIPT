' Force la d�claration des variables : on est oblig� de faire `Dim Variable`
Option Explicit

Const FICHIER       = "Creer_Dossiers_D.vbs"
Const DESCRIPTION   = "Cr�e sur D: les dossiers n�cessaires � l'installation d'un ordinateur EXPANSIA."
Const VERSION       = "3.3"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - Les noms des dossiers sont dans la constante LISTE_DOSSIERS, s�par�s par une virgule 
' - � enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la d�claration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo Creer_Dossiers_D.vbs > Creer_Dossiers_D.log


' Emp�che les erreurs de s'afficher (� supprimer lors du d�bogage)
' Doit �tre ajout� dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+
' 1) Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
' 2) Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
' 3) Const LISTE_DOSSIERS = "Z:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"

Const LISTE_DOSSIERS = "D:\INFORMATIQUE\,D:\MesDocuments\,D:\modele\,D:\PERSONNEL\,D:\Data\"

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
'| Nom         : CreerDossier                                                 |
'| Description : Cr�e un dossier � partir du chemin complet strNomDossier.    |
'| strNomDossier  : Chemin complet du dossier.                                |
'+----------------------------------------------------------------------------+

Sub CreerDossier(strNomDossier)
  On Error Resume Next
	Dim oFSO
	
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	
	If Not oFSO.FolderExists(strNomDossier) Then
	  oFSO.CreateFolder strNomDossier
    If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction CreateFolder (Num�ro: " & Err.Number & ", Description: " &  Err.Description & ", Dossier : " & strNomDossier & ")"
      Err.Clear
    Else
      WScript.Echo("Cr�ation du dossier """ & strNomDossier & """ : OK.")
    End If
	Else
		WScript.Echo("Le dossier """ & strNomDossier & """ existe d�j�.")
	End If
	
	set oFSO = Nothing

End Sub


'+----------------------------------------------------------------------------+
'|                              FIN DU SCRIPT                                 |
'+----------------------------------------------------------------------------+

'+----------------------------------------------------------------------------+
'|                                   TESTS                                    |
'+----------------------------------------------------------------------------+
'|                                                                            |
'| 1) Tout est correct.                                                       |
'| 2) Le dossier existe d�j�. (r�p�ter l'op�ration pr�c�dente).               |
'| 3) Le dossier n'existe pas (modifier le chemin en supprimant une lettre)   |
'|                                                                            |
'+----------------------------------------------------------------------------+

' 1) Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
' 2) Const LISTE_DOSSIERS = "D:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"
' 3) Const LISTE_DOSSIERS = "Z:\INFORMATIQUE1\,D:\MesDocuments1\,D:\modele1\,D:\PERSONNEL1\,D:\Data1\"

