' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

Const FICHIER       = "DECA_Verif_Sauvegarde.vbs"
Const DESCRIPTION   = "Vérifie que la sauvegarde de DECA a bien été effectuée."
Const VERSION       = "3.0"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - Le nom du fichier à vérifier est dans la constante FICHIER_A_VERIFIER 
' - À enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la déclaration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo DECA_Verif_Sauvegarde.vbs > DECA_Verif_Sauvegarde.log


' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+
' Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Granta.txt"
' Const FICHIER_A_VERIFIER = "C:\Users\brb6301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Granta.txt"
' Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Grant.txt"
' Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\LIMS_Sauvegardes_v1.vbs"
Const FICHIER_A_VERIFIER = "E:\SAUVEGARDES\DECAv7.10.32\BaseDECA.bak"

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

dim objFSO
dim dossier, fichier, dossierExiste, fichierExiste
dim dateFichier, dateHier
dim listeErreurs


' ------------------------------------------------------------
' -                     Initialisations                      -
' ------------------------------------------------------------

listeErreurs  = ""
dossier       = CheminDossierParent(FICHIER_A_VERIFIER)
fichier       = NomFichierSansChemin(FICHIER_A_VERIFIER)

' ------------------------------------------------------------
' -           Contrôle du dossier parent DECA                -
' ------------------------------------------------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")
dossierExiste = objFSO.FolderExists(dossier)

If dossierExiste Then

	' ------------------------------------------------------------
	' -                Test du fichier BaseDECA.bak              -
	' ------------------------------------------------------------

	fichierExiste = objFSO.FileExists(FICHIER_A_VERIFIER)

	if fichierExiste Then
		
		' - Test de la date de dernière modification
			
		dateFichier = DateDerniereModificationFichier(FICHIER_A_VERIFIER)
		
		dateHier = DateAdd("d",-1,Date) 'd: jour ; -1: un jour en moins; Date: la date à modifier
		
		If Not IsEmpty(dateFichier) Then
			if StrComp(dateFichier, dateHier) = 0 Then
			   'WScript.Echo "Les dates sont identiques"
         WScript.Echo "Sauvegarde DECA OK."
			else
			   'WScript.echo "Les dates ne sont pas identiques"
			   listeErreurs = "**ERREUR** : Le fichier " & fichier & " (" & dateFichier & ") n'est pas à la date d'hier (" & dateHier &")."
			end if
    Else
      listeErreurs = "**ERREUR** : La date du fichier " & fichier & " n'a pas pu être lue."
		End If
	Else
		listeErreurs = "**ERREUR** : Le fichier " & fichier & " n'existe pas."
	End If


Else
  listeErreurs = "**ERREUR** : Le dossier " & dossier & " n'existe pas."
end If

If listeErreurs <> "" Then
  WScript.Echo(listeErreurs)
End If

set objFSO = Nothing

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
'| Nom         : DateDerniereModificationFichier                              |
'| Description : Renvoi la date de dernière modification de filespec.         |
'| filespec    : Chemin complet du fichier.                                   |
'| retour      : Une date ou Empty s'il y a eu une erreur.                    |
'+----------------------------------------------------------------------------+

Function DateDerniereModificationFichier(filespec)
   On Error Resume Next ' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
   Dim objFSO, objFile, retour, strErrMsg, result
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.GetFile(filespec)
   If Err.Number <> 0 Then
      strErrMsg = "Erreur lors de l'appel de la fonction GetFile." & vbNewLine & "(Numéro: " & Err.Number & ", Description: " & Err.Description & ", Fichier : " & filespec & ")"
      Err.Clear
      ' result = MsgBox (strErrMsg, vbOKOnly+vbExclamation, "DateDerniereModificationFichier.vbs")
      WScript.Echo strErrMsg
   Else
      retour = FormatDateTime(objFile.DateLastModified, 2) ' vbShortDate - 2 - Display a date using the short date format specified in your computer's regional settings.
   End If
   Set objFSO = Nothing
   Set objFile = Nothing
   DateDerniereModificationFichier = retour
End Function


'+----------------------------------------------------------------------------+
'| Nom           : CheminDossierParent                                        |
'| Description   : Renvoi le chemin de strCheminComplet (terminé par un "\"). |
'| strCheminComplet : Nom complet de fichier ou de dossier.                   |
'+----------------------------------------------------------------------------+

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
    ' WScript.Echo "Il y a déjà un antislash à la fin"
		CheminDossierParent = strCheminDossierParent
	Else
    ' WScript.Echo "Il faut ajouter un antislash à la fin"
		CheminDossierParent = strCheminDossierParent  & "\" 
	End If
End Function


'+----------------------------------------------------------------------------+
'| Nom           : NomFichierSansChemin                                       |
'| Description   : Renvoie le nom du fichier (+extension) sans le chemin.     |
'| strNomComplet : Le nom complet du fichier : chemin + nom + extension.      |
'+----------------------------------------------------------------------------+

Function NomFichierSansChemin(strNomComplet)

  Dim objFSO, fullpath
  
  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  fullpath = objFSO.GetFileName(strNomComplet)
  ' Pas besoin de vérification d'erreur car GetFileName ne travaille
  ' pas sur des fichiers mais sur une chaîne de caractère.
  Set objFSO = Nothing
  NomFichierSansChemin = fullpath

End Function



'+----------------------------------------------------------------------------+
'|                              FIN DU SCRIPT                                 |
'+----------------------------------------------------------------------------+


'+----------------------------------------------------------------------------+
'|                                   TESTS                                    |
'+----------------------------------------------------------------------------+
'|                                                                            |
'| 1) Tout est correct                                                        |
'| 2) Le dossier n'existe pas (modifier le chemin en supprimant une lettre)   |
'| 3) Le fichier n'existe pas (modifier le nom du fichier en supprimant une   |
'|    lettre.                                                                 |
'| 4) Le fichier n'est pas à la bonne date (prendre un fichier quelconque)    |
'|                                                                            |
'+----------------------------------------------------------------------------+

' 1) Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Granta.txt"
' 2) Const FICHIER_A_VERIFIER = "C:\Users\brb6301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Granta.txt"
' 3) Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Grant.txt"
' 4) Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\LIMS_Sauvegardes_v1.vbs"
