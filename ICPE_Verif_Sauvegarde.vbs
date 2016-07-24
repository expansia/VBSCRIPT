' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

Const FICHIER       = "ICPE_Verif_Sauvegarde.vbs"
Const DESCRIPTION   = "Vérifie que la sauvegarde de ICPE a bien été effectuée."
Const VERSION       = "3.0"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - Le nom du fichier à vérifier est dans la constante FICHIER_A_VERIFIER 
' - À enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la déclaration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo ICPE_Verif_Sauvegarde.vbs > ICPE_Verif_Sauvegarde.log


' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+
' Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA"
' Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb6301\hubiC\EXPANSIA\Scripts EXPANSIA"
' Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA"
' Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA"
const DOSSIER_SAUVEGARDE_ICPE = "E:\SAUVEGARDES\ICPE"
const DEBUT_NOM_FICHIER_ICPE  = "3_"
const FIN_NOM_FICHIER_ICPE    = "icpestk.dbf"

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
dim FICHIER_A_VERIFIER


' ------------------------------------------------------------
' -                     Initialisations                      -
' ------------------------------------------------------------

listeErreurs  = ""
dossier       = DOSSIER_SAUVEGARDE_ICPE
fichier       = ""

' ------------------------------------------------------------
' -           Contrôle du dossier parent DECA                -
' ------------------------------------------------------------

Set objFSO    = CreateObject("Scripting.FileSystemObject")
dossierExiste = objFSO.FolderExists(dossier)

If dossierExiste Then

  ' ------------------------------------------------------------
  ' -                  Test du fichier ICPE                    -
  ' ------------------------------------------------------------

  dateHier            = DateAdd("d",-1,Date) 'd: jour ; -1: un jour en moins; Date: la date à modifier
  fichier             = DEBUT_NOM_FICHIER_ICPE & Year(dateHier) & LPad(Month(dateHier), "0", 2) & LPad(Day(dateHier), "0", 2) & FIN_NOM_FICHIER_ICPE
  FICHIER_A_VERIFIER  = DOSSIER_SAUVEGARDE_ICPE & "\" & fichier
  ' WScript.Echo "fichierICPE = " & FICHIER_A_VERIFIER


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
         WScript.Echo "Sauvegarde ICPE OK."
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
'| Nom         : LPad                                                         |
'| Description : Formate un nombre en ajoutant des 0 devant.                  |
'| str         : chaîne contenant le nombre.                                  |
'| pad         : le caractère à ajouter devant le nombre.                     |
'| length      : la longueur finale du nombre.                                |
'| retour      : Le nombre formaté.                                           |
'+----------------------------------------------------------------------------+

Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
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
'| 3) Le fichier n'existe pas (modifier FIN_NOM_FICHIER_ICPE en supprimant une|
'|    lettre).                                                                |
'| 4) Le fichier n'est pas à la bonne date (prendre un fichier quelconque et  |
'|    lui donner le nom d'un fichier icpe).                                   |
'+----------------------------------------------------------------------------+

' 1) Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Granta.txt"
' 2) Const FICHIER_A_VERIFIER = "C:\Users\brb6301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Granta.txt"
' 3) Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\historique_Grant.txt"
' 4) Const FICHIER_A_VERIFIER = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA\LIMS_Sauvegardes_v1.vbs"

' 1) Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA"
' 1) const DEBUT_NOM_FICHIER_ICPE  = "3_"
' 1) const FIN_NOM_FICHIER_ICPE    = "icpestk.dbf"

' 2) Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb6301\hubiC\EXPANSIA\Scripts EXPANSIA"
' 2) const DEBUT_NOM_FICHIER_ICPE  = "3_"
' 2) const FIN_NOM_FICHIER_ICPE    = "icpestk.dbf"

' 3) Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA"
' 3) const DEBUT_NOM_FICHIER_ICPE  = "3_"
' 3) const FIN_NOM_FICHIER_ICPE    = "icpest.dbf"

' 4) Const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\brb06301\hubiC\EXPANSIA\Scripts EXPANSIA"
' 4) const DEBUT_NOM_FICHIER_ICPE  = "3_"
' 4) const FIN_NOM_FICHIER_ICPE    = "icpestk.dbf"

