'******************************************************************************
'* Fichier     : ICPE_Verif_Sauvegarde.vbs                                    *
'* Auteur      : Bruno Boissonnet                                             *
'* Version     : 1.0                                                          *
'* Description : Script qui vérifie que la sauvegarde de ICPE a bien été      *
'*               effectuée.                                                   *
'* Remarques   :                                                              *
'*               - Vérifie le dossier                                         *
'*                    \\ARAMON02\e\SAUVEGARDES\ICPE                           *
'*               - Dans ce dossier le fichier 3_AAAAMMJJicpestk.dbf doit      *
'*                 avoir la date de la veille.                                *
'*               - Un fichier trace (constante FICHIER_TRACE) situé dans le   *
'*                 dossier du script permet de vérifier ce qu'il s'est passé. *
'******************************************************************************

' Force la déclaration des variables : on est obligé de faire : `Dim Variable`
Option Explicit

' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
'On Error Resume Next

' ------------------------------------------------------------
' -                        Constantes                        -
' ------------------------------------------------------------

const FICHIER_TRACE           = "ICPE_Sauvegarde.log"
'const DOSSIER_SAUVEGARDE_ICPE = "C:\Users\BRB06301\Desktop\SAUVEGARDES\ICPE"
const DOSSIER_SAUVEGARDE_ICPE = "E:\SAUVEGARDES\ICPE"
const DEBUT_NOM_FICHIER_ICPE  = "3_"
const FIN_NOM_FICHIER_ICPE    = "icpestk.dbf"



' ------------------------------------------------------------
' -                        Variables                         -
' ------------------------------------------------------------

dim objFSO
dim objShell
dim dossierICPE, fichierICPE
dim dateFichierICPE, dateHier
dim fichierTrace
dim listeErreurs
dim erreurTrouvee
dim dossierICPEExiste
dim fichierExiste


' ------------------------------------------------------------
' -                     Initialisations                      -
' ------------------------------------------------------------

fichierTrace  = CheminDossierParent(WScript.ScriptFullName) & FICHIER_TRACE
erreurTrouvee = False
listeErreurs  = ""


' ------------------------------------------------------------
' -                    Début du script                       -
' ------------------------------------------------------------
call Tracer(fichierTrace, "")
call Tracer(fichierTrace, "************************************************************************")
call Tracer(fichierTrace, ">>>>> Début du script   (" & WScript.ScriptFullName & ").")
call Tracer(fichierTrace, "")


' ------------------------------------------------------------
' -           Contrôle du dossier parent ICPE                -
' ------------------------------------------------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")
dossierICPEExiste = objFSO.FolderExists(DOSSIER_SAUVEGARDE_ICPE)

If dossierICPEExiste Then

	' ------------------------------------------------------------
	' -                  Test du fichier ICPE                    -
	' ------------------------------------------------------------

	dateHier = DateAdd("d",-1,Date) 'd: jour ; -1: un jour en moins; Date: la date à modifier
	fichierICPE = DEBUT_NOM_FICHIER_ICPE & Year(dateHier) & LPad(Month(dateHier), "0", 2) & LPad(Day(dateHier), "0", 2) & FIN_NOM_FICHIER_ICPE

	fichierICPE = DOSSIER_SAUVEGARDE_ICPE & "\" & fichierICPE
	'WScript.Echo "fichierICPE = " & fichierICPE

	'WScript.Quit

	fichierExiste = objFSO.FileExists(fichierICPE)

	if fichierExiste Then
		
		' - Test de la date de dernière modification
			
		dateFichierICPE = DateDerniereModificationFichier(fichierICPE)
		
		'dateHier = DateAdd("d",-1,Date) 'd: jour ; -1: un jour en moins; Date: la date à modifier
		
		If Not IsEmpty(dateFichierICPE) Then
			if StrComp(dateFichierICPE, dateHier) = 0 Then
			   'WScript.Echo "Les dates sont identiques"
			else
			   'WScript.echo "Les dates ne sont pas identiques"
			   listeErreurs = listeErreurs & "**ERREUR** : Le fichier " & fichierICPE & " (" & dateFichierICPE & ") n'est pas à la date d'hier (" & dateHier &")."
			   erreurTrouvee = True
			end if
		End If
	Else
		listeErreurs = listeErreurs & "**ERREUR** : Le fichier " & fichierICPE & " n'existe pas."
		erreurTrouvee = True
	End If

  If erreurTrouvee Then
    call Tracer(fichierTrace, "Dossier " & DOSSIER_SAUVEGARDE_ICPE & "            [NOK]")
    call Tracer(fichierTrace, listeErreurs)
  Else
    call Tracer(fichierTrace, "Dossier " & DOSSIER_SAUVEGARDE_ICPE & "            [OK]")
  end if
  call Tracer(fichierTrace, "")

Else
	call Tracer(fichierTrace, "Dossier " & DOSSIER_SAUVEGARDE_ICPE & "						[NOK]")
	call Tracer(fichierTrace, "**ERREUR** : Le dossier " & DOSSIER_SAUVEGARDE_ICPE & " n'existe pas.")
	call Tracer(fichierTrace, "")
	erreurTrouvee = True
end If

set objFSO = Nothing

' ------------------------------------------------------------
' -                      Fin du script                       -
' ------------------------------------------------------------

call Tracer(fichierTrace, "")
call Tracer(fichierTrace, ">>>>> Fin du script   (" & WScript.ScriptFullName & ").")
call Tracer(fichierTrace, "************************************************************************")
call Tracer(fichierTrace, "")

If erreurTrouvee Then
	WScript.echo "Script terminé avec des erreurs !"
	Set objShell = CreateObject("Wscript.Shell")
	objShell.Run "notepad.exe " & fichierTrace
	set objShell = Nothing
Else
	WScript.echo "Script terminé avec succès !"
end if


'******************************************************************************

' ***
' Nom         : LPad
' Description : Formate un nombre en ajoutant des 0 devant
' str         : chaîne contenant le nombre
' pad         : le caractère à ajouter devant le nombre
' length      : la longueur finale du nombre
' retour      : Le nombre formaté
' ***
Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function



' ***
' Nom         : DateDerniereModificationFichier
' Description : Renvoi la date de dernière modification du fichier filespec
' filespec    : Chemin complet du fichier
' retour      : Une date ou Empty s'il y a eu une erreur
' ***
Function DateDerniereModificationFichier(filespec)
   On Error Resume Next ' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
   Dim objFSO, objFile, retour, strErrMsg, result
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.GetFile(filespec)
   If Err.Number <> 0 Then
      strErrMsg = "Erreur lors de l'appel de la fonction GetFile." & vbNewLine & "(Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
      result = MsgBox (strErrMsg, vbOKOnly+vbExclamation, "DateDerniereModificationFichier.vbs")
   Else
      retour = FormatDateTime(objFile.DateLastModified, 2) ' vbShortDate - 2 - Display a date using the short date format specified in your computer's regional settings.
   End If
   Set objFSO = Nothing
   Set objFile = Nothing
   DateDerniereModificationFichier = retour
End Function


' ***
' Nom                     : LitDerniereLigneFichier
' Description             : Renvoi la dernière ligne lue dans le fichier passé en paramètre
' strCheminCompletFichier : chemin complet du fichier.
' retour                  : La dernière ligne du fichier.
' ***
Public Function LitDerniereLigneFichier(strCheminCompletFichier)
   On Error Resume Next
   Dim objFSO, objFile, objTextStream, S
   
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.GetFile(strCheminCompletFichier)

   If Err.Number <> 0 Then
      WScript.Echo "Erreur lors de l'appel de la fonction GetFile." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
      Err.Clear
   Else
      Set objTextStream = objFile.OpenAsTextStream(1) '1 = ForReading
      If Err.Number <> 0 Then
         WScript.Echo "Erreur lors de l'appel de la fonction OpenAsTextStream." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
         Err.Clear
      Else
         Do    While Not objTextStream.AtEndOfStream
            S = objTextStream.ReadLine
         Loop
         objTextStream.Close
      End If
   End If

   Set objFSO        = Nothing
   Set objFile       = Nothing
   Set objTextStream = Nothing

   LitDerniereLigneFichier = S

End Function

' ***
' Nom                          : Tracer.
' Description                  : Ecrit dans le fichier strCheminCompletFichierTrace la chaîne strTrace
' strCheminCompletFichierTrace : Chemin complet du fichier.
' strTrace                     : Ce qu'il faut écrire dans le fichier.
' ***
Public Sub Tracer(strCheminCompletFichierTrace, strTrace)
    On Error Resume Next
    Dim objFSO, objFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strCheminCompletFichierTrace, 8, True, -1) ' 8 = ForAppending, True pour créer le fichier s'il n'existe pas, -1 pour écrire au format Unicode
    
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de l'appel de la fonction OpenTextFile." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
        Err.Clear
    Else
        Dim MyVar
        MyVar = Now ' MyVar contains the current date and time.
        ' On écrit dans le fichier
        objFile.WriteLine MyVar & " " & strTrace

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


' ***
' Nom              : CheminDossierParent.
' Description      : Renvoi le chemin du dossier parent de strCheminComplet (terminé par un "\").
' strCheminComplet : chemin complet du fichier.
' retour           : Le chemin du dossier parent terminé par un "\".
' ***
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
