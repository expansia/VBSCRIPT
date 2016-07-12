'+----------------------------------------------------------------------------+
'| Fichier     : Liste_Montage_Disques_Reseau.vbs                             |
'+----------------------------------------------------------------------------+
'| Version     : 1.0                                                          |
'+----------------------------------------------------------------------------+
'| Description :                                                              |
'|                                                                            |
'| Renvoie la liste des disques réseau montés sur l'ordinateur.               |
'+----------------------------------------------------------------------------+


' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
'On Error Resume Next


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
'|                             PROCÉDURES/FONCTIONS                           |
'+----------------------------------------------------------------------------+

Sub Main()

  Dim strFichierTrace, strTrace, strComputer, objWMIService, colItems, objItem
  
  strTrace        = ""
  strFichierTrace = NomFichierTrace()
  strComputer     = "."
  
  
  Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  
  Set colItems = objWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk")
  
  For Each objItem in colItems
      ' strTrace = strTrace & "Compressed: " & objItem.Compressed & vbCRLF
      ' strTrace = strTrace & "Description: " & objItem.Description & vbCRLF
      ' strTrace = strTrace & "Device ID: " & objItem.DeviceID & vbCRLF
      ' strTrace = strTrace & "File System: " & objItem.FileSystem & vbCRLF
      ' strTrace = strTrace & "Free Space: " & objItem.FreeSpace & vbCRLF
      ' strTrace = strTrace & "Maximum Component Length: " & objItem.MaximumComponentLength & vbCRLF
      ' strTrace = strTrace & "Name: " & objItem.Name & vbCRLF
      ' strTrace = strTrace & "Provider Name: " & objItem.ProviderName & vbCRLF
      ' strTrace = strTrace & "Session ID: " & objItem.SessionID & vbCRLF
      ' strTrace = strTrace & "Size: " & objItem.Size & vbCRLF
      ' strTrace = strTrace & "Supports Disk Quotas: " & objItem.SupportsDiskQuotas & vbCRLF
      ' strTrace = strTrace & "Supports File-Based Compression: " & _
      '     objItem.SupportsFileBasedCompression & vbCRLF
      ' strTrace = strTrace & "Volume Name: " & objItem.VolumeName & vbCRLF
      ' strTrace = strTrace & "Volume Serial Number: " & objItem.VolumeSerialNumber & vbCRLF
      ' strTrace = strTrace & vbCRLF
      
  
      ' strTrace = strTrace & "Provider Name: " & objItem.ProviderName & vbCRLF
      ' strTrace = strTrace & "Name: " & objItem.Name & vbCRLF
      ' strTrace = strTrace & "Volume Name: " & objItem.VolumeName & vbCRLF

      strTrace = objItem.Name & "(" & objItem.VolumeName & ") <= " & objItem.ProviderName & "  (" & objItem.FreeSpace & "octets libres)"
      call Tracer(strFichierTrace, strTrace)
      strTrace = ""
  Next

  if colItems.Count = 0 Then
    call Tracer(strFichierTrace, "Il n'y a pas de disques réseau montés sur cet ordinateur.")
  End IF
End Sub

'------------------------------------------------------------------------------
' Nom         : Init
' Description : Ecrit un repère de début du script dans le fichier de trace.
'------------------------------------------------------------------------------

Sub Init()

  Banniere(" Début du script   (" & WScript.ScriptFullName & ").")

End Sub


'------------------------------------------------------------------------------
' Nom         : Terminate
' Description : Ecrit un repère de fin de script dans le fichier de trace.
'------------------------------------------------------------------------------

Sub Terminate()

  'Banniere(" Fin du script   (" & WScript.ScriptFullName & ").", 67)
  Banniere(" Fin du script   (" & WScript.ScriptFullName & ").")

End Sub


'------------------------------------------------------------------------------
' Nom         : Banniere
' Description : Ecrit un message dans le fichier de trace.
'------------------------------------------------------------------------------

Sub Banniere(strMessage)

  Dim fichierTrace

  fichierTrace  = NomFichierTrace()

  call Tracer(fichierTrace, "")
  call Tracer(fichierTrace, "------------------------------------------------------------------------")
  call Tracer(fichierTrace, strMessage)
  call Tracer(fichierTrace, "------------------------------------------------------------------------")
  call Tracer(fichierTrace, "")

End Sub


'------------------------------------------------------------------------------
' Nom         : Banniere
' Description : Ecrit un message dans le fichier de trace.
'------------------------------------------------------------------------------

Function NomFichierTrace()

  const FICHIER_TRACE_EXT = ".log"

  NomFichierTrace  = NomFichierSansExtension(WScript.ScriptFullName) & FICHIER_TRACE_EXT

End Function


' ---
' NomFichierSansExtension
' Renvoie le nom du fichier sans son extension.
' ---
Function NomFichierSansExtension(sNomAvecExt)

  Dim nPositionDernierPoint, nLongueurNomFichier 
  
  ' Pour récupérer le nom du fichier
  ' 1. On récupère la position du dernier point
  nPositionDernierPoint  = InStrRev(sNomAvecExt, ".")
  'WScript.Echo "nPositionDernierPoint = " & nPositionDernierPoint 
  ' 2. On calcule la longueur du nom du fichier à partir cette position
  nLongueurNomFichier = nPositionDernierPoint - 1
  'WScript.Echo "nLongueurNomFichier = " & nLongueurNomFichier
  ' 3. On récupère la chaîne de cette longueur à partir de la gauche
  NomFichierSansExtension = Left(sNomAvecExt, nLongueurNomFichier)
  'WScript.Echo "NomFichierSansExtension = " & NomFichierSansExtension

End Function


'------------------------------------------------------------------------------
' Nom                          : Tracer.
' Description                  : Ecrit dans le fichier strCheminCompletFichierTrace la chaîne strTrace
' strCheminCompletFichierTrace : Chemin complet du fichier.
' strTrace                     : Ce qu'il faut écrire dans le fichier.
'------------------------------------------------------------------------------

Public Sub Tracer(strCheminCompletFichierTrace, strTrace)
    On Error Resume Next
    Dim objFSO, objFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(strCheminCompletFichierTrace, 8, True, -1) ' 8 = ForAppending, True pour créer le fichier s'il n'existe pas, -1 pour écrire au format Unicode

    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de l'appel de la fonction OpenTextFile." & vbNewLine & " (Numéro: " & Err.Number & ", Description: " & Err.Description & ")"
        Err.Clear
    Else
        ' Dim MyVar
        ' MyVar = Now ' MyVar contains the current date and time.
        ' On écrit dans le fichier
        ' objFile.WriteLine MyVar & " " & strTrace
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

