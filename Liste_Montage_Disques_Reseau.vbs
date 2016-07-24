' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

Const FICHIER       = "Liste_Montage_Disques_Reseau.vbs"
Const DESCRIPTION   = "Renvoie la liste des disques réseau montés sur l'ordinateur."
Const VERSION       = "2.0"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques : 
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

  Dim strTrace, strComputer, objWMIService, colItems, objItem
  
  strTrace        = ""
  strComputer     = "."
  
  
  Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  
  Set colItems = objWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk")
  
  For Each objItem in colItems
      
      strTrace = objItem.Name & "(" & objItem.VolumeName & ") <= " & objItem.ProviderName & "  (" & objItem.FreeSpace & " octets libres)"
      ' ex: H:(DATA) <= \\192.168.9.5\Data  (16603258880 octets libres)

      WScript.Echo(strTrace)
      strTrace = ""

  Next

  if colItems.Count = 0 Then
    WScript.Echo("Il n'y a pas de disques réseau montés sur cet ordinateur.")
  End IF
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
'|                              FIN DU SCRIPT                                 |
'+----------------------------------------------------------------------------+



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
