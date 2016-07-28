' Force la déclaration des variables : on est obligé de faire `Dim Variable`
Option Explicit

Const FICHIER       = "Ordinateur_Infos.vbs"
Const DESCRIPTION   = "Récupère le fabriquant, le modèle et le numéro de série de l'ordinateur."
Const VERSION       = "3.0"
Const AUTEUR        = "Bruno Boissonnet"
Const DATE_CREATION = "22/07/2016"


' Remaques :
' - Les noms des dossiers sont dans la constante LISTE_DOSSIERS, séparés par une virgule 
' - À enregistrer avec l'encodage ANSI
' - Utiliser "option explicit" pour forcer la déclaration des variables
' - Si on ne souhaite pas utiliser l'interface graphique :
'     cscript.exe //NoLogo Ordinateur_Infos.vbs > Ordinateur_Infos.log


' Empêche les erreurs de s'afficher (à supprimer lors du débogage)
' Doit être ajouté dans chaque routine
' On Error Resume Next


'+----------------------------------------------------------------------------+
'|                                 CONSTANTES                                 |
'+----------------------------------------------------------------------------+
Const SYSTEM_NAME = "."

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

WScript.echo "Fabriquant : " & FabriquantOrdinateur(SYSTEM_NAME) & vbNewLine &_
              "Modèle : " & ModeleOrdinateur(SYSTEM_NAME) & vbNewLine &_
              "Numéro de série : " & NumeroSerieOrdinateur(SYSTEM_NAME)

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

'------------------------------------------------------------------------------
' Nom         : ModeleOrdinateur
' Description : Renvoie le modèle de l'ordinateur
' retour      : Le modèle de l'ordinateur
'------------------------------------------------------------------------------

Function ModeleOrdinateur(strSystemName)
 
   ' Déclaration des variables obligatoire
  Dim objComputerSystem, ordinateur, retour
  
  set objComputerSystem = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
  strSystemName & "\root\cimv2").InstancesOf ("Win32_ComputerSystem")
  for each ordinateur in objComputerSystem
    retour = trim(ordinateur.Model)
  Next
  
  Set objComputerSystem = Nothing
  Set ordinateur        = Nothing
  ModeleOrdinateur      = retour
   
End Function

'------------------------------------------------------------------------------
' Nom         : FabriquantOrdinateur
' Description : Renvoie le Fabriquant de l'ordinateur
' retour      : Le Fabriquant de l'ordinateur
'------------------------------------------------------------------------------

Function FabriquantOrdinateur(strSystemName)
 
   ' Déclaration des variables obligatoire
  Dim objComputerSystem, ordinateur, retour
  
  set objComputerSystem = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
  strSystemName & "\root\cimv2").InstancesOf ("Win32_ComputerSystem")
  for each ordinateur in objComputerSystem
    retour = trim(ordinateur.Manufacturer)
  Next
  
  Set objComputerSystem = Nothing
  Set ordinateur        = Nothing
  FabriquantOrdinateur  = retour
   
End Function

'------------------------------------------------------------------------------
' Nom         : NumeroSerieOrdinateur
' Description : Renvoie le numéro de série de l'ordinateur
' retour      : Le numéro de série de l'ordinateur
'------------------------------------------------------------------------------

Function NumeroSerieOrdinateur(strSystemName)
 
   ' Déclaration des variables obligatoire
  Dim objWMIService, colSMBIOS, objSMBIOS, retour
  
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strSystemName & "\root\cimv2")
  Set colSMBIOS = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
  for each objSMBIOS in colSMBIOS
    retour = objSMBIOS.SerialNumber
  Next
  
  Set objWMIService = Nothing
  Set colSMBIOS        = Nothing
  NumeroSerieOrdinateur = retour
   
End Function


'+----------------------------------------------------------------------------+
'|                              FIN DU SCRIPT                                 |
'+----------------------------------------------------------------------------+

'+----------------------------------------------------------------------------+
'|                                   TESTS                                    |
'+----------------------------------------------------------------------------+
'|                                                                            |
'| 1)                                                                         |
'| 2)                                                                         |
'| 3)                                                                         |
'|                                                                            |
'+----------------------------------------------------------------------------+
