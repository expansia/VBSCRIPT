'******************************************************************************
'* Fichier:	Ouvrir_Dossier.vbs                                            *
'* Auteur:	Bruno Boissonnet                                              *
'* Date:	08/10/2014                                                    *
'* Description: Script qui crée des dossiers.                                 *
'*                                                                            *
'* Remarques:   - Il faut renseigner les variables strDossier*                *
'******************************************************************************



Const strDossierInformatique = "D:\INFORMATIQUE1\"
Const strDossierMesDocuments = "D:\MesDocuments1\"
Const strDossierModele = "D:\modele1\"
Const strDossierPersonnel = "D:\PERSONNEL1\"
Const strTitre = "Création des dossiers sur D:."
'Const strFile = "\\server\folder\file.ext"
'Const Overwrite = True





CreerDossier(strDossierInformatique)
CreerDossier(strDossierMesDocuments)
CreerDossier(strDossierModele)
CreerDossier(strDossierPersonnel)

result = MsgBox ("Fin du script", _
	vbOK+vbExclamation, strTitre)

'oFSO.CopyFile strFile, strFolder, Overwrite

Sub CreerDossier(strNomDossier)

Dim oFSO

Set oFSO = CreateObject("Scripting.FileSystemObject")

If Not oFSO.FolderExists(strNomDossier) Then
  oFSO.CreateFolder strNomDossier
Else
	result = MsgBox ("Le dossier """ & strNomDossier & """ existe déjà.", _
	vbOK+vbExclamation, strTitre)
End If

set oFSO = Nothing

End Sub