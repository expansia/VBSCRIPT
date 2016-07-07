'******************************************************************************
'* Fichier:	Ouvrir_Dossier.vbs                                            *
'* Auteur:	Bruno Boissonnet                                              *
'* Date:	08/10/2014                                                    *
'* Description: Script qui cr�e des dossiers.                                 *
'*                                                                            *
'* Remarques:   - Il faut renseigner les variables strDossier*                *
'******************************************************************************



Const strDossierInformatique = "D:\INFORMATIQUE\"
Const strDossierMesDocuments = "D:\MesDocuments\"
Const strDossierModele = "D:\modele\"
Const strDossierPersonnel = "D:\PERSONNEL\"
Const strTitre = "Cr�ation des dossiers sur D:."
'Const strFile = "\\server\folder\file.ext"
'Const Overwrite = True
Dim oFSO


Set oFSO = CreateObject("Scripting.FileSystemObject")

'Set WShell = Wscript.CreateObject("Wscript.Shell")
'DTOPfolder = WShell.SpecialFolders("Desktop")

If Not oFSO.FolderExists(strDossierInformatique) Then
  oFSO.CreateFolder strDossierInformatique
Else
	result = MsgBox ("Le dossier """ & strDossierInformatique & """ existe d�j�.", _
	vbOK+vbExclamation, strTitre)
End If

If Not oFSO.FolderExists(strDossierMesDocuments) Then
  oFSO.CreateFolder strDossierMesDocuments
Else
	result = MsgBox ("Le dossier """ & strDossierMesDocuments & """ existe d�j�.", _
	vbOK+vbExclamation, strTitre)
End If

If Not oFSO.FolderExists(strDossierModele) Then
  oFSO.CreateFolder strDossierModele
Else
	result = MsgBox ("Le dossier """ & strDossierModele & """ existe d�j�.", _
	vbOK+vbExclamation, strTitre)
End If

If Not oFSO.FolderExists(strDossierPersonnel) Then
  oFSO.CreateFolder strDossierPersonnel
Else
	result = MsgBox ("Le dossier """ & strDossierPersonnel & """ existe d�j�.", _
	vbOK+vbExclamation, strTitre)
End If

result = MsgBox ("Fin du script", _
	vbOK+vbExclamation, strTitre)

'oFSO.CopyFile strFile, strFolder, Overwrite