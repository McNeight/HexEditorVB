VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =======================================================
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A complete hexadecimal editor for Windows ©
' (Editeur hexadécimal complet pour Windows ©)
'
' Copyright © 2006-2007 by Alain Descotes.
'
' This file is part of Hex Editor VB.
'
' Hex Editor VB is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' Hex Editor VB is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with Hex Editor VB; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' =======================================================


Option Explicit

Private clsFile As FileSystem
Private MyFolder As ClassFolder


'=======================================================
'Constantes
'=======================================================
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const OPEN_EXISTING                 As Long = 3
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const FILE_FLAG_BACKUP_SEMANTICS    As Long = &H2000000
Private Const FILE_SHARE_DELETE             As Long = &H4
Private Const INVALID_HANDLE_VALUE          As Long = -1
'=======================================================
'APIs
'=======================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'=======================================================
'Type privé contenant les infos sur ce fichier
'=======================================================
Private Type ClassFolder
    Path As String
    Drive As Drive
    DateCreated As String
    DateLastModified As String
    DateLastAccessed As String
    Attributes As FileAttributes
    IsHidden As Boolean
    IsArchive As Boolean
    IsCompressed As Boolean
    IsReadOnly As Boolean
    IsSystem As Boolean
    IsNormal As Boolean
    ShortPath As String
    DriveName As String
    ParentFolderName As String
End Type




'=======================================================
'Définit le path de cet objet
'=======================================================
Friend Function SetPath(Path As String, Optional ByVal RefreshInformations As Boolean = True)
    'on définit le path du fichier
    MyFolder.Path = Path
    If RefreshInformations Then Call RefreshInfos
End Function

'=======================================================
'//Propriétés de l'objet
'=======================================================
Public Property Get Drive() As Drive
    Set Drive = clsFile.GetDrive(Left$(Me.Path, 1))
End Property
Public Property Get Path() As String: Path = MyFolder.Path: End Property
Public Property Get Attributes() As FileAttributes: Attributes = MyFolder.Attributes: End Property
Public Property Get DateCreated() As String: DateCreated = MyFolder.DateCreated: End Property
Public Property Get DateLastModified() As String: DateLastModified = MyFolder.DateLastModified: End Property
Public Property Get DateLastAccessed() As String: DateLastAccessed = MyFolder.DateLastAccessed: End Property
Public Property Get IsHidden() As Boolean: IsHidden = MyFolder.IsHidden: End Property
Public Property Get IsReadOnly() As Boolean: IsReadOnly = MyFolder.IsReadOnly: End Property
Public Property Get IsSystem() As Boolean: IsSystem = MyFolder.IsSystem: End Property
Public Property Get IsNormal() As Boolean: IsNormal = MyFolder.IsNormal: End Property
Public Property Get ParentFolderName() As String: ParentFolderName = MyFolder.ParentFolderName: End Property
Public Property Get ShortPath() As String: ShortPath = MyFolder.ShortPath: End Property
Public Property Get IsCompressed() As Boolean: IsCompressed = MyFolder.IsCompressed: End Property
Public Property Get IsArchive() As Boolean: IsArchive = MyFolder.IsArchive: End Property
Public Property Get DriveName() As String: DriveName = MyFolder.DriveName: End Property




'=======================================================
'//Méthodes
'=======================================================

'=======================================================
'Rafraichit les infos
'=======================================================
Public Sub RefreshInfos()
Dim tDates As FileDates
Dim hFolder As Long

    'handle vers le dossier désiré
    hFolder = CreateFile(Me.Path, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
        
    If hFolder = INVALID_HANDLE_VALUE Then Exit Sub
    
    'récupère les infos sur les 3 types size, dates et exeinfos
    tDates = clsFile.GetFolderDates_HANDLE(hFolder)
    
    'stocke tout dans la variable temporaire
    With MyFolder
        .Attributes = clsFile.GetFolderAttributes_HANDLE(hFolder)
        .DateCreated = tDates.DateCreated
        .DateLastModified = tDates.DateLastModified
        .DateLastAccessed = tDates.DateLastAccessed
        .IsHidden = (.Attributes And Hidden) = Hidden
        .IsReadOnly = (.Attributes And ReadOnly) = ReadOnly
        .IsSystem = (.Attributes And System) = System
        .IsCompressed = (.Attributes And Compressed) = Compressed
        .IsArchive = (.Attributes And Archive) = Archive
        .IsNormal = Not (.IsHidden Or .IsReadOnly Or .IsSystem)
        .ShortPath = clsFile.GetShortPath(Me.Path & "\")
        .ParentFolderName = clsFile.GetParentFolderName(Me.Path)
        If Len(Me.ParentFolderName) >= 3 Then .DriveName = Left$(Me.Path, 3)
    End With
    
    'referme le handle
    Call CloseHandle(hFolder)
    
End Sub

'=======================================================
'Copie le dossier
'=======================================================
Public Function Copy(ByVal Destination As String)
    Copy = clsFile.CopyFolder(Me.Path, Destination)
End Function

'=======================================================
'Supprime le dossier
'=======================================================
Public Function Delete() As Long
    Delete = clsFile.DeleteEmptyFolder(Me.Path)
End Function

'=======================================================
'Déplace le dossier
'=======================================================
Public Function Move(ByVal Destination As String) As Long
    
    If clsFile.MoveFolder(Me.Path, Destination) Then
        'on refresh les infos
        Move = Me.SetPath(Destination, True)
    End If
    
End Function

'=======================================================
'Affiche la boite de dialogue propriétés
'=======================================================
Public Function ShowPropertyBox(ByVal hWnd As Long) As Long
    ShowPropertyBox = clsFile.ShowFolderProperty(Me.Path, hWnd)
End Function

'=======================================================
'Déplace vers la corbeille
'=======================================================
Public Function MoveToTrash() As Long
    MoveToTrash = clsFile.MoveToTrash(Me.Path)
End Function

'=======================================================
'Renomme
'=======================================================
Public Function Rename(ByVal NewName As String) As Long

    If clsFile.FolderExists(NewName) Then
        Rename = -1
        Exit Function
    End If
        
    If clsFile.Rename(Me.Path, NewName) Then
        'on refresh
        Call Me.SetPath(NewName, True)
    End If
    
End Function

'=======================================================
'Renvoie true si le dossier existe
'=======================================================
Public Function DoesFolderExist() As Boolean
    DoesFolderExist = clsFile.FolderExists(Me.Path)
End Function

'=======================================================
'Renvoie true si le dossier est accessible
'=======================================================
Public Function IsFolderAvailable() As Boolean
    IsFolderAvailable = clsFile.IsFolderAvailable(Me.Path)
End Function

'=======================================================
'Change les attributs du dossier
'=======================================================
Public Function ChangeAttributes(ByVal NewAttributes As FileAttributes)
Dim l As Long
    
    l = clsFile.ChangeFileAttributes(Me.Path, NewAttributes)
    If l Then
        'refresh les attributs
        With MyFolder
            .IsHidden = (.Attributes And Hidden) = Hidden
            .IsReadOnly = (.Attributes And ReadOnly) = ReadOnly
            .IsSystem = (.Attributes And System) = System
            .IsCompressed = (.Attributes And Compressed) = Compressed
            .IsArchive = (.Attributes And Archive) = Archive
            .IsNormal = Not (.IsHidden Or .IsReadOnly Or .IsSystem)
            .Attributes = NewAttributes
        End With
    End If
    
    ChangeAttributes = l
    
End Function

'=======================================================
'Récupère l'icone du fichier
'=======================================================
Public Function GetIcon(ByVal Size As IconSize) As IPictureDisp
    Set GetIcon = clsFile.GetIcon(Me.Path, Size)
End Function

'=======================================================
'Encrypte le dossier
'=======================================================
Public Function Crypt() As Long
    Crypt = clsFile.EncryptFolder(Me.Path)
End Function

'=======================================================
'Décrypte le dossier
'=======================================================
Public Function Decrypt() As Long
    Decrypt = clsFile.DecryptFolder(Me.Path)
End Function





'=======================================================
'Subs de la classe
'=======================================================
Private Sub Class_Initialize()
    'instancie la classe clsFile
    Set clsFile = New FileSystem
End Sub

Private Sub Class_Terminate()
    'libère la classe clsFile
    Set clsFile = Nothing
End Sub


