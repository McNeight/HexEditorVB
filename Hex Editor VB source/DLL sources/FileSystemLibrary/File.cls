VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
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
Private MyFile As ClassFile

'=======================================================
'Constantes
'=======================================================
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const OPEN_EXISTING                 As Long = 3
Private Const INVALID_HANDLE_VALUE          As Long = -1
'=======================================================
'APIs
'=======================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



'=======================================================
'Type privé contenant les infos sur ce fichier
'=======================================================
Private Type ClassFile
    Drive As Drive
    Path As String
    Attributes As FileAttributes
    FileExtension As String
    Folder As Folder
    FileType As String
    AssociatedExecutableProgram As String
    FileSize As Currency
    FileCompressedSize As Currency
    DateCreated As String
    DateLastModified As String
    DateLastAccessed As String
    IsHidden As Boolean
    IsReadOnly As Boolean
    IsSystem As Boolean
    IsNormal As Boolean
    IsArchive As Boolean
    IsCompressed As Boolean
    ShortName As String
    ShortPath As String
    FileVersionInfos As FileVersionInfos
    FolderName As String
    DriveName As String
End Type

'=======================================================
'Définit le path de cet objet
'=======================================================
Friend Function SetPath(Path As String, Optional ByVal RefreshInformations As Boolean = True)
    'on définit le path du fichier
    MyFile.Path = Path
    If RefreshInformations Then Call RefreshInfos
End Function

'=======================================================
'//Propriétés de l'objet
'=======================================================
Public Property Get Drive() As Drive
    Set Drive = clsFile.GetDrive(Left$(Me.Path, 1))
End Property
Public Property Get Path() As String: Path = MyFile.Path: End Property
Attribute Path.VB_UserMemId = 0
Public Property Get Attributes() As FileAttributes: Attributes = MyFile.Attributes: End Property
Public Property Get FileExtension() As String: FileExtension = MyFile.FileExtension: End Property
Public Property Get Folder() As Folder
    Set Folder = clsFile.GetFolder(clsFile.GetFolderName(Me.Path))
End Property
Public Property Get FileType() As String: FileType = MyFile.FileType: End Property
Public Property Get AssociatedExecutableProgram() As String: AssociatedExecutableProgram = MyFile.AssociatedExecutableProgram: End Property
Public Property Get FileSize() As Currency: FileSize = MyFile.FileSize: End Property
Public Property Get FileCompressedSize() As Currency: FileCompressedSize = MyFile.FileCompressedSize: End Property
Public Property Get DateCreated() As String: DateCreated = MyFile.DateCreated: End Property
Public Property Get DateLastModified() As String: DateLastModified = MyFile.DateLastModified: End Property
Public Property Get DateLastAccessed() As String: DateLastAccessed = MyFile.DateLastAccessed: End Property
Public Property Get IsHidden() As Boolean: IsHidden = MyFile.IsHidden: End Property
Public Property Get IsReadOnly() As Boolean: IsReadOnly = MyFile.IsReadOnly: End Property
Public Property Get IsSystem() As Boolean: IsSystem = MyFile.IsSystem: End Property
Public Property Get IsNormal() As Boolean: IsNormal = MyFile.IsNormal: End Property
Public Property Get ShortName() As String: ShortName = MyFile.ShortName: End Property
Public Property Get ShortPath() As String: ShortPath = MyFile.ShortPath: End Property
Public Property Get FileVersionInfos() As FileVersionInfos: FileVersionInfos = MyFile.FileVersionInfos: End Property
Public Property Get IsCompressed() As Boolean: IsCompressed = MyFile.IsCompressed: End Property
Public Property Get IsArchive() As Boolean: IsArchive = MyFile.IsArchive: End Property
Public Property Get FolderName() As String: FolderName = MyFile.FolderName: End Property
Public Property Get DriveName() As String: DriveName = MyFile.DriveName: End Property



'=======================================================
'//Méthodes
'=======================================================

'=======================================================
'Rafraichit les infos
'=======================================================
Public Sub RefreshInfos()
Dim tDates As FileDates
Dim tSizes As FileSizes
Dim hFile As Long

    'obtient le handle du fichier
    hFile = CreateFile(Me.Path, GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, _
        ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If hFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    'récupère les infos sur les 3 types size, dates et exeinfos
    With clsFile
        tDates = .GetFileDates_HANDLE(hFile)
        tSizes = .GetFileSizes_HANDLE(hFile, Me.Path)
    End With
    
    'stocke tout dans la variable temporaire
    With MyFile
        .Attributes = clsFile.GetFileAttributes_HANDLE(hFile)
        .FileExtension = clsFile.GetFileExtension(Me.Path)
        .FileType = clsFile.GetFileType(Me.Path)
        .AssociatedExecutableProgram = clsFile.GetAssociatedProgram(Me.Path)
        .FileSize = tSizes.FileSize
        .FileCompressedSize = tSizes.CompressedFileSize
        .DateCreated = tDates.DateCreated
        .DateLastModified = tDates.DateLastModified
        .DateLastAccessed = tDates.DateLastAccessed
        .IsHidden = (.Attributes And Hidden) = Hidden
        .IsReadOnly = (.Attributes And ReadOnly) = ReadOnly
        .IsSystem = (.Attributes And System) = System
        .IsCompressed = (.Attributes And Compressed) = Compressed
        .IsArchive = (.Attributes And Archive) = Archive
        .IsNormal = Not (.IsHidden Or .IsReadOnly Or .IsSystem)
        .ShortName = clsFile.GetShortName(Me.Path)
        .ShortPath = clsFile.GetShortPath(Me.Path)
        .FileVersionInfos = clsFile.GetFileVersionInfos(Me.Path)
        .FolderName = clsFile.GetFolderName(Me.Path)
        If Len(Me.Path) >= 3 Then .DriveName = Left$(Me.Path, 3)
    End With
    
    'referme le handle
    Call CloseHandle(hFile)
    
End Sub

'=======================================================
'Copie le fichier
'=======================================================
Public Function Copy(ByVal Destination As String) As Long
    Copy = clsFile.CopyFile(Me.Path, Destination)
End Function

'=======================================================
'Supprime le fichier
'=======================================================
Public Function Delete() As Long
    Delete = clsFile.DeleteFile(Me.Path)
End Function

'=======================================================
'Déplace le fichier
'=======================================================
Public Function Move(ByVal Destination As String) As Long
    
    If clsFile.MoveFile(Me.Path, Destination) Then
        'on refresh les infos
        Move = Me.SetPath(Destination, True)
    End If
    
End Function

'=======================================================
'Affiche la boite de dialogue propriétés
'=======================================================
Public Function ShowPropertyBox(ByVal hWnd As Long) As Long
    ShowPropertyBox = clsFile.ShowFileProperty(Me.Path, hWnd)
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

    If clsFile.FileExists(NewName) Then
        Rename = -1
        Exit Function
    End If
        
    If clsFile.Rename(Me.Path, NewName) Then
        'on refresh
        Call Me.SetPath(NewName, True)
    End If
    
End Function

'=======================================================
'Renvoie true si le fichier existe
'=======================================================
Public Function DoesFileExist() As Boolean
    DoesFileExist = clsFile.FileExists(Me.Path)
End Function

'=======================================================
'Renvoie true si le fichier est accessible
'=======================================================
Public Function IsFileAvailable() As Boolean
    IsFileAvailable = clsFile.IsFileAvailable(Me.Path)
End Function

'=======================================================
'Change les attributs du fichier
'=======================================================
Public Sub ChangeAttributes(ByVal NewAttributes As FileAttributes)
    
    If clsFile.ChangeFileAttributes(Me.Path, NewAttributes) Then
        'refresh les attributs
        With MyFile
        .IsHidden = (.Attributes And Hidden) = Hidden
        .IsReadOnly = (.Attributes And ReadOnly) = ReadOnly
        .IsSystem = (.Attributes And System) = System
        .IsCompressed = (.Attributes And Compressed) = Compressed
        .IsArchive = (.Attributes And Archive) = Archive
        .IsNormal = Not (.IsHidden Or .IsReadOnly Or .IsSystem)
        .Attributes = NewAttributes
        End With
    End If
    
End Sub

'=======================================================
'Lit des bytes dans le fichier
'=======================================================
Public Function ReadFileString(ByVal StartingOffset As Currency, _
    ByVal Size As Long) As String

    ReadFileString = clsFile.ReadFileString(Me.Path, Size, StartingOffset)
End Function

'=======================================================
'Ecrit des bytes dans le fichier
'=======================================================
Public Function WriteFileString(ByVal StartingOffset As Currency, _
    ByVal StringToWrite As String)
    
    WriteFileString = clsFile.WriteFileString(Me.Path, StringToWrite, _
        StartingOffset, False)
    
End Function

'=======================================================
'Récupère l'icone du fichier
'=======================================================
Public Function GetIcon(ByVal Size As IconSize) As IPictureDisp
    Set GetIcon = clsFile.GetIcon(Me.Path, Size)
End Function

'=======================================================
'Encrypte le fichier
'=======================================================
Public Function Crypt() As Long
    Crypt = clsFile.EncryptFile(Me.Path)
End Function

'=======================================================
'Décrypte le fichier
'=======================================================
Public Function Decrypt() As Long
    Decrypt = clsFile.DecryptFile(Me.Path)
End Function

'=======================================================
'Lance la sanitization... my god...
'=======================================================
Public Function Sanitize() As Long
    Sanitize = clsFile.SanitizeFile(Me.Path)
End Function

'=======================================================
'Affecte des dates au fichier
'=======================================================
Public Function SetDates(Optional ByVal DateCreated As _
    String = vbNullString, Optional ByVal DateLastAccess As String = vbNullString, _
    Optional ByVal DateLastModification As String = vbNullString) As Long
    
    SetDates = clsFile.SetFileDates(Me.Path, DateCreated, DateLastAccess, _
        DateLastModification)
End Function

'=======================================================
'Lance un découpage
'=======================================================
Public Function CutFile(DestinationFolder As String, Method As CutMethod) As Long
    CutFile = clsFile.CutFileFriend(Me.Path, DestinationFolder, Method)
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
