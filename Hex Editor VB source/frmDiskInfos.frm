VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmDiskInfos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Infos sur les disques"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   35
   Icon            =   "frmDiskInfos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5318
      Caption         =   "Partition logiques accessibles"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView LV2 
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   23
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Nom"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Taille"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Taille physique"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Espace utilis�"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Espace llibre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "% libre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Taille cluster"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Clusters utilis�s"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Clusters libres"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   9
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Clusters"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   10
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs cach�s"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   11
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs logiques"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   12
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs physiques"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   13
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   14
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Num�ro de s�rie"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   15
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Octets par secteur"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   16
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs par cluster"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   17
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Cylindres"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   18
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Pistes par cylindre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   19
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs par piste"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   20
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Offset de d�part"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   21
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Format de fichier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(23) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   22
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type de lecteur"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5318
      Caption         =   "Disques physiques accessibles"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView LV1 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Num�ro"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Nom"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Taille"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Cylindres"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Pistes par cylindre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs par piste"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Octets par secteur"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Secteurs physiques"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Menu rmnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Rafraichir"
      End
      Begin VB.Menu mnuTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveInfos 
         Caption         =   "&Enregistrer les informations..."
      End
      Begin VB.Menu mnuTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "frmDiskInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =======================================================
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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

'=======================================================
'FORM POUR AFFICHER DES INFOS SUR LES DISQUES
'=======================================================

Private Lang As New clsLang

Private Sub Form_Load()
        
    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on cr�� le fichier de langue fran�ais
                .Language = "French"
                .LangFolder = LANG_PATH
                .WriteIniFileFormIDEform
            End If
        #End If
        
        If App.LogMode = 0 Then
            'alors on est dans l'IDE
            .LangFolder = LANG_PATH
        Else
            .LangFolder = App.Path & "\Lang"
        End If
        
        'applique la langue d�sir�e aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        Call .LoadControlsCaption
    End With

    Call mnuRefresh_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRefresh_Click()
'refresh infos
Dim strDisk As String
Dim S As String
Dim cDrive As FileSystemLibrary.Drive
Dim cDisk As FileSystemLibrary.PhysicalDisk
Dim X As Byte
Dim Y As Byte
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_RetDisk"))
    
    LV1.ListItems.Clear: X = 0: Y = 0

    'obtient les infos sur les drives physiques
    For Each cDisk In cFile.PhysicalDisks
        
        Y = Y + 1
        
        'obtient les infos sur le drive
        If cDisk.IsDiskAvailable Then
            X = X + 1
            With LV1.ListItems
                .Add Text:=CStr(Y)
                .Item(X).SubItems(1) = cDisk.DiskName
                .Item(X).SubItems(2) = cDisk.TotalSpace
                .Item(X).SubItems(3) = cDisk.Cylinders
                .Item(X).SubItems(4) = cDisk.TracksPerCylinder
                .Item(X).SubItems(5) = cDisk.SectorsPerTrack
                .Item(X).SubItems(6) = cDisk.BytesPerSector
                .Item(X).SubItems(7) = cDisk.strMediaType
                .Item(X).SubItems(8) = cDisk.TotalPhysicalSectors
            End With
        End If
    Next cDisk
        
    LV2.ListItems.Clear: X = 0
    
    For Each cDrive In cFile.Drives
        
        If cDrive.IsDriveAvailable Then
            'le drive est accessible
            X = X + 1
            With LV2.ListItems
                .Add Text:=cDrive.VolumeName
                .Item(X).SubItems(1) = cDrive.TotalSpace
                .Item(X).SubItems(2) = cDrive.PartitionLength
                .Item(X).SubItems(3) = cDrive.UsedSpace
                .Item(X).SubItems(4) = cDrive.FreeSpace
                .Item(X).SubItems(5) = cDrive.PercentageFree
                .Item(X).SubItems(6) = cDrive.BytesPerCluster
                .Item(X).SubItems(7) = cDrive.UsedClusters
                .Item(X).SubItems(8) = cDrive.FreeClusters
                .Item(X).SubItems(9) = cDrive.TotalClusters
                .Item(X).SubItems(10) = cDrive.HiddenSectors
                .Item(X).SubItems(11) = cDrive.TotalLogicalSectors
                .Item(X).SubItems(12) = cDrive.TotalPhysicalSectors
                .Item(X).SubItems(13) = cDrive.strMediaType
                .Item(X).SubItems(14) = Hex$(cDrive.VolumeSerialNumber)
                .Item(X).SubItems(15) = cDrive.BytesPerSector
                .Item(X).SubItems(16) = cDrive.SectorPerCluster
                .Item(X).SubItems(17) = cDrive.Cylinders
                .Item(X).SubItems(18) = cDrive.TracksPerCylinder
                .Item(X).SubItems(19) = cDrive.SectorsPerTrack
                .Item(X).SubItems(20) = cDrive.StartingOffset
                .Item(X).SubItems(21) = cDrive.FileSystemName
                .Item(X).SubItems(22) = cDrive.strDriveType
            End With
        End If
   
    Next cDrive

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_RetOk"))
    
End Sub

Private Sub mnuSaveInfos_Click()
'enregistre les infos dans un fichier texte
Dim S As String

    On Error GoTo CancelPushed
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SaveInfos")
        .Filter = Lang.GetString("_HTMLfile") & "|*.html"
        .InitDir = cFile.GetSpecialFolder(CSIDL_PERSONAL)
        .FileName = vbNullString
        .ShowSave
        
        'fichier d�j� existant ==> pr�vient
        If cFile.FileExists(.FileName) Then
            If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + _
                vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
        End If
        
        'le recr��
        Call cFile.CreateEmptyFile(.FileName, True)
        
        'cr�� une string � enregistrer
        'format HTML
        S = CreateMeHtmlString(Me.LV1, Me.LV2)
        
        'colle la string dedans
        Call cFile.SaveDataInFile(.FileName, S, True)
        
    End With
    
CancelPushed:
End Sub
