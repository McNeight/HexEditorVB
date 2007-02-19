VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
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
   Icon            =   "frmDiskInfos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Partitions logiques accessibles"
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   9495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   1
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   9255
         TabIndex        =   3
         Top             =   240
         Width           =   9255
         Begin ComctlLib.ListView LV2 
            Height          =   2535
            Left            =   0
            TabIndex        =   5
            Top             =   120
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
               Text            =   "Espace utilisé"
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
               Text            =   "Clusters utilisés"
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
               Text            =   "Secteurs cachés"
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
               Text            =   "Numéro de série"
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
               Text            =   "Offset de départ"
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Disques physiques accessibles"
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   9255
         TabIndex        =   1
         Top             =   240
         Width           =   9255
         Begin ComctlLib.ListView LV1 
            Height          =   2535
            Left            =   0
            TabIndex        =   4
            Top             =   120
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Numéro"
               Object.Width           =   1411
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
               Text            =   "Cylindres"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Pistes par cylindre"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Secteurs par piste"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Octets par secteurs"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Type"
               Object.Width           =   2540
            EndProperty
         End
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

'=======================================================
'FORM POUR AFFICHER DES INFOS SUR LES DISQUES
'=======================================================

Private clsDisk As clsDiskInfos

Private Sub Form_Load()
    Set clsDisk = New clsDiskInfos
    mnuRefresh_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsDisk = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRefresh_Click()
'refresh infos
Dim strDisk As String
Dim s As String
Dim cDrive As clsDrive
Dim x As Byte, y As Byte
Dim s_t() As String

    On Error GoTo ErrGestion
    
    LV1.ListItems.Clear: x = 0

    'obtient les infos sur les drives physiques
    For y = 0 To clsDisk.NumberOfPhysicalDrives - 1
        
        'obtient les infos sur le drive
        Set cDrive = clsDisk.GetPhysicalDrive(y)
        
        If clsDisk.IsPhysicalDriveAccessible(y) Then
            x = x + 1
            With LV1.ListItems
                .Add Text:=CStr(y)
                .Item(x).SubItems(1) = cDrive.TotalSpace
                .Item(x).SubItems(2) = cDrive.Cylinders
                .Item(x).SubItems(3) = cDrive.TracksPerCylinder
                .Item(x).SubItems(4) = cDrive.SectorsPerTrack
                .Item(x).SubItems(5) = cDrive.BytesPerSector
                .Item(x).SubItems(6) = cDrive.strMediaType
            End With
        End If
    Next y

    
    'obtient la liste des drives physiques
    clsDisk.GetLogicalDrivesList s_t()

        
    LV2.ListItems.Clear: x = 0
    
    For y = 0 To UBound(s_t()) - 1
    
        Set cDrive = clsDisk.GetLogicalDrive(s_t(y))
        
        If clsDisk.IsLogicalDriveAccessible(s_t(y)) Then
            'le drive est accessible
            x = x + 1
            With LV2.ListItems
                .Add Text:=cDrive.VolumeName
                .Item(x).SubItems(1) = cDrive.TotalSpace
                .Item(x).SubItems(2) = cDrive.PartitionLength
                .Item(x).SubItems(3) = cDrive.UsedSpace
                .Item(x).SubItems(4) = cDrive.FreeSpace
                .Item(x).SubItems(5) = cDrive.PercentageFree
                .Item(x).SubItems(6) = cDrive.BytesPerCluster
                .Item(x).SubItems(7) = cDrive.UsedClusters
                .Item(x).SubItems(8) = cDrive.FreeClusters
                .Item(x).SubItems(9) = cDrive.TotalClusters
                .Item(x).SubItems(10) = cDrive.HiddenSectors
                .Item(x).SubItems(11) = cDrive.TotalLogicalSectors
                .Item(x).SubItems(12) = cDrive.TotalPhysicalSectors
                .Item(x).SubItems(13) = cDrive.strMediaType
                .Item(x).SubItems(14) = Hex$(cDrive.VolumeSerialNumber)
                .Item(x).SubItems(15) = cDrive.BytesPerSector
                .Item(x).SubItems(16) = cDrive.SectorPerCluster
                .Item(x).SubItems(17) = cDrive.Cylinders
                .Item(x).SubItems(18) = cDrive.TracksPerCylinder
                .Item(x).SubItems(19) = cDrive.SectorsPerTrack
                .Item(x).SubItems(20) = cDrive.StartingOffset
                .Item(x).SubItems(21) = cDrive.FileSystemName
                .Item(x).SubItems(22) = cDrive.strDriveType
            End With
        End If
   
    Next y
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmDiskInfos.mnuRefreshClick", True
End Sub

Private Sub mnuSaveInfos_Click()
'enregistre les infos dans un fichier texte
Dim s As String

    On Error GoTo CancelPushed
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "Sauvegarder les informations sur les disques"
        .Filter = "Fichier html|*.html"
        .InitDir = cFile.GetSpecialFolder(CSIDL_PERSONAL)
        .ShowSave
        
        'fichier déjà existant ==> prévient
        If cFile.FileExists(.Filename) Then
            If MsgBox("Voulez vous écraser le fichier déjà existant ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
        End If
        
        'le recréé
        cFile.CreateEmptyFile .Filename, True
        
        'créé une string à enregistrer
        'format HTML
        s = CreateMeHtmlString(Me.LV1, Me.LV2)
        
        'colle la string dedans
        cFile.SaveStringInfile .Filename, s, True
        
    End With
    
CancelPushed:
End Sub
