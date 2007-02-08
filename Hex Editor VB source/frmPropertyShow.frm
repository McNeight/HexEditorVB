VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmPropertyShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propriétés"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPropertyShow.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7695
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5775
         ScaleWidth      =   7455
         TabIndex        =   8
         Top             =   240
         Width           =   7455
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   5400
            Width           =   855
         End
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Caché"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   12
            Top             =   5400
            Width           =   1095
         End
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Système"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   11
            Top             =   5400
            Width           =   1095
         End
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lecture seule"
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   10
            Top             =   5400
            Width           =   1455
         End
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   5175
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   0
            Width           =   7455
         End
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   873
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fichier"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Disque"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Processus"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7695
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5775
         ScaleWidth      =   7455
         TabIndex        =   6
         Top             =   240
         Width           =   7455
         Begin VB.TextBox txtDisk 
            BorderStyle     =   0  'None
            Height          =   5895
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   0
            Width           =   7455
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7695
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   7455
         TabIndex        =   4
         Top             =   120
         Width           =   7455
         Begin VB.TextBox txtProcess 
            BorderStyle     =   0  'None
            Height          =   5895
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   0
            Width           =   7455
         End
      End
   End
   Begin VB.Menu mnuDisplayWindowsProp 
      Caption         =   "&Afficher les proriétés Windows"
   End
End
Attribute VB_Name = "frmPropertyShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'FORM D'AFFICHAGE DES PROPRIETES
'-------------------------------------------------------

Private Sub Form_Activate()
    RefreshProp
End Sub

Private Sub Form_Load()
Dim X As Long
    For X = 0 To 2
        Frame1(X).Top = 600
        Frame1(X).Left = 120
        Frame1(X).Visible = False
    Next X
End Sub

'-------------------------------------------------------
'rafraichit les propriétés
'-------------------------------------------------------
Private Sub RefreshProp()

    'affiche le bon TAB
    If TypeOfForm(frmContent.ActiveForm) = "Fichier" Then
        TabStrip1.Tabs(1).Selected = True
        Frame1(0).Visible = True
        ShowFileProp    'affiche les infos sur le fichier
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Disque" Then
        TabStrip1.Tabs(2).Selected = True
        Frame1(1).Visible = True
        ShowDiskProp    'affiche les infos sur le disque
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Processus" Then
        Frame1(2).Visible = True
        TabStrip1.Tabs(3).Selected = True
        ShowProcessProp 'affiche les infos sur le processus
    End If

    TabStrip1.Enabled = False
End Sub


'-------------------------------------------------------
'affiche les propriétés d'un disque
'-------------------------------------------------------
Private Sub ShowDiskProp()
Dim cDrive As clsDrive
Dim s As String

    On Error Resume Next
    
    If Not (frmContent.ActiveForm Is Nothing) Then
        'nom du disque
        txtDisk.Text = "[" & Trim$(Left$(Right$(frmContent.ActiveForm.Caption, 4), 3)) & "]"
    End If
    
    'récupère les infos sur le disque
    Set cDrive = cDisk.GetLogicalDrive(Mid$(txtDisk.Text, 2, Len(txtDisk.Text) - 2) & "\")
    
    'affiche tout çà
    With cDrive
        s = "Lecteur=[" & .VolumeLetter & "]"
        s = s & vbNewLine & "Nom du volume=[" & CStr(.VolumeName) & "]"
        s = s & vbNewLine & "Numéro de série=[" & Hex$(.VolumeSerialNumber) & "]"
        s = s & vbNewLine & "Système de fichier=[" & CStr(.FileSystemName) & "]"
        s = s & vbNewLine & "Type de lecteur=[" & CStr(.strDriveType) & "]"
        s = s & vbNewLine & "Type de média=[" & CStr(.strMediaType) & "]"
        s = s & vbNewLine & "Taille de la partition=[" & CStr(.PartitionLength) & "]"
        s = s & vbNewLine & "Taille totale=[" & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
        s = s & vbNewLine & "Taille libre=[" & CStr(.FreeSpace) & "  <--> " & FormatedSize(.FreeSpace, 10) & " ]"
        s = s & vbNewLine & "Taille utilisée=[" & CStr(.UsedSpace) & "  <--> " & FormatedSize(.UsedSpace, 10) & " ]"
        s = s & vbNewLine & "Pourcentage de taille libre=[" & CStr(.PercentageFree) & " %]"
        s = s & vbNewLine & "Nombre de secteurs logiques=[" & CStr(.TotalLogicalSectors) & "]"
        s = s & vbNewLine & "Nombre de secteurs physiques=[" & CStr(.TotalPhysicalSectors) & "]"
        s = s & vbNewLine & "Nombre de secteurs cachés=[" & CStr(.HiddenSectors) & "]"
        s = s & vbNewLine & "Octets par secteur=[" & CStr(.BytesPerSector) & "]"
        s = s & vbNewLine & "Secteurs par cluster=[" & CStr(.SectorPerCluster) & "]"
        s = s & vbNewLine & "Nombre de clusters=[" & CStr(.TotalClusters) & "]"
        s = s & vbNewLine & "Clusters libres=[" & CStr(.FreeClusters) & "]"
        s = s & vbNewLine & "Clusters utilisés=[" & CStr(.UsedClusters) & "]"
        s = s & vbNewLine & "Octets par cluster=[" & CStr(.BytesPerCluster) & "]"
        s = s & vbNewLine & "Nombre de cylindres=[" & CStr(.Cylinders) & "]"
        s = s & vbNewLine & "Pistes par cylindre=[" & CStr(.TracksPerCylinder) & "]"
        s = s & vbNewLine & "Secteurs par piste=[" & CStr(.SectorsPerTrack) & "]"
        s = s & vbNewLine & "Offset de départ=[" & CStr(.StartingOffset) & "]"
    End With
    
    txtDisk.Text = s
    
    'libère la mémoire
    Set cDrive = Nothing
    
End Sub

'-------------------------------------------------------
'affiche les propriétés d'un fichier
'-------------------------------------------------------
Private Sub ShowFileProp()
Dim cFic As clsFile
Dim s As String

    On Error Resume Next
    
    If Not (frmContent.ActiveForm Is Nothing) Then
        'nom du fichier
        txtFile.Text = "[" & frmContent.ActiveForm.Caption & "]"
    End If
    
    'récupère les infos sur le fichier
    Set cFic = cFile.GetFile(Mid$(txtFile.Text, 2, Len(txtFile.Text) - 2))
    
    'affiche tout çà
    With cFic
        s = "Fichier=[" & .File & "]"
        s = s & vbNewLine & "Taille=[" & CStr(.FileSize) & " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        s = s & vbNewLine & "Attribut=[" & CStr(.FileAttributes) & "]"
        s = s & vbNewLine & "Date de création=[" & .CreationDate & "]"
        s = s & vbNewLine & "Date de dernier accès=[" & .LastAccessDate & "]"
        s = s & vbNewLine & "Date de dernière modification=[" & .LastModificationDate & "]"
        s = s & vbNewLine & "Version=[" & .EXEFileVersion & "]"
        s = s & vbNewLine & "Description=[" & .EXEFileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .EXELegalCopyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .EXECompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .EXEInternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .EXEOriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .EXEProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .EXEProductVersion & "]"
        s = s & vbNewLine & "Taille compressée=[" & .FileCompressedSize & "]"
        s = s & vbNewLine & "Programme associé=[" & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & "Répertoire contenant=[" & .FileDirectory & "]"
        s = s & vbNewLine & "Lecteur contenant=[" & .FileDrive & "]"
        s = s & vbNewLine & "Type de fichier=[" & .FileType & "]"
        s = s & vbNewLine & "Extension du fichier=[" & .FileExtension & "]"
        s = s & vbNewLine & "Nom court=[" & .ShortName & "]"
        s = s & vbNewLine & "Chemin court=[" & .ShortPath & "]"
    
        chkAt(0).Value = Abs(.IsNormal)
        chkAt(1).Value = Abs(.IsHidden)
        chkAt(2).Value = Abs(.IsSystem)
        chkAt(3).Value = Abs(.IsReadOnly)
    End With
    
    txtFile.Text = s

    'libère mémoire
    Set cFic = Nothing

End Sub

'-------------------------------------------------------
'affiche les propriétés d'un processus
'-------------------------------------------------------
Private Sub ShowProcessProp()
Dim pProcess As ProcessItem
Dim cFic As clsFile
Dim s As String

    On Error Resume Next
    
    'vérifie l'existence du processus
    If cProc.DoesPIDExist(Val(frmContent.ActiveForm.Tag)) = False Then
        'existe pas
        txtProcess.Text = "Processus inaccessible"
        Exit Sub
    End If
    
    'récupère les infos sur le processus
    Set pProcess = cProc.GetProcess(Val(frmContent.ActiveForm.Tag), False, False, True)
    Set cFic = cFile.GetFile(pProcess.szImagePath)
    
    'infos fichier cible
    With cFic
        s = "-------------------------------------------"
        s = s & vbNewLine & "-------------- Fichier cible  -------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "Fichier=[" & .File & "]"
        s = s & vbNewLine & "Taille=[" & CStr(.FileSize) & " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        s = s & vbNewLine & "Attribut=[" & CStr(.FileAttributes) & "]"
        s = s & vbNewLine & "Date de création=[" & .CreationDate & "]"
        s = s & vbNewLine & "Date de dernier accès=[" & .LastAccessDate & "]"
        s = s & vbNewLine & "Date de dernière modification=[" & .LastModificationDate & "]"
        s = s & vbNewLine & "Version=[" & .EXEFileVersion & "]"
        s = s & vbNewLine & "Description=[" & .EXEFileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .EXELegalCopyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .EXECompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .EXEInternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .EXEOriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .EXEProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .EXEProductVersion & "]"
        s = s & vbNewLine & "Taille compressée=[" & .FileCompressedSize & "]"
        s = s & vbNewLine & "Programme associé=[" & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & "Répertoire contenant=[" & .FileDirectory & "]"
        s = s & vbNewLine & "Lecteur contenant=[" & .FileDrive & "]"
        s = s & vbNewLine & "Type de fichier=[" & .FileType & "]"
        s = s & vbNewLine & "Extension du fichier=[" & .FileExtension & "]"
        s = s & vbNewLine & "Nom court=[" & .ShortName & "]"
        s = s & vbNewLine & "Chemin court=[" & .ShortPath & "]"
    End With
    
    'info process
    With pProcess
        s = s & vbNewLine & vbNewLine & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "--------------- Processus --------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & "Processus parent=[" & .th32ParentProcessID & "   " & .procParentProcess.szImagePath & "]"
        s = s & vbNewLine & "Threads=[" & .cntThreads & "]"
        s = s & vbNewLine & "Priorité=[" & .pcPriClassBase & "]"
        s = s & vbNewLine & "Mémoire utilisée=[" & .procMemory.WorkingSetSize & "]"
        s = s & vbNewLine & "Pic de mémoire utilisée=[" & .procMemory.PeakWorkingSetSize & "]"
        s = s & vbNewLine & "Utilisation du SWAP=[" & .procMemory.PagefileUsage & "]"
        s = s & vbNewLine & "Pic d'utilisation du SWAP=[" & .procMemory.PeakPagefileUsage & "]"
        s = s & vbNewLine & "QuotaPagedPoolUsage=[" & .procMemory.QuotaPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaNonPagedPoolUsage=[" & .procMemory.QuotaNonPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakPagedPoolUsage=[" & .procMemory.QuotaPeakPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakNonPagedPoolUsage=[" & .procMemory.QuotaPeakNonPagedPoolUsage & "]"
        s = s & vbNewLine & "Erreurs de page=[" & .procMemory.PageFaultCount & "]"
    End With
    
    txtProcess.Text = s
    
    'libère mémoire
    Set cFic = Nothing
    Set pProcess = Nothing

End Sub

Private Sub mnuDisplayWindowsProp_Click()
'obtient les infos données par explorer
Dim clsP As clsProcess

    Set clsP = New clsProcess
    
    If TypeOfForm(frmContent.ActiveForm) = "Processus" Then
        'le PID est stocké dans le Tag
        cFile.DisplayFileProperty clsP.GetPathFromPID(Val(frmContent.ActiveForm.Tag)), Me.hwnd
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Disque" Then
        cFile.DisplayFileProperty Right$(frmContent.ActiveForm.Caption, 3), Me.hwnd
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Fichier" Then
        cFile.DisplayFileProperty frmContent.ActiveForm.Caption, Me.hwnd
    End If
    
    Set clsP = Nothing
End Sub
