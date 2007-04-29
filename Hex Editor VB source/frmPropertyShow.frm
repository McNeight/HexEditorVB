VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "lang_ok"
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   873
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fichier"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Disque"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Processus"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   7695
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5775
         ScaleWidth      =   7455
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   1
            Top             =   5400
            Width           =   855
         End
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Caché"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   2
            Top             =   5400
            Width           =   1095
         End
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Système"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   3
            Top             =   5400
            Width           =   1095
         End
         Begin VB.CheckBox chkAt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lecture seule"
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   4
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
            TabIndex        =   0
            Top             =   0
            Width           =   7455
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   7695
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5775
         ScaleWidth      =   7455
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
         Begin VB.TextBox txtDisk 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   7695
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   7455
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   7455
         Begin VB.TextBox txtProcess 
            BorderStyle     =   0  'None
            Height          =   5895
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   0
            Width           =   7455
         End
      End
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
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
'FORM D'AFFICHAGE DES PROPRIETES
'=======================================================

Private Sub Form_Activate()
    Call RefreshProp
End Sub

Private Sub Form_Load()
Dim x As Long

    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
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
        
        'applique la langue désirée aux controles
        .Language = cPref.env_Lang
        .LoadControlsCaption
    End With
    
    For x = 0 To 2
        Frame1(x).Top = 600
        Frame1(x).Left = 120
        Frame1(x).Visible = False
    Next x
    
    If TypeOfForm(frmContent.ActiveForm) = "Disque physique" Then _
        Me.mnuDisplayWindowsProp.Enabled = False
End Sub

'=======================================================
'rafraichit les propriétés
'=======================================================
Private Sub RefreshProp()

    'affiche le bon TAB
    If TypeOfForm(frmContent.ActiveForm) = "Fichier" Then
        TabStrip1.Tabs(1).Selected = True
        Frame1(0).Visible = True
        Call ShowFileProp    'affiche les infos sur le fichier
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Disque" Then
        TabStrip1.Tabs(2).Selected = True
        Frame1(1).Visible = True
        Call ShowDiskProp    'affiche les infos sur le disque
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Disque physique" Then
        TabStrip1.Tabs(2).Selected = True
        Frame1(1).Visible = True
        Call ShowDiskProp    'affiche les infos sur le disque
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Processus" Then
        Frame1(2).Visible = True
        TabStrip1.Tabs(3).Selected = True
        Call ShowProcessProp 'affiche les infos sur le processus
    End If

    TabStrip1.Enabled = False
End Sub


'=======================================================
'affiche les propriétés d'un disque
'=======================================================
Private Sub ShowDiskProp()
Dim cDrive As FileSystemLibrary.Drive
Dim s As String

    On Error Resume Next
    
    If Not (frmContent.ActiveForm Is Nothing) Then
        'nom du disque
        txtDisk.Text = Trim$(Left$(Right$(frmContent.ActiveForm.Caption, 4), 3))
    End If
    
    'récupère les infos sur le disque
    'TODO
    Set cDrive = cFile.GetDrive(Left$(txtDisk.Text, 1))
    
    'affiche tout çà
    With cDrive
        s = Lang.GetString("_Drive") & .VolumeLetter & "]"
        s = s & vbNewLine & Lang.GetString("_VolName") & CStr(.VolumeName) & "]"
        s = s & vbNewLine & Lang.GetString("_Serial") & Hex$(.VolumeSerialNumber) & "]"
        s = s & vbNewLine & Lang.GetString("_FileS") & CStr(.FileSystemName) & "]"
        s = s & vbNewLine & Lang.GetString("_DriveT") & CStr(.strDriveType) & "]"
        s = s & vbNewLine & Lang.GetString("_MedT") & CStr(.strMediaType) & "]"
        s = s & vbNewLine & Lang.GetString("_PartSize") & CStr(.PartitionLength) & "]"
        s = s & vbNewLine & Lang.GetString("_TotalSize") & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
        s = s & vbNewLine & Lang.GetString("_FreeSize") & CStr(.FreeSpace) & "  <--> " & FormatedSize(.FreeSpace, 10) & " ]"
        s = s & vbNewLine & Lang.GetString("_UsedSize") & CStr(.UsedSpace) & "  <--> " & FormatedSize(.UsedSpace, 10) & " ]"
        s = s & vbNewLine & Lang.GetString("_Percent") & CStr(.PercentageFree) & " %]"
        s = s & vbNewLine & Lang.GetString("_LogCount") & CStr(.TotalLogicalSectors) & "]"
        s = s & vbNewLine & Lang.GetString("_PhysCount") & CStr(.TotalPhysicalSectors) & "]"
        s = s & vbNewLine & Lang.GetString("_Hid") & CStr(.HiddenSectors) & "]"
        s = s & vbNewLine & Lang.GetString("_BPerSec") & CStr(.BytesPerSector) & "]"
        s = s & vbNewLine & Lang.GetString("_SPerClust") & CStr(.SectorPerCluster) & "]"
        s = s & vbNewLine & Lang.GetString("_Clust") & CStr(.TotalClusters) & "]"
        s = s & vbNewLine & Lang.GetString("_FreeClust") & CStr(.FreeClusters) & "]"
        s = s & vbNewLine & Lang.GetString("_UsedClust") & CStr(.UsedClusters) & "]"
        s = s & vbNewLine & Lang.GetString("_BPerClust") & CStr(.BytesPerCluster) & "]"
        s = s & vbNewLine & Lang.GetString("_Cyl") & CStr(.Cylinders) & "]"
        s = s & vbNewLine & Lang.GetString("_TPerCyl") & CStr(.TracksPerCylinder) & "]"
        s = s & vbNewLine & Lang.GetString("_SPerT") & CStr(.SectorsPerTrack) & "]"
        s = s & vbNewLine & Lang.GetString("_OffDep") & CStr(.StartingOffset) & "]"
    End With
    
    txtDisk.Text = s
    
    'libère la mémoire
    Set cDrive = Nothing
    
End Sub

'=======================================================
'affiche les propriétés d'un fichier
'=======================================================
Private Sub ShowFileProp()
Dim cFic As FileSystemLibrary.File
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
        s = Lang.GetString("_File") & .Path & "]"
        s = s & vbNewLine & Lang.GetString("_Size") & CStr(.FileSize) & " " & Lang.GetString("_Bytes") & "  -  " & CStr(Round(.FileSize / 1024, 3)) & " " & Lang.GetString("_Ko") & "]"
        s = s & vbNewLine & Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        s = s & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        s = s & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        s = s & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        s = s & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        s = s & vbNewLine & Lang.GetString("_Descr") & .FileVersionInfos.FileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .FileVersionInfos.CompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .FileVersionInfos.InternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .FileVersionInfos.OriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .FileVersionInfos.ProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .FileVersionInfos.ProductVersion & "]"
        s = s & vbNewLine & Lang.GetString("_CompS") & .FileCompressedSize & "]"
        s = s & vbNewLine & Lang.GetString("_AssocP") & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & Lang.GetString("_Fold") & .FolderName & "]"
        s = s & vbNewLine & Lang.GetString("_DriveC") & .DriveName & "]"
        s = s & vbNewLine & Lang.GetString("_FileType") & .FileType & "]"
        s = s & vbNewLine & Lang.GetString("_FileExt") & .FileExtension & "]"
        s = s & vbNewLine & Lang.GetString("_ShortN") & .ShortName & "]"
        s = s & vbNewLine & Lang.GetString("_ShortP") & .ShortPath & "]"
    
        chkAt(0).Value = Abs(.IsNormal)
        chkAt(1).Value = Abs(.IsHidden)
        chkAt(2).Value = Abs(.IsSystem)
        chkAt(3).Value = Abs(.IsReadOnly)
    End With
    
    txtFile.Text = s

    'libère mémoire
    Set cFic = Nothing

End Sub

'=======================================================
'affiche les propriétés d'un processus
'=======================================================
Private Sub ShowProcessProp()
Dim pProcess As ProcessItem
Dim cFic As FileSystemLibrary.File
Dim s As String

    On Error Resume Next
    
    'vérifie l'existence du processus
    If cProc.DoesPIDExist(Val(frmContent.ActiveForm.Tag)) = False Then
        'existe pas
        txtProcess.Text = Lang.GetString("_AccessDen")
        Exit Sub
    End If
    
    'récupère les infos sur le processus
    Set pProcess = cProc.GetProcess(Val(frmContent.ActiveForm.Tag), False, _
        False, True)
    Set cFic = cFile.GetFile(pProcess.szImagePath)
    
    'infos fichier cible
    With cFic
        s = "-------------------------------------------"
        s = s & vbNewLine & "-------------- " & Lang.GetString("_Cible") & "  -------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & Lang.GetString("_File") & .Path & "]"
        s = s & vbNewLine & Lang.GetString("_Size") & CStr(.FileSize) & " " & Lang.GetString("_Bytes") & "  -  " & CStr(Round(.FileSize / 1024, 3)) & " " & Lang.GetString("_Ko") & "]"
        s = s & vbNewLine & Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        s = s & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        s = s & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        s = s & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        s = s & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        s = s & vbNewLine & Lang.GetString("_Descr") & .FileVersionInfos.FileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .FileVersionInfos.CompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .FileVersionInfos.InternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .FileVersionInfos.OriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .FileVersionInfos.ProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .FileVersionInfos.ProductVersion & "]"
        s = s & vbNewLine & Lang.GetString("_CompS") & .FileCompressedSize & "]"
        s = s & vbNewLine & Lang.GetString("_AssocP") & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & Lang.GetString("_Fold") & .FolderName & "]"
        s = s & vbNewLine & Lang.GetString("_DriveC") & .DriveName & "]"
        s = s & vbNewLine & Lang.GetString("_FileType") & .FileType & "]"
        s = s & vbNewLine & Lang.GetString("_FileExt") & .FileExtension & "]"
        s = s & vbNewLine & Lang.GetString("_ShortN") & .ShortName & "]"
        s = s & vbNewLine & Lang.GetString("_ShortP") & .ShortPath & "]"
    End With
    
    'info process
    With pProcess
        s = s & vbNewLine & vbNewLine & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "--------------- " & Lang.GetString("_Process") & " --------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & Lang.GetString("_ParentP") & .th32ParentProcessID & "   " & .procParentProcess.szImagePath & "]"
        s = s & vbNewLine & "Threads=[" & .cntThreads & "]"
        s = s & vbNewLine & Lang.GetString("_Prior") & .pcPriClassBase & "]"
        s = s & vbNewLine & Lang.GetString("_MemUsed") & .procMemory.WorkingSetSize & "]"
        s = s & vbNewLine & Lang.GetString("_PicMemUsed") & .procMemory.PeakWorkingSetSize & "]"
        s = s & vbNewLine & Lang.GetString("_SwapU") & .procMemory.PagefileUsage & "]"
        s = s & vbNewLine & Lang.GetString("_PicSwapU") & .procMemory.PeakPagefileUsage & "]"
        s = s & vbNewLine & "QuotaPagedPoolUsage=[" & .procMemory.QuotaPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaNonPagedPoolUsage=[" & .procMemory.QuotaNonPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakPagedPoolUsage=[" & .procMemory.QuotaPeakPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakNonPagedPoolUsage=[" & .procMemory.QuotaPeakNonPagedPoolUsage & "]"
        s = s & vbNewLine & Lang.GetString("_PageF") & .procMemory.PageFaultCount & "]"
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
        Call cFile.ShowFileProperty(clsP.GetProcessPathFromPID(Val(frmContent.ActiveForm.Tag)), Me.hWnd)
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Disque" Then
        Call cFile.ShowFileProperty(Right$(frmContent.ActiveForm.Caption, 3), Me.hWnd)
    ElseIf TypeOfForm(frmContent.ActiveForm) = "Fichier" Then
        Call cFile.ShowFileProperty(frmContent.ActiveForm.Caption, Me.hWnd)
    End If
    
    Set clsP = Nothing
End Sub
