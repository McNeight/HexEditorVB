VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmPropertyShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propriétés"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   17
   Icon            =   "frmPropertyShow.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "lang_ok"
      Top             =   80
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
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   6135
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowTitle       =   0   'False
      Begin vkUserContolsXP.vkTextBox txtFile 
         Height          =   5535
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9763
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendText      =   "Informations sur le fichier"
         LegendForeColor =   12937777
         LegendType      =   1
      End
      Begin vkUserContolsXP.vkCheck chkAt 
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   7
         Top             =   5760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Lecture seule"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck chkAt 
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   6
         Top             =   5760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Système"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck chkAt 
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   5760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Caché"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vkUserContolsXP.vkCheck chkAt 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   5760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Normal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   6135
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowTitle       =   0   'False
      Begin vkUserContolsXP.vkTextBox txtProcess 
         Height          =   5895
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   10398
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendText      =   "Informations sur le processus"
         LegendForeColor =   12937777
         LegendType      =   1
      End
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   6135
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowTitle       =   0   'False
      Begin vkUserContolsXP.vkTextBox txtDisk 
         Height          =   5895
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   10398
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendText      =   "Informations sur le disque"
         LegendForeColor =   12937777
         LegendType      =   1
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
Private Lang As New clsLang

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
        Call .ActiveLang(Me): .Language = cPref.env_Lang
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
Dim S As String

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
        S = Lang.GetString("_Drive") & .VolumeLetter & "]"
        S = S & vbNewLine & Lang.GetString("_VolName") & CStr(.VolumeName) & "]"
        S = S & vbNewLine & Lang.GetString("_Serial") & Hex$(.VolumeSerialNumber) & "]"
        S = S & vbNewLine & Lang.GetString("_FileS") & CStr(.FileSystemName) & "]"
        S = S & vbNewLine & Lang.GetString("_DriveT") & CStr(.strDriveType) & "]"
        S = S & vbNewLine & Lang.GetString("_MedT") & CStr(.strMediaType) & "]"
        S = S & vbNewLine & Lang.GetString("_PartSize") & CStr(.PartitionLength) & "]"
        S = S & vbNewLine & Lang.GetString("_TotalSize") & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
        S = S & vbNewLine & Lang.GetString("_FreeSize") & CStr(.FreeSpace) & "  <--> " & FormatedSize(.FreeSpace, 10) & " ]"
        S = S & vbNewLine & Lang.GetString("_UsedSize") & CStr(.UsedSpace) & "  <--> " & FormatedSize(.UsedSpace, 10) & " ]"
        S = S & vbNewLine & Lang.GetString("_Percent") & CStr(.PercentageFree) & " %]"
        S = S & vbNewLine & Lang.GetString("_LogCount") & CStr(.TotalLogicalSectors) & "]"
        S = S & vbNewLine & Lang.GetString("_PhysCount") & CStr(.TotalPhysicalSectors) & "]"
        S = S & vbNewLine & Lang.GetString("_Hid") & CStr(.HiddenSectors) & "]"
        S = S & vbNewLine & Lang.GetString("_BPerSec") & CStr(.BytesPerSector) & "]"
        S = S & vbNewLine & Lang.GetString("_SPerClust") & CStr(.SectorPerCluster) & "]"
        S = S & vbNewLine & Lang.GetString("_Clust") & CStr(.TotalClusters) & "]"
        S = S & vbNewLine & Lang.GetString("_FreeClust") & CStr(.FreeClusters) & "]"
        S = S & vbNewLine & Lang.GetString("_UsedClust") & CStr(.UsedClusters) & "]"
        S = S & vbNewLine & Lang.GetString("_BPerClust") & CStr(.BytesPerCluster) & "]"
        S = S & vbNewLine & Lang.GetString("_Cyl") & CStr(.Cylinders) & "]"
        S = S & vbNewLine & Lang.GetString("_TPerCyl") & CStr(.TracksPerCylinder) & "]"
        S = S & vbNewLine & Lang.GetString("_SPerT") & CStr(.SectorsPerTrack) & "]"
        S = S & vbNewLine & Lang.GetString("_OffDep") & CStr(.StartingOffset) & "]"
    End With
    
    txtDisk.Text = S
    
    'libère la mémoire
    Set cDrive = Nothing
    
End Sub

'=======================================================
'affiche les propriétés d'un fichier
'=======================================================
Private Sub ShowFileProp()
Dim cFic As FileSystemLibrary.File
Dim S As String

    On Error Resume Next
    
    If Not (frmContent.ActiveForm Is Nothing) Then
        'nom du fichier
        txtFile.Text = "[" & frmContent.ActiveForm.Caption & "]"
    End If
    
    'récupère les infos sur le fichier
    Set cFic = cFile.GetFile(Mid$(txtFile.Text, 2, Len(txtFile.Text) - 2))
    
    'affiche tout çà
    With cFic
        S = Lang.GetString("_File") & .Path & "]"
        S = S & vbNewLine & Lang.GetString("_Size") & CStr(.FileSize) & " " & Lang.GetString("_Bytes") & "  -  " & CStr(Round(.FileSize / 1024, 3)) & " " & Lang.GetString("_Ko") & "]"
        S = S & vbNewLine & Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        S = S & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        S = S & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        S = S & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        S = S & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        S = S & vbNewLine & Lang.GetString("_Descr") & .FileVersionInfos.FileDescription & "]"
        S = S & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
        S = S & vbNewLine & "CompanyName=[" & .FileVersionInfos.CompanyName & "]"
        S = S & vbNewLine & "InternalName=[" & .FileVersionInfos.InternalName & "]"
        S = S & vbNewLine & "OriginalFileName=[" & .FileVersionInfos.OriginalFileName & "]"
        S = S & vbNewLine & "ProductName=[" & .FileVersionInfos.ProductName & "]"
        S = S & vbNewLine & "ProductVersion=[" & .FileVersionInfos.ProductVersion & "]"
        S = S & vbNewLine & Lang.GetString("_CompS") & .FileCompressedSize & "]"
        S = S & vbNewLine & Lang.GetString("_AssocP") & .AssociatedExecutableProgram & "]"
        S = S & vbNewLine & Lang.GetString("_Fold") & .FolderName & "]"
        S = S & vbNewLine & Lang.GetString("_DriveC") & .DriveName & "]"
        S = S & vbNewLine & Lang.GetString("_FileType") & .FileType & "]"
        S = S & vbNewLine & Lang.GetString("_FileExt") & .FileExtension & "]"
        S = S & vbNewLine & Lang.GetString("_ShortN") & .ShortName & "]"
        S = S & vbNewLine & Lang.GetString("_ShortP") & .ShortPath & "]"
    
        chkAt(0).Value = Abs(.IsNormal)
        chkAt(1).Value = Abs(.IsHidden)
        chkAt(2).Value = Abs(.IsSystem)
        chkAt(3).Value = Abs(.IsReadOnly)
    End With
    
    txtFile.Text = S

    'libère mémoire
    Set cFic = Nothing

End Sub

'=======================================================
'affiche les propriétés d'un processus
'=======================================================
Private Sub ShowProcessProp()
Dim pProcess As ProcessItem
Dim cFic As FileSystemLibrary.File
Dim S As String

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
        S = "-------------------------------------------"
        S = S & vbNewLine & "-------------- " & Lang.GetString("_Cible") & "  -------------"
        S = S & vbNewLine & "-------------------------------------------"
        S = S & vbNewLine & Lang.GetString("_File") & .Path & "]"
        S = S & vbNewLine & Lang.GetString("_Size") & CStr(.FileSize) & " " & Lang.GetString("_Bytes") & "  -  " & CStr(Round(.FileSize / 1024, 3)) & " " & Lang.GetString("_Ko") & "]"
        S = S & vbNewLine & Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        S = S & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        S = S & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        S = S & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        S = S & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        S = S & vbNewLine & Lang.GetString("_Descr") & .FileVersionInfos.FileDescription & "]"
        S = S & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
        S = S & vbNewLine & "CompanyName=[" & .FileVersionInfos.CompanyName & "]"
        S = S & vbNewLine & "InternalName=[" & .FileVersionInfos.InternalName & "]"
        S = S & vbNewLine & "OriginalFileName=[" & .FileVersionInfos.OriginalFileName & "]"
        S = S & vbNewLine & "ProductName=[" & .FileVersionInfos.ProductName & "]"
        S = S & vbNewLine & "ProductVersion=[" & .FileVersionInfos.ProductVersion & "]"
        S = S & vbNewLine & Lang.GetString("_CompS") & .FileCompressedSize & "]"
        S = S & vbNewLine & Lang.GetString("_AssocP") & .AssociatedExecutableProgram & "]"
        S = S & vbNewLine & Lang.GetString("_Fold") & .FolderName & "]"
        S = S & vbNewLine & Lang.GetString("_DriveC") & .DriveName & "]"
        S = S & vbNewLine & Lang.GetString("_FileType") & .FileType & "]"
        S = S & vbNewLine & Lang.GetString("_FileExt") & .FileExtension & "]"
        S = S & vbNewLine & Lang.GetString("_ShortN") & .ShortName & "]"
        S = S & vbNewLine & Lang.GetString("_ShortP") & .ShortPath & "]"
    End With
    
    'info process
    With pProcess
        S = S & vbNewLine & vbNewLine & vbNewLine & "-------------------------------------------"
        S = S & vbNewLine & "--------------- " & Lang.GetString("_Process") & " --------------"
        S = S & vbNewLine & "-------------------------------------------"
        S = S & vbNewLine & "PID=[" & .th32ProcessID & "]"
        S = S & vbNewLine & Lang.GetString("_ParentP") & .th32ParentProcessID & "   " & .procParentProcess.szImagePath & "]"
        S = S & vbNewLine & "Threads=[" & .cntThreads & "]"
        S = S & vbNewLine & Lang.GetString("_Prior") & .pcPriClassBase & "]"
        S = S & vbNewLine & Lang.GetString("_MemUsed") & .procMemory.WorkingSetSize & "]"
        S = S & vbNewLine & Lang.GetString("_PicMemUsed") & .procMemory.PeakWorkingSetSize & "]"
        S = S & vbNewLine & Lang.GetString("_SwapU") & .procMemory.PagefileUsage & "]"
        S = S & vbNewLine & Lang.GetString("_PicSwapU") & .procMemory.PeakPagefileUsage & "]"
        S = S & vbNewLine & "QuotaPagedPoolUsage=[" & .procMemory.QuotaPagedPoolUsage & "]"
        S = S & vbNewLine & "QuotaNonPagedPoolUsage=[" & .procMemory.QuotaNonPagedPoolUsage & "]"
        S = S & vbNewLine & "QuotaPeakPagedPoolUsage=[" & .procMemory.QuotaPeakPagedPoolUsage & "]"
        S = S & vbNewLine & "QuotaPeakNonPagedPoolUsage=[" & .procMemory.QuotaPeakNonPagedPoolUsage & "]"
        S = S & vbNewLine & Lang.GetString("_PageF") & .procMemory.PageFaultCount & "]"
    End With
    
    txtProcess.Text = S
    
    'libère mémoire
    Set cFic = Nothing
    Set pProcess = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
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
