VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Object = "{3AF19019-2368-4F9C-BBFC-FD02C59BD0EC}#1.0#0"; "DriveView_OCX.ocx"
Begin VB.Form frmSanitization 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sanitization"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSanitization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3135
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Index           =   2
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   3255
         TabIndex        =   13
         Top             =   120
         Width           =   3255
         Begin VB.CommandButton cmdGoPhys 
            Height          =   375
            Left            =   960
            TabIndex        =   15
            Top             =   2520
            Width           =   1215
         End
         Begin DriveView_OCX.DriveView DV2 
            Height          =   1935
            Left            =   0
            TabIndex        =   14
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   3413
            DisplayLogicalDrives=   0   'False
            LogicalDrivesString=   ""
         End
         Begin ProgressBar_OCX.pgrBar PGB3 
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   2160
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            BackColorTop    =   13027014
            BackColorBottom =   15724527
            Value           =   1
            BackPicture     =   "frmSanitization.frx":000C
            FrontPicture    =   "frmSanitization.frx":0028
         End
      End
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   240
      Top             =   360
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Disque logique"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fichiers"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Disque physique"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
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
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Index           =   0
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   3255
         TabIndex        =   2
         Top             =   120
         Width           =   3255
         Begin DriveView_OCX.DriveView DV 
            Height          =   1935
            Left            =   0
            TabIndex        =   11
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   3413
            DisplayPhysicalDrives=   0   'False
            PhysicalDrivesString=   ""
         End
         Begin VB.CommandButton cmdGo 
            Caption         =   "GO"
            Height          =   375
            Left            =   960
            TabIndex        =   3
            Top             =   2520
            Width           =   1215
         End
         Begin ProgressBar_OCX.pgrBar PGB 
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   2160
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            BackColorTop    =   13027014
            BackColorBottom =   15724527
            Value           =   1
            BackPicture     =   "frmSanitization.frx":0044
            FrontPicture    =   "frmSanitization.frx":0060
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Index           =   1
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   3045
         TabIndex        =   5
         Top             =   120
         Width           =   3050
         Begin ProgressBar_OCX.pgrBar PGB2 
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   2160
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   450
            BackColorTop    =   13027014
            BackColorBottom =   15724527
            Value           =   1
            BackPicture     =   "frmSanitization.frx":007C
            FrontPicture    =   "frmSanitization.frx":0098
         End
         Begin VB.CommandButton cmdSelFile 
            Caption         =   "Sélection de fichiers..."
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CommandButton cmdGo2 
            Caption         =   "GO"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   2520
            Width           =   735
         End
         Begin ComctlLib.ListView LV 
            Height          =   2055
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Fichier"
               Object.Width           =   14111
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmSanitization"
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
'FORM DE SANITIZATION DE DISQUE
'=======================================================

Private tDrive As DriveView_OCX.clsDrive

Private Sub cmdGoPhys_Click()
'lance la sanitization d'un disque physique

    'messages d'alerte
    With Lang
        If MsgBox(.GetString("_SanitWar"), vbCritical + vbYesNo, _
            .GetString("_War") & "  " & .GetString("_Disk") & " " & _
            Mid$(DV2.SelectedItem.Text, 3, 1)) <> vbYes Then Exit Sub
        If MsgBox(.GetString("_SanitWar2"), vbCritical + vbYesNo, _
            .GetString("_War") & "  " & .GetString("_Disk") & " " & _
            Mid$(DV2.SelectedItem.Text, 3, 1)) <> vbYes Then Exit Sub
        MsgBox .GetString("_SanitWarLau"), vbInformation, .GetString("_War")
    End With
    
    'on lance le processus de sanitization
    Call SanitPhysDiskNow(Val(Mid$(DV2.SelectedItem.Text, 3, 1)), Me.PGB2)
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'supprime si touche Delete
Dim x As Long

    If KeyCode = vbKeyDelete Then
        For x = LV.ListItems.Count To 1 Step -1
            If LV.ListItems.Item(x).Selected Then LV.ListItems.Remove x
        Next x
    End If
    
    cmdGo2.Enabled = CBool(LV.ListItems.Count)
End Sub

Private Sub TB_Click()
'change le frame visible

    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    
    Frame1(TB.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub cmdGo_Click()
'lance la sanitization d'un disque
    
    'messages d'alerte
    With Lang
        If MsgBox(.GetString("_SanitWar"), vbCritical + vbYesNo, _
            .GetString("_War") & "  " & .GetString("_Disk") & " " & _
            tDrive.VolumeLetter) <> vbYes Then Exit Sub
        If MsgBox(.GetString("_SanitWar2"), vbCritical + vbYesNo, _
            .GetString("_War") & "  " & .GetString("_Disk") & " " & _
            tDrive.VolumeLetter) <> vbYes Then Exit Sub
        MsgBox .GetString("_SanitWarLau"), vbInformation, .GetString("_War")
    End With
    
    cmdGO.Enabled = False
    
    'on lance le processus de sanitization
    Call SanitDiskNow(tDrive.VolumeLetter & ":\", Me.PGB)
        
End Sub

Private Sub cmdGo2_Click()
'lance la sanitization
    
    'messages d'alerte
    With Lang
        If MsgBox(.GetString("_SanitWarF"), vbCritical + vbYesNo, _
            .GetString("_War")) <> vbYes Then Exit Sub
        If MsgBox(.GetString("_SanitWarF2"), vbCritical + vbYesNo, _
            .GetString("_War")) <> vbYes Then Exit Sub
        MsgBox .GetString("_SanitWarLau"), vbInformation, .GetString("_War")
    End With
    
    cmdSelFile.Enabled = False
    cmdGo2.Enabled = False
    
    'on lance le processus de sanitization
    Call SanitFilesNow(LV, Me.PGB2)
    
    cmdGo2.Enabled = True
    cmdSelFile.Enabled = True
End Sub

Private Sub cmdSelFile_Click()
Dim cFile As FileSystemLibrary.FileSystem
Dim s() As String
Dim s2 As String
Dim x As Long

    'sélection d'un ou plusieurs fichiers
 
    Set cFile = New FileSystemLibrary.FileSystem
    
    ReDim s(0)
    s2 = cFile.ShowOpen(Lang.GetString("_SelFile"), Me.hWnd, _
        Lang.GetString("_All") & "|*.*", , , , , OFN_EXPLORER + _
        OFN_ALLOWMULTISELECT, s())
        
    For x = 1 To UBound(s())
        If cFile.FileExists(s(x)) Then LV.ListItems.Add Text:=s(x)
    Next x
    
    'dans le cas d'un fichier simple
    If cFile.FileExists(s2) Then LV.ListItems.Add Text:=s2
    
    cmdGo2.Enabled = CBool(LV.ListItems.Count)
    
End Sub

Private Sub DV_NodeClick(ByVal Node As ComctlLib.INode)
Dim s As String

    If DV.IsSelectedDriveAccessible Then
    
        'récupère le type de fichier du drive
        s = DV.GetSelectedDrive.FileSystemName
        
        If InStr(1, LCase$(s), "fat") Or InStr(1, LCase$(s), "ntfs") Then
            'on a choisi le drive
            Set tDrive = DV.GetSelectedDrive
            cmdGO.Enabled = True
        Else
            cmdGO.Enabled = False
        End If
    Else
        cmdGO.Enabled = False
    End If
End Sub

Private Sub Form_Activate()

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
        
        DV.LogicalDrivesString = .GetString("_LogString")
        DV.PhysicalDrivesString = .GetString("_PhysString")
    End With
    
    Call MarkNonDiskDrives(DV)  'marque les disques non NTFS ou FAT comme non accessibles
End Sub

Private Sub Form_Load()
Dim x As Long
    
    Set tDrive = New DriveView_OCX.clsDrive
    
    '//génère les tableaux
    'les redimensionne
    ReDim sH55(2097151)
    ReDim sHAA(2097151)
    'remplit
    For x = 0 To 2097151
        sH55(x) = 85
        sHAA(x) = 170
    Next x
    
    'récupère les pointeurs
    pAA = VarPtr(sHAA(0))
    p55 = VarPtr(sH55(0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tDrive = Nothing
    
    'vide les tableaux
    ReDim sH55(0)
    ReDim sHAA(0)
End Sub
