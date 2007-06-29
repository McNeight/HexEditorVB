VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3AF19019-2368-4F9C-BBFC-FD02C59BD0EC}#1.0#0"; "DriveView_OCX.ocx"
Object = "{2245E336-2835-4C1E-B373-2395637023C8}#1.0#0"; "ProcessView_OCX.ocx"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmHome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu principal"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   14
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand cmdQuit 
      Height          =   375
      Left            =   4260
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Quitter"
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
   Begin vkUserContolsXP.vkCommand cmdOk 
      Height          =   375
      Left            =   1140
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "OK"
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
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "lang_ok"
      Top             =   113
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir fichier"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un fichier"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir dossier"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un dossier de fichiers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir disque"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un disque"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir processus"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un processus"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Nouveau fichier"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   4335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
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
      Begin vkUserContolsXP.vkTextBox txtFileInfos 
         Height          =   3135
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5530
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
      Begin vkUserContolsXP.vkCommand cmdBrowseFile 
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         ToolTipText     =   "Choix du fichier à ouvrir"
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Caption         =   "..."
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
      Begin vkUserContolsXP.vkTextBox txtFile 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Fichier choisi pour l'ouverture"
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         LegendForeColor =   12937777
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choix du fichier à ouvrir :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4095
      End
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   4335
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
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
      Begin vkUserContolsXP.vkTextBox txtDiskInfos 
         Height          =   4095
         Left            =   2760
         TabIndex        =   27
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7223
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
      Begin DriveView_OCX.DriveView DV 
         Height          =   3615
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   6376
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choix du disque à ouvrir :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2775
      End
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   4335
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
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
      Begin vkUserContolsXP.vkTextBox txtProcessInfos 
         Height          =   3975
         Left            =   3000
         TabIndex        =   28
         Top             =   120
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   7011
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
      Begin ProcessView_OCX.ProcessView PV 
         Height          =   3495
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   6165
         Sorted          =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choix du processus à ouvrir :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3255
      End
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   4335
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
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
      Begin vkUserContolsXP.vkTextBox txtFolderInfos 
         Height          =   2415
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4260
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         LegendText      =   "Informations sur le dossier"
         LegendForeColor =   12937777
         LegendType      =   1
      End
      Begin vkUserContolsXP.vkOptionButton optFolderSub 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Ouvre également tous les fichiers des sous dossiers (lent)"
         Top             =   1080
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Ouvrir également les fichiers des sous-dossiers"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Group           =   1
      End
      Begin vkUserContolsXP.vkOptionButton optFolderSub 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Ne sélectionne que les fichiers qui sont dans la racine du dossier"
         Top             =   1440
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "N'ouvrir que les fichiers contenus directement dans le dossier"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Group           =   1
      End
      Begin vkUserContolsXP.vkCommand cmdBrowseFolder 
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         ToolTipText     =   "Choix du dossier à ouvrir"
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Caption         =   "..."
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
      Begin vkUserContolsXP.vkTextBox txtFolder 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Dossier choisi pour l'ouverture"
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         LegendForeColor =   12937777
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choix du dossier à ouvrir :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
   End
   Begin vkUserContolsXP.vkFrame Frame1 
      Height          =   4335
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
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
      Begin vkUserContolsXP.vkCommand cmdBrowseNew 
         Height          =   255
         Left            =   5640
         TabIndex        =   11
         ToolTipText     =   "Choix du fichier à créer"
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Caption         =   "..."
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
      Begin vkUserContolsXP.vkTextBox txtSize 
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         ToolTipText     =   "Taille"
         Top             =   1440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Text            =   "100"
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Alignment       =   2
         LegendForeColor =   12937777
      End
      Begin vkUserContolsXP.vkTextBox txtNewFile 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Nouveau fichier à créer"
         Top             =   840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         LegendForeColor =   12937777
      End
      Begin VB.ComboBox cdUnit 
         Height          =   315
         ItemData        =   "frmHome.frx":058A
         Left            =   3360
         List            =   "frmHome.frx":059A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "pref lang_ok"
         ToolTipText     =   "Unité"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Taille du fichier"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Création d'un nouveau fichier"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmHome"
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
'FORM DE DEMARRAGE (CHOIX DES ACTIONS A EFFECTUER)
'=======================================================

Private Lang As New clsLang
Private clsPref As clsIniForm
Private bFirst As Boolean

Private Sub cmdBrowseFile_Click()
'browse un fichier
    txtFile.Text = cFile.ShowOpen(Lang.GetString("_FileToOpen"), Me.hWnd, _
        Lang.GetString("_All") & "|*.*")
End Sub

Private Sub cmdBrowseFolder_Click()
'browse un dossier
    txtFolder.Text = cFile.BrowseForFolder(Lang.GetString("_DirToOpen"), Me.hWnd)
End Sub

Private Sub cmdBrowseNew_Click()
'browse un dossier
    txtNewFile.Text = cFile.ShowSave(Lang.GetString("_FileToCreate"), _
        Me.hWnd, Lang.GetString("_All") & "|*.*", App.Path)
End Sub

Private Sub cmdOk_Click()
'ouvre l'élément sélectionné
Dim m() As String
Dim Frm As Form
Dim X As Long
Dim sDrive As String
Dim lH As Long
Dim lFile As Long
Dim lLen As Double

    On Error GoTo ErrGestion

    Select Case TB.SelectedItem.Index
        Case 1
            'fichier
            If cFile.FileExists(txtFile.Text) = False Then Exit Sub
    
            'affiche une nouvelle fenêtre
            Set Frm = New Pfm
            Call Frm.GetFile(txtFile.Text)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
            
        Case 2
            'dossier
            
            'teste la validité du répertoire
            If cFile.FolderExists(txtFolder.Text) = False Then Exit Sub
            
            'liste les fichiers
            m() = cFile.EnumFilesStr(txtFolder.Text, optFolderSub(0).Value)
            If UBound(m()) < 1 Then Exit Sub
            
            'les ouvre un par un
            For X = 1 To UBound(m)
                If cFile.FileExists(m(X)) Then
                    Set Frm = New Pfm
                    Call Frm.GetFile(m(X))
                    Frm.Show
                    lNbChildFrm = lNbChildFrm + 1
                    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                    DoEvents
                End If
            Next X
    
        Case 3
        
            If DV.SelectedItem Is Nothing Then Exit Sub
                
            'on check si c'est un disque logique ou un disque physique
            If Left$(DV.SelectedItem.Key, 3) = "log" Then
            
                'disque logique
            
                'vérifie que le drive est accessible
                If DV.IsSelectedDriveAccessible = False Then Exit Sub
                
                'affiche une nouvelle fenêtre
                Set Frm = New diskPfm
                
                Call Frm.GetDrive(DV.SelectedItem.Text)  'renseigne sur le path sélectionné
                
                Unload Me   'quitte cette form
                
                Frm.Show    'affiche la nouvelle
                lNbChildFrm = lNbChildFrm + 1
                frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
            Else
            
                'disque physique
                
                'vérifie que le drive est accessible
                If DV.IsSelectedDriveAccessible = False Then Exit Sub
                
                'affiche une nouvelle fenêtre
                Set Frm = New physPfm
                
                Call Frm.GetDrive(Val(Mid$(DV.SelectedItem.Text, 3, 1)))   'renseigne sur le path sélectionné
                
                Unload Me   'quitte cette form
                
                Frm.Show    'affiche la nouvelle
                lNbChildFrm = lNbChildFrm + 1
                frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
            End If

        Case 4
            'processus

            'vérfie que le processus est ouvrable
            lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(PV.SelectedItem.Tag))
            Call CloseHandle(lH)
            
            If lH = 0 Then
                'pas possible
                Me.Caption = Lang.GetString("_AccessDen")
                Exit Sub
            End If
            
            'possible affiche une nouvelle fenêtre
            Set Frm = New MemPfm
            Call Frm.GetFile(Val(PV.SelectedItem.Tag))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
            
        Case 5
            
            'vérifie que le fichier n'existe pas
            If cFile.FileExists(txtNewFile.Text) Then
                'fichier déjà existant
                If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
            End If
            
            Call cFile.CreateEmptyFile(txtNewFile.Text, True)
            
            'création du fichier
            lLen = Abs(Val(txtSize.Text))
            With Lang
                If cdUnit.Text = .GetString("_Ko") Then lLen = lLen * 1024
                If cdUnit.Text = .GetString("_Mo") Then lLen = (lLen * 1024) * 1024
                If cdUnit.Text = .GetString("_Go") Then lLen = ((lLen * 1024) * 1024) * 1024
            End With
            
            'ajoute du texte à la console
            Call AddTextToConsole(Lang.GetString("_CreatingFile"))
    
            'obtient un numéro de fichier disponible
            lFile = FreeFile
            
            Open txtNewFile.Text For Binary Access Write As lFile
                Put lFile, , String$(lLen, vbNullChar)
            Close lFile
            
            'affiche une nouvelle fenêtre
            Set Frm = New Pfm
            Call Frm.GetFile(txtNewFile.Text)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
        
    End Select
    
    If cPref.general_CloseHomeWhenChosen Then
        'alors on ferme cette form car on a ouvert un fichier
        Unload Me
    End If
    
    Exit Sub
    
ErrGestion:
    clsERREUR.AddError "frmHome.cmdOk_Click", True
End Sub

Private Sub DV_DblClick()
    If DV.SelectedItem Is Nothing Then Exit Sub
    If DV.SelectedItem.Children <> 0 Then Exit Sub
    cmdOk_Click
End Sub

Private Sub DV_NodeClick(ByVal Node As ComctlLib.Node)
Dim cDrive As FileSystemLibrary.Drive
Dim cDisk As FileSystemLibrary.PhysicalDisk
Dim S As String
    
    If DV.SelectedItem Is Nothing Then Exit Sub
    
    'on check si c'est un disque logique ou un disque physique
    If Left$(DV.SelectedItem.Key, 3) = "log" Then
        
        'disque logique
        
        'vérifie la disponibilité du disque
        If cFile.IsDriveAvailable(Left$(Node.Text, 1)) = False Then
            'inaccessible
            txtDiskInfos.Text = Lang.GetString("_DiskNon")
            Exit Sub
        End If
        
        'affichage des infos disque
        Set cDrive = cFile.GetDrive(Left$(Node.Text, 1))
        
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
        
        Set cDrive = Nothing
    Else
        
        'disque physique
        
        'vérifie la disponibilité du disque
        If cFile.IsPhysicalDiskAvailable(Val(Mid$(Node.Text, 3, 1))) = False Then
            'inaccessible
            txtDiskInfos.Text = Lang.GetString("_DiskPNon")
            Exit Sub
        End If
        
        'affichage des infos disque
        Set cDisk = cFile.GetPhysicalDisk(Val(Mid$(Node.Text, 3, 1)))
    
        With cDisk
            S = Lang.GetString("_Drive") & .DiskNumber & "]"
            S = S & vbNewLine & Lang.GetString("_VolName") & CStr(.DiskName) & "]"
            S = S & vbNewLine & Lang.GetString("_MedT") & CStr(.strMediaType) & "]"
            S = S & vbNewLine & Lang.GetString("_TotalSize") & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
            S = S & vbNewLine & Lang.GetString("_PhysCount") & CStr(.TotalPhysicalSectors) & "]"
            S = S & vbNewLine & Lang.GetString("_BPerSec") & CStr(.BytesPerSector) & "]"
            S = S & vbNewLine & Lang.GetString("_Cyl") & CStr(.Cylinders) & "]"
            S = S & vbNewLine & Lang.GetString("_TPerCyl") & CStr(.TracksPerCylinder) & "]"
            S = S & vbNewLine & Lang.GetString("_SPerT") & CStr(.SectorsPerTrack) & "]"
        End With
        
        Set cDisk = Nothing
    End If
    
    txtDiskInfos.Text = S
End Sub

Private Sub Form_Activate()
Dim ND As Node
    
    On Error Resume Next
    
    Call MarkUnaccessibleDrives(Me.DV)  'marque les drives inaccessibles
    
    'on expand si c'est la première activation de la form
    If bFirst = False Then
        bFirst = True
        
        'on extend tous les noeuds
        With PV
            For Each ND In .Nodes
                ND.Expanded = True
            Next ND
            
            'met en surbrillance le premier
            .Nodes.Item(1).Selected = True
        End With
        
        With DV
            For Each ND In .Nodes
                ND.Expanded = True
            Next ND
            
            'met en surbrillance le premier
            .Nodes.Item(1).Selected = True
        End With
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\HomeWindow.ini", Me)
    Set clsPref = Nothing
    Set Lang = Nothing
End Sub

Private Sub cmdQuit_Click()
    'annule
    Unload Me
End Sub

'=======================================================
'FORM HOME ==> CHOIX DE L'OBJET A OUVRIR
'=======================================================
Private Sub Form_Load()
Dim X As Long

    Set clsPref = New clsIniForm
    
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
        
        DV.PhysicalDrivesString = .GetString("_PhysStr")
        DV.LogicalDrivesString = .GetString("_LogStr")
    End With
    
    bFirst = False
    
    'loading des preferences
    Call clsPref.GetFormSettings(App.Path & "\Preferences\HomeWindow.ini", Me)
    optFolderSub(1).Value = Not (optFolderSub(0).Value)
    
    'réorganise les Frames
    For X = 0 To Frame1.Count - 1
        With Frame1(X)
            .Left = 120
            .Top = 600
            .Width = 6615
            .Height = 4335
        End With
    Next X
    
    'affiche un seul frame
    Call MaskFrames(0)
End Sub

'=======================================================
'masque tous les frames sauf un
'=======================================================
Private Sub MaskFrames(ByVal lFrame As Long)
Dim X As Long

    For X = 0 To Frame1.Count - 1
        Frame1(X).Visible = False
    Next X
    Frame1(lFrame).Visible = True
    
    If lFrame = 4 Then
        'création de fichier
        cmdOK.Caption = Lang.GetString("_CreateAndOpen!")
    Else
        cmdOK.Caption = Lang.GetString("_Open!")
    End If

End Sub

Private Sub PV_NodeClick(ByVal Node As ComctlLib.Node)
'met à jour les infos
Dim pProcess As ProcessItem
Dim cFic As FileSystemLibrary.File
Dim S As String

    On Error Resume Next
    
    'vérifie l'existence du processus
    If cProc.DoesPIDExist(Val(Node.Tag)) = False Then
        'existe pas
        txtProcessInfos.Text = Lang.GetString("_AccDen")
        Exit Sub
    End If
    
    'récupère les infos sur le processus
    Set pProcess = cProc.GetProcess(Val(Node.Tag), False, False, True)
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
    
    txtProcessInfos.Text = S
    
    'libère mémoire
    Set cFic = Nothing
    Set pProcess = Nothing
End Sub

Private Sub TB_Click()
'change le frame visible
    Call MaskFrames(TB.SelectedItem.Index - 1)
End Sub

Private Sub txtFile_Change()
Dim cFil As FileSystemLibrary.File
Dim S As String

    'met à jour les infos du fichier si ce dernier existe
    If cFile.FileExists(txtFile.Text) = False Then
        'existe pas
        txtFileInfos.Text = Lang.GetString("_FileDoesnt")
        Exit Sub
    End If
    
    'alors le fichier existe
    'récupère les infos
    Set cFil = cFile.GetFile(txtFile.Text)
    
    With cFil
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
    End With
    
    txtFileInfos.Text = S
    
    Set cFil = Nothing  'libère
End Sub

Private Sub txtFolder_Change()
Dim cFol As FileSystemLibrary.Folder
Dim S As String

    'met à jour les infos du dossier si ce dernier existe
    If cFile.FolderExists(txtFolder.Text) = False Then
        'existe pas
        txtFolderInfos.Text = Lang.GetString("_DirDoesnt")
        Exit Sub
    End If
    
    'alors le dossier existe
    'récupère les infos
    Set cFol = cFile.GetFolder(txtFolder.Text)
    
    With cFol
        S = Lang.GetString("_DirIs") & .Path & "]"
        S = S & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        S = S & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        S = S & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        S = S & vbNewLine & Lang.GetString("_ShortN") & .ShortPath & "]"
        S = S & vbNewLine & Lang.GetString("_NormalAttr") & .IsNormal & "]"
        S = S & vbNewLine & Lang.GetString("_HidAttr") & .IsHidden & "]"
        S = S & vbNewLine & Lang.GetString("_ROAttr") & .IsReadOnly & "]"
        S = S & vbNewLine & Lang.GetString("_SysAttr") & .IsSystem & "]"
    End With
    
    
    txtFolderInfos.Text = S
    
    Set cFol = Nothing  'libère
    
End Sub
