VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5B5F5394-748F-414C-9FDD-08F3427C6A09}#3.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmCut 
   BackColor       =   &H00F9E5D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Découpeur/fusionneur de fichiers"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   36
   Icon            =   "frmCut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkBar PGB 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4560
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      Value           =   1
      BackPicture     =   "frmCut.frx":058A
      FrontPicture    =   "frmCut.frx":05A6
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
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Lancer"
      Height          =   495
      Left            =   765
      TabIndex        =   2
      Top             =   4908
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   3045
      TabIndex        =   1
      Top             =   4908
      Width           =   1215
   End
   Begin VB.TextBox txtBufSize 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2985
      TabIndex        =   0
      Text            =   "5"
      ToolTipText     =   "Taille du buffer"
      Top             =   4188
      Width           =   1935
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   75
      TabIndex        =   3
      Tag             =   "lang_ok"
      Top             =   88
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      TabWidthStyle   =   2
      TabFixedWidth   =   4322
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Découper"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fusionner"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   468
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   0
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtFileToCut 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   17
            ToolTipText     =   "Fichier à découper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdBrowseToCut 
            Caption         =   "..."
            Height          =   255
            Left            =   3840
            TabIndex        =   16
            ToolTipText     =   "Sélectionner le fichier à couper"
            Top             =   360
            Width           =   615
         End
         Begin VB.Frame Frame2 
            Height          =   2175
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   840
            Width           =   4695
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1815
               Index           =   0
               Left            =   120
               ScaleHeight     =   1815
               ScaleWidth      =   4455
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   240
               Width           =   4455
               Begin VB.TextBox txtFolderResult 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   0
                  TabIndex        =   14
                  ToolTipText     =   "Dossier du fichier groupeur"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.CommandButton cmdBrowseGroupFile 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   13
                  ToolTipText     =   "Sélectionner le dossier du fichier groupeur"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Taille des fichiers"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   12
                  Tag             =   "pref"
                  Top             =   960
                  Value           =   -1  'True
                  Width           =   2055
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Nombre de fichiers"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   11
                  Tag             =   "pref"
                  Top             =   1320
                  Width           =   1935
               End
               Begin VB.ComboBox cdUnit 
                  Height          =   315
                  ItemData        =   "frmCut.frx":05C2
                  Left            =   3000
                  List            =   "frmCut.frx":05D2
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Tag             =   "pref lang_ok"
                  ToolTipText     =   "Unité"
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.TextBox txtSize 
                  Alignment       =   2  'Center
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   9
                  Tag             =   "pref"
                  ToolTipText     =   "Taille"
                  Top             =   960
                  Width           =   975
               End
               Begin VB.TextBox txtNumberOfFiles 
                  Alignment       =   2  'Center
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   8
                  Tag             =   "pref"
                  ToolTipText     =   "Nombre"
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier du fichier groupeur"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   15
                  Top             =   0
                  Width           =   3735
               End
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Découpage de fichier"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Index           =   1
      Left            =   105
      TabIndex        =   19
      Top             =   468
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   1
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
         Begin VB.Frame Frame2 
            Height          =   1095
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   1560
            Width           =   4695
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   1
               Left            =   120
               ScaleHeight     =   735
               ScaleWidth      =   4455
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   240
               Width           =   4455
               Begin VB.CommandButton cmdBrowsePaste 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   26
                  ToolTipText     =   "Sélectionner le dossier du fichier fusionné"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox txtFolderFus 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   0
                  TabIndex        =   25
                  ToolTipText     =   "Dossier du fichier fusionné"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier résultat"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   27
                  Top             =   0
                  Width           =   3735
               End
            End
         End
         Begin VB.CommandButton cmdBrowseGrupFIle 
            Caption         =   "..."
            Height          =   255
            Left            =   3840
            TabIndex        =   22
            ToolTipText     =   "Sélectionner le fichier groupeur"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtGrFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   21
            ToolTipText     =   "Fichier à découper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Fichier groupeur"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Taille du buffer (valeur entière en Mo)"
      Height          =   255
      Left            =   105
      TabIndex        =   29
      Top             =   4188
      Width           =   2775
   End
End
Attribute VB_Name = "frmCut"
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
'FORM DE DECOUPAGE/COLLAGE DE FICHIERS
'=======================================================
Private Lang As New clsLang

Private Sub cmdBrowseGroupFile_Click()
'sélection du dossier (groupeur à sauvegarder)
    txtFolderResult.Text = cFile.BrowseForFolder(Lang.GetString("_SelectFolder"), Me.hWnd)
End Sub

Private Sub cmdBrowseGrupFIle_Click()
'sélection du fichier groupe
    txtGrFile.Text = cFile.ShowOpen(Lang.GetString("_GrupFileSel"), Me.hWnd, _
        Lang.GetString("_GrupFile") & "|*.grp")
    If txtFolderFus.Text = vbNullString Then txtFolderFus.Text = _
        cFile.GetFolderName(txtGrFile.Text)
End Sub

Private Sub cmdBrowsePaste_Click()
'sélection du dossier du fichier fusionné
    txtFolderFus.Text = cFile.BrowseForFolder(Lang.GetString("_FusionPathSel"), _
        Me.hWnd)
End Sub

Private Sub cmdBrowseToCut_Click()
'sélection du fichier à découper
    txtFileToCut.Text = cFile.ShowOpen(Lang.GetString("_FileToCutSel"), _
        Me.hWnd, "Tous|*.*", App.Path)
End Sub

Private Sub cmdProceed_Click()
Dim tMethod As CUT_METHOD
Dim lLen As Double
Dim lTime As Long
    
    lBufSize = Abs(Int(Val(txtBufSize.Text))) * 1024 * 1024
    
    If lBufSize < 1048576 Then
        'alors buffer trop faible
        MsgBox Lang.GetString("_BufferNotSuff"), vbCritical, _
            Lang.GetString("_War")
        Exit Sub
    ElseIf lBufSize > 31457280 Then
        'buffer trop grand (peut etre risques potentiels de dépassement de capacité)
        MsgBox Lang.GetString("_BufferTooSuff"), vbCritical, _
            Lang.GetString("_War")
        Exit Sub
    End If
    
    cmdQuit.Enabled = False
    cmdProceed.Enabled = False
    TB.Enabled = False
    txtGrFile.Enabled = False
    cmdBrowseGroupFile.Enabled = False
    txtFolderFus.Enabled = False
    cmdBrowsePaste.Enabled = False
    txtBufSize.Enabled = False
    txtFileToCut.Enabled = False
    txtFolderResult.Enabled = False
    cmdBrowseToCut.Enabled = False
    cmdBrowseGroupFile.Enabled = False
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    txtSize.Enabled = False
    txtNumberOfFiles.Enabled = False
    cdUnit.Enabled = False

    If Frame1(0).Visible Then
        'alors on découpe
        
        'ajoute du texte à la console
        Call AddTextToConsole(Lang.GetString("_DecLau"))
        
        If Option1(0).Value Then
            'taille
            tMethod.tMethod = [Taille fixe]
            lLen = Abs(Val(txtSize.Text))
            With Lang
                If cdUnit.Text = .GetString("_Ko") Then lLen = lLen * 1024
                If cdUnit.Text = .GetString("_Mo") Then lLen = (lLen * 1024) * 1024
                If cdUnit.Text = .GetString("_Go") Then lLen = ((lLen * 1024) * 1024) * 1024
            End With
            If lLen >= 2147483648# Then
                'alors fichiers trop grands
                MsgBox Lang.GetString("_MoreThanTwo"), vbCritical, Lang.GetString("_War")
                GoTo CannotCut
            End If
                
            tMethod.lParam = lLen
        Else
            'nombre
        
            'calcule la taille de chaque fichier
            lLen = Int(cFile.GetFileSize(txtFileToCut.Text) / Val(txtNumberOfFiles.Text))
            
            If lLen > 2147483648# Then
                'alors fichiers trop grands
                MsgBox Lang.GetString("_MoreThanTwo"), vbCritical, Lang.GetString("_War")
                GoTo CannotCut
            End If
            
            With tMethod
                .tMethod = [Nombre fichiers fixe]
                .lParam = Abs(Int(Val(txtNumberOfFiles.Text)))
            End With
        End If
        
        'lance la découpe
        lTime = CutFile(txtFileToCut.Text, txtFolderResult.Text, tMethod)
        
    Else
    
        'ajoute du texte à la console
        Call AddTextToConsole(Lang.GetString("_FusionLau"))
        
        'lance la fusion
        lTime = PasteFile(txtGrFile.Text, txtFolderFus.Text)
        
    End If

    Me.Caption = "Effectué en " & Trim$(Str$(lTime / 1000)) & " s"
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_OpeFin"))
    
CannotCut:
    cmdQuit.Enabled = True
    cmdProceed.Enabled = True
    TB.Enabled = True
    txtGrFile.Enabled = True
    cmdBrowseGroupFile.Enabled = True
    txtFolderFus.Enabled = True
    cmdBrowsePaste.Enabled = True
    txtBufSize.Enabled = True
    txtFileToCut.Enabled = True
    txtFolderResult.Enabled = True
    cmdBrowseToCut.Enabled = True
    cmdBrowseGroupFile.Enabled = True
    Option1(0).Enabled = True
    Option1(1).Enabled = True
    txtSize.Enabled = True
    txtNumberOfFiles.Enabled = True
    cdUnit.Enabled = True
End Sub

Private Sub cmdQuit_Click()
    'quitte
    Unload Me
End Sub

Private Sub Form_Load()

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
        Call .LoadControlsCaption
    End With
    
    'redimensionne et place les frames
    With Frame1(0)
        .Width = 4935
        .Top = 600
        .Left = 120
        .Height = 3615
    End With
    With Frame1(1)
        .Width = 4935
        .Top = 600
        .Left = 120
        .Height = 3615
    End With
    
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(1).Value = True Then
        'alors on découpe en un nombre fini de fichiers
        txtNumberOfFiles.Enabled = True
        txtSize.Enabled = False
        cdUnit.Enabled = False
    Else
        'découpage en taille fixe
        txtNumberOfFiles.Enabled = False
        txtSize.Enabled = True
        cdUnit.Enabled = True
    End If
        
End Sub

Private Sub TB_Click()
'change le Tab actif (et donc le frame visible)
    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Frame1(TB.SelectedItem.Index - 1).Visible = True
End Sub
