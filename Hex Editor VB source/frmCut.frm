VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmCut 
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
   Icon            =   "frmCut.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar pgb 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Etat d'avancement"
      Top             =   4560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmCut.frx":058A
      FrontPicture    =   "frmCut.frx":05A6
   End
   Begin VB.TextBox txtBufSize 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   15
      Text            =   "5"
      ToolTipText     =   "Taille du buffer"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   3060
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Lancer"
      Height          =   495
      Left            =   780
      TabIndex        =   0
      Top             =   4920
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   90
      TabIndex        =   28
      Tag             =   "lang_ok"
      Top             =   100
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
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   0
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
         Begin VB.Frame Frame2 
            Height          =   2175
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   840
            Width           =   4695
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1815
               Index           =   0
               Left            =   120
               ScaleHeight     =   1815
               ScaleWidth      =   4455
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   240
               Width           =   4455
               Begin VB.TextBox txtNumberOfFiles 
                  Alignment       =   2  'Center
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   10
                  Tag             =   "pref"
                  ToolTipText     =   "Nombre"
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.TextBox txtSize 
                  Alignment       =   2  'Center
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   7
                  Tag             =   "pref"
                  ToolTipText     =   "Taille"
                  Top             =   960
                  Width           =   975
               End
               Begin VB.ComboBox cdUnit 
                  Height          =   315
                  ItemData        =   "frmCut.frx":05C2
                  Left            =   3000
                  List            =   "frmCut.frx":05D2
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Tag             =   "pref lang_ok"
                  ToolTipText     =   "Unité"
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Nombre de fichiers"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   9
                  Tag             =   "pref"
                  Top             =   1320
                  Width           =   1935
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Taille des fichiers"
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   6
                  Tag             =   "pref"
                  Top             =   960
                  Value           =   -1  'True
                  Width           =   2055
               End
               Begin VB.CommandButton cmdBrowseGroupFile 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   5
                  ToolTipText     =   "Sélectionner le dossier du fichier groupeur"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox txtFolderResult 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   0
                  TabIndex        =   4
                  ToolTipText     =   "Dossier du fichier groupeur"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier du fichier groupeur"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   23
                  Top             =   0
                  Width           =   3735
               End
            End
         End
         Begin VB.CommandButton cmdBrowseToCut 
            Caption         =   "..."
            Height          =   255
            Left            =   3840
            TabIndex        =   3
            ToolTipText     =   "Sélectionner le fichier à couper"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFileToCut 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   2
            ToolTipText     =   "Fichier à découper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Découpage de fichier"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   1
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtGrFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   11
            ToolTipText     =   "Fichier à découper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdBrowseGrupFIle 
            Caption         =   "..."
            Height          =   255
            Left            =   3840
            TabIndex        =   12
            ToolTipText     =   "Sélectionner le fichier groupeur"
            Top             =   360
            Width           =   615
         End
         Begin VB.Frame Frame2 
            Height          =   1095
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Top             =   1560
            Width           =   4695
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   1
               Left            =   120
               ScaleHeight     =   735
               ScaleWidth      =   4455
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   240
               Width           =   4455
               Begin VB.TextBox txtFolderFus 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   0
                  TabIndex        =   13
                  ToolTipText     =   "Dossier du fichier fusionné"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.CommandButton cmdBrowsePaste 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   14
                  ToolTipText     =   "Sélectionner le dossier du fichier fusionné"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier résultat"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   26
                  Top             =   0
                  Width           =   3735
               End
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Fichier groupeur"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Taille du buffer (valeur entière en Mo)"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4200
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

Private Sub cmdBrowseGroupFile_Click()
'sélection du dossier (groupeur à sauvegarder)
    txtFolderResult.Text = cFile.BrowseForFolder("Choix du dossier résultant", Me.hWnd)
End Sub

Private Sub cmdBrowseGrupFIle_Click()
'sélection du fichier groupe
    txtGrFile.Text = cFile.ShowOpen("Selection du fichier groupeur", Me.hWnd, "Fichier groupeur|*.grp")
    If txtFolderFus.Text = vbNullString Then txtFolderFus.Text = cFile.GetFolderFromPath(txtGrFile.Text)
End Sub

Private Sub cmdBrowsePaste_Click()
'sélection du dossier du fichier fusionné
    txtFolderFus.Text = cFile.BrowseForFolder("Choix du dossier du fichier fusionné", Me.hWnd)
End Sub

Private Sub cmdBrowseToCut_Click()
'sélection du fichier à découper
    txtFileToCut.Text = cFile.ShowOpen("Fichier à découper", Me.hWnd, "Tous|*.*", App.Path)
End Sub

Private Sub cmdProceed_Click()
Dim tMethod As CUT_METHOD
Dim lLen As Double
Dim lTime As Long
    
    lBufSize = Abs(Int(Val(txtBufSize.Text))) * 1024 * 1024
    If lBufSize < 1048576 Then
        'alors buffer trop faible
        MsgBox "La taille du buffer n'est pas suffisante.", vbCritical, "Attention"
        Exit Sub
    ElseIf lBufSize > 31457280 Then
        'buffer trop grand (peut etre risques potentiels de dépassement de capacité)
        MsgBox "la taille du buffer est trop grande.", vbCritical, "Attention"
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
        Call AddTextToConsole("Découpage lancé...")
        
        If Option1(0).Value Then
            'taille
            tMethod.tMethode = [Taille fixe]
            lLen = Abs(Val(txtSize.Text))
            If cdUnit.Text = "Ko" Then lLen = lLen * 1024
            If cdUnit.Text = "Mo" Then lLen = (lLen * 1024) * 1024
            If cdUnit.Text = "Go" Then lLen = ((lLen * 1024) * 1024) * 1024
            If lLen >= 2147483648# Then
                'alors fichiers trop grands
                MsgBox "La taille maximale des fichiers résultants ne peut excéder 2Go.", vbCritical, "Attention"
                GoTo CannotCut
            End If
                
            tMethod.lParam = lLen
        Else
            'nombre
        
            'calcule la taille de chaque fichier
            lLen = Int(cFile.GetFileSize(txtFileToCut.Text) / Val(txtNumberOfFiles.Text))
            
            If lLen > 2147483648# Then
                'alors fichiers trop grands
                MsgBox "La taille maximale des fichiers résultants ne peut excéder 2Go.", vbCritical, "Attention"
                GoTo CannotCut
            End If
            
            tMethod.tMethode = [Nombre fichiers fixe]
            tMethod.lParam = Abs(Int(Val(txtNumberOfFiles.Text)))
        End If
        
        'lance la découpe
        lTime = CutFile(txtFileToCut.Text, txtFolderResult.Text, tMethod)
        
    Else
    
        'ajoute du texte à la console
        Call AddTextToConsole("Fusion lancée...")
        
        'lance la fusion
        lTime = PasteFile(txtGrFile.Text, txtFolderFus.Text)
        
    End If

    Me.Caption = "Effectué en " & Trim$(Str$(lTime / 1000)) & " s"
    
    'ajoute du texte à la console
    Call AddTextToConsole("Opération terminée")
    
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

    #If MODE_DEBUG Then
        If App.LogMode = 0 Then
            'on créé le fichier de langue français
            Lang.Language = "French"
            Lang.LangFolder = LANG_PATH
            Lang.WriteIniFileFormIDEform
        End If
    #End If
    
    If App.LogMode = 0 Then
        'alors on est dans l'IDE
        Lang.LangFolder = LANG_PATH
    Else
        Lang.LangFolder = App.Path & "\Lang"
    End If
    
    'applique la langue désirée aux controles
    Lang.Language = cPref.env_Lang
    Lang.LoadControlsCaption
    
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
