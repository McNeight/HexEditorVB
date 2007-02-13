VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Découpeur/fusionneur de fichiers"
   ClientHeight    =   4815
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
   ScaleHeight     =   4815
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   3090
      TabIndex        =   27
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Lancer"
      Height          =   495
      Left            =   810
      TabIndex        =   26
      Top             =   4200
      Width           =   1215
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   90
      TabIndex        =   25
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
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
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtGrFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   23
            ToolTipText     =   "Fichier à découper"
            Top             =   360
            Width           =   3615
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
         Begin VB.Frame Frame2 
            Height          =   1095
            Index           =   1
            Left            =   0
            TabIndex        =   17
            Top             =   1560
            Width           =   4695
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   735
               Index           =   1
               Left            =   120
               ScaleHeight     =   735
               ScaleWidth      =   4455
               TabIndex        =   18
               Top             =   240
               Width           =   4455
               Begin VB.TextBox txtFolderFus 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   0
                  TabIndex        =   20
                  ToolTipText     =   "Dossier du fichier fusionné"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.CommandButton cmdBrowsePaste 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   19
                  ToolTipText     =   "Sélectionner le dossier du fichier fusionné"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier résultat"
                  Height          =   255
                  Index           =   2
                  Left            =   0
                  TabIndex        =   21
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
            TabIndex        =   24
            Top             =   0
            Width           =   3735
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Index           =   0
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   4695
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         Begin VB.Frame Frame2 
            Height          =   2175
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   840
            Width           =   4695
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1815
               Index           =   0
               Left            =   120
               ScaleHeight     =   1815
               ScaleWidth      =   4455
               TabIndex        =   8
               Top             =   240
               Width           =   4455
               Begin VB.TextBox txtNumberOfFiles 
                  Alignment       =   2  'Center
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   16
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
                  TabIndex        =   15
                  Tag             =   "pref"
                  ToolTipText     =   "Taille"
                  Top             =   960
                  Width           =   975
               End
               Begin VB.ComboBox cdUnit 
                  Height          =   315
                  ItemData        =   "frmCut.frx":08CA
                  Left            =   3000
                  List            =   "frmCut.frx":08DA
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Tag             =   "pref"
                  ToolTipText     =   "Unité"
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Nombre de fichiers"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   13
                  Tag             =   "pref"
                  Top             =   1320
                  Width           =   1935
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
               Begin VB.CommandButton cmdBrowseGroupFile 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   10
                  ToolTipText     =   "Sélectionner le dossier du fichier groupeur"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox txtFolderResult 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   0
                  TabIndex        =   9
                  ToolTipText     =   "Dossier du fichier groupeur"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier du fichier groupeur"
                  Height          =   255
                  Index           =   1
                  Left            =   0
                  TabIndex        =   11
                  Top             =   0
                  Width           =   3735
               End
            End
         End
         Begin VB.CommandButton cmdBrowseToCut 
            Caption         =   "..."
            Height          =   255
            Left            =   3840
            TabIndex        =   6
            ToolTipText     =   "Sélectionner le fichier à couper"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFileToCut 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Fichier à découper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Découpage de fichier"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3735
         End
      End
   End
End
Attribute VB_Name = "frmCut"
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
'//FORM DE DECOUPAGE/COLLAGE DE FICHIERS
'-------------------------------------------------------

Private clsPref As clsIniForm

Private Sub cmdBrowseGroupFile_Click()
'sélection du dossier (groupeur à sauvegarder)
    txtFolderResult.Text = cFile.BrowseForFolder("Choix du dossier résultant", Me.hwnd)
End Sub

Private Sub cmdBrowseGrupFIle_Click()
'sélection du fichier groupe
    txtGrFile.Text = cFile.ShowOpen("Selection du fichier groupeur", Me.hwnd, "Fichier groupeur|*.grp")
    If txtFolderFus.Text = vbNullString Then txtFolderFus.Text = cFile.GetFolderFromPath(txtGrFile.Text)
End Sub

Private Sub cmdBrowsePaste_Click()
'sélection du dossier du fichier fusionné
    txtFolderFus.Text = cFile.BrowseForFolder("Choix du dossier du fichier fusionné", Me.hwnd)
End Sub

Private Sub cmdBrowseToCut_Click()
'sélection du fichier à découper
    txtFileToCut.Text = cFile.ShowOpen("Fichier à découper", Me.hwnd, "Tous|*.*", App.Path)
End Sub

Private Sub cmdProceed_Click()
'démarre
Dim tMethod As CUT_METHOD
Dim lLen As Double

    If Frame1(0).Visible Then
        'alors on découpe
        
        If Option1(0).Value Then
            'taille
            tMethod.tMethode = [Taille fixe]
            lLen = Abs(Val(txtSize.Text))
            If cdUnit.Text = "Ko" Then lLen = lLen * 1024
            If cdUnit.Text = "Mo" Then lLen = (lLen * 1024) * 1024
            If cdUnit.Text = "Go" Then lLen = ((lLen * 1024) * 1024) * 1024
            tMethod.lParam = lLen
        Else
            'nombre
            tMethod.tMethode = [Nombre fichiers fixe]
            tMethod.lParam = Val(txtNumberOfFiles.Text)
        End If
        
        'lance la découpe
        CutFile txtFileToCut.Text, txtFolderResult.Text, tMethod
        
    Else
    
        'lance la fusion
        PasteFile txtGrFile.Text, txtFolderFus.Text
        
    End If
    
End Sub

Private Sub cmdQuit_Click()
    'quitte
    Unload Me
End Sub

Private Sub Form_Load()

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
    
    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\CutFrm.ini", Me
    Option1(1).Value = Not (Option1(0).Value)
    
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

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\CutFrm.ini", Me
    Set clsPref = Nothing
End Sub
