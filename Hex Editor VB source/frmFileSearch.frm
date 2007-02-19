VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6ADE9E73-F694-428F-BF86-06ADD29476A5}#1.0#0"; "ProgressBar_OCX.ocx"
Begin VB.Form frmFileSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche de fichiers"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar pgb 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "Avancement de la recherche"
      Top             =   5520
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmFileSearch.frx":08CA
      FrontPicture    =   "frmFileSearch.frx":08E6
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder les résultats..."
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      ToolTipText     =   "Sauvegarde les résultats de la recherche"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   8400
      TabIndex        =   25
      ToolTipText     =   "Ferme la fenêtre de recherche"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Lancer la recherche"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      ToolTipText     =   "Lance la recherche"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Résultats"
      Height          =   2895
      Left            =   3120
      TabIndex        =   21
      Top             =   2520
      Width           =   6615
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   6375
         TabIndex        =   22
         Top             =   240
         Width           =   6375
         Begin ComctlLib.ListView LVres 
            Height          =   2535
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   4471
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
               Text            =   "Objet"
               Object.Width           =   12347
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Emplacements"
      Height          =   3735
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   2895
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3375
         ScaleWidth      =   2655
         TabIndex        =   18
         Top             =   240
         Width           =   2655
         Begin ComctlLib.ListView LV 
            Height          =   2895
            Left            =   0
            TabIndex        =   20
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   5106
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
               Text            =   "Dossier"
               Object.Width           =   8819
            EndProperty
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Ajouter un dossier..."
            Height          =   255
            Left            =   0
            TabIndex        =   19
            ToolTipText     =   "Ajouter un dossier à la liste des emplacements où il faut rechercher"
            Top             =   120
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critères de recherche"
      Height          =   1815
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   1
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   6375
         TabIndex        =   3
         Top             =   240
         Width           =   6375
         Begin VB.CommandButton cmdDate 
            Caption         =   "..."
            Height          =   255
            Left            =   4200
            TabIndex        =   29
            ToolTipText     =   "Sélectionner une date..."
            Top             =   1080
            Width           =   375
         End
         Begin VB.ComboBox cbDateType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":0902
            Left            =   4680
            List            =   "frmFileSearch.frx":090F
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "pref"
            ToolTipText     =   "Type de date"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   16
            ToolTipText     =   "Date (et heure)"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.ComboBox cbOpDate 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":093D
            Left            =   1200
            List            =   "frmFileSearch.frx":0950
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "pref"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "Par taille"
            Height          =   195
            Left            =   0
            TabIndex        =   14
            ToolTipText     =   "Ajoute le critère 'date' à la recherche"
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cbOpSize 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":0965
            Left            =   1200
            List            =   "frmFileSearch.frx":0978
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "pref"
            ToolTipText     =   "Opérateur de recherche"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Tag             =   "pref"
            Text            =   "100"
            ToolTipText     =   "Taille"
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cdUnit 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmFileSearch.frx":098D
            Left            =   3840
            List            =   "frmFileSearch.frx":099D
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "pref"
            ToolTipText     =   "Unité"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkSize 
            Caption         =   "Par taille"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            ToolTipText     =   "Ajoute le critère 'taille' à la recherche"
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkCasse 
            Caption         =   "Casse"
            Height          =   195
            Left            =   3840
            TabIndex        =   9
            ToolTipText     =   "Respecte ou non la casse"
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            ToolTipText     =   "Nom à rechercher"
            Top             =   120
            Width           =   2415
         End
         Begin VB.CheckBox chkName 
            Caption         =   "Par nom"
            Height          =   195
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Ajoute le critère 'nom' à la recherche"
            Top             =   120
            Value           =   1  'Checked
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de recherche"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   2655
         TabIndex        =   1
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher dans des fichiers"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   6
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que dans les fichiers (lent)"
            Top             =   840
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher des dossiers"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que des dossiers"
            Top             =   480
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechercher des fichiers"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Tag             =   "pref"
            ToolTipText     =   "Ne recherche que des fichiers"
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmFileSearch"
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
'FORM DE RECHERCHE DE FICHIERS DANS LES DISQUES DURS
'-------------------------------------------------------

Private Sub chkDate_Click()
    If chkDate.Value Then
        txtDate.Enabled = True
        cbOpDate.Enabled = True
    Else
        txtDate.Enabled = False
        cbOpDate.Enabled = False
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub chkName_Click()
    If chkName.Value Then
        txtName.Enabled = True
        chkCasse.Enabled = True
    Else
        txtName.Enabled = False
        chkCasse.Enabled = False
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub chkSize_Click()
    If chkSize.Value Then
        cbOpSize.Enabled = True
        txtSize.Enabled = True
        cdUnit.Enabled = True
    Else
        cbOpSize.Enabled = False
        txtSize.Enabled = False
        cdUnit.Enabled = False
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

Private Sub cmdAdd_Click()
Dim s As String
    s = cFile.BrowseForFolder("Ajouter un dossier", Me.hwnd)    'browse un dossier
    
    If cFile.FolderExists(s) Then
        'alors ajoute le dossier à la liste des emplacements
        LV.ListItems.Add Text:=s
    End If
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

'-------------------------------------------------------
'checke si la recherche est possible ou non
'-------------------------------------------------------
Private Sub CheckSearch()

    cmdGo.Enabled = False
    If LV.ListItems.Count = 0 Then Exit Sub 'aucune zone à chercher
    
    If chkName.Value Then
        cmdGo.Enabled = True
        Exit Sub
    End If
    If chkSize.Value Then
        If Len(txtSize.Text) > 0 And cdUnit.ListIndex >= 0 And cbOpSize.ListIndex >= 0 Then
            cmdGo.Enabled = True
            Exit Sub
        End If
    End If
    If chkDate.Value Then
        If Len(txtDate.Text) > 0 And cbOpDate.ListIndex >= 0 And cbDateType.ListIndex >= 0 Then
            cmdGo.Enabled = True
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cmdSave_Click()
'sauvegarde les résultats


End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'supprime les items sélectionnés
Dim r As Long

    If KeyCode = vbKeyDelete Then
        'touche suppr
        For r = LV.ListItems.Count To 1 Step -1
            If LV.ListItems.Item(r).Selected Then LV.ListItems.Remove r
        Next r
    End If

    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub

Private Sub txtDate_Change()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub txtName_Change()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub txtSize_Change()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cbOpDate_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cbOpSize_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
Private Sub cdUnit_Click()
    Call CheckSearch 'vérifie qu'une recherche est possible
End Sub
