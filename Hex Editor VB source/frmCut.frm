VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6ADE9E73-F694-428F-BF86-06ADD29476A5}#1.0#0"; "ProgressBar_OCX.ocx"
Begin VB.Form frmCut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D�coupeur/fusionneur de fichiers"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar pgb 
      Height          =   255
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Etat d'avancement"
      Top             =   4560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmCut.frx":08CA
      FrontPicture    =   "frmCut.frx":08E6
   End
   Begin VB.TextBox txtBufSize 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      TabIndex        =   29
      Text            =   "5"
      ToolTipText     =   "Taille du buffer"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   3060
      TabIndex        =   27
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Lancer"
      Height          =   495
      Left            =   780
      TabIndex        =   26
      Top             =   4920
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
            Caption         =   "D�couper"
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
                  ItemData        =   "frmCut.frx":0902
                  Left            =   3000
                  List            =   "frmCut.frx":0912
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Tag             =   "pref"
                  ToolTipText     =   "Unit�"
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
                  ToolTipText     =   "S�lectionner le dossier du fichier groupeur"
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
            ToolTipText     =   "S�lectionner le fichier � couper"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFileToCut 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Fichier � d�couper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "D�coupage de fichier"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3735
         End
      End
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
            ToolTipText     =   "Fichier � d�couper"
            Top             =   360
            Width           =   3615
         End
         Begin VB.CommandButton cmdBrowseGrupFIle 
            Caption         =   "..."
            Height          =   255
            Left            =   3840
            TabIndex        =   22
            ToolTipText     =   "S�lectionner le fichier groupeur"
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
                  ToolTipText     =   "Dossier du fichier fusionn�"
                  Top             =   360
                  Width           =   3615
               End
               Begin VB.CommandButton cmdBrowsePaste 
                  Caption         =   "..."
                  Height          =   255
                  Left            =   3840
                  TabIndex        =   19
                  ToolTipText     =   "S�lectionner le dossier du fichier fusionn�"
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label Label1 
                  Caption         =   "Dossier r�sultat"
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
   Begin VB.Label Label2 
      Caption         =   "Taille du buffer (valeur enti�re en Mo)"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   2775
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
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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
'FORM DE DECOUPAGE/COLLAGE DE FICHIERS
'-------------------------------------------------------

Private Sub cmdBrowseGroupFile_Click()
's�lection du dossier (groupeur � sauvegarder)
    txtFolderResult.Text = cFile.BrowseForFolder("Choix du dossier r�sultant", Me.hwnd)
End Sub

Private Sub cmdBrowseGrupFIle_Click()
's�lection du fichier groupe
    txtGrFile.Text = cFile.ShowOpen("Selection du fichier groupeur", Me.hwnd, "Fichier groupeur|*.grp")
    If txtFolderFus.Text = vbNullString Then txtFolderFus.Text = cFile.GetFolderFromPath(txtGrFile.Text)
End Sub

Private Sub cmdBrowsePaste_Click()
's�lection du dossier du fichier fusionn�
    txtFolderFus.Text = cFile.BrowseForFolder("Choix du dossier du fichier fusionn�", Me.hwnd)
End Sub

Private Sub cmdBrowseToCut_Click()
's�lection du fichier � d�couper
    txtFileToCut.Text = cFile.ShowOpen("Fichier � d�couper", Me.hwnd, "Tous|*.*", App.Path)
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
        'buffer trop grand (peut etre risques potentiels de d�passement de capacit�)
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
        'alors on d�coupe
        
        If Option1(0).Value Then
            'taille
            tMethod.tMethode = [Taille fixe]
            lLen = Abs(Val(txtSize.Text))
            If cdUnit.Text = "Ko" Then lLen = lLen * 1024
            If cdUnit.Text = "Mo" Then lLen = (lLen * 1024) * 1024
            If cdUnit.Text = "Go" Then lLen = ((lLen * 1024) * 1024) * 1024
            If lLen >= 2147483648# Then
                'alors fichiers trop grands
                MsgBox "La taille maximale des fichiers r�sultants ne peut exc�der 2Go.", vbCritical, "Attention"
                GoTo CannotCut
            End If
                
            tMethod.lParam = lLen
        Else
            'nombre
        
            'calcule la taille de chaque fichier
            lLen = Int(cFile.GetFileSize(txtFileToCut.Text) / Val(txtNumberOfFiles.Text))
            
            If lLen > 2147483648# Then
                'alors fichiers trop grands
                MsgBox "La taille maximale des fichiers r�sultants ne peut exc�der 2Go.", vbCritical, "Attention"
                GoTo CannotCut
            End If
            
            tMethod.tMethode = [Nombre fichiers fixe]
            tMethod.lParam = Abs(Int(Val(txtNumberOfFiles.Text)))
        End If
        
        'lance la d�coupe
        lTime = CutFile(txtFileToCut.Text, txtFolderResult.Text, tMethod)
        
    Else
    
        'lance la fusion
        lTime = PasteFile(txtGrFile.Text, txtFolderFus.Text)
        
    End If

    Me.Caption = "Effectu� en " & Trim$(Str$(lTime / 1000)) & " s"
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
        'alors on d�coupe en un nombre fini de fichiers
        txtNumberOfFiles.Enabled = True
        txtSize.Enabled = False
        cdUnit.Enabled = False
    Else
        'd�coupage en taille fixe
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
