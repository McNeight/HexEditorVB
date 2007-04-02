VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{9B9A881F-DBDC-4334-BC23-5679E5AB0DC6}#1.2#0"; "FileView_OCX.ocx"
Begin VB.Form frmRecoverFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "R�cup�ration de fichiers"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecoverFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4815
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4575
         Index           =   1
         Left            =   120
         ScaleHeight     =   4575
         ScaleWidth      =   6735
         TabIndex        =   4
         Top             =   120
         Width           =   6735
         Begin VB.CommandButton cmdRestoreValidFile 
            Caption         =   "Restaurer le fichier s�lectionn�..."
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   4200
            Width           =   3855
         End
         Begin VB.TextBox pctPath 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   3840
            Width           =   6615
         End
         Begin FileView_OCX.FileView LV 
            Height          =   3615
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   6376
            ShowEntirePath  =   0   'False
            AllowMultiSelect=   0   'False
            Path            =   "C:\"
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
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      Begin VB.PictureBox Picture1 
         Height          =   4575
         Index           =   0
         Left            =   120
         ScaleHeight     =   4515
         ScaleWidth      =   6675
         TabIndex        =   2
         Top             =   120
         Width           =   6735
      End
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fichiers effac�s"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fichiers existants"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Extraire des donn�es"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture1 
         Height          =   4575
         Index           =   2
         Left            =   120
         ScaleHeight     =   4515
         ScaleWidth      =   6675
         TabIndex        =   6
         Top             =   120
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmRecoverFiles"
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
' =======================================================


Option Explicit



'=======================================================
'FORM DE RECUPERATION DE FICHIERS
'=======================================================

Private Sub cmdRestoreValidFile_Click()
'restaure le fichier s�lectionn� dans le FV2
    MsgBox LV.ListItems.Item(LV.ListIndex).Text
End Sub

Private Sub Form_Load()
    Me.Width = 7275
    Me.Height = 5715
    
    'path par d�faut
    LV.Path = Left$(App.Path, 3)
    pctPath.Text = LV.Path
End Sub

Private Sub Form_Resize()
Dim x As Long

    'positionnement des frames
    For x = 0 To Frame1.Count - 1
        Frame1(x).Top = 480
        Frame1(x).Left = 120
    Next x
End Sub

Private Sub LV_PathChange(sOldPath As String, sNewPath As String)
    pctPath.Text = sNewPath
End Sub

Private Sub pctPath_Change()
    If cFile.FolderExists(cFile.GetFolderFromPath(pctPath.Text & "\")) = False Then
        'couleur rouge
        pctPath.ForeColor = RED_COLOR
    Else
        'c'est un path ok
        pctPath.ForeColor = GREEN_COLOR
    End If
End Sub

Private Sub pctPath_KeyDown(KeyCode As Integer, Shift As Integer)
'valide si entr�e
Dim s As String
    If KeyCode = vbKeyReturn Then
        s = pctPath.Text
        If cFile.FolderExists(pctPath.Text) Then LV.Path = pctPath.Text
        pctPath.Text = s
    End If
End Sub

Private Sub pctPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0 'vire le BEEP
End Sub

Private Sub TB_Click()
    Frame1(0).Visible = False
    Frame1(1).Visible = False
    Frame1(2).Visible = False
    Frame1(TB.SelectedItem.Index - 1).Visible = True
End Sub
