VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "AboutBox"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4860
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4485
         Left            =   120
         ScaleHeight     =   4485
         ScaleWidth      =   5655
         TabIndex        =   1
         Top             =   240
         Width           =   5655
         Begin VB.CommandButton cmdUnload 
            Caption         =   "Fermer"
            Height          =   375
            Left            =   900
            TabIndex        =   4
            Top             =   4080
            Width           =   1335
         End
         Begin VB.CommandButton cmdLicense 
            Caption         =   "Informations de licence"
            Height          =   375
            Left            =   2700
            TabIndex        =   3
            Top             =   4080
            Width           =   2055
         End
         Begin VB.TextBox txt 
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   1560
            Width           =   5415
         End
         Begin VB.Image Image1 
            Height          =   795
            Left            =   360
            Picture         =   "frmAbout.frx":0000
            Stretch         =   -1  'True
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblMain 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "File Renamer VB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblMain 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Coded by violent_ken"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   6
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label lblMain 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "v 1.0.2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   5
            Top             =   720
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =======================================================
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' An Windows utility which allows to rename lots of file (part of Hex Editor VB)
'
' Copyright (c) 2006-2007 by Alain Descotes.
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

Private Sub cmdLicense_Click()
'affiche le ReadMe

    If cFile.FileExists(App.Path & "\License.txt") = False Then Exit Sub
    
    ShellExecute Me.hwnd, "open", App.Path & "\License.txt", vbNullString, vbNullString, 1
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim s As String
    'mise à jour de la version
    lblMain(1).Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
    'écriture du texte
    s = "File Renamer VB par violent ken. Dernière version du 02/01/2007. Il s'agit d'un utilitaire permettant de gérer vos fichiers, en particulier les renommer de manière massive et automatisée." & vbNewLine & vbNewLine & "Ce logiciel est prévu pour Windows XP et une résolution minimale de 1024*768." & vbNewLine & vbNewLine & "Ce logiciel est sous licence GNU, veuillez lire le fichier de licence qui accompagne ce logiciel."
    txt.Text = s
    
End Sub
