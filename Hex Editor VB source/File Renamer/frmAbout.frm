VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00760401&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   6390
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      ToolTipText     =   "Fermer cette feuille"
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdLicense 
      Caption         =   "Informations de licence"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Afficher les informations sur la licence GNU GPL"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1660
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4040
      Width           =   6705
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2006-2007 Alain Descotes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   11
      Top             =   3000
      Width           =   4515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Renamer Tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2505
      TabIndex        =   10
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label lblVersionWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Pre Alpha version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   105
      TabIndex        =   8
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label lblWARNING 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement : ce logiciel est prot�g� par la license GNU General Public License"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3585
      Width           =   6855
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5970
      TabIndex        =   6
      Top             =   2685
      Width           =   795
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed for Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3915
      TabIndex        =   5
      Top             =   2340
      Width           =   2910
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex Editor VB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2265
      TabIndex        =   4
      Top             =   900
      Width           =   4275
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License accord�e � [NAME]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   165
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   2265
      Left            =   127
      Picture         =   "frmAbout.frx":5AAEA
      Stretch         =   -1  'True
      Top             =   982
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdLicense_Click()
'affiche le ReadMe

    If cFile.FileExists(App.Path & "\License.txt") = False Then Exit Sub
    cFile.ShellOpenFile App.Path & "\License.txt", Me.hWnd
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim s As String
    'mise � jour de la version et de l'USER
    lblLicenseTo.Caption = "License accord�e � " & GetUserName
    lblVersion.Caption = "Version " & Trim$(Str$(App.Major)) & "." & Trim$(Str$(App.Minor)) & "." & Trim$(Str$(App.Revision))
    
    '�criture du texte
    s = "Hex Editor VB" & vbNewLine & "Copyright (c) 2006-2007 Alain Descotes (violent_ken)" & vbNewLine & "Derni�re version du 16/02/2007. Il s'agit d'un �diteur h�xad�cimal complet permettant de modifier vos fichiers, vos disques et vos processus en m�moire tr�s facilement." & vbNewLine & vbNewLine & "Ce logiciel est pr�vu pour Windows XP/Vista et une r�solution minimale de 1024*768." & vbNewLine & vbNewLine & "Ce logiciel est sous license GNU General Public License, veuillez lire le contrat de licence ci dessous."
    s = s & vbNewLine & vbNewLine & cFile.LoadFileInString(App.Path & "\License.txt")
    txt.Text = s
    
End Sub
