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
      TabIndex        =   2
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
      TabIndex        =   0
      Top             =   4040
      Width           =   6705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   105
      TabIndex        =   9
      Top             =   3960
      Width           =   6855
   End
   Begin VB.Label lblCopyright 
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
      Height          =   255
      Left            =   3937
      TabIndex        =   8
      Top             =   2872
      Width           =   3015
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Avertissement : ce logiciel est protégé par la license GNU General Public License"
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
      Left            =   5977
      TabIndex        =   6
      Top             =   2512
      Width           =   795
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed for Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3307
      TabIndex        =   5
      Top             =   2152
      Width           =   3525
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hex Editor VB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   2257
      TabIndex        =   4
      Top             =   1072
      Width           =   4500
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License accordée à [NAME]"
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


'---------------------------------------------------------------------------
'                            HEX EDITOR VB
'                         CODED BY VIOLENT_KEN
'                Dernière version v1.3 du 08/02/2007
'
'
'
'                        DESCRIPTION DU PROGRAMME
'
' Ce logiciel vous permet d'éditer le contenu de vos fichiers
' ainsi que le contenu de la mémoire virtuelle de vos
' processus et le contenu de vos disques physiques.
'
' Ce programme est conçu pour Windows XP/Vista, avec une résolution
' optimale minimale de 1024*768.
'
'
'
'
'                         HISTORIQUE DES VERSIONS
'
' v1.3 support de l'historique
' v1.2 Ajout de la gestion des disques
' v1.1 Ajout de la gestion de la modification des processus
'   en mémoire
' v1.0 Initial release
'---------------------------------------------------------------------------


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
    'mise à jour de la version et de l'USER
    lblLicenseTo.Caption = "License accordée à " & GetUserName
    lblVersion.Caption = "Version " & Str$(App.Major) & "." & Str$(App.Minor) & "." & Str$(App.Revision)
    
    'écriture du texte
    s = "HexEditor VB par violent ken. Dernière version du 02/01/2007. Il s'agit d'un éditeur héxadécimal complet permettant de modifier vos fichiers, vos disques et vos processus en mémoire très facilement." & vbNewLine & vbNewLine & "Ce logiciel est prévu pour Windows XP et une résolution minimale de 1024*768." & vbNewLine & vbNewLine & "Ce logiciel est sous license GNU, veuillez lire le fichier de licence ci dessous."
    s = s & vbNewLine & vbNewLine & cFile.LoadFileInString(App.Path & "\License.txt")
    txt.Text = s
    
End Sub
