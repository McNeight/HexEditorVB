VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
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
   ScaleHeight     =   4845
   ScaleWidth      =   5895
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
         Begin VB.PictureBox pctIcon 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   120
            Picture         =   "frmAbout.frx":000C
            ScaleHeight     =   960
            ScaleWidth      =   960
            TabIndex        =   5
            Top             =   0
            Width           =   960
         End
         Begin VB.TextBox txt 
            Height          =   2175
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1560
            Width           =   5415
         End
         Begin VB.CommandButton cmdLicense 
            Caption         =   "Informations de licence"
            Height          =   375
            Left            =   2700
            TabIndex        =   3
            ToolTipText     =   "Afficher les informations sur la licence GNU GPL"
            Top             =   4080
            Width           =   2055
         End
         Begin VB.CommandButton cmdUnload 
            Caption         =   "Fermer"
            Height          =   375
            Left            =   900
            TabIndex        =   2
            ToolTipText     =   "Fermer cette feuille"
            Top             =   4080
            Width           =   1335
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
            TabIndex        =   7
            Top             =   720
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
            Caption         =   "Hex Editor VB"
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
            TabIndex        =   8
            Top             =   240
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
    'mise à jour de la version
    lblMain(1).Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
    'écriture du texte
    s = "HexEditor VB par violent ken. Dernière version du 02/01/2007. Il s'agit d'un éditeur héxadécimal complet permettant de modifier vos fichiers, vos disques et vos processus en mémoire très facilement." & vbNewLine & vbNewLine & "Ce logiciel est prévu pour Windows XP et une résolution minimale de 1024*768." & vbNewLine & vbNewLine & "Ce logiciel est sous license GNU, veuillez lire le fichier de licence ci dessous."
    s = s & vbNewLine & vbNewLine & cFile.LoadFileInString(App.Path & "\License.txt")
    txt.Text = s
    
End Sub
