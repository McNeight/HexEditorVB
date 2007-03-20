VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmDisAsm 
   BackColor       =   &H8000000C&
   Caption         =   "Désassembleur d'exécutables"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   7425
   Icon            =   "frmDisAsm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5385
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14993
            MinWidth        =   14993
            Text            =   "Status=[Ready]"
            TextSave        =   "Status=[Ready]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Menu rmnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuDisAsm 
         Caption         =   "&Désassembler un fichier..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu rnuSave 
         Caption         =   "&Enregistrer"
         Begin VB.Menu mnuSaveASM 
            Caption         =   "&Liste des instructions ASM..."
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu rmnuWindow 
      Caption         =   "&Fenêtres"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&En cascade"
      End
      Begin VB.Menu mnuMH 
         Caption         =   "Mosaïque &horizontale"
      End
      Begin VB.Menu mnuMV 
         Caption         =   "Mosaïque &verticale"
      End
      Begin VB.Menu mnuReorganize 
         Caption         =   "&Réorganiser les icones"
      End
   End
   Begin VB.Menu rmnuHelp 
      Caption         =   "&Aide"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Aide..."
      End
      Begin VB.Menu mnuTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&A propos"
      End
   End
End
Attribute VB_Name = "frmDisAsm"
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
'PROJET DE DESASSEMBLAGE D'EXECUTABLES
'=======================================================

Private Sub mnuQuit_Click()
    Unload Me
End Sub
