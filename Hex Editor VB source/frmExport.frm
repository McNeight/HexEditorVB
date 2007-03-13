VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporter"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Ne pas sauvegarder"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Lancer la sauvegarde"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Lancer la sauvegarde"
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format d'export"
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
      Begin VB.CheckBox chkOffset 
         Caption         =   "Ajouter les offsets"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Permer d'ajouter les offsets au fichier (en hexa ou décimal, selon les préférences)"
         Top             =   840
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   1
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3975
         TabIndex        =   5
         Top             =   240
         Width           =   3975
         Begin VB.CheckBox chkString 
            Caption         =   "Ajouter les valeurs ASCII"
            Height          =   195
            Left            =   0
            TabIndex        =   10
            ToolTipText     =   "Permet l'ajout des valeurs ASCII au fichier"
            Top             =   840
            Width           =   3495
         End
         Begin VB.ComboBox cbFormat 
            Height          =   315
            ItemData        =   "frmExport.frx":000C
            Left            =   0
            List            =   "frmExport.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Tag             =   "pref"
            ToolTipText     =   "Format d'exportation"
            Top             =   120
            Width           =   3975
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fichier à sauvegarder"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3975
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   3480
            TabIndex        =   3
            ToolTipText     =   "Choix du fichier à sauvegarder"
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   2
            ToolTipText     =   "Emplacement du fichier à sauvegarder"
            Top             =   120
            Width           =   3255
         End
      End
   End
End
Attribute VB_Name = "frmExport"
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
'FORM D'EXPORT MULTIFORMATS
'=======================================================

Private bEntireFile As Boolean

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'lance la sauvegarde

    Select Case cbFormat.Text
        Case "HTML"
        
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsHTML(txtFile.Text, CBool(chkOffset.Value), CBool(chkString.Value), _
                    frmContent.ActiveForm.Caption, -1)
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsHTML(txtFile.Text, CBool(chkOffset.Value), CBool(chkString.Value), _
                    "az", 1, 1)
            End If
            
        Case "RTF"
            
        Case "Texte"
            
        Case "Source c"
            
        Case "Source VB"
            
        Case "Source JAVA"
            
    End Select
            
End Sub

Private Sub Form_Load()
    bEntireFile = False
End Sub

'=======================================================
'à appeler si on veut sauver un fichier entier
'=======================================================
Public Sub IsEntireFile()
    bEntireFile = True
End Sub
