VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRecoverFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Récupération de fichiers"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7185
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
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
            Caption         =   "Fichiers effacés"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fichiers existants"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Extraire des données"
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
   Begin VB.Frame Frame1 
      Height          =   4815
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6975
      Begin VB.PictureBox Picture1 
         Height          =   4575
         Index           =   1
         Left            =   120
         ScaleHeight     =   4515
         ScaleWidth      =   6675
         TabIndex        =   4
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
'FORM DE RECUPERATION DE FICHIERS
'=======================================================

Private Sub Form_Resize()
Dim x As Long

    'positionnement des frames
    For x = 0 To Frame1.Count - 1
        Frame1(x).Top = 480
        Frame1(x).Left = 120
    Next x
End Sub
