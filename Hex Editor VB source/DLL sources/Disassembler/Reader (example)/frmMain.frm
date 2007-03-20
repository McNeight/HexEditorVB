VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Test Form"
   ClientHeight    =   6990
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2415
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4260
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":000C
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
   Begin VB.Menu mnuDisassembleNow 
      Caption         =   "&Désassembler"
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "frmMain"
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
'FORM DE TEST DE LA DLL DE DESASSEMBLAGE
'=======================================================

'=======================================================
'APIS
'=======================================================
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Sub Form_Resize()
    With RTB
        .Left = 0
        .Top = 0
        .Width = Me.Width - 290
        .Height = Me.Height - 890
    End With
End Sub

Private Sub mnuDisassembleNow_Click()
'lance le désassemblage ==> appel à la fonction de la dll
Dim l As Long
Dim s As String
Dim cFile As New clsFileInfos
    
    s = cFile.ShowOpen("Choix du fichier à désassembler (un répertoire sera créé dans le dossier de ce programme)", Me.hWnd, "Executables|*.exe|Dll|*.dll")
    
    If cFile.FileExists(s) = False Then Exit Sub
    
    Me.Caption = "Désassemblage en cours...."
    
    l = GetTickCount
    
    'appelle la dll et lance le désassemblage du fichier
    Call DisassembleWin32Executable(s, cFile.GetFolderFromPath(s) & "\DisAsm_Dir\")
    
    Me.Caption = "Désassemblage terminé en " & Trim$(Str$(GetTickCount - l)) & " ms"
    
    'affiche le résultat (instructions ASM uniquement)
    Call RTB.LoadFile(cFile.GetFolderFromPath(s) & "\DisAsm_Dir\" & Left$(cFile.GetFileFromPath(s), Len(cFile.GetFileFromPath(s)) - 4) & ".asm")
    
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub
