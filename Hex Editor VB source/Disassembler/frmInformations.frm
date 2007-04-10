VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmInformations 
   Caption         =   "Informations sur le fichier"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInformations.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   8895
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView LV 
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   2880
      Width           =   5000
      _ExtentX        =   8811
      _ExtentY        =   1296
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Characteristics"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PointerToRawData"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SizeOfRawData"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "VirtualAddress"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "VirtualSize"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PointerToLinenumbers"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "NumberOfLinenumbers"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PointerToRelocations"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "NumberOfRelocations"
         Object.Width           =   2469
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmInformations.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmInformations"
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
'AFFICHAGE DES INFORMATIONS SUR LA STRCTURE DU FICHIER
'=======================================================

Private sStr As String

Private Sub Form_Load()
    #If MODE_DEBUG Then
        If App.LogMode = 0 And CREATE_FRENCH_FILE Then
            'on créé le fichier de langue français
            Lang.Language = "French"
            Lang.LangFolder = LANG_PATH
            Lang.WriteIniFileFormIDEform
        End If
    #End If
    
    If App.LogMode = 0 Then
        'alors on est dans l'IDE
        Lang.LangFolder = LANG_PATH
    Else
        Lang.LangFolder = App.Path & "\Lang\Disassembler\"
    End If
    
    'applique la langue désirée aux controles
    Lang.Language = cPref.env_Lang
    Lang.LoadControlsCaption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmDisAsm.mnuShowInfos.Checked = False
    Me.Hide
    Cancel = 1
End Sub

'=======================================================
'récupère le path d'un fichier et affiche les infos
'=======================================================
Public Sub GetFileInfos(ByVal sFile As String)
Dim l As Long
Dim s() As String
Dim x As Long
Dim s2 As String

    On Error Resume Next
    
    'récupère le fichier en mémoire
    sStr = cFile.LoadFileInString(sFile)
    
    'récupère la position de la partie "section"
    l = InStr(1, sStr, "Number of Sections :")
    
    'alors on affiche comme infos tout ce qui est avant dans le RTB
    RTB.Text = Left$(sStr, l - 152)
    
    'maintenant on remplit les sections
    'récupère la partie de fin du fichier
    s2 = Right$(sStr, Len(sStr) - (l - 3))
    'on split pour subdiviser à chaque saut de ligne
    s() = Split(s2, vbNewLine, , vbBinaryCompare)
    
    'structure :
        '0
        '1Number of Sections :         4
        '2
        '3Name: .Text
        'Characteristics :            1610612768
        'PointerToRawData :          00000400H
        'SizeOfRawData :              373760  byte(s)
        'VirtualAddress :            01001000H
        'VirtualSize :                373538  byte(s)
        'PointerToLinenumbers :      00000000H
        'NumberOfLinenumbers :        0
        'PointerToRelocations :      00000000H
        'NumberOfRelocations :        0
        ';------------------------------------------------------------------
        
    With LV
        .ListItems.Clear
        For x = 3 To UBound(s()) Step 11
        
            l = x - 3 + 11
            
            '/!\ IMPORTANT : DO NOT REMOVE !
            If x + 9 > UBound(s()) Then Exit Sub
            
            'Name
            .ListItems.Add Text:=Right$(s(x), Len(s(x)) - 28)
            'Characteristics
            .ListItems.Item(.ListItems.Count).SubItems(1) = Right$(s(x + 1), Len(s(x + 1)) - 28)
            'PointerToRawData
            .ListItems.Item(.ListItems.Count).SubItems(2) = Right$(s(x + 2), Len(s(x + 2)) - 28)
            'SizeOfRawData
            .ListItems.Item(.ListItems.Count).SubItems(3) = Right$(s(x + 3), Len(s(x + 3)) - 28)
            'VirtualAddress
            .ListItems.Item(.ListItems.Count).SubItems(4) = Right$(s(x + 4), Len(s(x + 4)) - 28)
            'VirtualSize
            .ListItems.Item(.ListItems.Count).SubItems(5) = Right$(s(x + 5), Len(s(x + 5)) - 28)
            'PointerToLinenumbers
            .ListItems.Item(.ListItems.Count).SubItems(6) = Right$(s(x + 6), Len(s(x + 6)) - 28)
            'NumberOfLinenumbers
            .ListItems.Item(.ListItems.Count).SubItems(7) = Right$(s(x + 7), Len(s(x + 7)) - 28)
            'PointerToRelocations
            .ListItems.Item(.ListItems.Count).SubItems(8) = Right$(s(x + 8), Len(s(x + 8)) - 28)
            'NumberOfRelocations
            .ListItems.Item(.ListItems.Count).SubItems(9) = Right$(s(x + 9), Len(s(x + 9)) - 28)
        Next x
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With RTB
        .Left = 0
        .Top = 0
        .Width = Me.Width - 290
        .Height = Me.Height - 2340
    End With
    With Label1
        .Left = 0
        .Top = RTB.Height
        .Width = Me.Width
        .Height = 255
    End With
    With LV
        .Left = 0
        .Top = Label1.Top + 255
        .Width = Me.Width - 290
        .Height = 1500
    End With

End Sub
