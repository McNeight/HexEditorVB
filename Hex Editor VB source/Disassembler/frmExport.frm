VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmExport 
   Caption         =   "Exporter"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6975
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   6975
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstDll 
      Height          =   2010
      ItemData        =   "frmExport.frx":08CA
      Left            =   0
      List            =   "frmExport.frx":08CC
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin ComctlLib.ListView LV 
      Height          =   2055
      Left            =   3360
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nom"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Hint"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Adresse"
         Object.Width           =   3528
      EndProperty
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin LanguageTranslator.ctrlLanguage ctrlLanguage1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
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
'AFFICHAGE DES EXPORTS
'=======================================================

Private nbDll As Long
Private sD() As String
Private sF() As String

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
    
    nbDll = 0
    ReDim sF(0)
    ReDim sD(0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmDisAsm.mnuShowExports.Checked = False
    Me.Hide
    Cancel = 1
End Sub

Private Sub Form_Resize()
    With lstDll
        .Left = 0
        .Top = 0
        .Width = 2000
        .Height = Me.Height
    End With
    With txt
        .Top = 0
        .Left = 2020
        .Height = 1500
        .Width = Me.Width - 2300
    End With
    With LV
        .Top = 1505
        .Left = 2020
        .Height = Me.Height - 2055
        .Width = Me.Width - 2270
    End With
End Sub

'=======================================================
'récupère le path d'un fichier et affiche les infos
'=======================================================
Public Sub GetFileInfosExp(ByVal sFile As String)
Dim l As Long
Dim s() As String
Dim x As Long
Dim s2 As String
Dim dep As Long
Dim sStr As String
Dim sDll As String

    On Error Resume Next
    
    'récupère le fichier en mémoire
    sStr = cFile.LoadFileInString(sFile)
    
    'pour chaque dll
    dep = 1
    lstDll.Clear
    While InStr(dep, sStr, ";Export From ", vbBinaryCompare)
    
        nbDll = nbDll + 1
        
        '//récupère le nom de la dll
        'récupère la position de la partie "Import from"
        l = InStr(dep, sStr, ";Export From ")
        x = InStr(l + 1, sStr, vbNewLine, vbBinaryCompare)
        sDll = Mid$(sStr, l + 13, x - l - 13)
        lstDll.AddItem sDll 'ajoute la dll à la liste
        
        '//on récupère maintenant le texte pour chaque dll
        ReDim Preserve sD(UBound(sD()) + 1)
        l = InStr(x + 20, sStr, vbNewLine)
        x = InStr(x + 220, sStr, vbNewLine, vbBinaryCompare)
        sD(UBound(sD())) = Replace$(Mid$(sStr, l + 2, x - l - 2), ";", vbNullString)
        
        '//on récupère maintenant la string avec le nom de toutes les fonctions importées
        ReDim Preserve sF(UBound(sD))
        l = InStr(x + 60, sStr, vbNewLine)
        x = InStr(x + 60, sStr, ";=", vbBinaryCompare)
        sF(UBound(sF())) = Mid$(sStr, l + 2, x - l - 2)
        'vire tous les caractères inutiles
        sF(UBound(sF())) = Replace$(sF(UBound(sF())), ";Exported ", vbNullString, , , vbBinaryCompare)
        sF(UBound(sF())) = Replace$(sF(UBound(sF())), " (", "|", , , vbBinaryCompare)
        sF(UBound(sF())) = Replace$(sF(UBound(sF())), ") at address ", "|", , , vbBinaryCompare)
        
        dep = l + 1
    Wend
    
    'on affiche les infos sur la première dll
    lstDll.ListIndex = 0
    Call DllDisplay
    
End Sub

'=======================================================
'affiche les infos sur la dll sélectionnée
'=======================================================
Private Sub DllDisplay()
Dim l As Long
Dim x As Long
Dim l2 As Long
Dim s As String
Dim s2() As String

    On Error Resume Next
    
    'affichage du texte dans le txt
    txt.Text = sD(lstDll.ListIndex + 1)
    
    'on affiche maintenant les infos sur les fonctions associées à la DLL dans le LV
    'texte de la forme
    
    'EnumChildWindows by Name|0x0||0050D65CH
    'EnumChildWindows by Name|0x0|0050D65CH
    'EnumChildWindows by Name|0x0|0050D65CH
    
    'tant qu'on voit des saut de lignes
    LV.Visible = False
    With LV.ListItems
    
        .Clear
        
        'sépare chaque ligne
        s2() = Split(sF(lstDll.ListIndex + 1), vbNewLine, , vbBinaryCompare)
        
        For x = 0 To UBound(s2()) - 1
                   
            'récupère les positions des "|" pour en extraire chaque function, son hint
            'et son adresse
            
            s = s2(x) & "|"
            
            l = InStr(1, s, "|", vbBinaryCompare)
            .Add Text:=Mid$(s, 1, l - 1)
            
            l2 = InStr(l + 1, s, "|", vbBinaryCompare)
            .Item(.Count).SubItems(1) = Mid$(s, l + 1, l2 - l - 1)
            
            l = InStr(l2 + 1, s, "|", vbBinaryCompare)
            .Item(.Count).SubItems(2) = Mid$(s, l2 + 1, l - l2 - 1)
            
        Next x
    End With
    LV.Visible = True
    
End Sub

Private Sub lstDll_Click()
    Call DllDisplay
End Sub

