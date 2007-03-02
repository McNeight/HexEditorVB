VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmShredd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Effacement définitif de fichiers"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShredd.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1763
      TabIndex        =   2
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Supprimer définitivement"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2543
      TabIndex        =   1
      ToolTipText     =   "Détruit les fichiers (/!\ suppression IRRECUPERABLE)"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Ajouter des fichiers..."
      Height          =   375
      Left            =   143
      TabIndex        =   0
      ToolTipText     =   "Permet l'ajout de fichiers à détruire"
      Top             =   3600
      Width           =   2175
   End
   Begin ComctlLib.ListView LV 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fichier"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "frmShredd"
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
'FORM POUR LA SUPPRESSION DEFINITIVE DE FICHIER
'=======================================================

Private Sub cmdAddFile_Click()
'ajoute un fichier à la liste à supprimer
Dim s() As String
Dim s2 As String
Dim X As Long

    s2 = cFile.ShowOpen("Choix des fichiers à supprimer", Me.hWnd, "Tous|*.*", , , , , _
        OFN_EXPLORER + OFN_ALLOWMULTISELECT, 4096, s())
    
    For X = 1 To UBound(s())
        If cFile.FileExists(s(X)) Then
            LV.ListItems.Add Text:=s(X) 'ajoute l'élément
        End If
    Next X
    
    'dans le cas d'un fichier simple
    If cFile.FileExists(s2) Then LV.ListItems.Add Text:=s2
    
    CheckBtn    'enable ou non le cmdProceed

ErrCancel:
End Sub

Private Sub cmdProceed_Click()
'procède à la suppression définitive
Dim X As Long

    'affiche un advertissement
    X = MsgBox("Les fichiers sélectionnés seront IRRECUPERABLES." & vbNewLine & "Procéder à la suppression ?", vbYesNo + vbInformation, "Attention")
    
    If Not (X = vbYes) Then Exit Sub
    
    
    For X = LV.ListItems.Count To 1 Step -1
        DoEvents    'rend quand même la main, si bcp de fichiers, c'est utile
        If ShreddFile(LV.ListItems.Item(X)) Then    'procède à la suppression
            LV.ListItems.Remove (X) 'enlève l'item si la suppression à échoué
        End If
    Next
    
    'affichage des résultats
    If LV.ListItems.Count > 0 Then
        'alors il reste au moins un fichier
        MsgBox "Au moins un des fichiers n'a pas pu être supprimé.", vbInformation, "Attention"
    Else
        'OK
        MsgBox "Fichiers supprimés avec succès", vbOKOnly, "Suppression réussie"
    End If

End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.SelectedItem Is Nothing Then Exit Sub
        
    If KeyCode = vbKeyDelete Then
        'alors enleve le fichiers sélectionnés
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
    CheckBtn    'enable ou non le cmdProceed

End Sub

'=======================================================
'vérifie que le bouton de suppression est enabled ou pas
'=======================================================
Private Sub CheckBtn()
    Me.cmdProceed.Enabled = (LV.ListItems.Count > 0)
End Sub

Private Sub LV_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim i As Long

    'gestion de la dépose des fichiers sur le listview

    On Error GoTo BadFormat 'pas de drag and drop de fichier
    
    'ajoute les fichers du drag and drop à la liste
    For i = 1 To Data.Files.Count
        If cFile.FileExists(Data.Files.Item(i)) Then LV.ListItems.Add Text:=Data.Files.Item(i)
    Next i
    
BadFormat:
End Sub
