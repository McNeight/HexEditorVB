VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "     S�lectionner une zone"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1485
      TabIndex        =   5
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "S�lectionner"
      Height          =   375
      Left            =   165
      TabIndex        =   4
      ToolTipText     =   "Proc�der � la restriction"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtTo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Offset sup�rieur"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtFrom 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Offset inf�rieur"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "jusqu'au byte"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "A partir du byte"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelect"
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'FORM POUR SELECTIONNER UNE ZONE PARTICULIERE
'-------------------------------------------------------

Private byteFunc As Byte

Private Sub cmdOk_Click()
'valide
Dim lFrom As Long
Dim lTo As Long
Dim x As Long

    On Error GoTo ErrGestion
    
    'r�cup�re les valeurs num�riques
    lFrom = FormatedVal(txtFrom.Text) + 1
    lTo = FormatedVal(txtTo.Text)
    
    'fait en sorte que lFrom soit le plus petit
    If lFrom > lTo Then
        x = lFrom
        lFrom = lTo
        lTo = x
    End If
        
    If byteFunc = 0 Then    'il s'agit d'une s�lection param�tr�e

        'v�rifie que la plage est OK
        If lFrom < frmContent.ActiveForm.HW.FirstOffset Or lTo > frmContent.ActiveForm.HW.MaxOffset Then
            Unload Me
            Exit Sub
        End If
        If lFrom > frmContent.ActiveForm.HW.MaxOffset Or lTo < frmContent.ActiveForm.HW.FirstOffset Then
            Unload Me
            Exit Sub
        End If
        
        'fait la s�lection d�sir�e
        frmContent.ActiveForm.HW.SelectZone 16 - (By16(lFrom) - lFrom), By16(lFrom) - 16, 17 - (By16(lTo) - lTo), By16(lTo) - 16
        
        'refresh le label qui contient la taille de la s�lection
        frmContent.ActiveForm.Sb.Panels(4).Text = "S�lection=[" & CStr(frmContent.ActiveForm.HW.NumberOfSelectedItems) & " bytes]"
        frmContent.ActiveForm.Label2(9) = frmContent.ActiveForm.Sb.Panels(4).Text
        
    ElseIf byteFunc = 1 Then 'il s'agit d'une restriction d'affichage
    
        'formate les valeurs en terme d'offset
        lFrom = By16(lFrom)
        
        'v�rifie que la plage est OK
        '/!\ v�rifie la plage � l'aide des TAGS CURRENCY du HW (valeurs maximales autoris�es)
        If lFrom < frmContent.ActiveForm.HW.curTag1 Or lTo > frmContent.ActiveForm.HW.curTag2 Then
            Unload Me
            Exit Sub
        End If
        If lFrom > frmContent.ActiveForm.HW.curTag2 Or lTo < frmContent.ActiveForm.HW.curTag1 Then
            Unload Me
            Exit Sub
        End If
        
        'ajoute une entr�e � l'historique
        frmContent.ActiveForm.AddHistoFrm actRestArea, , , frmContent.ActiveForm.HW.FirstOffset, frmContent.ActiveForm.HW.MaxOffset
        
        'change les valeurs dans la ActiveForm
        frmContent.ActiveForm.HW.FirstOffset = lFrom
        frmContent.ActiveForm.HW.MaxOffset = lTo
        frmContent.ActiveForm.VS.Min = lFrom / 16
        frmContent.ActiveForm.VS.Max = By16(lTo / 16)
        Call frmContent.ActiveForm.VS_Change(frmContent.ActiveForm.VS.Min)

    End If
    
    Unload Me
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmSelect.cmdOkClick", True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    If frmContent.ActiveForm Is Nothing Then Unload Me
    
    'affiche l'�l�ment actuellement s�lectionn� dans l'activeform
    txtFrom.Text = CStr(frmContent.ActiveForm.HW.Item.Offset + frmContent.ActiveForm.HW.Item.Col) - 1
End Sub

'-------------------------------------------------------
'sub permettant de r�cup�rer un nombre qui va sp�cifier
'a quoi la s�lection servira
'0 = s�lection param�tr�e
'1 = affichage restreint
'-------------------------------------------------------
Public Sub GetEditFunction(ByVal btFunction As Byte)
    byteFunc = btFunction
End Sub
