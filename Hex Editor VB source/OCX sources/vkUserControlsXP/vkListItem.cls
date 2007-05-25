VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vkListItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private lBackColor As Long
Private bChecked As Boolean
Private tFont As StdFont
Private lForeColor As Long
Private tIcon As Variant
Private lIndex As Long
Private sKey As String
Private bSelected As Boolean
Private sText As String
Private lHeight As Long
Private tAlig As AlignmentConstants
Private lpxlIconHeight As Long
Private lpxlIconWidth As Long
Private lSelColor As Long
Private lBorderSelColor As Long

'=======================================================
'properties
'=======================================================
Public Property Get BackColor() As Long: BackColor = lBackColor: End Property
Public Property Let BackColor(BackColor As Long): lBackColor = BackColor: End Property
Public Property Get Checked() As Boolean: Checked = bChecked: End Property
Public Property Let Checked(Checked As Boolean): bChecked = Checked: End Property
Public Property Get Font() As StdFont: Set Font = tFont: End Property
Public Property Let Font(Font As StdFont): Set tFont = Font: End Property
Public Property Get ForeColor() As Long: ForeColor = lForeColor: End Property
Public Property Let ForeColor(ForeColor As Long): lForeColor = ForeColor: End Property
Public Property Get Icon() As Variant:  Icon = tIcon: End Property
Public Property Let Icon(Icon As Variant):  tIcon = Icon: End Property
Public Property Get Index() As Long: Index = lIndex: End Property
Public Property Let Index(Index As Long): lIndex = Index: End Property
Public Property Get Key() As String: Key = sKey: End Property
Public Property Let Key(Key As String): sKey = Key: End Property
Public Property Get Selected() As Boolean: Selected = bSelected: End Property
Public Property Let Selected(Selected As Boolean): bSelected = Selected: End Property
Public Property Get Text() As String: Text = sText: End Property
Public Property Let Text(Text As String): sText = Text: End Property
Public Property Get Height() As Long: Height = lHeight: End Property
Public Property Let Height(Height As Long)
lHeight = Height
End Property
Public Property Get Alignment() As AlignmentConstants: Alignment = tAlig: End Property
Public Property Let Alignment(Alignment As AlignmentConstants): tAlig = Alignment: End Property
Public Property Get pxlIconHeight() As Long: pxlIconHeight = lpxlIconHeight: End Property
Public Property Let pxlIconHeight(pxlIconHeight As Long): lpxlIconHeight = pxlIconHeight: End Property
Public Property Get pxlIconWidth() As Long: pxlIconWidth = lpxlIconWidth: End Property
Public Property Let pxlIconWidth(pxlIconWidth As Long): lpxlIconWidth = pxlIconWidth: End Property
Public Property Get SelColor() As Long: SelColor = lSelColor: End Property
Public Property Let SelColor(SelColor As Long): lSelColor = SelColor: End Property
Public Property Get BorderSelColor() As Long: BorderSelColor = lBorderSelColor: End Property
Public Property Let BorderSelColor(BorderSelColor As Long): lBorderSelColor = BorderSelColor: End Property

Private Sub Class_Initialize()
'on affecte les valeurs par défaut
    With Me
        .BackColor = &HFFFFFF
        .ForeColor = vbBlack
        .Alignment = vbLeftJustify
        .SelColor = 16768444
        '.Font =
        .BorderSelColor = 16419097
        .pxlIconHeight = 16
        .pxlIconWidth = 16
        .Height = 255
    End With
End Sub
