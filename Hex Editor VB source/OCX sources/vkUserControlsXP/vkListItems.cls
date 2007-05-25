VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vkListItems"
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

Private colItems As Collection


'=======================================================
'Permet d'utiliser For Each
'=======================================================
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    'on renvoie l'objet d'énumération

    Set NewEnum = colItems.[_NewEnum]

End Function
 
'Renvoie l'élément d'index iCurrent dans la variable var
'elle doit renvoyer 0 si tout vabien
'1 s'il n'y a plus d'éléments dans la collection
Public Function ForEach(ByVal iCurrent As Long, var As Variant) As Long
'

End Function



'=======================================================
'Properiétés de la classe
'=======================================================
Public Property Get Count() As Long: Count = colItems.Count: End Property
Public Property Get Item(Index As Long) As vkListItem:
    If Index <= colItems.Count And Index > 0 Then Set Item = colItems(Index)
End Property


'=======================================================
'Subs d'initialisations
'=======================================================
Private Sub Class_Initialize(): Set colItems = New Collection: End Sub
Private Sub Class_Terminate(): Set colItems = Nothing: End Sub


'=======================================================
'Ajout d'un objet dans la collection
'=======================================================
Public Sub AddItem(ByRef Item As vkListItem, Optional Key As Variant, _
    Optional Before As Variant, Optional After As Variant)
    
    Call colItems.Add(Item, Key, Before, After)
End Sub

'=======================================================
'Suppression d'un objet de la collection
'=======================================================
Public Sub RemoveItem(Index As Long)
    Call colItems.Remove(Index)
End Sub

'=======================================================
'Suppression de tous les ojbets
'=======================================================
Public Sub Clear()
Dim Obj As Object

    For Each Obj In colItems
        Call colItems.Remove(Obj.Index)
    Next Obj
    
End Sub


