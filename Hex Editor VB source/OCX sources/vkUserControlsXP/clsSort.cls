VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'PUBLIC FUNCTIONS
'=======================================================

Public Function SortList(ByRef Col As clsFastCollection, Optional _
    ByVal Asc As Boolean = True)
Attribute SortList.VB_Description = "Sort a list of vkListItem"

Dim s() As vkListItem
Dim x As Long
    
    'on prépare une liste de vkListItems
    ReDim s(Col.Count - 1)

    For x = 1 To Col.Count
        Set s(x - 1) = Col.Item(x)
    Next x

    'on lance le tri
    If Asc Then
        Call QuickSortObj(s(), 0, UBound(s()))
    Else
        Call QuickSortDescObj(s(), 1, UBound(s()))
    End If
    
    'on récupère notre collection en la recréant
    Call Col.Clear
    For x = 1 To UBound(s()) + 1
        Call Col.Add(s(x - 1))
    Next x
    
    'vire le tableau
    Erase s
    
End Function

'=======================================================
'fonction récursive de tri ordre alphabétique
'=======================================================
'Private Sub QuickSort(c() As String, ByVal First As Long, ByVal Last As Long)
'Dim Low As Long
'Dim High As Long
'Dim MidValue As String
'
'    Low = First: High = Last: MidValue = c((First + Last) \ 2)
'
'    Do
'        While c(Low) < MidValue
'            Low = Low + 1
'        Wend
'        While c(High) > MidValue
'            High = High - 1
'        Wend
'        If Low <= High Then
'            Call Swap(c(Low), c(High))
'            Low = Low + 1
'            High = High - 1
'        End If
'    Loop While Low <= High
'
'    If First < High Then Call QuickSort(c, First, High)
'    If Low < Last Then Call QuickSort(c, Low, Last)
'End Sub
'
''=======================================================
''fonction récursive de tri inverse ordre alphabétique
''=======================================================
'Private Sub QuickSortDesc(c() As String, ByVal First As Long, ByVal Last As Long)
'Dim Low As Long
'Dim High As Long
'Dim MidValue As String
'
'    Low = First: High = Last: MidValue = c((First + Last) \ 2)
'
'    Do
'        While c(Low) > MidValue
'            Low = Low + 1
'        Wend
'        While c(High) < MidValue
'            High = High - 1
'        Wend
'        If Low <= High Then
'            Call Swap(c(Low), c(High))
'            Low = Low + 1
'            High = High - 1
'        End If
'    Loop While Low <= High
'
'    If First < High Then Call QuickSortDesc(c, First, High)
'    If Low < Last Then Call QuickSortDesc(c, Low, Last)
'End Sub
'
'
''=======================================================
''permute deux valeurs
''=======================================================
'Private Sub Swap(ByRef A As String, ByRef B As String)
'Dim T As String
'    T = A: A = B: B = T
'End Sub





'=======================================================
'fonction récursive de tri ordre alphabétique
'=======================================================
Private Sub QuickSortObj(c() As vkListItem, ByVal First As Long, ByVal Last As Long)
Dim Low As Long
Dim High As Long
Dim MidValue As String
    
    Low = First: High = Last: MidValue = c((First + Last) \ 2).Text
    
    Do
        While c(Low).Text < MidValue
            Low = Low + 1
        Wend
        While c(High).Text > MidValue
            High = High - 1
        Wend
        If Low <= High Then
            Call SwapObj(c(Low), c(High))
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then Call QuickSortObj(c, First, High)
    If Low < Last Then Call QuickSortObj(c, Low, Last)
End Sub

'=======================================================
'fonction récursive de tri inverse ordre alphabétique
'=======================================================
Private Sub QuickSortDescObj(c() As vkListItem, ByVal First As Long, ByVal Last As Long)
Dim Low As Long
Dim High As Long
Dim MidValue As String
    
    Low = First: High = Last: MidValue = c((First + Last) \ 2).Text
    
    Do
        While c(Low).Text > MidValue
            Low = Low + 1
        Wend
        While c(High).Text < MidValue
            High = High - 1
        Wend
        If Low <= High Then
            Call SwapObj(c(Low), c(High))
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    
    If First < High Then Call QuickSortDescObj(c, First, High)
    If Low < Last Then Call QuickSortDescObj(c, Low, Last)
End Sub


'=======================================================
'permute deux valeurs
'=======================================================
Private Sub SwapObj(ByRef A As vkListItem, ByRef B As vkListItem)
Dim T As vkListItem
    Set T = A: Set A = B: Set B = T
End Sub
