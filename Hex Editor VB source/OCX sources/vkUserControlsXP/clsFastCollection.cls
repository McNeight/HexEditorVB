VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFastCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'APIS
'=======================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'=======================================================
'variables privées
'=======================================================
Private lCount As Long
Private lPtr() As Long

Private Const USTEP As Long = 1000


'=======================================================
'PUBLIC FUNCTIONS & PROPERTIES
'=======================================================

'=======================================================
'ajoute un item
'=======================================================
Public Sub Add(Obj As Object, Optional ByVal Index As Long = -1)
Dim n As Long
Dim tmp As Object
    
    Set tmp = Obj
    
    If Index = -1 Then
        'alors on ajoute simplement à la suite
        
        lCount = lCount + 1
        
        'redimensionne le tableau, pas à chaque fois, mais quand
        'c'est nécessaire (tous les 1000)
        If UBound(lPtr()) < lCount Then
            ReDim Preserve lPtr(UBound(lPtr()) + USTEP)
        End If
        
        'recupère le pointeur sur l'objet
        lPtr(lCount) = ObjPtr(Obj)
        
        Call CopyMemory(ByVal VarPtr(Obj), 0, 4)
        Set Obj = tmp

    ElseIf Index > 0 And Index <= lCount Then
    
        'alors on insère dans la liste
        Call AddItemPr(ObjPtr(Obj), Index)
        
        Call CopyMemory(ByVal VarPtr(Obj), 0, 4)
        Set Obj = tmp
                
    End If
    
End Sub

'=======================================================
'enlève un item
'=======================================================
Public Sub Remove(Index As Long)
    If Index > 0 And Index < lCount Then
        Call RemoveItemPr(Index)
        lCount = lCount - 1
    End If
End Sub

'=======================================================
'clear la collection
'=======================================================
Public Sub Clear()
Dim x As Long

    'supprime tous les objets
    For x = 1 To lCount
        Call DestroyObj(lPtr(x))
    Next x
    
    'supprime le tableau
    ReDim lPtr(0)
    lCount = 0
 
End Sub

'=======================================================
'récupère l'item pointé par l'index
'=======================================================
Public Function Item(Index As Long) As Object
Attribute Item.VB_UserMemId = 0
    
    If Index > 0 And Index <= lCount Then _
        Set Item = ObjFromPtr(lPtr(Index))
        
End Function

'=======================================================
'récupère le Count de la collection
'=======================================================
Public Property Get Count() As Long: Count = lCount: End Property





'=======================================================
'PRIVATE SUBS
'=======================================================

'=======================================================
'récupère un objet depuis un pointeur
'=======================================================
Private Function ObjFromPtr(ByVal lngPtr As Long) As Object
Dim Obj As Object

   Call CopyMemory(Obj, lngPtr, 4)      '4 octets
   Set ObjFromPtr = Obj     'récupère l'objet en résultat
   Call CopyMemory(Obj, 0&, 4)      'détruit l'objet temporaire

End Function

'=======================================================
'libère un objet depuis un pointeur
'=======================================================
Private Sub DestroyObj(ByVal lngPtr As Long)
Dim Obj As Object
    Call CopyMemory(Obj, lngPtr, 4)
    Set Obj = Nothing
End Sub

'=======================================================
'ajoute un item dans la liste
'=======================================================
Private Sub AddItemPr(ByVal Ptr As Long, ByVal Index As Long)
Dim tmp As Long

    'récupère le dernier item
    tmp = lPtr(lCount)
    
    lCount = lCount + 1
    
    'redimensionne le tableau, pas à chaque fois, mais quand
    'c'est nécessaire (tous les 1000)
    If UBound(lPtr()) < lCount Then
        ReDim Preserve lPtr(UBound(lPtr()) + USTEP)
    End If

    'on copie la partie APRES index en décalant d'une case
    Call CopyMemory(lPtr(Index + 1), lPtr(Index), (lCount - Index) * 4)

    'on insère le nouvel item
    lPtr(Index) = Ptr
    
    lPtr(lCount) = tmp

 End Sub
 
'=======================================================
'supprime un item de la liste
'=======================================================
Private Sub RemoveItemPr(ByVal Index As Long)

    'libère l'objet correspondant à l'index à virer
    Call DestroyObj(lPtr(Index))

    'on décale de 1 vers la gauche le tableau
    Call CopyMemory(lPtr(Index), lPtr(Index + 1), (lCount - Index) * 4)

    'redimensionne de 1 en moins le tableau
    ReDim Preserve lPtr(UBound(lPtr()) - 1)
    
End Sub




'=======================================================
'SUB DE LA CLASSE
'=======================================================

'=======================================================
'initialisation et destruction
'=======================================================
Private Sub Class_Initialize()
    lCount = 0
    ReDim lPtr(0)
End Sub
Private Sub Class_Terminate()
    Call Clear
End Sub
