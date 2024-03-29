VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrmSubClass"
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
' =======================================================


Option Explicit


'=======================================================
'//MODULE DE GESTION DU SUBCLASSING DE FORM (frmContent)
'=======================================================

Private AddrWndProc As Long   'adresse de la routine standart de traitement des events
Private sCaption As String
Private bAlsoForPopup As Boolean
Private cEvent As IOverMenuEvent

'=======================================================
'r�cup�ration de l'adresse de la routine standart
'=======================================================
Public Property Get AddressWndProc() As Long
    AddressWndProc = AddrWndProc
End Property

'=======================================================
'r�cup�ration de la caption du menu survol�
'=======================================================
Public Property Get OverCaption() As String
    OverCaption = sCaption
End Property

'=======================================================
'fonction qui active le hook de la form
'=======================================================
Public Function HookFormMenu(ByRef voForm As Object, Optional ByVal AlsoForPopup As Boolean = True) As Long
    
    bAlsoForPopup = AlsoForPopup
    
    'instancie la classe de l'event
    Set cEvent = voForm
    
    'r�cup�re l'adresse de la routine standart
    AddrWndProc = SetWindowLong(voForm.hWnd, GWL_WNDPROC, AddressOf MaWndProc)
    
    HookFormMenu = AddrWndProc
End Function

'=======================================================
'd�sactive le hook de la form
'=======================================================
Public Function UnHookFormMenu(ByVal hWnd As Long) As Long
    If AddrWndProc Then
        UnHookFormMenu = SetWindowLong(hWnd, GWL_WNDPROC, AddrWndProc)  'redonne l'adresse de la routine standart
        AddrWndProc = 0
    End If
End Function

'=======================================================
'renverra l'event qui contient le nom du menu surlign�
'=======================================================
Public Sub OnMenuSelect(ByVal hWnd As Long)
Dim hMenu As Long
Dim hSubMenu As Long
Dim sStr As String
Dim X As Long
Dim mnuCount As Long
Dim e As Long
Dim oForm As Form
Dim oControl As Control
    
    If bAlsoForPopup Then e = 0 Else e = MF_BYPOSITION
    
    sStr = vbNullString
    sCaption = vbNullString

    'r�cup�re le handle du menu
    hMenu = GetMenu(hWnd)

    'si il y a effectivement un menu
    If hMenu <> 0 Then
        
        'r�cup�re le nombre d'item du menu
        mnuCount = GetMenuItemCount(hMenu)

        'r�cup�re le menu concern�
        For X = 0 To mnuCount - 1
 
            If GetMenuState(hMenu, X, e) And MF_HILITE Then
                'alors s�lectionn�
                
                'alors on r�cup�re le sous menu
                hSubMenu = GetSubMenu(hMenu, X)

                'r�cup�re l'�l�ment s�lectionn� du sous menu
                Call CheckSub(hSubMenu)
                sStr = Me.OverCaption
                
            End If
        Next X
    End If
    
    'on lib�re l'event
    Call cEvent.OnMenuOver(sStr)
End Sub

'=======================================================
'r�cup�re l'item s�lectionn� du sous menu
'=======================================================
Private Sub CheckSub(hSubMenu As Long)
Dim mnuCount As Long
Dim X As Long
Dim lRet As Long
Dim hSubSubMenu As Long
Dim sBuf As String
Dim e As Long

    'r�cup�re le nombre d'items
    mnuCount = GetMenuItemCount(hSubMenu)

    'pour chaque item
    For X = 0 To mnuCount - 1
            
        If GetMenuState(hSubMenu, X, MF_BYPOSITION) And MF_HILITE Then
            
            'on r�cup�re le sous menu
            hSubSubMenu = GetSubMenu(hSubMenu, X)

            'on continue r�cursivement dans les sous menus jusqu'� trouver un item selectionn� sans sous menu
            If hSubSubMenu And IsAMenuSelected(hSubSubMenu) Then
            
                'alors il y a un sub menu ==> on continue
                Call CheckSub(hSubSubMenu)
                
            Else
                'ok, plus de submenu
                sBuf = Space$(255)    'un buffer
                
                'r�cup�re le caption
                lRet = GetMenuString(hSubMenu, X, sBuf, Len(sBuf), MF_BYPOSITION)

                'formate la string
                sBuf = Left$(sBuf, lRet)
                
                sCaption = sBuf
                Exit Sub
            End If
        End If
    Next X

End Sub

'=======================================================
'v�rifie si on a un �l�ment s�lectionn�
'=======================================================
Private Function IsAMenuSelected(hSubSubMenu As Long) As Long
Dim X As Long
Dim mnuCount As Long

    IsAMenuSelected = 0
    
    'nombre d'items
    mnuCount = GetMenuItemCount(hSubSubMenu)

    'pour chaque menu
    For X = 0 To mnuCount - 1
        'on checke si le menu est selected
        If GetMenuState(hSubSubMenu, X, MF_BYPOSITION) And MF_HILITE Then
            IsAMenuSelected = 1   'oui, il l'est
            Exit Function
        End If
    Next X

    'si on est l�, pas d'item s�lectionn� (normalement car pas de sous items)
End Function

'=======================================================
'on lib�re la m�moire occup�e par la classe de l'Event
'=======================================================
Private Sub Class_Terminate()
    Set cEvent = Nothing
End Sub
