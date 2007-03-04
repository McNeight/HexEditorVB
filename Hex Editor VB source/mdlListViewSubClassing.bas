Attribute VB_Name = "mdlListViewSubClassing"
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
'//MODULE POUR SUBCLASSER LES LISTVIEW LVICON
'=======================================================

Private AddrWndProc As Long   'adresse de la routine standart de traitement des events
Private lvHwnd As Long


'=======================================================
'fonction qui active le hook de la form
'=======================================================
Public Function HookLVDragAndDrop(ByRef hWnd As Long) As Long
Dim tET As TRACKMOUSEEVENTTYPE

    'récupère les infos sur la picturebox
    lvHwnd = hWnd
    
    'récupère l'adresse de la routine standart
    AddrWndProc = SetWindowLong(lvHwnd, GWL_WNDPROC, AddressOf ProcLVSubClassProc)
    
    HookLVDragAndDrop = AddrWndProc
End Function

'=======================================================
'désactive le hook de la form
'=======================================================
Public Function UnHookLVDragAndDrop(ByVal hWnd As Long) As Long
    If AddrWndProc Then
         'redonne l'adresse de la routine standart
        UnHookLVDragAndDrop = SetWindowLong(hWnd, GWL_WNDPROC, AddrWndProc)
        AddrWndProc = 0
    End If
End Function

'=======================================================
'routine de remplacement pour l'interception des messages ==> subclassing
'=======================================================
Public Function ProcLVSubClassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
   Debug.Print uMsg & "  " & wParam & "  " & lParam
    Select Case uMsg
        
        Case 4111, 4137, 4110
            'alors on ne fait rien
            'ne rien faire, et donc ne pas appeler la routine standart, empêche de pouvoir
            'déplacer les icones présentes dans le listview

        Case Else
           'appel de la routine standard pour les autres messages
           ProcLVSubClassProc = CallWindowProc(AddrWndProc, hWnd, uMsg, wParam, lParam)
    End Select
    
End Function


