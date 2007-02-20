Attribute VB_Name = "mdlSubClassFrmContent"
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
'//MODULE DE GESTION DU SUBCLASSING DE FORM (frmContent)
'=======================================================

Private AddrWndProc As Long 'adresse de la routine standart

'=======================================================
'fonction qui active le hook de la form
'=======================================================
Public Function HookFormMenu(ByVal hWnd As Long) As Long
    AddrWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MaWndProc)
    HookFormMenu = AddrWndProc
End Function

'=======================================================
'routine de remplacement pour l'interception des messages
'=======================================================
Private Function MaWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    Debug.Print uMsg & "  " & wParam & "  " & lParam
    
    Select Case uMsg
 
        Case 287
            'alors on a capté un changement de menu
            
            DisplayAssociatedMenu wParam    'change le text de la statusbar de frmContent
                        
        Case Else
           MaWndProc = CallWindowProc(AddrWndProc, hWnd, uMsg, wParam, lParam) 'Appel de la routine standard pour les autres messages
    End Select
End Function

'=======================================================
'désactive le hook
'=======================================================
Public Function UnHookFormMenu(ByVal hWnd As Long) As Long
    UnHookFormMenu = SetWindowLong(hWnd, GWL_WNDPROC, AddrWndProc)
End Function

'=======================================================
'affiche le menu correspondant au paramètre
'=======================================================
Private Sub DisplayAssociatedMenu(ByVal wParam As Long)
Dim s As String
    
    Select Case wParam
        Case -2139095037
            s = "1"
        Case -2139095036
            s = "2"
        Case Else
            s = "Status=[Ready]"
    End Select
    
    frmContent.Caption = wParam
    
    
    frmContent.Sb.Panels(1).Text = s
End Sub


