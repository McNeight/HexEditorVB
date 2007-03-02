Attribute VB_Name = "mdlPictureSubClassing"
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
'//MODULE POUR SUBCLASSER LA PICTURE DE FRMCONTENT CONTENANT LE FV
'//PERMET LE RESIZE
'=======================================================

Private AddrWndProc As Long   'adresse de la routine standart de traitement des events
Private pctHwnd As Long
Private pc As PictureBox
Private IsSized As Boolean


'=======================================================
'fonction qui active le hook de la form
'=======================================================
Public Function HookPictureResizement(ByRef pct As PictureBox) As Long
Dim tET As TRACKMOUSEEVENTTYPE

    'récupère les infos sur la picturebox
    pctHwnd = pct.hWnd
    Set pc = pct
    
    'démarre le tracking de l'event MOUSE_LEAVE
    With tET    'prépare la structure
        .cbSize = Len(tET)
        .hwndTrack = pctHwnd
        .dwFlags = TME_LEAVE
    End With
    'lance le tracking
    Call TrackMouseEvent(tET)


    IsSized = False 'pas de resize
    
    'récupère l'adresse de la routine standart
    AddrWndProc = SetWindowLong(pctHwnd, GWL_WNDPROC, AddressOf ProcPictureSubClassProc)
    
    HookPictureResizement = AddrWndProc
End Function

'=======================================================
'désactive le hook de la form
'=======================================================
Public Function UnHookPictureResizement(ByVal hWnd As Long) As Long
    If AddrWndProc Then
         'redonne l'adresse de la routine standart
        UnHookPictureResizement = SetWindowLong(hWnd, GWL_WNDPROC, AddrWndProc)
        AddrWndProc = 0
    End If
End Function

'=======================================================
'routine de remplacement pour l'interception des messages ==> subclassing
'=======================================================
Public Function ProcPictureSubClassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Dim cur As POINTAPI 'en pixels, donc *15 pour remettre dans la bonne unité
Dim i As Long
Dim rec As RECT
Dim tET As TRACKMOUSEEVENTTYPE

    On Error Resume Next    'évite les erreurs de resizing
    
    Select Case uMsg
        
        Case WM_LBUTTONDOWN

            'récupère la position du curseur
            Call GetCursorPos(cur)
            
            'vérifie si on est dans le bas de la picturebox
            i = pc.Parent.Top + pc.Top + pc.Height - cur.y * 15
            If i < -675 Then
                'alors dans la position pour faire le drag
                IsSized = True
                
                'on change le curseur
                pc.MousePointer = 7
                
            End If

           
            ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hWnd, uMsg, wParam, lParam)
        
        Case WM_MOUSEMOVE
        
            'redémarre le tracking de l'event MOUSE_LEAVE
            With tET    'prépare la structure
                .cbSize = Len(tET)
                .hwndTrack = pctHwnd
                .dwFlags = TME_LEAVE
            End With
            'relance le tracking
            Call TrackMouseEvent(tET)
            
            If IsSized Then
                'alors drag ==> on change la taille
                
                'récupère la position du curseur
                Call GetCursorPos(cur)

                'récupère la taille à affecter au picturebox
                pc.Height = cur.y * 15 - pc.Top - pc.Parent.Top - 710
            
            Else
                'alors pas drag, on checke juste si on est en position d'afficher
                'le nouveau curseur ou pas
                
                'récupère la position du curseur
                Call GetCursorPos(cur)
            
                'vérifie si on est dans le bas de la picturebox
                i = pc.Parent.Top + pc.Top + pc.Height - cur.y * 15
                If i < -675 Then pc.MousePointer = 7 Else pc.MousePointer = 0
                
            End If
        
            ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hWnd, uMsg, wParam, lParam)
        
        Case WM_LBUTTONUP
        
            IsSized = False 'plus de drag
                
            'remet le curseur normal
            pc.MousePointer = 0
            
            ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hWnd, uMsg, wParam, lParam)
        
        Case WM_MOUSELEAVE
            
            If IsSized = False Then
                'alors on remet le curseur normal car on quitte le composant sans resize
                pc.MousePointer = 0
            End If
        
        Case Else
           'appel de la routine standard pour les autres messages
           ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hWnd, uMsg, wParam, lParam)
    End Select
    
End Function
