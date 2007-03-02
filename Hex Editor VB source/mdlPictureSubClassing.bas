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



Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long



Private AddrWndProc As Long   'adresse de la routine standart de traitement des events
Private pctHwnd As Long
Private pc As PictureBox
Private IsSized As Boolean


'=======================================================
'fonction qui active le hook de la form
'=======================================================
Public Function HookPictureResizement(ByRef pct As PictureBox) As Long
    
    'récupère les infos sur la picturebox
    pctHwnd = pct.hwnd
    Set pc = pct
    
    IsSized = False 'pas de resize
    
    'récupère l'adresse de la routine standart
    AddrWndProc = SetWindowLong(pctHwnd, GWL_WNDPROC, AddressOf ProcPictureSubClassProc)
    
    HookPictureResizement = AddrWndProc
End Function

'=======================================================
'désactive le hook de la form
'=======================================================
Public Function UnHookPictureResizement(ByVal hwnd As Long) As Long
    If AddrWndProc Then
         'redonne l'adresse de la routine standart
        UnHookPictureResizement = SetWindowLong(hwnd, GWL_WNDPROC, AddrWndProc)
        AddrWndProc = 0
    End If
End Function

'=======================================================
'routine de remplacement pour l'interception des messages ==> subclassing
'=======================================================
Public Function ProcPictureSubClassProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Dim cur As POINTAPI
Dim i As Long

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

           
            ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hwnd, uMsg, wParam, lParam)
        
        Case WM_MOUSEMOVE
        
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
        
            ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hwnd, uMsg, wParam, lParam)
        
        Case WM_LBUTTONUP
        
            IsSized = False 'plus de drag
            
            'remet le curseur normal
            pc.MousePointer = 0
            
            ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hwnd, uMsg, wParam, lParam)
        
        Case Else
           'appel de la routine standard pour les autres messages
           ProcPictureSubClassProc = CallWindowProc(AddrWndProc, hwnd, uMsg, wParam, lParam)
    End Select
    
End Function
