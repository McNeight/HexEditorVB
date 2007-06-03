Attribute VB_Name = "mdlDeclarations"
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
'CONSTANTES
'=======================================================
Public Const MK_LBUTTON                     As Long = &H1&
Public Const MK_RBUTTON                     As Long = &H2&
Public Const MK_SHIFT                       As Long = &H4&
Public Const MK_CONTROL                     As Long = &H8&
Public Const MK_MBUTTON                     As Long = &H10&
Public Const WM_KEYDOWN                     As Long = &H100
Public Const WM_KEYFIRST                    As Long = &H100
Public Const WM_KEYLAST                     As Long = &H108
Public Const WM_KEYUP                       As Long = &H101
Public Const WM_LBUTTONDBLCLK               As Long = &H203
Public Const WM_LBUTTONDOWN                 As Long = &H201
Public Const WM_LBUTTONUP                   As Long = &H202
Public Const WM_MBUTTONDBLCLK               As Long = &H209
Public Const WM_MBUTTONDOWN                 As Long = &H207
Public Const WM_MBUTTONUP                   As Long = &H208
Public Const WM_MOUSEFIRST                  As Long = &H200
Public Const WM_MOUSELAST                   As Long = &H209
Public Const WM_MOUSEHOVER                  As Long = &H2A1
Public Const WM_MOUSELEAVE                  As Long = &H2A3
Public Const WM_MOUSEMOVE                   As Long = &H200
Public Const WM_RBUTTONDBLCLK               As Long = &H206
Public Const WM_RBUTTONDOWN                 As Long = &H204
Public Const WM_RBUTTONUP                   As Long = &H205
Public Const WM_MOUSEWHEEL                  As Long = &H20A
Public Const WM_PAINT                       As Long = &HF
Public Const GWL_WNDPROC                    As Long = -4&
Public Const TME_LEAVE                      As Long = &H2&
Public Const TME_HOVER                      As Long = &H1&
Public Const DT_CENTER                      As Long = &H1&
Public Const DT_LEFT                        As Long = &H0&
Public Const DT_RIGHT                       As Long = &H2&
Public Const DI_MASK                        As Long = &H1
Public Const DI_IMAGE                       As Long = &H2
Public Const DI_NORMAL                      As Long = DI_MASK Or DI_IMAGE
Public Const SRCCOPY                        As Long = 13369376
Public Const TIME_ONESHOT                   As Long = 0
Public Const TIME_PERIODIC                  As Long = 1
Public Const TIME_CALLBACK_EVENT_PULSE      As Long = &H20
Public Const TIME_CALLBACK_EVENT_SET        As Long = &H10
Public Const TIME_CALLBACK_FUNCTION         As Long = &H0




'=======================================================
'APIs
'=======================================================
Public Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Public Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hDc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Attribute BitBlt.VB_MemberFlags = "40"
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDc As Long, lpRect As RECT) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, Lppoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function FrameRgn Lib "gdi32" (ByVal hDc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hDc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long



'=======================================================
'TYPES
'=======================================================
Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type RGB_COLOR
    R As Long
    G As Long
    B As Long
End Type
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type


'=======================================================
'FUNCTIONS PUBLIQUES
'=======================================================
'=======================================================
'conversion de nombres
'=======================================================
Public Function LoWord(DWord As Long) As Long
    If DWord And &H8000& Then ' &H8000& = &H00008000
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function
Public Function HiWord(DWord As Long) As Long
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
'=======================================================
'convertit une couleur en long vers RGB
'=======================================================
Public Sub LongToRGB(ByVal Color As Long, ByRef R As Long, ByRef G As Long, ByRef B As Long)
    R = Color And &HFF&
    G = (Color And &HFF00&) \ &H100&
    B = Color \ &H10000
End Sub
