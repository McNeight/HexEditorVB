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


Private AddrWndProc As Long   'adresse de la routine standart de traitement des events
Private cEvent As clsFrmSubClassEvent   'contient la classe (définie WithEvents) qui libèrera l'event




'=======================================================
'récupèration de l'adresse de la routine standart
'=======================================================
Public Property Get AddressWndProc() As Long
    AddressWndProc = AddrWndProc
End Property

'=======================================================
'fonction qui active le hook de la form
'=======================================================
Public Function HookFormMenu(ByVal hWnd As Long, ByVal cFrmSubClassEvent As clsFrmSubClassEvent) As Long
    
    'instancie la classe de l'event
    Set cEvent = cFrmSubClassEvent
    
    'récupère l'adresse de la routine standart
    AddrWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MaWndProc)
    
    HookFormMenu = AddrWndProc
End Function

'=======================================================
'désactive le hook de la form
'=======================================================
Public Function UnHookFormMenu(ByVal hWnd As Long) As Long
    UnHookFormMenu = SetWindowLong(hWnd, GWL_WNDPROC, AddrWndProc)  'redonne l'adresse de la routine standart
End Function

'=======================================================
'sub qui sera activée avec le message capté
'=======================================================
Public Sub OnMenuSelect(hWnd As Long, nItemID As Integer, nFlags As Integer, hSysMenu As Long, wParam As Long)
Dim tMenuInfo As MENUITEMINFO_STRINGDATA
Dim hMenu As Long
Dim stBuffer As String * 255

    'récupère le handle du menu
    hMenu = GetMenu(hWnd&)

    With tMenuInfo
        .cbSize = Len(tMenuInfo)    'taille du type
        .dwTypeData = stBuffer & vbNullChar    'null terminated string
        .fType = MF_STRING  'type String
        .cch = 255  'taille du buffer
        .fState = MFS_DEFAULT   'non checked par défaut
        .fMask = MIIM_ID Or MIIM_TYPE Or MIIM_STATE
    
        'vérifie que c'est un menu racine ou pas
        If nFlags <> -32624 Then
            'pas de sous menu
            Call GetMenuItemInfoStr(hMenu&, nItemID, False, tMenuInfo)
            frmContent.Caption = nFlags & "   " & hSysMenu & "   " & nItemID & "   " & hMenu & "   " & wParam
        Else
            'alors on a un sous menu
            'Call GetMenuItemInfoStr(hMenu&, nItemID, False, tMenuInfo)
            frmContent.Caption = nFlags & "   " & hSysMenu & "   " & nItemID & "   " & hMenu & "   " & wParam
        End If
        
        'vire les &H0 de la string
        .dwTypeData = Replace$(.dwTypeData, vbNullChar, vbNullString)
    
        'libère l'event dans la classe cEvent
        Call cEvent.RaiseEventCaption(.dwTypeData)
    End With

End Sub

'=======================================================
'on libère la mémoire occupée par la classe de l'Event
'=======================================================
Private Sub Class_Terminate()
    Set cEvent = Nothing
End Sub
