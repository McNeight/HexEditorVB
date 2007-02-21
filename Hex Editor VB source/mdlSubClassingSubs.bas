Attribute VB_Name = "mdlSubClassingSubs"
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
'//MODULE DE GESTION DU SUBCLASSING ET DES ICONES
'=======================================================


Public cSub As clsFrmSubClass   'contient la classe de suclassing
Private IMG As ImageList    'contient un ImageList



'=======================================================
'ajoute les icones au menu
'=======================================================
Public Sub AddIconsToMenus(ByVal Frm As Form, tIMG As ImageList)
Dim hMenu As Long
    
    'récupère le handle du menu de la form
    hMenu = GetMenu(Frm.hWnd)
    
    'vérifie qu'il existe bien un menu
    If hMenu = 0 Then
        Exit Sub
    End If
    
    'récupère l'imagelist en privé (car nécessaire ailleurs)
    Set IMG = tIMG
    
    Call AddSub(hMenu)  'liste récursivement les menus
                
End Sub

'=======================================================
'liste récursivement les menus et y ajoute des icones
'=======================================================
Private Function AddSub(hMenu As Long, Optional ByVal hParentMenu As Long, Optional ByVal iP As Long) As Long
Dim hSubMenu As Long
Dim tMenuInfo As MENUITEMINFO_STRINGDATA
Dim i As Long
Dim dwItemData As Long
Dim Text As String
Dim mnuC As Long
Dim Text2 As String
Dim stBuffer As String * 255

    'récupère le nombre d'items dans le menu
    mnuC = GetMenuItemCount(hMenu)


    'pour chaque item
    For i = 0 To mnuC - 1
    
        'récupère le premier sous menu
        hSubMenu = GetSubMenu(hMenu, i)
        
        If hSubMenu <> 0 Then
            'il y a encore un sous menu ==> récursif
            Call AddSub(hSubMenu, hMenu, i)
        End If
        
        'récupère le caption du menu
        With tMenuInfo
            .cbSize = Len(tMenuInfo)    'taille de la structure
            .dwTypeData = stBuffer & vbNullChar     'finie par un &H0
            .fType = MF_STRING  'type String
            .cch = 255   'taille du buffer
            .fState = MFS_DEFAULT   'non checked
            .fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU
            
            'récupère le caption
            Call GetMenuItemInfoStr(hMenu&, i&, MF_BYPOSITION, tMenuInfo)
            'supprime les &H0
            .dwTypeData = Replace$(.dwTypeData, vbNullChar, vbNullString)
            'sauvegarde la string
            Text2 = .dwTypeData
            
            'identique, mais avec le menu parent
        
            .cbSize = Len(tMenuInfo)
            .dwTypeData = stBuffer & Chr$(0)
            .fType = MF_STRING
            .cch = 255
            .fState = MFS_DEFAULT
            .fMask = MIIM_ID Or MIIM_STATE Or MIIM_TYPE Or MIIM_SUBMENU
            Call GetMenuItemInfoStr(hParentMenu&, iP&, MF_BYPOSITION, tMenuInfo)
            .dwTypeData = Replace$(.dwTypeData, vbNullChar, vbNullString)
            Text = .dwTypeData
    End With

    'ajoute les bitmaps du IMG au menu
    Call AddBitmap(hMenu, i, Text & "|" & Text2)
        
  Next i
  
End Function

'=======================================================
'ajoute une bitmap au menu spécifié
'=======================================================
Private Sub AddBitmap(ByVal hSubMenu As Long, lPos As Long, id As String)
    
    On Error Resume Next    'pas d'erreur si jamais pas de bitmap définie
    
    'affecte la bitmap
    SetMenuItemBitmaps hSubMenu, lPos, MF_BYPOSITION, IMG.ListImages.Item(id).Picture, _
        IMG.ListImages.Item(id).Picture
End Sub

'=======================================================
'routine de remplacement pour l'interception des messages ==> subclassing
'=======================================================
Public Function MaWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
    
    Debug.Print uMsg & "  " & wParam & "  " & lParam
    
    Select Case uMsg
                
        Case WM_MENUSELECT    'selection d'un menu (surbrillance)

            Call cSub.OnMenuSelect(hWnd, LoWord(wParam), HiWord(wParam), lParam, lParam)
    
        Case Else
           'appel de la routine standard pour les autres messages
           MaWndProc = CallWindowProc(cSub.AddressWndProc, hWnd, uMsg, wParam, lParam)
    End Select
    
End Function

'=======================================================
'fonctions de conversion de long en integer
'=======================================================
Private Function LoWord(ByVal DWord As Long) As Long
  If DWord And &H8000& Then
    LoWord = DWord Or &HFFFF0000
  Else
    LoWord = DWord And &HFFFF&
  End If
End Function
Private Function HiWord(ByVal DWord As Long) As Long
  HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

'Private Function HiWord(LongIn As Long) As Integer
'    Call CopyMemory(HiWord, ByVal VarPtr(LongIn) + 2, 2)
'End Function
'Private Function LoWord(LongIn As Long) As Integer
'    Call CopyMemory(LoWord, LongIn, 2)
'End Function
