Attribute VB_Name = "mdlColorMenus"
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
'MODULE POUR LA COLORISATION DE MENUS
'=======================================================

'=======================================================
'colorise les menus de la form
'=======================================================
Public Sub ColorFormMenu(Form As Object, ByVal Color As Long)
Dim hMenu As Long
Dim hBrush As Long
Dim lbBrushInfo As LOGBRUSH
Dim miMenuInfo As tagMENUINFO
    
    'définit le brush qui sera utilisé
    With lbBrushInfo
        .lbStyle = BS_SOLID
        .lbColor = Color
        .lbHatch = 0
    End With
    
    hBrush = CreateBrushIndirect(lbBrushInfo) 'We create our brush
    hMenu = GetMenu(Form.hWnd) 'Get the handle to the menu that we are modifying (note we pass the form's hWnd because it is the owner of the menu)
    miMenuInfo.cbSize = Len(miMenuInfo) 'Set the MenuInfo structure size so that we don't get errors
    Call GetMenuInfo(hMenu, miMenuInfo) 'Go and get the actual menu info should return non-zero if successful
    miMenuInfo.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND 'Set the mask for the changes (changing the background for menu and all sub-menus)
    miMenuInfo.hbrBack = hBrush 'Assign our brush to the menu info
    Call SetMenuInfo(hMenu, miMenuInfo)  'Write our info back to the menu and we're done. (should return non-zero if successful)
    
End Sub

'=======================================================
'applique une image dans un toolbar
'=======================================================
Public Sub ColorToolbar(TB As Object, hPicture As Long)
Dim lTBWnd As Long
Dim LngNew As Long

    'récupère le handle de la picturebox contenant l'image
    LngNew = CreatePatternBrush(hPicture)
    
    'applique le changement
    Call DeleteObject(SetClassLong(TB.hWnd, GCL_HBRBACKGROUND, LngNew))
    
    'refresh
    Call InvalidateRect(0&, 0&, False)
    
End Sub
