Attribute VB_Name = "Declarations"
' =======================================================
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A Windows utility which allows to rename lots of file (part of Hex Editor VB)
'
' Copyright (c) 2006-2007 by Alain Descotes.
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
'//DECLARATIONS DES APIS/TYPES/CONSTANTES/ENUMS
'=======================================================

'=======================================================
'//CONSTANTES
'=======================================================

'constantes contenant mes couleurs publiques
Public Const GREEN_COLOR                        As Long = &HC000&
Public Const RED_COLOR                          As Long = &HC0&


'=======================================================
'//ENUMS
'=======================================================

Public Enum TYPE_OF_MODIFICATION
    Style = 1
    Compteur = 2
    Remplacer = 3
    Base = 4
    Audio = 5
    Video = 6
    Image = 7
End Enum


'=======================================================
'//TYPES
'=======================================================


'=======================================================
'//APIS
'=======================================================

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)
Public Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

