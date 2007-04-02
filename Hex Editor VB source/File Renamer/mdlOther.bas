Attribute VB_Name = "mdlOther"
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
'//MODULE DE SUBS ET PROC DIVERSES
'=======================================================

'=======================================================
'récupère le nom de l'utilisateur
'=======================================================
Public Function GetUserName() As String
Dim strS As String
Dim Ret As Long

    'créé un buffer
    strS = String$(200, 0)
    
    'récupère le Name
    Ret = GetUserNameA(strS, 199)
    If Ret <> 0 Then GetUserName = Left$(strS, 199) Else GetUserName = vbNullString
End Function
