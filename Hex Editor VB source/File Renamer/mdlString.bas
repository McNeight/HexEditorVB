Attribute VB_Name = "mdlString"
' -----------------------------------------------
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
'
' An Windows utility which allows to rename lots of file (part of Hex Editor VB)
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
' -----------------------------------------------


Option Explicit


'-----------------------------------------------
'//MODULE DE GESTION DES STRINGS
'-----------------------------------------------

'-----------------------------------------------
'renvoie si oui ou non une string est convenable pour un nom de fichier
'-----------------------------------------------
Public Function IsFileNameOK(ByVal sFileName As String) As Boolean
Dim s As String
Dim x As Long
    IsFileNameOK = False
    For x = 1 To Len(sFileName)
        s = Mid$(sFileName, x, 1)
        If s = Chr$(34) Or s = "\" Or s = "/" Or s = ":" Or s = "*" Or s = "?" Or s = "<" Or _
            s = ">" Or s = "|" Then
            Exit Function
        End If
    Next
    IsFileNameOK = True
End Function

'-----------------------------------------------
'la fonction qui est au coeur de tout : le renommage
'prend directement en paramètres le composant listbox concerné
'ne renomme RIEN directement, renvoie en sortie un tableau avec les nouveaux noms
'-----------------------------------------------
Public Sub RenameMyFiles(Lst As ListBox, OldNames() As String, ByRef NewNames() As String)
     NewNames = OldNames
End Sub
