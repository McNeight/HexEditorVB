Attribute VB_Name = "Demarrage"
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

Public AfManifest As New AfClsManifest  'classe appliquant le style XP
Public cFile As clsFileInfos


'=======================================================
'sub de démarrage du programme
'=======================================================
Sub Main()

    '//on instancie les classes
    Set AfManifest = New AfClsManifest
    Set cFile = New clsFileInfos
    
    '//application du style XP
    AfManifest.Run
    
    '//lecture des préférences
    
    '//lance la form principale
    frmMain.Show
    
    '//gère le Command()
    
End Sub

'=======================================================
'quitte le programme
'=======================================================
Public Sub EndProg()

    '//désinstancie les classes
    Set cFile = Nothing
    Set AfManifest = Nothing
    
    '//quitte
    End
    
End Sub
