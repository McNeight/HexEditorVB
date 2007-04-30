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

'=======================================================
'MODULE CONTENANT LES PROC DE DEMARRAGE
'=======================================================

Public AfManifest As New AfClsManifest  'classe appliquant le style XP
Public Chr_(255) As String  'contient la liste des char, pour gagner en vitesse
Public cFile As FileSystemLibrary.FileSystem


'=======================================================
'sub de d�marrage du programme
'=======================================================
Sub Main()
Dim x As Long

    '//on instancie les classes
    Set AfManifest = New AfClsManifest
    Set cFile = New FileSystemLibrary.FileSystem
            
    '//application du style XP
    Call AfManifest.Run
    
    '//on remplit le tableau Chr_()
    For x = 0 To 255
        Chr_(x) = Chr$(x)
    Next x
        
    '//lecture des pr�f�rences
    
    '//lance la form principale
    frmMain.Show
    
    '//g�re le Command()
    
End Sub

'=======================================================
'quitte le programme
'=======================================================
Public Sub EndProg()

    '//d�sinstancie les classes
    Set cFile = Nothing
    Set AfManifest = Nothing
    
    '//quitte
    End
    
End Sub
