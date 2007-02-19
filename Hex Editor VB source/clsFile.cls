VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
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
'//PERSONNAL TYPE
'INFOS SUR UN FICHIER
'=======================================================


'=======================================================
'variables de la structure clsFile
'=======================================================
Public File As String
Public FileDrive As String
Public FileExtension As String
Public FileDirectory As String
Public FileType As String
Public AssociatedExecutableProgram As String
Public FileSize As Currency
Public FileCompressedSize As Currency
Public CreationDate As String
Public LastModificationDate As String
Public LastAccessDate As String
Public FileAttributes As Long
Public IsHidden As Boolean
Public IsReadOnly As Boolean
Public IsSystem As Boolean
Public IsNormal As Boolean
Public ShortName As String
Public ShortPath As String
Public EXECompanyName As String
Public EXEFileDescription As String
Public EXEFileVersion As String
Public EXEInternalName As String
Public EXELegalCopyright As String
Public EXEOriginalFileName As String
Public EXEProductName As String
Public EXEProductVersion As String
