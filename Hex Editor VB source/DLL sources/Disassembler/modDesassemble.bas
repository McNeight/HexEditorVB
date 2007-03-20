Attribute VB_Name = "modDesassemble"
' =======================================================
'
' Disassembler DLL
' Coded by ShareVB
'
' =======================================================
'
' Copyright © 2006-2007 by ShareVB.
'
' This file is part of Disassembler DLL.
'
' Disassembler DLL is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' Disassembler DLL is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with Disassembler DLL; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' =======================================================

Option Explicit

Private Sub MakeDir(szDir As String)
On Local Error Resume Next
MkDir szDir
End Sub

Public Sub DisassembleFile(szFilename As String, sDir As String)
Dim szOutPattern As String, pos1 As Long, pos2 As Long

'szOutPattern = Mid$(szFilename, 1, InStrRev(szFilename, "\"))
pos1 = InStrRev(szFilename, "\")
If pos1 = 0 Then pos1 = 1 Else pos1 = pos1 + 1
pos2 = InStrRev(szFilename, ".")
If pos2 = 0 Then pos2 = Len(szFilename) + 1

szOutPattern = sDir & Mid$(szFilename, pos1, pos2 - pos1) 'szFilename & "_desam\" & Mid$(szFilename, pos1, pos2 - pos1)
MakeDir sDir 'szFilename & "_desam\"

If IsLIB(szFilename) Then
  '  If MsgBox("Ce fichier est une library. Son désassemblage peut produire un très grand nombre de fichiers." & vbCrLf & "Voulez-vous continuer ?", vbExclamation Or vbYesNo) = vbYes Then
        DysLIBFile szFilename, szOutPattern
   ' End If
ElseIf IsPE(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable PE.", vbInformation
    DysPE szFilename, szOutPattern, True
ElseIf IsNE(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable NE (format non supporté)", vbCritical
    'TODO
ElseIf IsIntelLE(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable LE (VxD).", vbInformation
    DysLEFile szFilename, szOutPattern
ElseIf IsMZ(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable MZ (MS-DOS).", vbInformation
    DysMZ szFilename, szOutPattern, True
ElseIf IsCOFF(szFilename) Then
    'MsgBox "Ce fichier est un fichier objet COFF.", vbInformation
    DysCOFF szFilename, szOutPattern, True
'TODO OMF
Else
   ' MsgBox "Ce fichier est dans un format non supporté. Désolé.", vbCritical
End If
End Sub
