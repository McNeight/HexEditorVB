Attribute VB_Name = "modImportLib"
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

Private Const IMPORT_CODE As Long = 0
Private Const IMPORT_DATA As Long = 1
Private Const IMPORT_CONST As Long = 2

Private Const IMPORT_ORDINAL As Long = 0
Private Const IMPORT_NAME As Long = 1
Private Const IMPORT_NAME_NOPREFIX As Long = 2
Private Const IMPORT_NAME_UNDECORATE As Long = 3

Private Type ImportHeader
    Sig1 As Integer         'Must be IMAGE_FILE_MACHINE_UNKNOWN. See Section 3.3.1, "Machine Types, " for more information.
    Sig2 As Integer         'Must be 0xFFFF.
    Version As Integer
    Machine As Integer      'Number identifying type of target machine. See Section 3.3.1, "Machine Types, " for more information.
    TimeDateStamp As Long   'Time and date the file was created.
    SizeOfData As Long      'Size of the strings following the header.
    OrdinalHint As Integer  'Either the ordinal or the hint for the import, determined by the value in the Name Type field.
    BitType As Integer      'The import type. See Section 8.2 Import Type for specific values and descriptions.
                            'The Import Name Type. See Section 8.3. Import Name Type for specific values and descriptions.
                            'Reserved. Must be zero.
End Type

Public Sub DysImport(ByVal strExeName As String, ByVal strOutFilePattern As String, ByVal iFileNum As Integer)
Dim IH As ImportHeader
Dim szRawImportName As String, szImportName As String
Dim szDllName As String

    getUnkOffset 0, VarPtr(IH), Len(IH)
    
    setPointerOffset Len(IH)
    szRawImportName = GetSZString
    szDllName = GetSZString
    With IH
        Select Case (.BitType And &H1C) \ 4
            Case IMPORT_NAME_NOPREFIX
                Select Case Asc(Mid$(szRawImportName, 1, 1))
                    Case 63, 64, 95 '?,@,_
                        szImportName = Mid$(szRawImportName, 2)
                End Select
            Case IMPORT_NAME_UNDECORATE
                Select Case Asc(Mid$(szRawImportName, 1, 1))
                    Case 63, 64, 95 '?,@,_
                        szImportName = Mid$(szRawImportName, 2)
                End Select
                szImportName = Mid$(szImportName, 1, InStr(szImportName, "@") - 1)
        End Select
        
        Print #iFileNum, "----------------------------------------------------------------------"
        Print #iFileNum, "Import Library :", strExeName
        Print #iFileNum, "Import "; szImportName; " ("; szRawImportName; ") from "; szDllName
        Print #iFileNum, "----------------------------------------------------------------------"
        
        Print #iFileNum, "Machine :", .Machine
        
        Select Case (.BitType And &H1C) \ 4
            Case IMPORT_ORDINAL
                Print #iFileNum, "Ordinal :", .OrdinalHint
            Case IMPORT_NAME
                Print #iFileNum, "Name :", .OrdinalHint
        End Select
        
        Print #iFileNum, "Version :", .Version
        Print #iFileNum, "Time Date Stamp :", .TimeDateStamp
        Print #iFileNum, "Size Of Data :", .SizeOfData; " byte(s)"
        
        Print #iFileNum, "Import Type :",
        Select Case (.BitType And &H3&)
            Case IMPORT_CODE
                Print #iFileNum, "CODE"
            Case IMPORT_CONST
                Print #iFileNum, "CONST"
            Case IMPORT_DATA
                Print #iFileNum, "DATA"
        End Select
    End With
End Sub
