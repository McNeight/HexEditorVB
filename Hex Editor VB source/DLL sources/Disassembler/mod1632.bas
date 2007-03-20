Attribute VB_Name = "mod1632"
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

Public bOperandSizeOverride As Byte
Public bAddressSizeOverride As Byte
Public dwInitESP As Long
'Public dwAddressSizeBytes As Byte
'Public dwAddressSizeBits As Byte

Public Sub Set16BitsDecode()
bOperandSizeOverride = &H0
bAddressSizeOverride = &H0
dwInitESP = 2
End Sub

Public Sub Set32BitsDecode()
bOperandSizeOverride = &H66
bAddressSizeOverride = &H67
dwInitESP = 4
End Sub

