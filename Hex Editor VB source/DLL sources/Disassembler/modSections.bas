Attribute VB_Name = "modSections"
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

Private Const IMAGE_SCN_CNT_CODE As Long = &H20&
Private Const IMAGE_SCN_MEM_EXECUTE As Long = &H20000000
Private Const IMAGE_SCN_MEM_16BIT As Long = &H20000

Public Type SectionHeader
    SecName(0 To 7) As Byte
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

Public dwImageBase As Long
Public retSectionTables() As SectionHeader
'TODO BSS : uninitialized data

Public Function GetSectionVA(ByVal va As Long) As Long
Dim X As Long 'var de contrôle
Dim ubSec As Long

ubSec = UBound(retSectionTables)
'pour chaque section
For X = 0 To ubSec
    With retSectionTables(X)
        'si l'adresse virtuelle se trouve dans la plage des adresses virtuelles de la sections
        If (va >= .VirtualAddress) And (va < (.VirtualAddress + .VirtualSize)) Then
            GetSectionVA = X
            Exit Function
        End If
    End With
Next
GetSectionVA = -1
End Function

Public Function Offset2VA(ByVal Offset As Long) As Long
Dim X As Long 'var de contrôle
Dim ubSec As Long

ubSec = UBound(retSectionTables)
'pour chaque section
For X = 0 To ubSec
    With retSectionTables(X)
        'si l'adresse virtuelle se trouve dans la plage des adresses virtuelles de la sections
        If (Offset >= .PointerToRawData) And (Offset < (.PointerToRawData + .SizeOfRawData)) Then
            Offset2VA = Offset - .PointerToRawData + .VirtualAddress '- 1
            Exit Function
        End If
    End With
Next
Offset2VA = 0
End Function

Public Function VA2Offset(ByVal va As Long) As Long
Dim X As Long 'var de contrôle

X = GetSectionVA(va)
If X <> -1 Then
    With retSectionTables(X)
        'alors l'adresse virtuelle appartient à la section
        'l'offset est RVA - VA de base de la section + offset de la section
        '(+ 1 pour VB qui prend ses offsets à partir de 1)
        VA2Offset = va - .VirtualAddress + .PointerToRawData '+ 1
    End With
Else
    VA2Offset = -1
End If
End Function

Public Function GetSectionRVA(ByVal rva As Long) As Long
GetSectionRVA = GetSectionVA(rva + dwImageBase)
End Function

Public Function RVA2Offset(ByVal rva As Long) As Long
    RVA2Offset = VA2Offset(rva + dwImageBase)
End Function

Public Function Offset2RVA(ByVal Offset As Long) As Long
    Offset2RVA = Offset2VA(Offset) - dwImageBase
End Function

'renvoit le type RWX d'une adresse virtuelle relative
'================================================================================
'IN RVA : (Relative Virtual Address) adresse virtuelle relative dont on veut le type RWX
Public Function IsCodeVA(ByVal va As Long) As Boolean
Dim X As Long 'var de contrôle

X = GetSectionVA(va)
If X <> -1 Then
    With retSectionTables(X)
        IsCodeVA = ((.Characteristics And IMAGE_SCN_CNT_CODE) = IMAGE_SCN_CNT_CODE) ' And ((.Characteristics And IMAGE_SCN_MEM_EXECUTE) = IMAGE_SCN_MEM_EXECUTE)
    End With
End If
End Function

'renvoit le type RWX 16bits d'une adresse virtuelle relative
'================================================================================
'IN RVA : (Relative Virtual Address) adresse virtuelle relative dont on veut le type RWX 16bits
Public Function IsCode16VA(ByVal va As Long) As Boolean
Dim X As Long 'var de contrôle

X = GetSectionVA(va)
If X <> -1 Then
    With retSectionTables(X)
        IsCode16VA = ((.Characteristics And IMAGE_SCN_CNT_CODE) = IMAGE_SCN_CNT_CODE) And ((.Characteristics And IMAGE_SCN_MEM_16BIT) = IMAGE_SCN_MEM_16BIT)
    End With
End If
End Function

'indique si la RVA est valide
Public Function CheckVA(ByVal va As Long) As Boolean
    CheckVA = (GetSectionVA(va) <> -1)
End Function

Public Sub ProcessSections()
    Dim X As Long, addrdeb As Long, addrfin As Long, addr As Long, dw As Long, ubs As Long
    
    ubs = UBound(retSectionTables)
   ' With frmProgress.pbSection1
     '   .Min = 0
     '   .Max = ubs
        For X = 0 To ubs
            With retSectionTables(X)
                addrdeb = .PointerToRawData + .VirtualSize
                addrfin = .PointerToRawData + .SizeOfRawData - 1
                
                For addr = addrdeb To addrfin
                    setMapOffset addr, 255
                Next
                If ((.Characteristics And IMAGE_SCN_CNT_CODE) = IMAGE_SCN_CNT_CODE) Then 'And _
                   ((.Characteristics And IMAGE_SCN_MEM_EXECUTE) = IMAGE_SCN_MEM_EXECUTE) Then
                    addrdeb = .PointerToRawData
                    addrfin = .PointerToRawData + .VirtualSize
                    For addr = addrdeb To addrfin Step 4
                        dw = getDwordOffset(addr)
                        If CheckVA(dw) Then
                            setMapOffset addr, 4
                        End If
                    Next
                End If
            End With
          '  .value = X
        Next
   ' End With
End Sub

