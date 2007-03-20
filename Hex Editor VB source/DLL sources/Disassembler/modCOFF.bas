Attribute VB_Name = "modCOFF"
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

Private Const UNDNAME_COMPLETE As Long = &H0
Private Declare Function UnDecorateSymbolName Lib "dbghelp.dll" (ByVal DecoratedName As String, ByVal UnDecoratedName As String, ByVal UndecoratedLength As Long, ByVal Flags As Long) As Long

Public Enum SymBaseType
IMAGE_SYM_TYPE_NULL = 0
IMAGE_SYM_TYPE_VOID = 1
IMAGE_SYM_TYPE_CHAR = 2
IMAGE_SYM_TYPE_SHORT = 3
IMAGE_SYM_TYPE_INT = 4
IMAGE_SYM_TYPE_LONG = 5
IMAGE_SYM_TYPE_FLOAT = 6
IMAGE_SYM_TYPE_DOUBLE = 7
IMAGE_SYM_TYPE_STRUCT = 8
IMAGE_SYM_TYPE_UNION = 9
IMAGE_SYM_TYPE_ENUM = 10
IMAGE_SYM_TYPE_MOE = 11
IMAGE_SYM_TYPE_BYTE = 12
IMAGE_SYM_TYPE_WORD = 13
IMAGE_SYM_TYPE_UINT = 14
IMAGE_SYM_TYPE_DWORD = 15
End Enum

Public Enum SymComplexType
IMAGE_SYM_DTYPE_NULL = 0
IMAGE_SYM_DTYPE_POINTER = 1
IMAGE_SYM_DTYPE_FUNCTION = 2
IMAGE_SYM_DTYPE_ARRAY = 3
End Enum

Public Enum StorageClass
IMAGE_SYM_CLASS_END_OF_FUNCTION = -1
IMAGE_SYM_CLASS_NULL = 0
IMAGE_SYM_CLASS_AUTOMATIC = 1
IMAGE_SYM_CLASS_EXTERNAL = 2
IMAGE_SYM_CLASS_STATIC = 3
IMAGE_SYM_CLASS_REGISTER = 4
IMAGE_SYM_CLASS_EXTERNAL_DEF = 5
IMAGE_SYM_CLASS_LABEL = 6
IMAGE_SYM_CLASS_UNDEFINED_LABEL = 7
IMAGE_SYM_CLASS_MEMBER_OF_STRUCT = 8
IMAGE_SYM_CLASS_ARGUMENT = 9
IMAGE_SYM_CLASS_STRUCT_TAG = 10
IMAGE_SYM_CLASS_MEMBER_OF_UNION = 11
IMAGE_SYM_CLASS_UNION_TAG = 12
IMAGE_SYM_CLASS_TYPE_DEFINITION = 13
IMAGE_SYM_CLASS_UNDEFINED_STATIC = 14
IMAGE_SYM_CLASS_ENUM_TAG = 15
IMAGE_SYM_CLASS_MEMBER_OF_ENUM = 16
IMAGE_SYM_CLASS_REGISTER_PARAM = 17
IMAGE_SYM_CLASS_BIT_FIELD = 18
IMAGE_SYM_CLASS_BLOCK = 100
IMAGE_SYM_CLASS_FUNCTION = 101
IMAGE_SYM_CLASS_END_OF_STRUCT = 102
IMAGE_SYM_CLASS_FILE = 103
IMAGE_SYM_CLASS_SECTION = 104
IMAGE_SYM_CLASS_WEAK_EXTERNAL = 105
End Enum

Public Enum SectionNumberValues
IMAGE_SYM_UNDEFINED = 0
IMAGE_SYM_ABSOLUTE = -1
IMAGE_SYM_DEBUG = -2
End Enum

Public Enum RelocationType
IMAGE_REL_I386_ABSOLUTE = &H0&    'This relocation is ignored.
IMAGE_REL_I386_DIR16 = &H1&       'Not supported.
IMAGE_REL_I386_REL16 = &H2&       'Not supported.
IMAGE_REL_I386_DIR32 = &H6&       'The target's 32-bit virtual address.
IMAGE_REL_I386_DIR32NB = &H7&     'The target's 32-bit relative virtual address.
IMAGE_REL_I386_SEG12 = &H9&       'Not supported.
IMAGE_REL_I386_SECTION = &HA&     'The 16-bit-section index of the section containing the target. This is used to support debugging information.
IMAGE_REL_I386_SECREL = &HB&      'The 32-bit offset of the target from the beginning of its section. This is used to support debugging information as well as static thread local storage.
IMAGE_REL_I386_REL32 = &H14&      'The 32-bit relative displacement to the target. This supports the x86 relative branch and call instructions.
End Enum

Public Type RawCOFFSymbol
    SymName(7) As Byte
    value As Long
    SectionNumber As Integer
    BaseType As Byte
    ComplexType As Byte
    StorageClass As Byte
    NumberOfAuxSymbols As Byte
End Type

'Public Type RawSection
'    SecName(7) As Byte
'    VirtualSize As Long
'    VirtualAddress As Long
'    SizeOfRawData As Long
'    PointerToRawData As Long
'    PointerToRelocations As Long
'    PointerToLinenumbers As Long
'    NumberOfRelocations As Integer
'    NumberOfLinenumbers As Integer
'    Characteristics As Long
'End Type

Public Type RawCOFFRelocation
    VirtualAddress As Long
    SymbolTableIndex As Long
    Type As Integer
End Type

Public Type RawCOFFLineNumber
Type  As Long
Linenumber As Integer
End Type

Public Type RawOptionalHeader
Magic As Integer
MajorLinkerVersion  As Byte
MinorLinkerVersion  As Byte
SizeOfCode As Long
SizeOfInitializedData   As Long
SizeOfUninitializedData  As Long
AddressOfEntryPoint As Long
BaseOfCode As Long
'PE32 contains this additional field, absent in PE32+, following BaseOfCode:
BaseOfData As Long
End Type

''relocalisation du code
''cette structure n'existe pas comme ca dans le fichier exe
'Public Type COFFRelocation
'VirtualAddress As Long ' RVA du code
'SymbolTableIndex As Long ' symbole
'Symbol As COFFSymbol 'symbole
'Type As Integer ' type de relocalisation
'End Type
'
''information de numéro de ligne
''cette structure n'existe pas comme ca dans le fichier exe
'Public Type COFFLineNumber
'SymbolTableIndex As Long 'symbole
'VirtualAddress As Long 'RVA
'Linenumber As Integer 'numéro de ligne
'Symbol As COFFSymbol 'symbole
'End Type
'
''entete de section COFF
''cette structure n'existe pas comme ca dans le fichier exe
'Public Type COFFSection
'Name As String * 8 'nom de section
'LongName As String 'nom long de section si > 8 caractères
'VirtualSize As Long 'taille de la section une fois chargée en mémoire
'VirtualAddress As Long 'RVA de la section chargée
'SizeOfRawData As Long 'taille des données dans l'exe
'PointerToRawData As Long 'offset des données de la section
'PointerToRelocations As Long 'offset des données de relocalisation
'PointerToLinenumbers As Long 'offset des données de numéro de lignes
'NumberOfRelocations As Integer 'nombre de relocalisations
'NumberOfLinenumbers As Integer 'nombre de numéro de ligne
'Characteristics As SectionCharacteristics 'caracteristiques
'LineNumbers() As COFFLineNumber 'numéro de ligne
'Relocations() As COFFRelocation 'relocatlisations
'Data() As Byte 'données de la sections
'End Type
'
''entete PE optionelle
'Public Type PEOptionalHeader
'Magic As Integer 'identification
'MajorLinkerVersion  As Byte 'version de l'éditeur de liens
'MinorLinkerVersion  As Byte 'version de l'éditeur de liens
'SizeOfCode As Long 'taille du code
'SizeOfInitializedData   As Long 'taille des données initialisées
'SizeOfUninitializedData  As Long 'taille des données non initialisées
'AddressOfEntryPoint As Long ' adresse du poijnt d'entrée du code
'BaseOfCode As Long
'BaseOfData As Long
'End Type

''structure représentant un fichier COFF .obj
''cette structure n'existe pas comme ca dans le fichier exe
'Public Type COFFFile
'Machine As Machine 'identificateur
'NumberOfSections As Integer 'nombre de sections
'TimeDateStamp As Long 'date de compilation
'PointerToSymbolTable As Long 'offset de la table des symboles
'NumberOfSymbols As Long 'nombre de symboles dans la table
'SizeOfOptionalHeader As Integer 'taille de l'entete PE optionnelle
'Characteristics As Characteristic 'caracteristiques du fichier
'OptionalHeader As PEOptionalHeader 'entete PE Optionnelle
'Sections() As COFFSection 'sections du fichier
'SymbolTable() As COFFSymbol 'table des symboles
'StringTableLenght As Long 'taille de la table de chaines de caractères
'StringTable() As String 'table de chaines de caracteres
'End Type

'fonction permettant de copier une zone mémoire dans une autre
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private EntryPoints As Collection
Private EntryNames As Collection
Private Jumps As Collection

Public Sub InitCOFF()
Set EntryNames = New Collection
Set EntryPoints = New Collection
Set Jumps = New Collection
End Sub

Private Sub ProcessBType(Symbol As RawCOFFSymbol, ByVal iFileNum As Integer, ByVal bSetType As Boolean)
Dim b As Byte
With Symbol
    Select Case (.BaseType And &HF)
        Case IMAGE_SYM_TYPE_BYTE
            Print #iFileNum, "BYTE"
            b = 30
        Case IMAGE_SYM_TYPE_CHAR
            Print #iFileNum, "CHAR"
            b = 5
        Case IMAGE_SYM_TYPE_DOUBLE
            Print #iFileNum, "DOUBLE"
            b = 3
        Case IMAGE_SYM_TYPE_DWORD
            Print #iFileNum, "DWORD"
            b = 32
        Case IMAGE_SYM_TYPE_ENUM
            Print #iFileNum, "ENUM"
            b = 3
        Case IMAGE_SYM_TYPE_FLOAT
            Print #iFileNum, "FLOAT"
            b = 3
        Case IMAGE_SYM_TYPE_INT
            Print #iFileNum, "INT"
            b = 32
        Case IMAGE_SYM_TYPE_LONG
            Print #iFileNum, "LONG"
            b = 32
        Case IMAGE_SYM_TYPE_MOE
            Print #iFileNum, "MOE"
            b = 3
        Case IMAGE_SYM_TYPE_NULL
            If bSetType Then
                b = GetDataType(Offset2VA(retSectionTables(.SectionNumber - 1).PointerToRawData + .value), 0)
            End If
            Print #iFileNum, "NULL"
        Case IMAGE_SYM_TYPE_SHORT
            Print #iFileNum, "SHORT"
            b = 31
        Case IMAGE_SYM_TYPE_STRUCT
            Print #iFileNum, "STRUCT"
            b = 3
        Case IMAGE_SYM_TYPE_UINT
            Print #iFileNum, "UINT"
            b = 32
        Case IMAGE_SYM_TYPE_UNION
            Print #iFileNum, "UNION"
            b = 3
        Case IMAGE_SYM_TYPE_VOID
            Print #iFileNum, "VOID"
            bSetType = False
        Case IMAGE_SYM_TYPE_WORD
            Print #iFileNum, "WORD"
            b = 31
    End Select
    If bSetType Then
        setMapOffset retSectionTables(.SectionNumber - 1).PointerToRawData + .value, b
    End If
End With
End Sub

Private Sub ProcessType(Symbol As RawCOFFSymbol, COFF As RawCOFFHeader, ByVal iFileNum As Integer, ByVal bImport As Boolean)
With Symbol
    Select Case (.BaseType And &HF0) \ &H10
        Case IMAGE_SYM_DTYPE_ARRAY
            Print #iFileNum, "ARRAY OF ";
            ProcessBType Symbol, iFileNum, True
        Case IMAGE_SYM_DTYPE_FUNCTION
            Print #iFileNum, "FUNCTION RETURN ";
            ProcessBType Symbol, iFileNum, False
            If bImport = False Then
                EntryPoints.Add Offset2VA(retSectionTables(.SectionNumber - 1).PointerToRawData + .value)
                EntryNames.Add GetSymbolName(Symbol, COFF)
            End If
        Case IMAGE_SYM_DTYPE_NULL
            Print #iFileNum, "VAR ";
            ProcessBType Symbol, iFileNum, True And (Not bImport)
        Case IMAGE_SYM_DTYPE_POINTER
            Print #iFileNum, "POINTER TO ";
            ProcessBType Symbol, iFileNum, False
            setMapOffset retSectionTables(.SectionNumber - 1).PointerToRawData + .value, 4
    End Select
End With
End Sub

Private Sub ProcessStorage(Symbol As RawCOFFSymbol, COFF As RawCOFFHeader, ByVal iFileNum As Integer)
With Symbol
    Select Case .StorageClass
        'Case IMAGE_SYM_CLASS_BLOCK
            'A .bb (beginning of block) or .eb (end of block) record. Value is the relocatable address of the code location.
            'TODO
        Case IMAGE_SYM_CLASS_EXTERNAL
            If .SectionNumber = IMAGE_SYM_UNDEFINED Then
                'Value = size
                Print #iFileNum, ";IMPORTS "; GetSymbolName(Symbol, COFF); " ";
                ProcessType Symbol, COFF, iFileNum, True
            Else
                'Value offset within the section
                Print #iFileNum, ";EXPORTS "; GetSymbolName(Symbol, COFF); " ";
                ProcessType Symbol, COFF, iFileNum, False
            End If
        'Case IMAGE_SYM_CLASS_EXTERNAL_DEF
            'EXTERNAL SYMBOL : Value = nothing
            'TODO
        Case IMAGE_SYM_CLASS_LABEL
            'code label : Value = offset within the section
            Jumps.Add Offset2VA(retSectionTables(.SectionNumber - 1).PointerToRawData + .value)
        Case IMAGE_SYM_CLASS_STATIC
            
            'If .value = 0 Then
                'section name
            'Else
                'Value = offset within the section
                Print #iFileNum, ";STATIC "; GetSymbolName(Symbol, COFF); " ";
                ProcessType Symbol, COFF, iFileNum, False
            'End If
    
    '                    Case IMAGE_SYM_CLASS_ARGUMENT
    '                        'Value = number of the argument of a function
    '                    Case IMAGE_SYM_CLASS_AUTOMATIC
    '                        'automatic var : stack based (offset in Value)
    '                    Case IMAGE_SYM_CLASS_BIT_FIELD
    '                        'Value = nth bit in a bitfield
    '                    Case IMAGE_SYM_CLASS_END_OF_FUNCTION
    '                        'for debug
    '                    Case IMAGE_SYM_CLASS_END_OF_STRUCT
    '                        'end of struct
    '                    Case IMAGE_SYM_CLASS_ENUM_TAG
    '                        'Enum tag name
    '                    Case IMAGE_SYM_CLASS_FILE
    '                        'source file name
    '                    Case IMAGE_SYM_CLASS_FUNCTION
    '                        'extent of a function
    '                    Case IMAGE_SYM_CLASS_MEMBER_OF_ENUM
    '                        'Value = number of the member in an enum
    '                    Case IMAGE_SYM_CLASS_MEMBER_OF_STRUCT
    '                        'Value = number of the member in a structure
    '                    Case IMAGE_SYM_CLASS_MEMBER_OF_UNION
    '                        'Value = number of the member in an union
    '                    Case IMAGE_SYM_CLASS_NULL
    '                        'pas de classe de stockage
    '                    Case IMAGE_SYM_CLASS_REGISTER
    '                        'register var : Value = register number
    '                    Case IMAGE_SYM_CLASS_REGISTER_PARAM
    '                        'register param
    '                    Case IMAGE_SYM_CLASS_SECTION
    '                        'section
    '                    Case IMAGE_SYM_CLASS_STRUCT_TAG
    '                        'structure tag name
    '                    Case IMAGE_SYM_CLASS_TYPE_DEFINITION
    '                        'typedef
    '                    Case IMAGE_SYM_CLASS_UNDEFINED_LABEL
    '                        'not defined label
    '                    Case IMAGE_SYM_CLASS_UNDEFINED_STATIC
    '                        'static data declaration
    '                    Case IMAGE_SYM_CLASS_UNION_TAG
    '                        'union tag name
    '                    Case IMAGE_SYM_CLASS_WEAK_EXTERNAL
    '                        'weak reference
    End Select
End With
End Sub

Private Sub ProcessSectionNumber(Symbol As RawCOFFSymbol, COFF As RawCOFFHeader, ByVal iFileNum As Integer)
With Symbol
    Select Case .SectionNumber
        Case IMAGE_SYM_UNDEFINED
            If .value = 0 Then 'EXTERNAL SYMBOL
                ProcessStorage Symbol, COFF, iFileNum
            Else 'COMMON SYMBOL
                'TODO
                Debug.Assert False
            End If
        Case IMAGE_SYM_DEBUG
            'debug
        Case IMAGE_SYM_ABSOLUTE
            'not an address
        Case Else
            ProcessStorage Symbol, COFF, iFileNum
    End Select
End With
End Sub

Private Sub ProcessSymbolTable(ByVal iFileCoff As Integer, ByVal iFileIO As Integer, COFF As RawCOFFHeader)
    Dim COFFSym As RawCOFFSymbol, X As Long, off As Long, szSymName As String, Y As Long
    's'il ya des symboles dans le fichier
    If COFF.NumberOfSymbols Then
        Print #iFileCoff, "====================================================================="
        Print #iFileCoff, "Symbol Table"
        Print #iFileCoff, "====================================================================="
        Print #iFileCoff, "Number of Symbols :", COFF.NumberOfSymbols
        
        X = 0
        Do While X < COFF.NumberOfSymbols
            Print #iFileCoff, "---------------------------------------------------------------------"
            
            'on récupère le symbole
            off = COFF.PointerToSymbolTable + X * 18
            getUnkOffset off, VarPtr(COFFSym), Len(COFFSym)
            
            ProcessSectionNumber COFFSym, COFF, iFileIO
            
            With COFFSym
                szSymName = GetSymbolName(COFFSym, COFF)
                AddName off, szSymName
                AddSubName off, szSymName
                
                Print #iFileCoff, "Name :", szSymName, "->", GetUndecoratedName(szSymName)
                Print #iFileCoff, "Index :", X
                Print #iFileCoff, "---------------------------------------------------------------------"
                Print #iFileCoff, "Number Of Aux Symbols :", .NumberOfAuxSymbols
                
                Print #iFileCoff, "Section Number :", .SectionNumber
                Print #iFileCoff, "Storage Class :", .StorageClass
                Print #iFileCoff, "Value :", , .value
                
                Print #iFileCoff, "Base Type :", , .BaseType
                Print #iFileCoff, "Complex Type :", .ComplexType
                
                'on passe les symboles aux
                X = X + .NumberOfAuxSymbols + 1
            End With
        Loop
        Print #iFileCoff, "====================================================================="
    End If
End Sub

Private Sub ProcessStringTable(iFileCoff As Integer, COFF As RawCOFFHeader)
    Dim strtemp As String 'buffer chaine
    Dim StringTableLenght As Long, cb As Long, c As Byte
    
    Print #iFileCoff, "====================================================================="
    Print #iFileCoff, "String Table"
    Print #iFileCoff, "====================================================================="
    
    setPointerOffset COFF.PointerToSymbolTable + COFF.NumberOfSymbols * 18
    
    'on récupère la taille de la table des chaines de caracteres
    StringTableLenght = getDword(0) - 4
    
    'si elle contient des chaines
    If StringTableLenght > 0 Then
        'on les récupère
        Do
            strtemp = vbNullString
            setMap 5
            c = getByte(0)
            Do While c
                strtemp = strtemp & Chr$(c)
                c = getByte(0)
                StringTableLenght = StringTableLenght - 1
            Loop
            StringTableLenght = StringTableLenght - 1
            Print #iFileCoff, strtemp
        Loop While StringTableLenght
    End If
    Print #iFileCoff, "====================================================================="
End Sub

Private Sub ProcessOptionalHeader(iFileCoff As Integer, COFF As RawCOFFHeader)
    Dim OptionalHeader As RawOptionalHeader
    's'il y a de une entete Optionelle
    If COFF.SizeOfOptionalHeader Then
        'on la récupère
        getUnkOffset 20, VarPtr(OptionalHeader), Len(OptionalHeader)
    End If
End Sub

Private Sub ProcessSectionTable(iFileCoff As Integer, COFF As RawCOFFHeader)
Dim X As Long, Y As Long, rawReloc As RawCOFFRelocation, off As Long, value As Long
If COFF.NumberOfSections Then
    Print #iFileCoff, "====================================================================="
    Print #iFileCoff, "Sections"
    Print #iFileCoff, "====================================================================="
    Print #iFileCoff, "Number of Sections :", COFF.NumberOfSections
    
    ReDim retSectionTables(COFF.NumberOfSections)
    
    With retSectionTables(COFF.NumberOfSections)
        .Characteristics = 0
        .PointerToRawData = COFF.PointerToSymbolTable
        .SizeOfRawData = COFF.NumberOfSymbols * 18
        .VirtualAddress = .PointerToRawData
        .VirtualSize = .SizeOfRawData
        For Y = .PointerToRawData To .PointerToRawData + .SizeOfRawData - 1
            setMapOffset Y, 255
        Next
        .SecName(0) = Asc("U")
        .SecName(1) = Asc("N")
        .SecName(2) = Asc("D")
        .SecName(3) = Asc("E")
        .SecName(4) = Asc("F")
        .SecName(5) = 0
    End With
    
    'on récupère les entetes  des sections
    getUnkOffset 20 + COFF.SizeOfOptionalHeader, VarPtr(retSectionTables(0)), 40 * COFF.NumberOfSections
    
    'ici modif
    'With frmProgress.pbSection1
    '    .Min = 0
    '    .Max = COFF.NumberOfSections - 1
        'pour chaque section
        For X = 0 To COFF.NumberOfSections - 1
            Print #iFileCoff, "---------------------------------------------------------------------"
            With retSectionTables(X)
                'on analyse la section
                Print #iFileCoff, "Name :", , StrConv(.SecName, vbUnicode)
                Print #iFileCoff, "Index :", X + 1
                Print #iFileCoff, "---------------------------------------------------------------------"
                Print #iFileCoff, "Characteristics :", .Characteristics
                
                Print #iFileCoff, "Pointer To RawData :", .PointerToRawData
                Print #iFileCoff, "Size Of Raw Data :", .SizeOfRawData
                
                Print #iFileCoff, "Virtual Address :", .VirtualAddress
                Print #iFileCoff, "Virtual Size :", .VirtualSize
                
                Print #iFileCoff, "Pointer To Linenumbers :", .PointerToLinenumbers
                Print #iFileCoff, "Number Of Linenumbers :", .NumberOfLinenumbers
                
                Print #iFileCoff, "Pointer To Relocations :", .PointerToRelocations
                Print #iFileCoff, "Number Of Relocations :", .NumberOfRelocations
                            
                .VirtualAddress = .VirtualAddress + .PointerToRawData
                .VirtualSize = .SizeOfRawData
                
                'les relocalisations
                If .NumberOfRelocations Then
                    Print #iFileCoff, "---------------------------------------------------------------------"
                    Print #iFileCoff, "Relocations"
                    Print #iFileCoff, "---------------------------------------------------------------------"
                    For Y = 0 To .NumberOfRelocations - 1
                        getUnkOffset .PointerToRelocations + Y * 10, VarPtr(rawReloc), 10&
                        
                        Print #iFileCoff, , "Symbol Table Index :", rawReloc.SymbolTableIndex
                        Print #iFileCoff, , "Type :", rawReloc.Type
                        Print #iFileCoff, , "Virtual Address :", rawReloc.VirtualAddress
                        
                        off = .PointerToRawData + rawReloc.VirtualAddress
                        value = getDwordOffset(off)
                        Select Case rawReloc.Type
                            'Case IMAGE_REL_I386_ABSOLUTE
                            'Case IMAGE_REL_I386_DIR16
                            'Case IMAGE_REL_I386_REL16
                            'Case IMAGE_REL_I386_SEG12
                            Case IMAGE_REL_I386_DIR32NB '32-bit RVA
                                value = value + COFF.PointerToSymbolTable + rawReloc.SymbolTableIndex * 18
                            Case IMAGE_REL_I386_DIR32 '32-bit VA
                                value = value + COFF.PointerToSymbolTable + rawReloc.SymbolTableIndex * 18
                            Case IMAGE_REL_I386_REL32 '32-bit relative displacement
                                value = value + COFF.PointerToSymbolTable + rawReloc.SymbolTableIndex * 18 - (rawReloc.VirtualAddress + .PointerToRawData + 4)
                            'Case IMAGE_REL_I386_SECREL '32-bit offset of the target from the beginning of its section
                            '    Debug.Assert False
                            'Case IMAGE_REL_I386_SECTION '16-bit-section index of the section containing the target
                            '    Debug.Assert False
                        End Select
                        setDwordOffset off, value
                    Next
                    Print #iFileCoff, "---------------------------------------------------------------------"
                End If
        
        '        If retSectionTables(X).NumberOfLinenumbers Then
        '            'les numéros de ligne
        '            COFFNums = GetCOFFLineNumber(retSectionTables(X))
        '            ReDim File.Sections(X).LineNumbers(retSectionTables(X).NumberOfLinenumbers - 1)
        '            For Y = 0 To retSectionTables(X).NumberOfLinenumbers - 1
        '                If COFFNums(Y).Linenumber = 0 Then
        '                    File.Sections(X).LineNumbers(Y).Linenumber = 0
        '                    File.Sections(X).LineNumbers(Y).SymbolTableIndex = COFFNums(Y).Type
        '                    File.Sections(X).LineNumbers(Y).Symbol = File.SymbolTable(COFFNums(Y).Type)
        '                    File.Sections(X).LineNumbers(Y).VirtualAddress = 0
        '                Else
        '                    File.Sections(X).LineNumbers(Y).Linenumber = COFFNums(Y).Linenumber
        '                    File.Sections(X).LineNumbers(Y).VirtualAddress = COFFNums(Y).Type
        '                    File.Sections(X).LineNumbers(Y).SymbolTableIndex = 0
        '                End If
        '            Next
        '        End If
            End With
           ' .value = X
        Next
        Print #iFileCoff, "====================================================================="
    'End With
End If
End Sub

'cette fonction permet d'extraire les informations sur un fichier COFF .obj
'==================================================================
'renvoie une structure COFFFile
'FileName : nom du fichier à analyser
Public Sub ProcessCOFFFile(szOutPattern As String, szFilename As String)
Dim COFF As RawCOFFHeader 'entete brute du fichier
Dim iFileCoff As Integer, iFileIO As Integer

iFileCoff = FreeFile
Open szOutPattern & ".cof" For Output As #iFileCoff
iFileIO = FreeFile
Open szOutPattern & ".sym" For Output As #iFileIO
    'on récupère l'entete
    getUnkOffset 0, VarPtr(COFF), Len(COFF)
    
    'on analyse l'entete
    Print #iFileCoff, "====================================================================="
    Print #iFileCoff, "Header of "; szFilename
    Print #iFileCoff, "====================================================================="
    
    Print #iFileCoff, "Characteristics :", COFF.Characteristics
    Print #iFileCoff, "Machine :", COFF.Machine
    Print #iFileCoff, "Number Of Sections :", COFF.NumberOfSections
    Print #iFileCoff, "Number Of Symbols :", COFF.NumberOfSymbols
    'Print #iFileCoff, "Pointer To Symbol Table :", COFF.PointerToSymbolTable
    Print #iFileCoff, "Size Of OptionalHeader :", COFF.SizeOfOptionalHeader
    Print #iFileCoff, "TimeDate Stamp :", COFF.TimeDateStamp
    
    Print #iFileCoff, "====================================================================="
    
    ProcessOptionalHeader iFileCoff, COFF
 'ici modif
'frmProgress.lblState.Caption = "Traitement des sections..."
DoEvents
    
    ProcessSectionTable iFileCoff, COFF
    
'frmProgress.imSection.Visible = True
'frmProgress.lblState.Caption = "Traitement des symboles..."
DoEvents
    
    ProcessSymbolTable iFileCoff, iFileIO, COFF
    
'frmProgress.imSym1.Visible = True
'frmProgress.lblState.Caption = "Traitement de la table de chaines..."
DoEvents

    ProcessStringTable iFileCoff, COFF

Close #iFileIO
Close #iFileCoff
End Sub

Private Function GetSymbolName(Symbol As RawCOFFSymbol, COFF As RawCOFFHeader) As String
With Symbol
    If (.SymName(0) = 0) And _
       (.SymName(1) = 0) And _
       (.SymName(2) = 0) And _
       (.SymName(3) = 0) Then
        GetSymbolName = GetSymbolLongName(Symbol, COFF)
    Else
        GetSymbolName = StrConv(.SymName, vbUnicode)
    End If
End With
End Function

Private Function GetUndecoratedName(szSymbol As String) As String
Dim buff As String * 255
UnDecorateSymbolName szSymbol, buff, 255, UNDNAME_COMPLETE
GetUndecoratedName = Mid$(buff, 1, InStr(buff, vbNullChar) - 1)
End Function

'renvoie le nom long du symbole
Private Function GetSymbolLongName(Symbol As RawCOFFSymbol, COFF As RawCOFFHeader) As String
If COFF.PointerToSymbolTable Then
    Dim Offset As Long, c As Byte
    
    Offset = Symbol.SymName(4) + Symbol.SymName(5) * &H100& + Symbol.SymName(6) * &H10000 + Symbol.SymName(7) * &H1000000
    setPointerOffset COFF.PointerToSymbolTable + 18 * COFF.NumberOfSymbols + Offset
    
    setMap 5
    c = getByte(0)
    Do While c
        GetSymbolLongName = GetSymbolLongName & Chr$(c)
        c = getByte(0)
    Loop
End If
End Function

''renvoie les numéros de lignes du fichier
'Private Function GetCOFFLineNumber(Section As RawSection) As RawCOFFLineNumber()
'Dim tmp() As RawCOFFLineNumber
'If (Section.NumberOfLinenumbers > 0) And Section.PointerToLinenumbers Then
'    ReDim tmp(Section.NumberOfLinenumbers - 1)
'    Get #1, Section.PointerToLinenumbers + 1, tmp
'    GetCOFFLineNumber = tmp
'End If
'End Function

Public Sub DysCOFF(ByVal strExeName As String, ByVal strOutFilePattern As String, Optional bProcessCall As Boolean = False)
Dim lpEntry As Long, iFileNum As Integer, X As Long, iLog As Integer

'ici modif
'Load frmProgress
'frmProgress.InitCOFF
'frmProgress.Show

'frmProgress.lblFile.Caption = "Filename : " & strExeName
'frmProgress.lblState.Caption = "Chargement..."
DoEvents

Init

Set32BitsDecode

'chargement du fichier
If LoadFile2(strExeName) = 0 Then Exit Sub

'frmProgress.lblState.Caption = "Traitement de l'entête..."
DoEvents

ProcessCOFFFile strOutFilePattern, strExeName

iFileNum = FreeFile
Open strOutFilePattern & ".asm" For Output As #iFileNum
iLog = FreeFile
Open strOutFilePattern & ".log" For Output As #iLog
   ' frmProgress.lblState.Caption = "Traitement des points d'entrée..."
    DoEvents

    If EntryPoints.Count Then
       ' With frmProgress.pbSym
           ' .Min = 0
           ' .Max = EntryPoints.Count
            For X = 1 To EntryPoints.Count
                lpEntry = EntryPoints(X)
                DysCode iFileNum, lpEntry, True, EntryNames(X)
              '  .value = X
            Next
       ' End With
    End If
    
    'frmProgress.imSym2.Visible = True
    DoEvents
    
    'frmProgress.lblState.Caption = "Traitement des données..."
    DoEvents
    
    ProcessData iFileNum, iFileNum, iLog

    'frmProgress.imData.Visible = True
    DoEvents

Close #iLog
Close #iFileNum

'frmProgress.lblState.Caption = "File disassembled in " & Format$(StopTimer, "#.##") & " seconds"

UnloadFile2
End Sub

Public Function DysCOFF2(ByVal strExeName As String, ByVal strOutFilePattern As String, Optional bProcessCall As Boolean = False)
Dim lpEntry As Long, iFileNum As Integer, X As Long, iLog As Integer

'Load frmProgress
'frmProgress.InitCOFF
'frmProgress.Show

'frmProgress.lblFile.Caption = "Filename : " & strExeName
'frmProgress.lblState.Caption = "Chargement..."
DoEvents

Init2

Set32BitsDecode

'frmProgress.lblState.Caption = "Traitement de l'entête..."
DoEvents

ProcessCOFFFile strOutFilePattern, strExeName

iFileNum = FreeFile
Open strOutFilePattern & ".asm" For Output As #iFileNum
iLog = FreeFile
Open strOutFilePattern & ".log" For Output As #iLog
    'frmProgress.lblState.Caption = "Traitement des points d'entrée..."
    DoEvents
    
    If EntryPoints.Count Then
        'With frmProgress.pbSym
            '.Min = 0
           ' .Max = EntryPoints.Count
            For X = 1 To EntryPoints.Count
                lpEntry = EntryPoints(X)
                DysCode iFileNum, lpEntry, True, EntryNames(X)
                '.value = X
            Next
       ' End With
    End If
    
   ' frmProgress.imSym2.Visible = True
    DoEvents
    
   ' frmProgress.lblState.Caption = "Traitement des données..."
    DoEvents
    
    ProcessData iFileNum, iFileNum, iLog

    'frmProgress.imData.Visible = True
    DoEvents

Close #iLog
Close #iFileNum
End Function

