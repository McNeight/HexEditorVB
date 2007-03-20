Attribute VB_Name = "modPE"
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
'pour plus d'infos
'http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx

Private Type LIST_ENTRY
    FLink As Long
    BLink As Long
End Type

Private Type IMAGE_DEBUG_INFORMATION
    List As LIST_ENTRY
    Size As Long
    MappedBase As Long
    Machine As Integer
    Characteristics As Integer
    CheckSum As Long
    ImageBase As Long
    SizeOfImage As Long
    NumberOfSections As Long
    Sections As Long 'PIMAGE_SECTION_HEADER
    ExportedNamesSize As Long
    ExportedNames As Long 'PSTR
    NumberOfFunctionTableEntries As Long
    FunctionTableEntries As Long 'PIMAGE_FUNCTION_ENTRY
    LowestFunctionStartingAddress As Long
    HighestFunctionEndingAddress As Long
    NumberOfFpoTableEntries As Long
    FpoTableEntries As Long 'PFPO_DATA
    SizeOfCoffSymbols As Long
    CoffSymbols As Long 'PIMAGE_COFF_SYMBOLS_HEADER
    SizeOfCodeViewSymbols As Long
    CodeViewSymbols As Long 'PVOID
    ImageFilePath As Long 'PSTR
    ImageFileName As Long 'PSTR
    ReservedDebugFilePath As Long 'PSTR
    ReservedTimeDateStamp As Long
    ReservedRomImage As Long 'BOOL
    ReservedDebugDirectory As Long 'PIMAGE_DEBUG_DIRECTORY
    ReservedNumberOfDebugDirectories As Long
    ReservedOriginalFunctionTableBaseAddress As Long
    Reserved(1 To 2) As Long
End Type

Public Type COFFFileHeader
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type OptionalHeaderStandardFields
    MagicPE As Integer
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

Public Type OptionalHeaderWinNTSpecificFields
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DLLCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
End Type

Private Type PE
    Signature As Long
    rawCOFF As COFFFileHeader
    rawOptional As OptionalHeaderStandardFields
    rawNT As OptionalHeaderWinNTSpecificFields
End Type

'toutes les structures internes d'un exe
Public Type IMAGE_DATA_DIRECTORY
    rva As Long
    Size As Long
End Type

Public Type RawDelayLoadDirectoryEntry
    Attributes As Long
    Name As Long
    ModuleHandle As Long
    DelayImportAddressTable As Long
    DelayImportNameTable As Long
    BoundDelayImportTable As Long
    UnloadDelayImportTable As Long
    TimeStamp As Long
End Type

Public Type RawExportDirectory
    ExportFlags As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    NameRVA As Long
    OrdinalBase As Long
    AddressTableEntries As Long
    NumberofNamePointers As Long
    ExportAddressTableRVA As Long
    NamePointerRVA As Long
    OrdinalTableRVA As Long
End Type

Public Type RawImportDirectoryEntry
    ImportLookupTableRVA As Long
    TimeDateStamp As Long
    FowarderChain As Long
    NameRVA As Long
    ImportAddressTableRVA As Long
End Type

Public Type RawCOFFHeader
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

''pour plus d'infos
''http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx
'
''type de machine cible de l'exe
'Public Enum Machine
'IMAGE_FILE_MACHINE_UNKNOWN = &H0
'IMAGE_FILE_MACHINE_ALPHA = &H184
'IMAGE_FILE_MACHINE_ARM = &H1C0
'IMAGE_FILE_MACHINE_ALPHA64 = &H284
'IMAGE_FILE_MACHINE_I386 = &H14C
'IMAGE_FILE_MACHINE_IA64 = &H200
'IMAGE_FILE_MACHINE_M68K = &H268
'IMAGE_FILE_MACHINE_MIPS16 = &H266
'IMAGE_FILE_MACHINE_MIPSFPU = &H366
'IMAGE_FILE_MACHINE_MIPSFPU16 = &H466
'IMAGE_FILE_MACHINE_POWERPC = &H1F0
'IMAGE_FILE_MACHINE_R3000 = &H162
'IMAGE_FILE_MACHINE_R4000 = &H166
'IMAGE_FILE_MACHINE_R10000 = &H168
'IMAGE_FILE_MACHINE_SH3 = &H1A2
'IMAGE_FILE_MACHINE_SH4 = &H1A6
'IMAGE_FILE_MACHINE_THUMB = &H1C2
'End Enum
'
''caractéristique du fichier exe
'Public Enum Characteristic
'IMAGE_FILE_RELOCS_STRIPPED = &H1
'IMAGE_FILE_EXECUTABLE_IMAGE = &H2
'IMAGE_FILE_LINE_NUMS_STRIPPED = &H4
'IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8
'IMAGE_FILE_AGGRESSIVE_WS_TRIM = &H10
'IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20
'IMAGE_FILE_16BIT_MACHINE = &H40
'IMAGE_FILE_BYTES_REVERSED_LO = &H80
'IMAGE_FILE_32BIT_MACHINE = &H100
'IMAGE_FILE_DEBUG_STRIPPED = &H200
'IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400
'IMAGE_FILE_SYSTEM = &H1000
'IMAGE_FILE_DLL = &H2000
'IMAGE_FILE_UP_SYSTEM_ONLY = &H4000
'IMAGE_FILE_BYTES_REVERSED_HI = &H8000
'End Enum
'
''type de sous système NT attendu par l'exe
'Public Enum NTSubSystem
'IMAGE_SUBSYSTEM_UNKNOWN = 0
'IMAGE_SUBSYSTEM_NATIVE = 1
'IMAGE_SUBSYSTEM_WINDOWS_GUI = 2
'IMAGE_SUBSYSTEM_WINDOWS_CUI = 3
'IMAGE_SUBSYSTEM_POSIX_CUI = 7
'IMAGE_SUBSYSTEM_WINDOWS_CE_GUI = 9
'IMAGE_SUBSYSTEM_EFI_APPLICATION = 10
'IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER = 11
'IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER = 12
'End Enum
'
''caractéristique d'une Dll
'Public Enum DLLCharacteristic
'IMAGE_DLLCHARACTERISTICS_NO_BIND = &H800
'IMAGE_DLLCHARACTERISTICS_WDM_DRIVER = &H2000
'IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE = &H8000
'End Enum
'
''caractéristique d'une section
'Public Enum SectionCharacteristics
'IMAGE_SCN_TYPE_REG = &H0
'IMAGE_SCN_TYPE_DSECT = &H1
'IMAGE_SCN_TYPE_NOLOAD = &H2
'IMAGE_SCN_TYPE_GROUP = &H4
'IMAGE_SCN_TYPE_NO_PAD = &H8
'IMAGE_SCN_TYPE_COPY = &H10
'IMAGE_SCN_CNT_CODE = &H20
'IMAGE_SCN_CNT_INITIALIZED_DATA = &H40
'IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80
'IMAGE_SCN_LNK_OTHER = &H100
'IMAGE_SCN_LNK_INFO = &H200
'IMAGE_SCN_TYPE_OVER = &H400
'IMAGE_SCN_LNK_REMOVE = &H800
'IMAGE_SCN_LNK_COMDAT = &H1000
'IMAGE_SCN_MEM_FARDATA = &H8000
'IMAGE_SCN_MEM_PURGEABLE = &H20000
'IMAGE_SCN_MEM_16BIT = &H20000
'IMAGE_SCN_MEM_LOCKED = &H40000
'IMAGE_SCN_MEM_PRELOAD = &H80000
'IMAGE_SCN_ALIGN_1BYTES = &H100000
'IMAGE_SCN_ALIGN_2BYTES = &H200000
'IMAGE_SCN_ALIGN_4BYTES = &H300000
'IMAGE_SCN_ALIGN_8BYTES = &H400000
'IMAGE_SCN_ALIGN_16BYTES = &H500000
'IMAGE_SCN_ALIGN_32BYTES = &H600000
'IMAGE_SCN_ALIGN_64BYTES = &H700000
'IMAGE_SCN_ALIGN_128BYTES = &H800000
'IMAGE_SCN_ALIGN_256BYTES = &H900000
'IMAGE_SCN_ALIGN_512BYTES = &HA00000
'IMAGE_SCN_ALIGN_1024BYTES = &HB00000
'IMAGE_SCN_ALIGN_2048BYTES = &HC00000
'IMAGE_SCN_ALIGN_4096BYTES = &HD00000
'IMAGE_SCN_ALIGN_8192BYTES = &HE00000
'IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000
'IMAGE_SCN_MEM_DISCARDABLE = &H2000000
'IMAGE_SCN_MEM_NOT_CACHED = &H4000000
'IMAGE_SCN_MEM_NOT_PAGED = &H8000000
'IMAGE_SCN_MEM_SHARED = &H10000000
'IMAGE_SCN_MEM_EXECUTE = &H20000000
'IMAGE_SCN_MEM_READ = &H40000000
'IMAGE_SCN_MEM_WRITE = &H80000000
'End Enum
'
'Public Enum SectionNumberValues
'IMAGE_SYM_UNDEFINED = 0
'IMAGE_SYM_ABSOLUTE = -1
'IMAGE_SYM_DEBUG = -2
'End Enum
'
'Public Enum SymBaseType
'IMAGE_SYM_TYPE_NULL = 0
'IMAGE_SYM_TYPE_VOID = 1
'IMAGE_SYM_TYPE_CHAR = 2
'IMAGE_SYM_TYPE_SHORT = 3
'IMAGE_SYM_TYPE_INT = 4
'IMAGE_SYM_TYPE_LONG = 5
'IMAGE_SYM_TYPE_FLOAT = 6
'IMAGE_SYM_TYPE_DOUBLE = 7
'IMAGE_SYM_TYPE_STRUCT = 8
'IMAGE_SYM_TYPE_UNION = 9
'IMAGE_SYM_TYPE_ENUM = 10
'IMAGE_SYM_TYPE_MOE = 11
'IMAGE_SYM_TYPE_BYTE = 12
'IMAGE_SYM_TYPE_WORD = 13
'IMAGE_SYM_TYPE_UINT = 14
'IMAGE_SYM_TYPE_DWORD = 15
'End Enum
'
'Public Enum SymComplexType
'IMAGE_SYM_DTYPE_NULL = 0
'IMAGE_SYM_DTYPE_POINTER = 1
'IMAGE_SYM_DTYPE_FUNCTION = 2
'IMAGE_SYM_DTYPE_ARRAY = 3
'End Enum
'
'Public Enum StorageClass
'IMAGE_SYM_CLASS_END_OF_FUNCTION = -1
'IMAGE_SYM_CLASS_NULL = 0
'IMAGE_SYM_CLASS_AUTOMATIC = 1
'IMAGE_SYM_CLASS_EXTERNAL = 2
'IMAGE_SYM_CLASS_STATIC = 3
'IMAGE_SYM_CLASS_REGISTER = 4
'IMAGE_SYM_CLASS_EXTERNAL_DEF = 5
'IMAGE_SYM_CLASS_LABEL = 6
'IMAGE_SYM_CLASS_UNDEFINED_LABEL = 7
'IMAGE_SYM_CLASS_MEMBER_OF_STRUCT = 8
'IMAGE_SYM_CLASS_ARGUMENT = 9
'IMAGE_SYM_CLASS_STRUCT_TAG = 10
'IMAGE_SYM_CLASS_MEMBER_OF_UNION = 11
'IMAGE_SYM_CLASS_UNION_TAG = 12
'IMAGE_SYM_CLASS_TYPE_DEFINITION = 13
'IMAGE_SYM_CLASS_UNDEFINED_STATIC = 14
'IMAGE_SYM_CLASS_ENUM_TAG = 15
'IMAGE_SYM_CLASS_MEMBER_OF_ENUM = 16
'IMAGE_SYM_CLASS_REGISTER_PARAM = 17
'IMAGE_SYM_CLASS_BIT_FIELD = 18
'IMAGE_SYM_CLASS_BLOCK = 100
'IMAGE_SYM_CLASS_FUNCTION = 101
'IMAGE_SYM_CLASS_END_OF_STRUCT = 102
'IMAGE_SYM_CLASS_FILE = 103
'IMAGE_SYM_CLASS_SECTION = 104
'IMAGE_SYM_CLASS_WEAK_EXTERNAL = 105
'End Enum
'
'Public Enum SelectionCOMDAT
'IMAGE_COMDAT_SELECT_NODUPLICATES = 1
'IMAGE_COMDAT_SELECT_ANY = 2
'IMAGE_COMDAT_SELECT_SAME_SIZE = 3
'IMAGE_COMDAT_SELECT_EXACT_MATCH = 4
'IMAGE_COMDAT_SELECT_ASSOCIATIVE = 5
'IMAGE_COMDAT_SELECT_LARGEST = 6
'End Enum
'
'Public Enum FixupType
'IMAGE_REL_BASED_ABSOLUTE = 0
'IMAGE_REL_BASED_HIGH = 1
'IMAGE_REL_BASED_LOW = 2
'IMAGE_REL_BASED_HIGHLOW = 3
'IMAGE_REL_BASED_HIGHADJ = 4
'IMAGE_REL_BASED_MIPS_JMPADDR = 5
'IMAGE_REL_BASED_SECTION = 6
'IMAGE_REL_BASED_REL32 = 7
'IMAGE_REL_BASED_MIPS_JMPADDR16 = 9
'IMAGE_REL_BASED_DIR64 = 10
'IMAGE_REL_BASED_HIGH3ADJ = 11
'End Enum
'
'Public Enum DebugType
'IMAGE_DEBUG_TYPE_UNKNOWN = 0
'IMAGE_DEBUG_TYPE_COFF = 1
'IMAGE_DEBUG_TYPE_CODEVIEW = 2
'IMAGE_DEBUG_TYPE_FPO = 3
'IMAGE_DEBUG_TYPE_MISC = 4
'IMAGE_DEBUG_TYPE_EXCEPTION = 5
'IMAGE_DEBUG_TYPE_FIXUP = 6
'IMAGE_DEBUG_TYPE_OMAP_TO_SRC = 7
'IMAGE_DEBUG_TYPE_OMAP_FROM_SRC = 8
'IMAGE_DEBUG_TYPE_BORLAND = 9
'End Enum

Private Enum SymTagEnum
    SymTagFunction = 5
    SymTagData = 7
    SymTagPublicSymbol = 10
    SymTagUDT = 11
    SymTagEnum = 12
    SymTagTypedef = 17
End Enum

Private Enum BasicType
    btNoType = 0
    btVoid = 1
    btChar = 2
    btWChar = 3
    btInt = 6
    btUInt = 7
    btFloat = 8
    btBCD = 9
    btBool = 10
    btLong = 13
    btULong = 14
    btCurrency = 25
    btDate = 26
    btVariant = 27
    btComplex = 28
    btBit = 29
    btBSTR = 30
    btHresult = 31
End Enum
Private Type SYMBOL_INFO
    SizeOfStruct As Long
    TypeIndex As Long
    Reserved0 As Long
    Reserved1 As Long
    Reserved2 As Long
    Reserved3 As Long
    Index As Long
    Size As Long
    ModBaseLo As Long
    ModBaseHi As Long
    Flags As Long
    
    padding As Long
    
    ValueLo As Long
    ValueHi As Long
    AddressLo As Long
    AddressHi As Long
    Register As Long
    Scope As Long
    Tag As Long
    NameLen As Long
    MaxNameLen As Long
    Name As Byte
End Type

Private Declare Function SymInitialize Lib "dbghelp.dll" (ByVal hProcess As Long, ByVal UserSearchPath As String, ByVal fInvadeProcess As Long) As Long
Private Declare Function SymLoadModule Lib "dbghelp.dll" (ByVal hProcess As Long, ByVal hFile As Long, ByVal ImageName As String, ByVal ModuleName As Long, ByVal BaseOfDll As Long, ByVal SizeOfDll As Long) As Long
'Private Declare Function SymEnumerateSymbolsW Lib "dbghelp.dll" (ByVal hProcess As Long, ByVal BaseOfDll As Long, ByVal EnumSymbolsCallback As Long, ByVal UserContext As Long) As Long
Private Declare Function SymEnumSymbols Lib "dbghelp.dll" (ByVal hProcess As Long, ByVal BaseOfDllLo As Long, ByVal BaseOfDllHi As Long, ByVal Mask As String, ByVal EnumSymbolsCallback As Long, ByVal UserContext As Long) As Long
Private Declare Function SysAllocString Lib "oleaut32.dll" (ByRef pOlechar As Byte) As String
'Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal lpSZ As Long, ByVal cbLen As Long) As String
Private Declare Sub SysFreeString Lib "oleaut32.dll" (ByRef bstr As String)
Private Declare Function SymUnloadModule Lib "dbghelp.dll" (ByVal hProcess As Long, ByVal BaseOfDll As Long) As Long
Private Declare Function SymCleanup Lib "dbghelp.dll" (ByVal hProcess As Long) As Long

Private Const TI_GET_CHILDRENCOUNT As Long = 13&
Private Const TI_GET_LENGTH As Long = 2&
Private Const TI_GET_COUNT As Long = 12&
Private Const TI_GET_BASETYPE As Long = 5&
Private Declare Function SymGetTypeInfo Lib "dbghelp.dll" (ByVal hProcess As Long, ByVal ModBaseLo As Long, ByVal ModBaseHi As Long, ByVal TypeId As Long, ByVal GetType As Long, ByRef pInfo As Any) As Long

Private Const SYMOPT_UNDNAME As Long = &H2
Private Declare Function SymGetOptions Lib "dbghelp.dll" () As Long
Private Declare Function SymSetOptions Lib "dbghelp.dll" (ByVal SymOptions As Long) As Long

Private Sub ProcessImports(szFilePattern As String, lpImportTable As IMAGE_DATA_DIRECTORY, lpDelayImportDescriptor As IMAGE_DATA_DIRECTORY)
Dim iFileImp As Integer, X As Long
Dim strDllName As String, strName As String
Dim lngOrdinal As Long, lngHint As Long, lngAddress As Long
Dim c As Byte, ILTEntry As Long, Address As Long, oldRVA As Long
    Dim off As Long, rawImp As RawImportDirectoryEntry

iFileImp = FreeFile
Open szFilePattern & ".imp" For Output As #iFileImp
    
    's'il y a un dossier Delay Import
    If lpDelayImportDescriptor.rva Then
        'infos Delay Import brutes
        Dim Delay As RawDelayLoadDirectoryEntry
        
        off = RVA2Offset(lpDelayImportDescriptor.rva)
        Do
            getUnkOffset off, VarPtr(Delay), Len(Delay)
            With Delay
                If (.Attributes = 0) And _
                   (.BoundDelayImportTable = 0) And _
                   (.DelayImportAddressTable = 0) And _
                   (.DelayImportNameTable = 0) And _
                   (.ModuleHandle = 0) And _
                   (.Name = 0) And _
                   (.TimeStamp = 0) And _
                   (.UnloadDelayImportTable = 0) Then
                   Exit Do
                Else
                    setMapRVA .ModuleHandle, 32
                    
                    'le nom de dll
                    strDllName = vbNullString
                    setPointerRVA .Name
                    setMap 5
                    c = getByte(0)
                    Do While c
                        strDllName = strDllName & Chr$(c)
                        c = getByte(0)
                    Loop
                    
                    Print #iFileImp, ";=================================================================="
                    Print #iFileImp, ";Delay Import From "; strDllName
                    Print #iFileImp, ";=================================================================="
                    
                    Print #iFileImp, ";Attributes :", .Attributes
                    Print #iFileImp, ";Module Handle RVA :", getNumber(.ModuleHandle, 8)
                    Print #iFileImp, ";TimeStamp :", .TimeStamp
                    Print #iFileImp, ";------------------------------------------------------------------"
                    
                    'si la table de nom est présente
                    If .DelayImportNameTable Then
                        'la table de nom
                        setPointerRVA .DelayImportNameTable
                    'sinon, on regarde dans la table d'adresse qui doit être présente
                    Else
                        'la table d'adresse
                        setPointerRVA .DelayImportAddressTable
                    End If
                    
                    Do
                        'l'adresse dans tous les cas
                        lngAddress = getPointerVA
                        ILTEntry = getDword(0)
                        If ILTEntry = 0 Then
                            Exit Do
                        Else
                            'si l'import est par ordinal
                            If (ILTEntry And &H80000000) = &H80000000 Then
                                lngOrdinal = (ILTEntry And &H7FFFFFFF)
                                strName = "_dimp_Ordinal_" & getNumber(ILTEntry And &H7FFFFFFF, 8)
                                
                                Print #iFileImp, ";Imported "; strName; " by Ordinal (0x"; Hex$(lngOrdinal); ") at address "; getNumber(lngAddress, 8)
                            'sinon par nom
                            Else
                                Address = ILTEntry And &H7FFFFFFF
                                
                                oldRVA = setPointerRVA(Address + 2)
                                strName = vbNullString
                                setMap 5
                                c = getByte(0)
                                Do While c
                                    strName = strName & Chr$(c)
                                    c = getByte(0)
                                Loop
                                setPointerRVA oldRVA
                                
                                lngHint = getWordRVA(Address)
                                
                                Print #iFileImp, ";Imported "; strName; " by Name (Hint 0x"; Hex$(lngHint); ") at address "; getNumber(lngAddress, 8)
                            End If
                            
                            AddSubName lngAddress, strName
                        End If
                    Loop
                    Print #iFileImp, ";=================================================================="
                End If
            End With
            off = off + Len(Delay)
        Loop
    End If
    
    's'il y a un dossier Import
    If lpImportTable.rva Then
        'on passe au dossier d'import
        off = RVA2Offset(lpImportTable.rva)
        Do
            'on récupère les infos sur chaque Dll importée
            getUnkOffset off, VarPtr(rawImp), Len(rawImp)
            With rawImp
                If (.FowarderChain = 0) And _
                   (.ImportAddressTableRVA = 0) And _
                   (.ImportLookupTableRVA = 0) And _
                   (.NameRVA = 0) And _
                   (.TimeDateStamp = 0) Then
                    Exit Do
                Else
                    'le nom de dll
                    strDllName = vbNullString
                    setPointerRVA .NameRVA
                    setMap 5
                    c = getByte(0)
                    Do While c
                        strDllName = strDllName & Chr$(c)
                        c = getByte(0)
                    Loop
                    
                    Print #iFileImp, ";=================================================================="
                    Print #iFileImp, ";Import From "; strDllName
                    Print #iFileImp, ";=================================================================="
                    
                    Print #iFileImp, ";Fowarder Chain :", .FowarderChain
                    Print #iFileImp, ";Time Date Stamp :", .TimeDateStamp
                    Print #iFileImp, ";------------------------------------------------------------------"
                        
                    'si la table de nom est présente
                    If .ImportLookupTableRVA Then
                        'la table de nom
                        setPointerRVA .ImportLookupTableRVA
                    'sinon, on regarde dans la table d'adresse qui doit être présente
                    Else
                        'la table d'adresse
                        setPointerRVA .ImportAddressTableRVA
                    End If
                    Do
                        'et l'adresse dans tous les cas
                        lngAddress = getPointerVA
                        ILTEntry = getDword(0)
                        If ILTEntry = 0 Then
                            Exit Do
                        Else
                            'si l'import est par ordinal
                            If (ILTEntry And &H80000000) = &H80000000 Then
                                lngOrdinal = (ILTEntry And &H7FFFFFFF)
                                strName = "_imp_Ordinal_" & getNumber(ILTEntry And &H7FFFFFFF, 8)
                            
                                Print #iFileImp, ";Imported "; strName; " by Ordinal (0x"; Hex$(lngOrdinal); ") at address "; getNumber(lngAddress, 8)
                            'sinon par nom
                            Else
                                Address = ILTEntry And &H7FFFFFFF
                                oldRVA = setPointerRVA(Address + 2)
                                strName = vbNullString
                                setMap 5
                                c = getByte(0)
                                Do While c
                                    strName = strName & Chr$(c)
                                    c = getByte(0)
                                Loop
                                setPointerRVA oldRVA
                                
                                lngHint = getWordRVA(Address)
                                
                                Print #iFileImp, ";Imported "; strName; " by Name (Hint 0x"; Hex$(lngHint); ") at address "; getNumber(lngAddress, 8)
                            End If
            
                            AddSubName lngAddress, strName
                        End If
                    Loop
                    Print #iFileImp, ";=================================================================="
                End If
            End With
            off = off + Len(rawImp)
        Loop
    End If
Close #iFileImp
End Sub

Private Sub ProcessExports(szFilePattern As String, lpExportTable As IMAGE_DATA_DIRECTORY)
Dim iFileExp As Integer, ExpDir As RawExportDirectory, X As Long, c As Byte
Dim ENPTEntry As Long, EOTEntry As Integer, EATEntry As Long
Dim strName As String, lngAddress As Long, lngOrdinal As Long, strDllName As String, strForwarderName As String

iFileExp = FreeFile
Open szFilePattern & ".exp" For Output As #iFileExp
    'si le fichier contient un dossier Export
    If lpExportTable.rva Then
        'le dossier lui-même
        getUnkRVA lpExportTable.rva, VarPtr(ExpDir), Len(ExpDir)
        
        With ExpDir
            'le nom de la dll
            setPointerRVA .NameRVA
            setMap 5
            c = getByte(0)
            Do While c
                strDllName = strDllName & Chr$(c)
                c = getByte(0)
            Loop
            
            Print #iFileExp, ";=================================================================="
            Print #iFileExp, ";Export From "; strDllName
            Print #iFileExp, ";=================================================================="
            
            Print #iFileExp, ";Export Flags :", .ExportFlags
            Print #iFileExp, ";Version :", .MajorVersion; "."; .MinorVersion
            Print #iFileExp, ";Number of Exports :", .NumberofNamePointers
            Print #iFileExp, ";Ordinal Base :", .OrdinalBase
            Print #iFileExp, ";TimeDateStamp :", .TimeDateStamp
            Print #iFileExp, ";------------------------------------------------------------------"
            
            If .NamePointerRVA And .OrdinalTableRVA Then
                'pour chaque export
                For X = 0 To .NumberofNamePointers - 1
                    EOTEntry = getWordRVA(.OrdinalTableRVA + X * 2)
                    ENPTEntry = getDwordRVA(.NamePointerRVA + X * 4)
                    
                    'on récupère le nom
                    setPointerRVA ENPTEntry
                    strName = vbNullString
                    setMap 5
                    c = getByte(0)
                    Do While c
                        strName = strName & Chr$(c)
                        c = getByte(0)
                    Loop
                
                    EATEntry = getDwordRVA(.ExportAddressTableRVA + EOTEntry * 4)
                    setMapRVA .ExportAddressTableRVA + EOTEntry * 4, 32
                    
                    'l'adresse
                    lngAddress = EATEntry
                    'le numéro d'ordre
                    lngOrdinal = EOTEntry + .OrdinalBase
                    
                    If lngAddress Then
                        If (lngAddress <= lpExportTable.rva) Or _
                           (lngAddress >= (lpExportTable.rva + lpExportTable.Size)) Then
                            'ceci est un export
                            Print #iFileExp, ";Exported "; strName; " by Name (Ordinal 0x"; Hex$(lngOrdinal); ") at address "; getNumber(lngAddress + dwImageBase, 8)
                            AddExport strName, lngAddress
                        Else
                            'ceci est un Forwarder
                            strForwarderName = vbNullString
                            setPointerRVA lngAddress
                            setMap 5
                            c = getByte(0)
                            Do While c
                                strForwarderName = strForwarderName & Chr$(c)
                                c = getByte(0)
                            Loop
                            
                            Print #iFileExp, ";Forwarder "; strName; "(Ordinal 0x"; Hex$(lngOrdinal); ") link to "; strForwarderName; " at address "; getNumber(lngAddress + dwImageBase, 8)
                        End If
                        AddSubName lngAddress, strName
                    End If
                Next
                'pour chaque export
                For X = 0 To .AddressTableEntries - 1
                    If getMapRVA(.ExportAddressTableRVA + X * 4) = 0 Then
                        EATEntry = getDwordRVA(.ExportAddressTableRVA + X * 4)
                        'l'adresse
                        lngAddress = EATEntry
                        'le numéro d'ordre
                        lngOrdinal = X + .OrdinalBase
            
                        strName = "_exp_Ordinal_" & Hex$(lngOrdinal)
                        
                        If lngAddress Then
                            If (lngAddress <= lpExportTable.rva) Or _
                               (lngAddress >= (lpExportTable.rva + lpExportTable.Size)) Then
                                'ceci est un export
                                Print #iFileExp, ";Exported "; strName; " by Ordinal (Ordinal 0x"; Hex$(lngOrdinal); ") at address "; getNumber(lngAddress + dwImageBase, 8)
                                AddExport strName, lngAddress
                            Else
                                'ceci est un Forwarder
                                strForwarderName = vbNullString
                                setPointerRVA lngAddress
                                setMap 5
                                c = getByte(0)
                                Do While c
                                    strForwarderName = strForwarderName & Chr$(c)
                                    c = getByte(0)
                                Loop
                                
                                Print #iFileExp, ";Forwarder "; strName; "(Ordinal 0x"; Hex$(lngOrdinal); ") link to "; strForwarderName; " at address "; getNumber(lngAddress + dwImageBase, 8)
                            End If
                            
                            AddSubName lngAddress, strName
                        End If
                    Else
                        setMapRVA .ExportAddressTableRVA + X * 4, 0
                    End If
                Next
            Else
                'pour chaque export
                For X = 0 To .AddressTableEntries - 1
                    EATEntry = getDwordRVA(.ExportAddressTableRVA + X * 4)
                    'l'adresse
                    lngAddress = EATEntry
                    'le numéro d'ordre
                    lngOrdinal = X + .OrdinalBase
        
                    strName = "_exp_Ordinal_" & Hex$(lngOrdinal)
                    
                    If lngAddress Then
                        If (lngAddress <= lpExportTable.rva) Or _
                           (lngAddress >= (lpExportTable.rva + lpExportTable.Size)) Then
                            'ceci est un export
                            Print #iFileExp, ";Exported "; strName; " by Ordinal (Ordinal 0x"; Hex$(lngOrdinal); ") at address "; getNumber(lngAddress + dwImageBase, 8)
                            AddExport strName, lngAddress
                        Else
                            'ceci est un Forwarder
                            strForwarderName = vbNullString
                            setPointerRVA lngAddress
                            setMap 5
                            c = getByte(0)
                            Do While c
                                strForwarderName = strForwarderName & Chr$(c)
                                c = getByte(0)
                            Loop
                            
                            Print #iFileExp, ";Forwarder "; strName; "(Ordinal 0x"; Hex$(lngOrdinal); ") link to "; strForwarderName; " at address "; getNumber(lngAddress + dwImageBase, 8)
                        End If
                        
                        AddSubName lngAddress, strName
                    End If
                Next
            End If
            Print #iFileExp, ";=================================================================="
        End With
    End If
Close #iFileExp
End Sub

'renvoie des informations sur un fichier exe, dll, ocx
'=====================================================
Public Function ProcessPE(szFilePattern As String, szFilename As String) As Long
Dim iFilePE As Integer
Dim Offset As Long, X As Long
'les entêtes
Dim rawPE As PE
'les "dossiers de données"
Dim retDataDirectory() As IMAGE_DATA_DIRECTORY

'on récupère l'offset de l'entête PE à l'offset 0x3C
Offset = getDwordOffset(&H3C)

setPointerOffset Offset

'on récupère les entêtes PE
getUnk ByVal VarPtr(rawPE), Len(rawPE)

'et la place pour les dossiers
ReDim retDataDirectory(rawPE.rawNT.NumberOfRvaAndSizes - 1)
'les dossiers
getUnk ByVal VarPtr(retDataDirectory(0)), rawPE.rawNT.NumberOfRvaAndSizes * 8
    
'la place pour les sections
ReDim retSectionTables(rawPE.rawCOFF.NumberOfSections - 1)
'les sections
getUnk ByVal VarPtr(retSectionTables(0)), rawPE.rawCOFF.NumberOfSections * 40

iFilePE = FreeFile
Open szFilePattern & ".pe" For Output As #iFilePE
    dwImageBase = rawPE.rawNT.ImageBase
        
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Image Information : "; szFilename
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Characteristics :", rawPE.rawCOFF.Characteristics
    Print #iFilePE, ";DLL Characteristics :", rawPE.rawNT.DLLCharacteristics
    Print #iFilePE, ";Machine :", , rawPE.rawCOFF.Machine
    Print #iFilePE, ";Image Version :", rawPE.rawNT.MajorImageVersion; "."; rawPE.rawNT.MinorImageVersion
    Print #iFilePE, ";TimeDate Stamp :", rawPE.rawCOFF.TimeDateStamp
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Entry Point"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Address of Entry Point :", getNumber(rawPE.rawOptional.AddressOfEntryPoint + dwImageBase, 8)
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Image in Memory"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Image Base :", , getNumber(rawPE.rawNT.ImageBase, 8)
    Print #iFilePE, ";Size Of Image :", rawPE.rawNT.SizeOfImage
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Image Code"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Base of Code :", getNumber(rawPE.rawOptional.BaseOfCode + dwImageBase, 8)
    Print #iFilePE, ";Size of Code :", rawPE.rawOptional.SizeOfCode; " byte(s)"
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Image Data"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Base of Data :", getNumber(rawPE.rawOptional.BaseOfData + dwImageBase, 8)
    Print #iFilePE, ";Size of Initialized Data :", rawPE.rawOptional.SizeOfInitializedData; " byte(s)"
    Print #iFilePE, ";Size of Uninitialized Data :"; rawPE.rawOptional.SizeOfUninitializedData; " byte(s)"
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Linker Information"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Linker Version :", rawPE.rawOptional.MajorLinkerVersion; "."; rawPE.rawOptional.MinorLinkerVersion
    Print #iFilePE, ";CheckSum :", , getNumber(rawPE.rawNT.CheckSum, 8)
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";SO Information"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Operating System Version :", rawPE.rawNT.MajorOperatingSystemVersion; "."; rawPE.rawNT.MinorOperatingSystemVersion
    Print #iFilePE, ";Subsystem :", , rawPE.rawNT.Subsystem
    Print #iFilePE, ";Subsystem Version :", rawPE.rawNT.MajorSubsystemVersion; "."; rawPE.rawNT.MinorSubsystemVersion
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Image Alignment"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";File Alignment :", rawPE.rawNT.FileAlignment; " byte(s) boundary"
    Print #iFilePE, ";Section Alignment :", rawPE.rawNT.SectionAlignment; " byte(s) boundary"
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Default Heap"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Size Of HeapCommit :", rawPE.rawNT.SizeOfHeapCommit; " byte(s)"
    Print #iFilePE, ";Size Of HeapReserve :", rawPE.rawNT.SizeOfHeapReserve; " byte(s)"
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Default Stack"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Size Of StackCommit :", rawPE.rawNT.SizeOfStackCommit; " byte(s)"
    Print #iFilePE, ";Size Of StackReserve :", rawPE.rawNT.SizeOfStackReserve; " byte(s)"
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Other Information"
    Print #iFilePE, ";=================================================================="
    
    Print #iFilePE, ";Loader Flags :", rawPE.rawNT.LoaderFlags
    Print #iFilePE, ";Number Of Rva And Sizes :", rawPE.rawNT.NumberOfRvaAndSizes
    Print #iFilePE, ";Size Of Headers :", rawPE.rawNT.SizeOfHeaders; " byte(s)"
    Print #iFilePE, ";Number of Symbols :", rawPE.rawCOFF.NumberOfSymbols
    
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, ";Sections"
    Print #iFilePE, ";=================================================================="
    Print #iFilePE, "Number of Sections :", rawPE.rawCOFF.NumberOfSections
    Print #iFilePE,
    
    'pour chaque section
    For X = 0 To rawPE.rawCOFF.NumberOfSections - 1
        With retSectionTables(X)
            .VirtualAddress = .VirtualAddress + dwImageBase
            
            Print #iFilePE, "Name :", , StrConv(.SecName, vbUnicode)
            Print #iFilePE, "Characteristics :", .Characteristics
            
            Print #iFilePE, "PointerToRawData :", getNumber(.PointerToRawData, 8)
            Print #iFilePE, "SizeOfRawData :", .SizeOfRawData; " byte(s)"
            
            Print #iFilePE, "VirtualAddress :", getNumber(.VirtualAddress, 8)
            Print #iFilePE, "VirtualSize :", , .VirtualSize; " byte(s)"
            
            Print #iFilePE, "PointerToLinenumbers :", getNumber(.PointerToLinenumbers, 8)
            Print #iFilePE, "NumberOfLinenumbers :", .NumberOfLinenumbers
            
            Print #iFilePE, "PointerToRelocations :", getNumber(.PointerToRelocations, 8)
            Print #iFilePE, "NumberOfRelocations :", .NumberOfRelocations
            
            Print #iFilePE, ";------------------------------------------------------------------"
        End With
    Next
    
    Dim lpExportTable As IMAGE_DATA_DIRECTORY
    Dim lpImportTable As IMAGE_DATA_DIRECTORY
    Dim lpDelayImportDescriptor As IMAGE_DATA_DIRECTORY
    
    'on recopie les infos sur les dossiers standards de données
    If rawPE.rawNT.NumberOfRvaAndSizes >= 1 Then lpExportTable = retDataDirectory(0)
    If rawPE.rawNT.NumberOfRvaAndSizes >= 2 Then lpImportTable = retDataDirectory(1)
'    File.StdDataDirectories.ResourceTable = retDataDirectory(2)
'    File.StdDataDirectories.ExceptionTable = retDataDirectory(3)
'    File.StdDataDirectories.CertificateTable = retDataDirectory(4)
'    File.StdDataDirectories.BaseRelocationTable = retDataDirectory(5)
'    File.StdDataDirectories.Debug = retDataDirectory(6)
'    File.StdDataDirectories.Architecture = retDataDirectory(7)
'    File.StdDataDirectories.GlobalPtr = retDataDirectory(8)
'    File.StdDataDirectories.TLSTable = retDataDirectory(9)
'    File.StdDataDirectories.LoadConfigTable = retDataDirectory(10)
'    File.StdDataDirectories.BoundImport = retDataDirectory(11)
'    If rawPE.rawNT.NumberOfRvaAndSizes >= 13 Then lpIAT = retDataDirectory(12)
    If rawPE.rawNT.NumberOfRvaAndSizes >= 14 Then lpDelayImportDescriptor = retDataDirectory(13)
'    File.StdDataDirectories.COMPlusRuntimeHeader = retDataDirectory(14)

Close #iFilePE

ProcessImports szFilePattern, lpImportTable, lpDelayImportDescriptor

ProcessExports szFilePattern, lpExportTable

'on renvoie l'adresse du point d'entrée
ProcessPE = dwImageBase + rawPE.rawOptional.AddressOfEntryPoint
End Function

'Public Sub ProcessOffsets()
'    Dim X As Long, addrfin As Long, addr As Long, dw As Long
'
'    For X = 0 To UBound(retSectionTables)
'        With retSectionTables(X)
'            If ((.Characteristics And IMAGE_SCN_CNT_CODE) = IMAGE_SCN_CNT_CODE) And _
'               ((.Characteristics And IMAGE_SCN_MEM_EXECUTE) = IMAGE_SCN_MEM_EXECUTE) Then
'                addr = .PointerToRawData
'                addrfin = .PointerToRawData + .VirtualSize
'                Do While addr <= addrfin
'                    dw = getDwordOffset(addr)
'                    If CheckVA(dw) Then
'                        If getMapOffset(addr) = 0 Then
'                            setMapOffset addr, 4
'                            tryCallCol.Add addr
'                        End If
'                        addr = addr + 4
'                    Else
'                        addr = addr + 1
'                    End If
'                Loop
'            End If
'        End With
'    Next
'End Sub

Private Function GetChildrenCount(ByVal ModBaseLo As Long, ByVal TypeIndex As Long) As Long
    SymGetTypeInfo 0, ModBaseLo, 0, TypeIndex, TI_GET_CHILDRENCOUNT, GetChildrenCount
End Function

Private Function GetCount(ByVal ModBaseLo As Long, ByVal TypeIndex As Long) As Long
    SymGetTypeInfo 0, ModBaseLo, 0, TypeIndex, TI_GET_COUNT, GetCount
End Function

Private Function GetLength(ByVal ModBaseLo As Long, ByVal TypeIndex As Long) As Long
    Dim l(1) As Long
    SymGetTypeInfo 0, ModBaseLo, 0, TypeIndex, TI_GET_LENGTH, l(0)
    GetLength = l(0)
End Function

Private Function GetBaseType(ByVal ModBaseLo As Long, ByVal TypeIndex As Long) As BasicType
    SymGetTypeInfo 0, ModBaseLo, 0, TypeIndex, TI_GET_BASETYPE, GetBaseType
End Function

Private Function SymEnumSymbolsProc(ByRef pSymInfo As SYMBOL_INFO, ByVal SymbolSize As Long, ByVal UserContext As Long) As Long
Dim szSymName As String, symRVA As Long, symVA As Long
Dim symBT As BasicType, symChildren As Long, symCount As Long, symLen As Long

Dim pos As Long

szSymName = SysAllocString(pSymInfo.Name)
pos = InStr(szSymName, vbNullChar)
If pos > 0 Then szSymName = Mid$(szSymName, 1, pos - 1)
symRVA = pSymInfo.AddressLo - UserContext
symVA = symRVA + dwImageBase

If (pSymInfo.Tag = SymTagData) Or (pSymInfo.Tag = SymTagPublicSymbol) Then
    AddName symVA, szSymName
    
    symBT = GetBaseType(pSymInfo.ModBaseLo, pSymInfo.TypeIndex)
    symChildren = GetChildrenCount(pSymInfo.ModBaseLo, pSymInfo.TypeIndex)
    symCount = GetCount(pSymInfo.ModBaseLo, pSymInfo.TypeIndex)
    symLen = GetLength(pSymInfo.ModBaseLo, pSymInfo.TypeIndex)
    
    If symChildren Then
        setMapRVA symRVA, 3
        'TODO UDT
    ElseIf symCount Then
        setMapRVA symRVA, 3
        'TODO ARRAY
    Else
        Select Case symBT
            Case btBCD
                setMapRVA symRVA, 3
            Case btBit
                setMapRVA symRVA, 3
            Case btBool
                setMapRVA symRVA, 31
            Case btBSTR
                setMapRVA symRVA, 10
            Case btChar
                If IsValidNullString(symVA) Then
                    setMapRVA symRVA, 5
                Else
                    setMapRVA symRVA, 30
                End If
            Case btComplex
                setMapRVA symRVA, 3
            Case btCurrency
                setMapRVA symRVA, 33
            Case btDate
                setMapRVA symRVA, 3
            Case btFloat
                setMapRVA symRVA, 3
            Case btHresult
                setMapRVA symRVA, 32
            Case btInt
                setMapRVA symRVA, 32
            Case btLong
                setMapRVA symRVA, 32
            Case btUInt
                setMapRVA symRVA, 32
            Case btULong
                setMapRVA symRVA, 32
            Case btVariant
                setMapRVA symRVA, 3
            Case btWChar
                If IsValidUnicodeString(symVA) Then
                    setMapRVA symRVA, 10
                Else
                    setMapRVA symRVA, 31
                End If
            Case Else
                If symLen = 4 Then
                    setMapRVA symRVA, 4
                Else
                    setMapRVA symRVA, 3
                End If
        End Select
    End If
'ElseIf pSymInfo.Tag = SymTagEnum Then
ElseIf pSymInfo.Tag = SymTagFunction Then
    AddSubName symVA, szSymName
    tryCallCol.Add symVA
'ElseIf pSymInfo.Tag = SymTagTypedef Then
ElseIf pSymInfo.Tag = SymTagUDT Then
    AddName symVA, szSymName
    setMapRVA symRVA, 3
End If
SysFreeString szSymName

SymEnumSymbolsProc = 1
End Function

'Public Function EnumSymbolsCallback(ByVal SymbolName As Long, ByVal SymbolAddress As Long, ByVal SymbolSize As Long, ByVal UserContext As Long) As Long
'Dim szSymName As String
'
'szSymName = SysAllocString(SymbolName)
'AddName SymbolAddress - UserContext + dwImageBase, szSymName
'SysFreeString szSymName
'
'EnumSymbolsCallback = 1
'End Function
    
Private Sub ProcessSymbols(szExename As String, ByVal lpMapped As Long)
Dim dwOpt As Long, ret As Long
Dim szPath As String

szPath = Mid$(szExename, 1, InStrRev(szExename, "\") - 1)

ret = SymInitialize(0, szPath, 0&)

dwOpt = SymGetOptions
dwOpt = dwOpt Or SYMOPT_UNDNAME
SymSetOptions dwOpt

ret = SymLoadModule(0, 0, szExename, 0&, lpMapped, 0)

'ret = SymEnumerateSymbolsW(0, lpMapped, AddressOf EnumSymbolsCallback, ByVal lpMapped)
ret = SymEnumSymbols(0, lpMapped, 0&, vbNullString, AddressOf SymEnumSymbolsProc, lpMapped)

ret = SymUnloadModule(0, lpMapped)
ret = SymCleanup(0)
End Sub

'désassemble un executable
'=========================
'strExeName : nom et chemin de l'exécutable à désassembler
'strOutASMName : nom du fichier du listing produit
'dwStartingAddress : adresse de départ du désassemblage
'dwRVABase : indique le base des adresses virtuelles relatives de l'exécutable
'bProcessCall :  indique s'il faut descendre dans les procédures rencontrées
Public Function DysPE(ByVal strExeName As String, ByVal strOutFilePattern As String, Optional bProcessCall As Boolean = False)
Dim iCodeFileNum As Integer, iDataFileNum As Integer, iLogFileNum As Integer, X As Long, dw As Long, addr As Long, b As Byte
Dim dwStartingAddress As Long, ExpCount As Long, lpMapped As Long

Set32BitsDecode

'Load frmProgress
'frmProgress.InitPE
'frmProgress.Show

'frmProgress.lblFile.Caption = "Filename : " & strExeName
'frmProgress.lblState.Caption = "Chargement..."
DoEvents

Init
'chargement du fichier
lpMapped = LoadFile(strExeName)
If lpMapped = 0 Then Exit Function

'frmProgress.lblState.Caption = "Traitement de l'entête..."
DoEvents
dwStartingAddress = ProcessPE(strOutFilePattern, strExeName)
'dwStartingAddress = &H1A71C

'frmProgress.lblState.Caption = "Traitement des sections..."
DoEvents
ProcessSections
'frmProgress.imSection.Visible = True
DoEvents

'frmProgress.lblState.Caption = "Traitement des symboles..."
DoEvents
ProcessSymbols strExeName, lpMapped
'frmProgress.imSym1.Visible = True
DoEvents

iCodeFileNum = FreeFile
Open strOutFilePattern & ".asm" For Output As #iCodeFileNum
iDataFileNum = FreeFile
Open strOutFilePattern & ".dat" For Output As #iDataFileNum
iLogFileNum = FreeFile
Open strOutFilePattern & ".log" For Output As #iLogFileNum

    'frmProgress.lblState.Caption = "Traitement des fonctions trouvées dans les symboles..."
    DoEvents
   ' With frmProgress.pbSym
        X = 1
        Do While X <= tryCallCol.Count
           ' .Max = tryCallCol.Count
            
            addr = tryCallCol(X)
            DysCode iCodeFileNum, addr, bProcessCall, GetSubName(addr)
           ' .value = X
            DoEvents
            X = X + 1
        Loop
        Set tryCallCol = New Collection
   ' End With
   ' frmProgress.imSym2.Visible = True
    DoEvents

   ' frmProgress.lblState.Caption = "Traitement du point d'entrée..."
    DoEvents
    DysCode iCodeFileNum, dwStartingAddress, bProcessCall, "start"
   ' frmProgress.imStart.Visible = True
    DoEvents
    
    ExpCount = GetExportsCount
    If ExpCount Then
       ' frmProgress.lblState.Caption = "Traitement des fonctions exportées..."
        DoEvents
       ' With frmProgress.pbExp
        '    .Min = 0
         '   .Max = ExpCount
            For X = 1 To ExpCount
                DysCode iCodeFileNum, GetExportAddr(X), bProcessCall, GetExportName(X)
             '   .value = X
                DoEvents
            Next
        'End With
    End If
   ' frmProgress.imExp.Visible = True
    DoEvents
    
  '  frmProgress.lblState.Caption = "Traitement des offsets..."
    DoEvents
  '  With frmProgress.pbTry
     '   .Min = 0
        X = 1
        Do While X <= tryCallCol.Count
         '   .Max = tryCallCol.Count
            
            addr = tryCallCol(X)
            b = getMapVA(addr)
            If b = 0 Then
                dw = GetAddrSize(addr)
                If dw = 1 Then
                    setMapVA addr, 30
                ElseIf dw = 2 Then
                    setMapVA addr, 31
                ElseIf dw = 4 Then
                    setMapVA addr, 32
                Else
                    If IsValidUnicodeString(addr) Then
                        setMapVA addr, 10
                    ElseIf IsValidNullString(addr) Then
                        setMapVA addr, 5
                    ElseIf IsValidPascalString(addr) Then
                        setMapVA addr, 7
                    Else 'numérique
                        dw = getDwordVA(addr)
                        If CheckVA(dw) Then
                            'pointeur
                            setMapVA addr, 4
                            ProcessPointer addr
                        ElseIf dw Then
                            'code
                            setMapVA addr, 0
                            DysCode iCodeFileNum, addr, bProcessCall, GetSubName(addr)
                            Print #iLogFileNum, "Disassembling from offset at :", getNumber(addr, 8)
                        Else
                            setMapVA addr, 3
                        End If
                    End If
                End If
'            ElseIf b = 4 Then
'                ProcessPointer addr
            End If
         '   .value = X
            DoEvents
            X = X + 1
        Loop
   ' End With
   ' frmProgress.imOff.Visible = True
    DoEvents
    
    'Set tryCallCol = New Collection

    'ProcessOffsets
    
    'TODO desassemble dead code
    
  '  frmProgress.lblState.Caption = "Traitement des données..."
    DoEvents
    ProcessData iCodeFileNum, iDataFileNum, iLogFileNum
  '  frmProgress.imData.Visible = True
    DoEvents
    
Close #iCodeFileNum
Close #iDataFileNum
Close #iLogFileNum

'fermeture de l'exécutable
UnloadFile

'frmProgress.lblState.Caption = "File disassembled in " & Format$(StopTimer, "#.##") & " seconds"
End Function

