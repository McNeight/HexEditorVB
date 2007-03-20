Attribute VB_Name = "modMaps"
Option Explicit

'pointeur vers le début de la zone mémoire de la carte
Dim lpBaseMap As Long

Dim hFile As Long, hFileMap As Long

Dim lpDI As Long
'offset de la position dans le fichier exécutable en mémoire (par rapport à la VA Base)
Dim lpPosition As Long
'pointeur vers la base de l'exécutable en mémoire
Dim lpBase As Long
'taille de l'exécutable en mémoire
Dim cbLenght As Long

Private Declare Function MapDebugInformation Lib "dbghelp.dll" (ByVal FileHandle As Long, ByVal FileName As String, ByVal SymbolPath As String, ByVal ImageBase As Long) As Long
Private Declare Function UnmapDebugInformation Lib "dbghelp.dll" (ByVal DebugInfo As Long) As Long

'copie une zone de mémoire dans une autre
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Const MEM_COMMIT As Long = &H1000
Private Const PAGE_READWRITE As Long = &H4
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Const MEM_RELEASE As Long = &H8000
Private Declare Function VirtualFree Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_MAP_COPY As Long = &H1
Private Const PAGE_WRITECOPY As Long = &H8

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

'permet de formatter un message d'erreur système
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' du système
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
' dans la langue du système
Private Const LANG_NEUTRAL = &H0

'valide le pointeur de position dans le fichier
'pour ne pas dépasser la zone (GPF)
Public Function CheckPointer() As Boolean
    'si on est bien dans zone
    CheckPointer = (lpPosition >= 0) And (lpPosition < (cbLenght))
End Function

'Fonctions relative à la position courante

'renvoie un octet de l'executable à la position actuelle et avance de 1
Public Function getByte(Optional ByVal bSetMap As Byte = 1) As Byte
    'vérification du pointeur
    If CheckPointer = False Then Exit Function
    
    'indique le type de l'octet qui va être lu
    If bSetMap Then setMapOffset lpPosition, bSetMap
    
    'copie de l'octet
    CopyMemory getByte, ByVal lpBase + lpPosition, 1&
    'incrémentation d'une position
    lpPosition = lpPosition + 1
End Function

'récupère la valeur de l'octet à la position actuelle sans avancer d'une position
Public Function peekByte() As Byte
    'validation du pointeur
    If CheckPointer = False Then Exit Function
    
    'copie de l'octet à la position en cours
    CopyMemory peekByte, ByVal lpBase + lpPosition, 1&
End Function

'récupère la valeur de l'octet à la position actuelle + 1 sans avancer
Public Function peekByte2() As Byte
    'validation du pointeur
    If CheckPointer = False Then Exit Function
    
    'copie de l'octet
    CopyMemory peekByte2, ByVal lpBase + lpPosition + 1, 1&
End Function

'pour un mot
Public Function getWord(Optional ByVal bSetMap As Byte = 1) As Integer
    If CheckPointer = False Then Exit Function
    
    If bSetMap Then
        setMapOffset lpPosition, bSetMap
        setMapOffset lpPosition + 1, bSetMap
    End If
    CopyMemory getWord, ByVal lpBase + lpPosition, 2&
    lpPosition = lpPosition + 2
End Function

'pour un mot
Public Function getUWord(Optional ByVal bSetMap As Byte = 1) As Long
    If CheckPointer = False Then Exit Function
    
    If bSetMap Then
        setMapOffset lpPosition, bSetMap
        setMapOffset lpPosition + 1, bSetMap
    End If
    CopyMemory getUWord, ByVal lpBase + lpPosition, 2&
    lpPosition = lpPosition + 2
End Function

'pour un double mot
Public Function getDword(Optional ByVal bSetMap As Byte = 1) As Long
    If CheckPointer = False Then Exit Function
    
    If bSetMap Then
        setMapOffset lpPosition, bSetMap
        setMapOffset lpPosition + 1, bSetMap
        setMapOffset lpPosition + 2, bSetMap
        setMapOffset lpPosition + 3, bSetMap
    End If
    
    CopyMemory getDword, ByVal lpBase + lpPosition, 4&
    lpPosition = lpPosition + 4
End Function

'pour un double mot
Public Sub getUnk(ByVal lpBuffer As Long, ByVal cbBuffer As Long)
    If CheckPointer = False Then Exit Sub
    
    CopyMemory ByVal lpBuffer, ByVal lpBase + lpPosition, cbBuffer
    lpPosition = lpPosition + cbBuffer
End Sub

'Fonctions de lecture et de changement de la position courante


'définit l'adresse virtuelle en cours (déplacement far)
Public Function setPointerRVA(ByVal rva As Long) As Long
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then Exit Function

    setPointerRVA = Offset2RVA(lpPosition)
    lpPosition = rva
End Function

'définit l'adresse virtuelle en cours (déplacement far)
Public Function setPointerVA(ByVal va As Long) As Long
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then Exit Function

    setPointerVA = Offset2VA(lpPosition)
    lpPosition = va
End Function

'définit l'adresse virtuelle en cours (déplacement far)
Public Function setPointerOffset(ByVal Offset As Long) As Long
    If (Offset < 0) Or (Offset >= cbLenght) Then Exit Function

    setPointerOffset = lpPosition
    lpPosition = Offset
End Function

'renvoie l'adresse virtuelle en cours
Public Function getPointerRVA() As Long
    getPointerRVA = Offset2RVA(lpPosition)
End Function

'renvoie l'adresse virtuelle en cours
Public Function getPointerVA() As Long
    getPointerVA = Offset2VA(lpPosition)
End Function

'renvoie l'adresse virtuelle en cours
Public Function getPointerOffset() As Long
    getPointerOffset = lpPosition
End Function

'Lecture à des emplacement quelconques

'pour un octet
Public Function getByteRVA(ByVal rva As Long) As Byte
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then Exit Function
    
    CopyMemory getByteRVA, ByVal lpBase + rva, 1&
End Function

'pour un mot
Public Function getWordRVA(ByVal rva As Long) As Integer
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then Exit Function
    
    CopyMemory getWordRVA, ByVal lpBase + rva, 2&
End Function

'pour un double mot
Public Function getDwordRVA(ByVal rva As Long) As Long
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then Exit Function
    
    CopyMemory getDwordRVA, ByVal lpBase + rva, 4&
End Function

'pour un double mot
Public Sub getUnkRVA(ByVal rva As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long)
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then Exit Sub
    
    CopyMemory ByVal lpBuffer, ByVal lpBase + rva, cbBuffer
End Sub

'pour un octet
Public Function getByteVA(ByVal va As Long) As Byte
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then Exit Function
    
    CopyMemory getByteVA, ByVal lpBase + va, 1&
End Function

'pour un mot
Public Function getWordVA(ByVal va As Long) As Integer
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then Exit Function
    
    CopyMemory getWordVA, ByVal lpBase + va, 2&
End Function

'pour un double mot
Public Function getDwordVA(ByVal va As Long) As Long
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then Exit Function
    
    CopyMemory getDwordVA, ByVal lpBase + va, 4&
End Function

'pour un double mot
Public Sub getUnkVA(ByVal va As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long)
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then Exit Sub
    
    CopyMemory ByVal lpBuffer, ByVal lpBase + va, cbBuffer
End Sub

'pour un octet
Public Function getByteOffset(ByVal Offset As Long) As Byte
    If (Offset < 0) Or (Offset >= cbLenght) Then Exit Function
    
    CopyMemory getByteOffset, ByVal lpBase + Offset, 1&
End Function

'pour un mot
Public Function getWordOffset(ByVal Offset As Long) As Integer
    If (Offset < 0) Or (Offset >= cbLenght) Then Exit Function
    
    CopyMemory getWordOffset, ByVal lpBase + Offset, 2&
End Function

'pour un double mot
Public Function getDwordOffset(ByVal Offset As Long) As Long
    If (Offset < 0) Or (Offset >= cbLenght) Then Exit Function
    
    CopyMemory getDwordOffset, ByVal lpBase + Offset, 4&
End Function

'pour un double mot
Public Function getUnkOffset(ByVal Offset As Long, ByVal lpBuffer As Long, ByVal cbBuffer As Long) As Boolean
    If (Offset < 0) Or (Offset >= cbLenght) Then getUnkOffset = False: Exit Function
    
    CopyMemory ByVal lpBuffer, ByVal lpBase + Offset, cbBuffer
    getUnkOffset = True
End Function

'définit le type d'octet à l'adresse indiquée
Public Sub setMapVA(ByVal va As Long, ByVal bType As Byte)
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then Exit Sub
    
    CopyMemory ByVal lpBaseMap + va, bType, 1&
End Sub

'renvoie le type d'octet à l'adresse indiquée
Public Function getMapVA(ByVal va As Long) As Byte
    va = VA2Offset(va)
    If (va < 0) Or (va >= cbLenght) Then getMapVA = 255: Exit Function

    CopyMemory getMapVA, ByVal lpBaseMap + va, 1&
End Function

'définit le type d'octet à l'adresse indiquée
Public Sub setMapRVA(ByVal rva As Long, ByVal bType As Byte)
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then Exit Sub
    
    CopyMemory ByVal lpBaseMap + rva, bType, 1&
End Sub

'renvoie le type d'octet à l'adresse indiquée
Public Function getMapRVA(ByVal rva As Long) As Byte
    rva = RVA2Offset(rva)
    If (rva < 0) Or (rva >= cbLenght) Then getMapRVA = 255: Exit Function

    CopyMemory getMapRVA, ByVal lpBaseMap + rva, 1&
End Function

'définit le type d'octet à l'adresse indiquée
Public Sub setMapOffset(ByVal Offset As Long, ByVal bType As Byte)
    If (Offset < 0) Or (Offset >= cbLenght) Then Exit Sub
    
    CopyMemory ByVal lpBaseMap + Offset, bType, 1&
End Sub

'renvoie le type d'octet à l'adresse indiquée
Public Function getMapOffset(ByVal Offset As Long) As Byte
    If (Offset < 0) Or (Offset >= cbLenght) Then getMapOffset = 255: Exit Function

    CopyMemory getMapOffset, ByVal lpBaseMap + Offset, 1&
End Function

'définit le type d'octet à l'adresse indiquée
Public Sub setMap(ByVal bType As Byte)
    If CheckPointer = False Then Exit Sub
    
    CopyMemory ByVal lpBaseMap + lpPosition, bType, 1&
End Sub

'renvoie le type d'octet à l'adresse indiquée
Public Function getMap() As Byte
    If CheckPointer = False Then getMap = 255: Exit Function

    CopyMemory getMap, ByVal lpBaseMap + lpPosition, 1&
End Function

'chargement d'un fichier
Public Function LoadFile(strFilename As String) As Long
Dim b As Byte

lpDI = MapDebugInformation(-1, strFilename, vbNullString, 0)

If lpDI = 0 Then MsgBox FormatErrorMessage(Err.LastDllError): Exit Function

CopyMemory lpBase, ByVal lpDI + 12, 4&

'copie de la taille du fichier
cbLenght = FileLen(strFilename)

lpBaseMap = VirtualAlloc(ByVal 0&, cbLenght + 1, MEM_COMMIT, PAGE_READWRITE) 'MapViewOfFile(hMap, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, 0)

If lpBaseMap = 0 Then MsgBox FormatErrorMessage(Err.LastDllError): UnloadFile: Exit Function

b = 255
CopyMemory ByVal lpBaseMap + cbLenght, b, 1

LoadFile = lpBase
End Function

'fermeture du fichier
Public Function UnloadFile()
VirtualFree lpBaseMap, 0, MEM_RELEASE

UnmapDebugInformation lpDI
End Function

'chargement d'un fichier
Public Function LoadFile2(strFilename As String) As Long
Dim b As Byte

hFile = CreateFile(strFilename, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
If hFile = -1 Then MsgBox FormatErrorMessage(Err.LastDllError): Exit Function

hFileMap = CreateFileMapping(hFile, 0&, PAGE_WRITECOPY, 0&, 0&, vbNullString)
If hFileMap = -1 Then MsgBox FormatErrorMessage(Err.LastDllError): Exit Function

lpBase = MapViewOfFile(hFileMap, FILE_MAP_COPY, 0&, 0&, 0&)
If lpBase = 0 Then MsgBox FormatErrorMessage(Err.LastDllError): Exit Function
 
'copie de la taille du fichier
cbLenght = FileLen(strFilename)

lpBaseMap = VirtualAlloc(ByVal 0&, cbLenght + 1, MEM_COMMIT, PAGE_READWRITE) 'MapViewOfFile(hMap, FILE_MAP_READ Or FILE_MAP_WRITE, 0, 0, 0)

If lpBaseMap = 0 Then MsgBox FormatErrorMessage(Err.LastDllError): UnloadFile: Exit Function

b = 255
CopyMemory ByVal lpBaseMap + cbLenght, b, 1

LoadFile2 = lpBase
End Function

'fermeture du fichier
Public Function UnloadFile2()
VirtualFree lpBaseMap, 0, MEM_RELEASE

UnmapViewOfFile lpBase
CloseHandle hFileMap
CloseHandle hFile
End Function

'pour un double mot
Public Function setDwordOffset(ByVal Offset As Long, ByVal dw As Long) As Long
    If (Offset < 0) Or (Offset >= cbLenght) Then Exit Function
    
    CopyMemory ByVal lpBase + Offset, dw, 4&
End Function

Public Function GetSZString() As String
Dim c As Byte

setMap 5
c = getByte(0)
Do While c
    GetSZString = GetSZString & Chr$(c)
    c = getByte(0)
Loop
End Function

Public Function getImageBase() As Long
    getImageBase = lpBase
End Function

Public Function setImageBase(ByVal lpNewBase As Long) As Long
    setImageBase = lpBase
    lpBase = lpNewBase
End Function

Public Function getImageLength() As Long
    getImageLength = cbLenght
End Function

Public Function setImageLength(ByVal cbNewLength As Long) As Long
    setImageLength = cbLenght
    cbLenght = cbNewLength
End Function

Public Function getMapBase() As Long
    getMapBase = lpBaseMap
End Function

Public Function setMapBase(ByVal lpNewBase As Long) As Long
    setMapBase = lpBaseMap
    lpBaseMap = lpNewBase
End Function

'permet de formater un message correspondant à un code d'erreur
'==============================================================
'ErrCode : code de l'erreur
'renvoie un chaine descriptive de l'erreur
Public Function FormatErrorMessage(ByVal ErrCode As Long) As String

    Dim sBuffer As String ' Définit un buffer pour contenir le message
    Dim nBufferSize As Long ' Taille du buffer

    'taille du buffer
    nBufferSize = 1024
    'on fait la place pour 1024 caractères
    sBuffer = String$(nBufferSize, Chr$(0))

    'demande de la chaine descriptive
    nBufferSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrCode, LANG_NEUTRAL, sBuffer, nBufferSize, ByVal 0&)
    
    'si erreur connue
    If nBufferSize > 0 Then
        'on supprime les zéros terminaux en trop
        FormatErrorMessage = Left$(sBuffer, nBufferSize)
    'sinon erreur inconnue
    ElseIf nBufferSize = 0 Then
        'on le signal dans la chaine descriptive
        FormatErrorMessage = "Erreur " & ErrCode & " non définie."
    End If
End Function
