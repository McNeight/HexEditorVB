Attribute VB_Name = "modData"
Option Explicit
Private Const IMAGE_SCN_CNT_CODE As Long = &H20&
Private Const IMAGE_SCN_MEM_EXECUTE As Long = &H20000000

'définition pour connaitre le type de donnée d'un octet
Private Function IsEscapeSeq(ByVal c As Byte) As Boolean
'\a Bell (alert),\b Backspace,\f Formfeed,\n New line,\r Carriage return,\t Horizontal tab,\v Vertical tab
IsEscapeSeq = ((c = 7) Or (c = 8) Or (c = 12) Or (c = 10) Or (c = 13) Or (c = 9) Or (c = 11))
End Function

Private Function IsPrintExtended(ByVal c As Byte) As Boolean
IsPrintExtended = ((c >= 32) And (c <= &H7F)) Or (c = 130) Or (c = 133) Or (c = 138) Or (c = 135) Or (c = 136)
End Function

Private Function IsPrint(ByVal c As Byte) As Boolean
IsPrint = ((c >= 32) And (c <= 127))
End Function

Private Function IsAlphaNum(ByVal c As Byte) As Boolean
c = Asc(UCase$(Chr$(c)))
IsAlphaNum = (((c >= 65) And (c <= 90)) Or ((c >= 48) And (c <= 57)))
End Function

Private Function IsValidCaracter(ByVal c As Byte) As Boolean
IsValidCaracter = IsEscapeSeq(c) Or IsPrintExtended(c)
End Function

Private Function IsValidEuropeanUnicode(ByVal w As Integer) As Boolean
Dim cl As Byte, ch As Byte
cl = (w And &HFF&)
ch = (w And &HFF00&) \ 256
IsValidEuropeanUnicode = IsValidCaracter(cl) And (ch = 0)
End Function

Public Function IsValidNullString(ByVal lpAddress As Long) As Boolean
Dim oldVA As Long, l As Long

If CheckVA(lpAddress) = False Then
    IsValidNullString = False
    Exit Function
End If

'If (IsValidUnicodeString(lpAddress) = True) Or (getMapRVA(lpAddress) > 0) Then
If (getMapVA(lpAddress) > 0) Then
    IsValidNullString = False
    Exit Function
End If

oldVA = setPointerVA(lpAddress)

l = 0
While (getMap = 0) And (IsValidCaracter(getByte(0)) = True): l = l + 1: Wend

setPointerOffset getPointerOffset - 1
IsValidNullString = ((l > 1) And (getMap = 0) And (getByte(0) = 0))

setPointerVA oldVA
End Function

Public Function IsValidDOSString(ByVal lpAddress As Long) As Boolean
Dim oldVA As Long, l As Long, b As Byte

If CheckVA(lpAddress) = False Then
    IsValidDOSString = False
    Exit Function
End If

If (getMapVA(lpAddress) > 0) Then
    IsValidDOSString = False
    Exit Function
End If

oldVA = setPointerVA(lpAddress)

l = 0
b = getByte(0)
While (getMap = 0) And (b <> 36) And (IsValidCaracter(b) = True)
    l = l + 1
    b = getByte(0)
Wend

IsValidDOSString = ((l > 0) And (getMap = 0) And (b = 0))

setPointerVA oldVA
End Function

Public Function IsValidPascalString(ByVal lpAddress As Long) As Boolean
Dim oldVA As Long, cb As Byte, X As Long

If CheckVA(lpAddress) = False Then Exit Function

If getMapVA(lpAddress) > 0 Then Exit Function

oldVA = setPointerVA(lpAddress)

cb = getByte(0)
For X = 0 To cb - 1
    If (getMap > 0) Or (IsValidCaracter(getByte(0)) = False) Then
        IsValidPascalString = False
        setPointerVA oldVA
        Exit Function
    End If
Next

IsValidPascalString = (cb > 0)
setPointerVA oldVA
End Function

Public Function IsValidUnicodeString(ByVal lpAddress As Long) As Boolean
Dim oldVA As Long, l As Long

If CheckVA(lpAddress) = False Then Exit Function

oldVA = setPointerVA(lpAddress)

l = 0
While (getMap = 0) And (getMapOffset(getPointerOffset + 1) = 0) And (IsValidEuropeanUnicode(getWord(0)) = True): l = l + 1: Wend

setPointerOffset getPointerOffset - 2
IsValidUnicodeString = ((l > 1) And (getMap = 0) And (getMapOffset(getPointerOffset + 1) = 0) And (getWord(0) = 0))

setPointerVA oldVA
End Function

Private Function GetCharPrintable(ByVal bChar As Long) As String
    Select Case bChar
        Case 7 '\a Bell (alert)
            GetCharPrintable = "\a"
        Case 8 '\b Backspace
            GetCharPrintable = "\b"
        Case 12 '\f Formfeed
            GetCharPrintable = "\f"
        Case 10 '\n New line
            GetCharPrintable = "\n"
        Case 13 '\r Carriage return
            GetCharPrintable = "\r"
        Case 9 '\t Horizontal tab
            GetCharPrintable = "\t"
        Case 11 '\v Vertical tab
            GetCharPrintable = "\v"
        Case Else
            GetCharPrintable = ChrW$(bChar)
    End Select
End Function

Public Function GetDataType(ByVal va As Long, ByVal dwSize As Long) As Long
Dim b As Byte, dw As Long

b = getMapVA(va)
Select Case b
    Case 1 To 13, 30 To 33, 255, 250
        GetDataType = b
    Case 0
        If IsCodeVA(va) Then
            dw = getDwordVA(va)
            If CheckVA(dw) Then
                GetDataType = 4
            Else
                If IsValidUnicodeString(va) Then
                    GetDataType = 10
                ElseIf IsValidNullString(va) Then
                    GetDataType = 5
                ElseIf IsValidPascalString(va) Then
                    GetDataType = 7
                Else 'numérique
                    GetDataType = 0
                End If
            End If
        Else
            Select Case dwSize
                Case 0
                    If IsValidUnicodeString(va) Then
                        GetDataType = 10
                    ElseIf IsValidNullString(va) Then
                        GetDataType = 5
                    ElseIf IsValidPascalString(va) Then
                        GetDataType = 7
                    Else 'numérique
                        dw = getDwordVA(va)
                        If CheckVA(dw) Then
                            'pointeur
                            GetDataType = 4
                        Else
                            'numérique taille inconnue
                            GetDataType = 3
                        End If
                    End If
                Case 8
                    If IsValidUnicodeString(va) Then
                        GetDataType = 10
                    ElseIf IsValidNullString(va) Then
                        GetDataType = 5
                    ElseIf IsValidPascalString(va) Then
                        GetDataType = 7
                    Else 'numérique
                        GetDataType = 3
                    End If
                Case 16
                    If IsValidUnicodeString(va) Then
                        GetDataType = 10
                    Else 'numérique
                        GetDataType = 3
                    End If
                Case 32
                    dw = getDwordVA(va)
                    If CheckVA(dw) Then
                        'pointeur
                        GetDataType = 4
                    Else
                        'numérique
                        GetDataType = 3
                    End If
            End Select
        End If
End Select
End Function

Public Sub ProcessData(ByVal iCodeFileNum As Integer, ByVal iDataFileNum As Integer, ByVal iLogFileNum As Integer)
Dim oldVA As Long, addr As Long, iFileNum As Integer
Dim pt As Long, ptr As Long
Dim X As Long, addrfin As Long, dw As Long
Dim off As Long

oldVA = getPointerVA

'ici modif
'With frmProgress.pbData
  '  .Min = 0
   ' .Max = UBound(retSectionTables)
    For X = 0 To UBound(retSectionTables)
        With retSectionTables(X)
            If ((.Characteristics And IMAGE_SCN_CNT_CODE) = IMAGE_SCN_CNT_CODE) Then ' And _
               ((.Characteristics And IMAGE_SCN_MEM_EXECUTE) = IMAGE_SCN_MEM_EXECUTE) Then
                iFileNum = iCodeFileNum
            Else
                iFileNum = iDataFileNum
            End If
            
            'les pointeurs d'abord
            off = .PointerToRawData
            addr = .VirtualAddress
            addrfin = .VirtualAddress + .VirtualSize
            Print #iFileNum, getNumber(addr, 8), StrConv(.SecName, vbUnicode), "segment"
            Do While addr < addrfin
                Select Case getMapOffset(off)
                    Case 4
                        ptr = getDwordVA(addr)
                        pt = GetDataType(ptr, 0)
                        Select Case pt
                            Case 0
                                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
                                dw = GetAddrSize(ptr)
                                If dw = 1 Then
                                    setMapVA ptr, 30
                                    Print #iFileNum, "offset "; getAddrName(ptr, 8)
                                ElseIf dw = 2 Then
                                    setMapVA ptr, 31
                                    Print #iFileNum, "offset "; getAddrName(ptr, 16)
                                ElseIf dw = 4 Then
                                    setMapVA ptr, 32
                                    Print #iFileNum, "offset "; getAddrName(ptr, 32)
                                Else
                                    If IsCodeVA(ptr) Then
                                        Print #iFileNum, "offset sub_"; getNumber(ptr, 8)
                                        Print #iLogFileNum, "Code not disassembled at :", getNumber(ptr, 8)
                                        'DysCode iCodeFileNum, ptr, True
                                    Else
                                        Print #iFileNum, "offset unk_"; getNumber(ptr, 8)
                                    End If
                                End If
                            Case 1 To 10
                                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
                                Print #iFileNum, "offset "; getAddrName(ptr, 0)
'                            Case 2
'                                Print #iFilenum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
'                                Print #iFilenum, "offset loc_"; getNumber(ptr, 8)
'                            Case 3
'                                Print #iFilenum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
'                                Print #iFilenum, "offset unk_"; getNumber(ptr, 8)
'                            Case 4
'                                Print #iFilenum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
'                                Print #iFilenum, "offset ptr_"; getNumber(ptr, 8)
'                            Case 5
'                                Print #iFilenum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
'                                Print #iFilenum, "offset sz_"; getNumber(ptr, 8)
'                            Case 7
'                                Print #iFilenum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
'                                Print #iFilenum, "offset pascal_"; getNumber(ptr, 8)
'                            Case 10
'                                Print #iFilenum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
'                                Print #iFilenum, "offset uni_"; getNumber(ptr, 8)
                            Case 30
                                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
                                Print #iFileNum, "offset "; getAddrName(ptr, 8)
                            Case 31
                                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
                                Print #iFileNum, "offset "; getAddrName(ptr, 16)
                            Case 32
                                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
                                Print #iFileNum, "offset "; getAddrName(ptr, 32)
                            Case 33
                                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ",
                                Print #iFileNum, "offset qword_"; getNumber(ptr, 8)
                        End Select
                        addr = addr + 4
                        off = off + 4
                    Case Else
                        addr = addr + 1
                        off = off + 1
                End Select
            Loop
            addr = .VirtualAddress
            addrfin = .VirtualAddress + .VirtualSize
            Do While addr < addrfin
                Select Case getMapVA(addr)
                    Case 0
                        Print #iFileNum, getNumber(addr, 8), , , "DB ";
                        Call getDataDx(iFileNum, addr, 8, addrfin)
                    'Case 1
                    'Case 2
                    Case 3
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DB ";
                        Call getDataDx(iFileNum, addr, 8, addrfin)
                    Case 4
                        addr = addr + 4
                    Case 5
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DB "; getString(addr, 5); ",0"
                    Case 7
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DB "; getString(addr, 7)
                    Case 10
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DW "; getString(addr, 10); ",0"
                    Case 30
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DB ";
                        Call getDataDx(iFileNum, addr, 8, addrfin, 1)
                    Case 31
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DW ";
                        Call getDataDx(iFileNum, addr, 16, addrfin, 1)
                    Case 32
                        Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0), "DD ";
                        Call getDataDx(iFileNum, addr, 32, addrfin, 1)
                    Case 33
                        Print #iFileNum, getNumber(addr, 8), , "qword_"; getNumber(addr, 8), "DD ";
                        Call getDataDx(iFileNum, addr, 32, addrfin, 2)
                    'Case 34
                    Case Else
                        addr = addr + 1
                End Select
            Loop
            Print #iFileNum, getNumber(addrfin, 8), StrConv(.SecName, vbUnicode), "ends"
        End With
      '  .value = X
        DoEvents
    Next
'End With
setPointerVA oldVA
End Sub

Public Function getString(lpAddress As Long, ByVal bType As Byte) As String
Dim oldVA As Long, c As Byte, w As Integer, cbLength As Long, X As Long

oldVA = setPointerVA(lpAddress)

getString = "'"
Select Case bType
    Case 5 'nullstring
        Do
            c = getByte(5)
            getString = getString & GetCharPrintable(CLng(c))
            lpAddress = lpAddress + 1
        Loop Until c = 0
        getString = Mid$(getString, 1, Len(getString) - 1) & "'"
    Case 6 'DOS "$"
        Do
            c = getByte(6)
            getString = getString & GetCharPrintable(CLng(c))
            lpAddress = lpAddress + 1
        Loop Until c = 36
        getString = getString & "'"
    Case 7 'pascal
        cbLength = getByte(7)
        For X = 0 To cbLength - 1
            getString = getString & GetCharPrintable(getByte(7))
        Next
        lpAddress = lpAddress + cbLength + 1
        getString = getString & "'"
    Case 8 'wide pascal
        cbLength = getWord(8)
        For X = 0 To cbLength - 1
            getString = getString & GetCharPrintable(getByte(8))
        Next
        lpAddress = lpAddress + cbLength + 2
        getString = getString & "'"
    Case 9 'delphi
        cbLength = getDword(9)
        For X = 0 To cbLength - 1
            getString = getString & GetCharPrintable(getByte(9))
        Next
        lpAddress = lpAddress + cbLength + 4
        getString = getString & "'"
    Case 10 'unicode
        Do
            w = getWord(10)
            getString = getString & GetCharPrintable(CLng(w))
            lpAddress = lpAddress + 2
        Loop Until w = 0
        getString = Mid$(getString, 1, Len(getString) - 1) & "',0"
    Case 11 'unicode pascal
    Case 12 'unicode wide pascal
    Case 13 'character terminated
End Select
setPointerVA oldVA
End Function

Public Sub getDataDx(ByVal iFileNum As Integer, lpAddress As Long, ByVal dwDataSize As Long, ByVal lpMaxAddr As Long, Optional dwNumber As Long = -1)
    Select Case dwDataSize
        Case 8
            Call getDataDB(iFileNum, lpAddress, dwNumber, lpMaxAddr)
        Case 16
            Call getDataDW(iFileNum, lpAddress, dwNumber, lpMaxAddr)
        Case 32
            Call getDataDD(iFileNum, lpAddress, dwNumber, lpMaxAddr)
        Case Else
            Call getDataDB(iFileNum, lpAddress, dwNumber, lpMaxAddr)
    End Select
End Sub

Private Sub getDataDB(iFileNum As Integer, lpAddress As Long, dwNumber As Long, ByVal lpMaxAddr As Long)
    Dim b As Byte, X As Long
    Dim oldVA As Long
    
    oldVA = setPointerVA(lpAddress)
    
    If dwNumber = -1 Then
        b = getByte(0)
        Print #iFileNum, Hex$(b), "; "; GetCharPrintable(b)
        lpAddress = lpAddress + 1
        Do While (getMap = 0) And (lpAddress < lpMaxAddr)
            b = getByte(0)
            
            Print #iFileNum, getNumber(lpAddress, 8), , "DB ";
            Print #iFileNum, Hex$(b), "; "; GetCharPrintable(b)
            
            lpAddress = lpAddress + 1
        Loop
    Else
        b = getByte(0)
        Print #iFileNum, Hex$(b), "; "; GetCharPrintable(b)
        lpAddress = lpAddress + 1
        For X = 1 To dwNumber - 1
            b = getByte(0)
            
            Print #iFileNum, getNumber(lpAddress, 8), , "DB ";
            Print #iFileNum, Hex$(b), "; "; GetCharPrintable(b)
            
            lpAddress = lpAddress + 1
        Next
    End If
    
    setPointerVA oldVA
End Sub

Private Sub getDataDW(iFileNum As Integer, lpAddress As Long, dwNumber As Long, ByVal lpMaxAddr As Long)
    Dim w As Integer, X As Long
    Dim oldVA As Long, off As Long
    
    oldVA = setPointerVA(lpAddress)
    
    If dwNumber = -1 Then
        off = VA2Offset(lpAddress + 1)
        
        w = getWord(0)
        Print #iFileNum, getNumber(w, 4)
        lpAddress = lpAddress + 2
        Do While (getMap = 0) And (getMapOffset(off) = 0) And (lpAddress < lpMaxAddr)
            w = getWord(0)
            
            Print #iFileNum, getNumber(lpAddress, 8), , "DW ";
            Print #iFileNum, getNumber(w, 4)
            
            lpAddress = lpAddress + 2
            off = off + 2
        Loop
    Else
        w = getWord(0)
        Print #iFileNum, getNumber(w, 4)
        lpAddress = lpAddress + 2
        For X = 1 To dwNumber - 1
            w = getWord(0)
            
            Print #iFileNum, getNumber(lpAddress, 8), , "DW ";
            Print #iFileNum, getNumber(w, 4)
            
            lpAddress = lpAddress + 2
        Next
    End If
    
    setPointerVA oldVA
End Sub

Private Sub getDataDD(iFileNum As Integer, lpAddress As Long, dwNumber As Long, ByVal lpMaxAddr As Long)
    Dim dw As Long, X As Long
    Dim oldVA As Long, off As Long
    
    oldVA = setPointerVA(lpAddress)
    
    If dwNumber = -1 Then
        off = VA2Offset(lpAddress) + 1
            
        dw = getDword(0)
        Print #iFileNum, getNumber(dw, 8)
        lpAddress = lpAddress + 4
        Do While (getMap = 0) And (getMapOffset(off) = 0) And (getMapOffset(off + 1) = 0) And (getMapOffset(off + 2) = 0) And (lpAddress < lpMaxAddr)
            dw = getDword(0)
            
            Print #iFileNum, getNumber(lpAddress, 8), , "DD ";
            Print #iFileNum, getNumber(dw, 8)
            
            lpAddress = lpAddress + 4
        Loop
    
    Else
        dw = getDword(0)
        Print #iFileNum, getNumber(dw, 8)
        lpAddress = lpAddress + 4
        For X = 1 To dwNumber - 1
            dw = getDword(0)
            
            Print #iFileNum, getNumber(lpAddress, 8), , "DD ";
            Print #iFileNum, getNumber(dw, 8)
            
            lpAddress = lpAddress + 4
        Next
    End If
    
    setPointerVA oldVA
End Sub

Public Function GetAddrSize(ByVal va As Long) As Long
    Dim b As Byte, X As Long
    
    GetAddrSize = 1
    
    va = va + 1
    va = VA2Offset(va)
    b = getMapOffset(va)
    For X = 0 To 5
        If b = 0 Then
            GetAddrSize = GetAddrSize + 1
            va = va + 1
            b = getMapOffset(va)
        Else
            Exit Function
        End If
    Next
End Function
