Attribute VB_Name = "modMZ"
Option Explicit

Private Const IMAGE_SCN_MEM_16BIT As Long = &H20000
Private Type MZHeader
    cMZ(1) As Byte
    cbLastPage As Integer        'Number of bytes in last 512-byte page
                                 'of executable
    cPages As Integer            'Total number of 512-byte pages in executable
                                 '(including the last page)
    cRelocations As Integer      'Number of relocation entries
    cbHeaderSize As Integer      'Header size in paragraphs
    cMinParagraph As Integer     'Minimum paragraphs of memory allocated in
                                 'addition to the code size
    cMaxParagraph As Integer     'Maximum number of paragraphs allocated in
                                 'addition to the code size
    wInitSS As Integer           'Initial SS relative to start of executable
    wInitSP As Integer           'Initial SP
    wCheckSum As Integer         'Checksum (or 0) of executable
    dwCSIPEntryPoint As Long     'CS:IP relative to start of executable
                                 '(entry point)
    wOffsetRelocTable As Integer 'Offset of relocation table;
                                 '40h for new-(NE,LE,LX,W3,PE etc.) executable
    wOverlay As Integer          'Overlay number (0h = main program)
End Type
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public regEAX As Long          '
Public regDS As Long          '

Private Function ProcessMZ(strOutFilePattern As String, strFilename As String) As Long
Dim MZH As MZHeader, iFileMZ As Integer, wOff As Long, wSeg As Long
Dim X As Long, ptr As Long
Dim wRawOff As Integer, wRawSeg As Integer, w As Long

getUnkOffset 0, VarPtr(MZH), Len(MZH)

iFileMZ = FreeFile
Open strOutFilePattern & ".mz" For Output As #iFileMZ
    With MZH
        Print #iFileMZ, ";=================================================================="
        Print #iFileMZ, ";MS-DOS Executable Information : "; strFilename
        Print #iFileMZ, ";=================================================================="
    
        Print #iFileMZ, "Header size :", .cbHeaderSize * 16; " byte(s)"
        Print #iFileMZ, "Number of 512-bytes Pages :", .cPages
        Print #iFileMZ, "Last Page :", .cbLastPage; " byte(s)"
        
        Print #iFileMZ, "Entry Point CS:IP :", getNumber(.dwCSIPEntryPoint, 8)
        Print #iFileMZ, "Initial SP :", , getNumber(.wInitSP, 4)
        Print #iFileMZ, "Initial SS :", , getNumber(.wInitSS, 4)
        
        Print #iFileMZ, "Min Extra Alloc :", .cMinParagraph * 16; " byte(s)"
        Print #iFileMZ, "Max Extra Alloc :", .cMaxParagraph * 16; " byte(s)"
        
        Print #iFileMZ, "Offset to Relocation Table :", getNumber(.wOffsetRelocTable, 4)
        Print #iFileMZ, "Number of Relocations :", .cRelocations
        
        Print #iFileMZ, "CheckSum :", , getNumber(.wCheckSum, 4)
        Print #iFileMZ, "Overlay :", , .wOverlay
        
        If .wInitSP Or .wInitSS Then
            CopyMemory wOff, .wInitSP, 2&
            CopyMemory wSeg, .wInitSS, 2&
            
            For ptr = .cbHeaderSize * 16 + wOff + wSeg * 16 - 1 To .cbHeaderSize * 16 Step -1
                setMapOffset ptr, 250
            Next
        End If
        
        ReDim retSectionTables(0)
        
        With retSectionTables(0)
            .Characteristics = IMAGE_SCN_MEM_16BIT
            .PointerToRawData = MZH.cbHeaderSize * 16
            .SizeOfRawData = CLng(MZH.cPages - 1) * 512 + MZH.cbLastPage - MZH.cbHeaderSize * 16
            .VirtualAddress = &H10000
            .VirtualSize = .SizeOfRawData
            setMapOffset .PointerToRawData + .SizeOfRawData, 255
        End With
        
'        setPointerOffset .wOffsetRelocTable
'        For X = 0 To .cRelocations - 1
'            wRawOff = getWord(0)
'            wRawSeg = getWord(0)
'
'            CopyMemory wOff, wRawOff, 2&
'            CopyMemory wSeg, wRawSeg, 2&
'
'            ptr = lpMapped + .cbHeaderSize * 16 + wOff + wSeg * 16
'            CopyMemory w, ByVal ptr, 2&
'            w = w + .cbHeaderSize * 16
'            CopyMemory ByVal ptr, w, 2&
'        Next
    
        CopyMemory wOff, ByVal VarPtr(.dwCSIPEntryPoint), 2&
        CopyMemory wSeg, ByVal VarPtr(.dwCSIPEntryPoint) + 2, 2&
        
        ProcessMZ = wSeg * 16 + wOff + &H10000
        dwImageBase = &H10000
        'regDS = &H1000
    End With
Close #iFileMZ
End Function

Private Sub ProcessData16(ByVal iFileNum As Integer)
Dim oldVA As Long, addr As Long
Dim pt As Long, ptr As Long
Dim X As Long, addrfin As Long, dw As Long
Dim off As Long, cnt As Long

With retSectionTables(0)
    'les pointeurs d'abord
    off = .PointerToRawData
    addr = .VirtualAddress
    addrfin = .VirtualAddress + .VirtualSize
    
    Do While addr < addrfin
        Select Case getMapVA(addr)
            Case 0, 4
                Print #iFileNum, getNumber(addr, 8), , , "DB ";
                Call getDataDx(iFileNum, addr, 8, addrfin)
            Case 3
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DB ";
                Call getDataDx(iFileNum, addr, 8, addrfin)
            Case 5
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DB "; getString(addr, 5); ",0"
            Case 7
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DB "; getString(addr, 7)
            Case 10
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DW "; getString(addr, 10); ",0"
            Case 30
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DB ";
                Call getDataDx(iFileNum, addr, 8, addrfin, 1)
            Case 31
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DW ";
                Call getDataDx(iFileNum, addr, 16, addrfin, 1)
            Case 32
                Print #iFileNum, getNumber(addr, 8), , getAddrName(addr, 0, 5), "DD ";
                Call getDataDx(iFileNum, addr, 32, addrfin, 1)
            Case 33
                Print #iFileNum, getNumber(addr, 8), , "qword_"; getNumber(addr, 5), "DD ";
                Call getDataDx(iFileNum, addr, 32, addrfin, 2)
            Case 250
                Print #iFileNum, getNumber(addr, 8), "STACK segment"
                cnt = 0
                Do While getMapVA(addr) = 250
                    addr = addr + 1
                    cnt = cnt + 1
                Loop
                Print #iFileNum, getNumber(addr, 8), , "DB "; getNumber(cnt, 5); " DUP (0)"
                Print #iFileNum, getNumber(addr, 8), "STACK ends"
                
                Print #iFileNum, getNumber(addr + 1, 8), "seg000 segment"
                If regDS Then
                    Print #iFileNum, getNumber(addr + 1, 8), , "assume CS:seg000,DS:"; regDS
                Else
                    Print #iFileNum, getNumber(addr + 1, 8), , "assume CS:seg000,DS:seg000"
                End If
            Case Else
                addr = addr + 1
        End Select
       ' frmProgress.pbData.value = 100 / addrfin * addr
        DoEvents
    Loop
    Print #iFileNum, getNumber(addrfin - 1, 8), "seg000 ends"
    Print #iFileNum, getNumber(addrfin, 8), , "end start"
End With
End Sub

Public Function DysMZ(ByVal strExeName As String, ByVal strOutFilePattern As String, Optional bProcessCall As Boolean = False)
Dim lpEntry As Long, iFileNum As Integer

'Load frmProgress
'frmProgress.InitMZ
'frmProgress.Show

'frmProgress.lblFile.Caption = "Filename : " & strExeName
'frmProgress.lblState.Caption = "Chargement..."
DoEvents

Init

Set16BitsDecode

'chargement du fichier
If LoadFile2(strExeName) = 0 Then Exit Function

'frmProgress.lblState.Caption = "Traitement de l'entête..."
DoEvents

lpEntry = ProcessMZ(strOutFilePattern, strExeName)

'frmProgress.imSection.Visible = True
DoEvents

'frmProgress.lblState.Caption = "Traitement du point d'entrée..."
DoEvents

iFileNum = FreeFile
Open strOutFilePattern & ".asm" For Output As #iFileNum
    DysCode iFileNum, lpEntry, True, "start"
    
   ' frmProgress.imStart.Visible = True
    DoEvents
    
  '  frmProgress.lblState.Caption = "Traitement des données..."
    DoEvents
    
    ProcessData16 iFileNum

  '  frmProgress.imData.Visible = True
    DoEvents
Close #iFileNum

'frmProgress.lblState.Caption = "File disassembled in " & Format$(StopTimer, "#.##") & " seconds"

UnloadFile2
End Function
