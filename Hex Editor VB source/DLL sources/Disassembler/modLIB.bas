Attribute VB_Name = "modLIB"
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

Private Type LIBFileSignature
    Signature(7) As Byte
End Type

Private Type LIBMemberHeader
    MemberName(15) As Byte
    MemberDate(11) As Byte
    MemberUserID(5) As Byte
    MemberGroupID(5) As Byte
    MemberMode(7) As Byte
    MemberSize(9) As Byte
    MemberEnd(1) As Byte
End Type
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Function BigEndian2LittleEndian(ByVal dw As Long) As Long
    Dim ptr1 As Long, ptr2 As Long
    
    ptr1 = VarPtr(BigEndian2LittleEndian)
    ptr2 = VarPtr(dw)
    CopyMemory ByVal ptr1 + 0, ByVal ptr2 + 3, 1&
    CopyMemory ByVal ptr1 + 1, ByVal ptr2 + 2, 1&
    CopyMemory ByVal ptr1 + 2, ByVal ptr2 + 1, 1&
    CopyMemory ByVal ptr1 + 3, ByVal ptr2 + 0, 1&
End Function

Private Function GetMemberName(MH As LIBMemberHeader, ByVal offLN As Long) As String
    Dim oldOff As Long
    If (MH.MemberName(0) = 47) And ((MH.MemberName(1) >= 48) And (MH.MemberName(1) <= 57)) Then 'long name
        oldOff = setPointerOffset(offLN + CLng(Val(Mid$(StrConv(MH.MemberName, vbUnicode), 2))))
        GetMemberName = GetSZString()
        setPointerOffset oldOff
    Else
        GetMemberName = Trim$(StrConv(MH.MemberName, vbUnicode))
    End If
End Function

Private Function GetMemberSize(MH As LIBMemberHeader) As Long
    GetMemberSize = CLng(Val(StrConv(MH.MemberSize, vbUnicode)))
End Function

Private Sub ProcessMemberHeader(ByVal iFileLib As Integer, ByRef MemberHeader As LIBMemberHeader, ByVal offLN As Long)
With MemberHeader
    Print #iFileLib, "Name :", GetMemberName(MemberHeader, offLN)
    Print #iFileLib, "Date :", Trim$(StrConv(.MemberDate, vbUnicode))
    Print #iFileLib, "User ID :", Trim$(StrConv(.MemberUserID, vbUnicode))
    Print #iFileLib, "Group ID :", Trim$(StrConv(.MemberGroupID, vbUnicode))
    Print #iFileLib, "Mode :", Trim$(StrConv(.MemberMode, vbUnicode))
    Print #iFileLib, "Size :", Trim$(StrConv(.MemberSize, vbUnicode))
End With
End Sub

Private Sub ProcessFirstMember(ByVal iFileLib As Integer, ByVal Offset As Long, ByVal OffsetLN As Long)
Dim NumberOfSymbols As Long, offStringTable As Long
Dim X As Long, off As Long, MH As LIBMemberHeader

'Private Type LIBFirstArchiveMember
'    MemberHeader As LIBMemberHeader
'    NumberOfSymbols As Long
'    'Offsets(NumberOfSymbols - 1) As Long 'big-endian
'    'StringTable : series of null-terminated string
'End Type

NumberOfSymbols = getDwordOffset(Offset)
NumberOfSymbols = BigEndian2LittleEndian(NumberOfSymbols)

If NumberOfSymbols Then
    setPointerOffset Offset + 4 + NumberOfSymbols * 4
    For X = 0 To NumberOfSymbols - 1
        off = getDwordOffset(Offset + 4 + X * 4)
        off = BigEndian2LittleEndian(off)
        
        getUnkOffset off, VarPtr(MH), Len(MH)
        Print #iFileLib, "Symbol "; GetSZString; " exported from module "; GetMemberName(MH, OffsetLN); " at offset "; off
    Next
End If
End Sub

Private Sub ProcessSecondMember(ByVal iFileLib As Integer, ByVal Offset As Long, ByVal OffsetLN As Long)
Dim NumberOfMembers As Long, offnbSymbols As Long, NumberOfSymbols As Long, offStringTable As Long
Dim X As Long, idx As Integer, off As Long, MH As LIBMemberHeader

'Private Type LIBSecondArchiveMember
'    MemberHeader As LIBMemberHeader
'    NumberOfMembers As Long
'    'Offsets(NumberOfMembers-1) As Long
'    'NumberOfSymboles As Long
'    'Indices(NumberOfSymbols - 1) As Integer
'    'StringTable : series of null-terminated string
'End Type

NumberOfMembers = getDwordOffset(Offset)
If NumberOfMembers Then
    offnbSymbols = Offset + 4 + 4 * NumberOfMembers
    NumberOfSymbols = getDwordOffset(offnbSymbols)
    
    If NumberOfSymbols Then
        setPointerOffset offnbSymbols + 4 + 2 * NumberOfSymbols
        For X = 0 To NumberOfSymbols - 1
            idx = getWordOffset(offnbSymbols + 4 + X * 2)
            off = getDwordOffset(Offset + 4 + 4 * CLng(idx - 1))
            
            getUnkOffset off, VarPtr(MH), Len(MH)
            Print #iFileLib, "Symbol "; GetSZString; " exported from module "; GetMemberName(MH, OffsetLN); " (index "; idx; ") at offset "; off
        Next
    End If
End If
End Sub

Private Sub ProcessObjects(ByVal iFileLib As Integer, ByVal Offset As Long, ByVal OffsetLN As Long, szOutPattern As String)
Dim off As Long, MH As LIBMemberHeader, num As Long, szMemberName As String
Dim oldBase As Long, oldLength As Long, oldMapBase As Long, sig As Long, iFileNum As Integer

off = Offset
num = 0
Do While getUnkOffset(off, VarPtr(MH), Len(MH))
    szMemberName = GetMemberName(MH, OffsetLN)
    
    'ici modif
   ' frmProgress.lblFile.Caption = "Filename :" & szOutPattern & " -> " & num & " -> " & szMemberName
    DoEvents
    
    Print #iFileLib, "----------------------------------------------------------------------"
    Print #iFileLib, num; "th Object Member :", szMemberName
    Print #iFileLib, "----------------------------------------------------------------------"
    ProcessMemberHeader iFileLib, MH, OffsetLN
    
    off = off + Len(MH)
    oldBase = getImageBase
    setImageBase oldBase + off
    oldLength = setImageLength(GetMemberSize(MH))
    oldMapBase = getMapBase
    setMapBase oldMapBase + off
    
    If InStr(szMemberName, "/") > 0 Then
        szMemberName = Mid$(szMemberName, 1, InStr(szMemberName, "/") - 1)
    End If
    
    sig = getDwordOffset(0)
    If sig = &HFFFF0000 Then
        If iFileNum = 0 Then
            iFileNum = FreeFile
            Open szOutPattern & ".sil" For Output As #iFileNum
        End If
        Print #iFileLib, "This member is an import member"
        DysImport szOutPattern & "_" & num & "_" & szMemberName, szOutPattern & "_" & num & "_" & szMemberName, iFileNum
    Else
        Print #iFileLib, "This member is an object file"
        DysCOFF2 szOutPattern & "_" & num & "_" & szMemberName, szOutPattern & "_" & num & "_" & szMemberName, True
    End If
    
    setImageBase oldBase
    setImageLength oldLength
    setMapBase oldMapBase
    
    num = num + 1
    off = off + GetMemberSize(MH)
    If (off Mod 2) = 1 Then off = off + 1
Loop

If iFileNum Then
    Close #iFileNum
End If
End Sub

Public Sub DysLIBFile(szLibFile As String, szOutPattern As String)
Dim iFileLib As Integer
Dim sign As LIBFileSignature
'Dim fam As LIBFirstArchiveMember, sam As LIBSecondArchiveMember
Dim MH1 As LIBMemberHeader, MH2 As LIBMemberHeader, MH3 As LIBMemberHeader
Dim offFAM As Long, offSAM As Long, offTAM As Long, offOBJ As Long

'Load frmProgress
'frmProgress.InitCOFF
'frmProgress.Show

'frmProgress.lblFile.Caption = "Filename : " & szLibFile
'frmProgress.lblState.Caption = "Chargement..."
DoEvents

Init

Set32BitsDecode

'chargement du fichier
If LoadFile2(szLibFile) = 0 Then Exit Sub

'frmProgress.lblState.Caption = "Traitement de l'entête..."
DoEvents

iFileLib = FreeFile
Open szOutPattern & ".lct" For Output As #iFileLib
    Print #iFileLib, "======================================================================"
    Print #iFileLib, "LIB File : "; szLibFile
    Print #iFileLib, "======================================================================"
    
    getUnkOffset 0, VarPtr(sign), Len(sign)
    
    Print #iFileLib, "Signature :", StrConv(sign.Signature, vbUnicode)
    
    offFAM = Len(sign)
    getUnkOffset offFAM, VarPtr(MH1), Len(MH1)
    offFAM = offFAM + Len(MH1)
    
    offSAM = offFAM + GetMemberSize(MH1)
    getUnkOffset offSAM, VarPtr(MH2), Len(MH2)
    offSAM = offSAM + Len(MH2)
    
    offTAM = offSAM + GetMemberSize(MH2)
    getUnkOffset offTAM, VarPtr(MH3), Len(MH3)
    offTAM = offTAM + Len(MH3)
    
    offOBJ = offTAM + GetMemberSize(MH3)
    
    If (MH1.MemberName(0) = 47) And (MH1.MemberName(1) = 32) Then

       ' frmProgress.lblState.Caption = "Traitement du premier membre du linker..."
        DoEvents
        
        Print #iFileLib, "----------------------------------------------------------------------"
        Print #iFileLib, "First Linker Member"
        Print #iFileLib, "----------------------------------------------------------------------"
        
        ProcessMemberHeader iFileLib, MH1, offTAM
        ProcessFirstMember iFileLib, offFAM, offTAM
        
        If (MH2.MemberName(0) = 47) And (MH2.MemberName(1) = 32) Then
            Print #iFileLib, "----------------------------------------------------------------------"
            Print #iFileLib, "Second Linker Member"
            Print #iFileLib, "----------------------------------------------------------------------"
            
           ' frmProgress.lblState.Caption = "Traitement du second membre du linker..."
            DoEvents
            
            ProcessMemberHeader iFileLib, MH2, offTAM
            ProcessSecondMember iFileLib, offSAM, offTAM
            
            If (MH3.MemberName(0) = 47) And (MH3.MemberName(1) = 47) And (MH3.MemberName(2) = 32) Then
               ' frmProgress.lblState.Caption = "Traitement du troisième membre..."
                DoEvents
                
                Print #iFileLib, "----------------------------------------------------------------------"
                Print #iFileLib, "Long Names Member"
                Print #iFileLib, "----------------------------------------------------------------------"
                
                ProcessMemberHeader iFileLib, MH3, offTAM
            Else
                offOBJ = offTAM - 60
                offTAM = 0
            End If
        Else
            offOBJ = offSAM - 60
            offSAM = 0
            offTAM = 0
        End If
    Else
        offOBJ = offFAM - 60
        offFAM = 0
        offSAM = 0
        offTAM = 0
    End If
    
'frmProgress.lblState.Caption = "Traitement des membres objets..."
DoEvents
    
    Print #iFileLib, "----------------------------------------------------------------------"
    Print #iFileLib, "Object Module Members"
    Print #iFileLib, "----------------------------------------------------------------------"
    
    ProcessObjects iFileLib, offOBJ, offTAM, szOutPattern
Close #iFileLib

'frmProgress.lblState.Caption = "File disassembled in " & Format$(StopTimer, "#.##") & " seconds"

'UnloadFile2
End Sub
