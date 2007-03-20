Attribute VB_Name = "modNames"
Option Explicit

Private fctsNamesColl As Collection
Private NamesColl As Collection
Private ExportNamesColl As Collection
Private ExportAddrColl As Collection

Public Sub InitNames()
Set fctsNamesColl = New Collection
Set NamesColl = New Collection
Set ExportAddrColl = New Collection
Set ExportNamesColl = New Collection
End Sub

Public Sub AddExport(strExportName As String, ByVal lngExportAddress As Long)
On Error GoTo Fin
    ExportNamesColl.Add strExportName, "SUB" & CStr(lngExportAddress)
    ExportAddrColl.Add lngExportAddress
Fin:
End Sub

Public Function GetExportsCount() As Long
    GetExportsCount = ExportAddrColl.Count
End Function

Public Function GetExportAddr(ByVal X As Long) As Long
    GetExportAddr = ExportAddrColl(X)
End Function

Public Function GetExportName(ByVal X As Long) As String
    GetExportName = ExportNamesColl(X)
End Function

Public Function AddSubName(ByVal dwAddress As Long, Optional szSubName As String = "")
    On Error Resume Next
    If Len(szSubName) Then
        fctsNamesColl.Add szSubName, "SUB" & dwAddress
    Else
        fctsNamesColl.Add "sub_" & getNumber(dwAddress, 8), "SUB" & dwAddress
    End If
End Function

Public Function GetSubName(ByVal dwAddress As Long) As String
On Error GoTo NotInList

GetSubName = fctsNamesColl("SUB" & CStr(dwAddress))

Exit Function
NotInList:
    GetSubName = GetName(dwAddress)
End Function

Public Function AddName(ByVal dwAddress As Long, szName As String)
    On Error Resume Next
    NamesColl.Add szName, "N" & CStr(dwAddress)
End Function

Public Function GetName(ByVal dwAddress As Long) As String
On Error GoTo NotInList

GetName = NamesColl("N" & CStr(dwAddress))

Exit Function
NotInList:
    GetName = vbNullString
End Function

Public Function getAddrName(ByVal lpAddress As Long, ByVal dwSize As Long, Optional cDigit As Long = 8) As String
Dim dt As Long, szSubName As String, szName As String

szSubName = GetSubName(lpAddress)
If Len(szSubName) Then
    getAddrName = szSubName
Else
    szName = GetName(lpAddress)
    If Len(szName) Then
        getAddrName = szName
    Else
        dt = GetDataType(lpAddress, dwSize)
        Select Case dt
            Case 0, 1
                szSubName = GetSubName(lpAddress)
                If Len(szSubName) Then
                    getAddrName = szSubName
                Else
                    getAddrName = "sub_" & getNumber(lpAddress, cDigit)
                End If
            Case 2
                getAddrName = "loc_" & getNumber(lpAddress, cDigit)
            Case 3
                getAddrName = "unk_" & getNumber(lpAddress, cDigit)
            Case 4
                getAddrName = "ptr_" & getNumber(lpAddress, cDigit)
            Case 5
                getAddrName = "sz_" & getNumber(lpAddress, cDigit)
            Case 7
                getAddrName = "pascal_" & getNumber(lpAddress, cDigit)
            Case 10
                getAddrName = "uni_" & getNumber(lpAddress, cDigit)
            Case 30
                getAddrName = "byte_" & getNumber(lpAddress, cDigit)
            Case 31
                getAddrName = "word_" & getNumber(lpAddress, cDigit)
            Case 32
                getAddrName = "dword_" & getNumber(lpAddress, cDigit)
            Case 33
                getAddrName = "qword_" & getNumber(lpAddress, cDigit)
            'Case 254
            '    getAddrName = getNumber(lpAddress, cDigit)
            Case Else
                getAddrName = "unk_" & getNumber(lpAddress, cDigit)
        End Select
    End If
End If
End Function

