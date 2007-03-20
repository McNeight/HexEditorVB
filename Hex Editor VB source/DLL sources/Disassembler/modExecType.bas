Attribute VB_Name = "modExecType"
Option Explicit

Public Function IsPE(szFilename As String) As Boolean
Dim Offset As Long, Signature As Long
Dim iFileNum As Integer

iFileNum = FreeFile
Open szFilename For Binary As #iFileNum
    'on récupère l'offset de l'entête PE à l'offset 0x3C
    Get #1, &H3C + 1, Offset
    'on récupère l'entête PE
    Get #1, Offset + 1, Signature
Close #iFileNum

'si PE
IsPE = (Signature = &H4550&)
End Function

Public Function IsNE(szFilename As String) As Boolean
Dim Offset As Long, Signature As Integer
Dim iFileNum As Integer

iFileNum = FreeFile
Open szFilename For Binary As #iFileNum
    'on récupère l'offset de l'entête PE à l'offset 0x3C
    Get #1, &H3C + 1, Offset
    'on récupère l'entête PE
    Get #1, Offset + 1, Signature
Close #iFileNum

'si PE
IsNE = (Signature = &H454E)
End Function

Public Function IsIntelLE(szFilename As String) As Boolean
Dim Offset As Long, Signature As Long
Dim iFileNum As Integer

iFileNum = FreeFile
Open szFilename For Binary As #iFileNum
    'on récupère l'offset de l'entête PE à l'offset 0x3C
    Get #1, &H3C + 1, Offset
    'on récupère l'entête PE
    Get #1, Offset + 1, Signature
Close #iFileNum

'si PE
IsIntelLE = (Signature = &H454C&)
End Function

'indique si le fichier contient un exe MS-DOS
'remarque : renvoie True pour un PE car un contient un programme stub MS-DOS
Public Function IsMZ(szFilename As String) As Boolean
Dim Signature As Integer
Dim iFileNum As Integer

iFileNum = FreeFile
Open szFilename For Binary As #iFileNum
    'on récupère la signature
    Get #1, 1, Signature
Close #iFileNum

'si MZ
IsMZ = (Signature = &H5A4D)
End Function

'indique si le fichier contient un fichier objet COFF I386
Public Function IsCOFF(szFilename As String) As Boolean
Dim Signature As Integer
Dim iFileNum As Integer

iFileNum = FreeFile
Open szFilename For Binary As #iFileNum
    'on récupère la signature
    Get #1, 1, Signature
Close #iFileNum

'si COFF I386
IsCOFF = (Signature = &H14C)
End Function

'indique si le fichier contient un fichier LIB
Public Function IsLIB(szFilename As String) As Boolean
Dim Signature(1) As Long
Dim iFileNum As Integer

iFileNum = FreeFile
Open szFilename For Binary As #iFileNum
    'on récupère la signature
    Get #1, 1, Signature
Close #iFileNum

'si COFF I386
IsLIB = ((Signature(0) = &H72613C21) And (Signature(1) = &HA3E6863))
End Function

