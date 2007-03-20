Attribute VB_Name = "modDesassemble"
Option Explicit

Private Sub MakeDir(szDir As String)
On Local Error Resume Next
MkDir szDir
End Sub

Public Sub DisassembleFile(szFilename As String, sDir As String)
Dim szOutPattern As String, pos1 As Long, pos2 As Long

'szOutPattern = Mid$(szFilename, 1, InStrRev(szFilename, "\"))
pos1 = InStrRev(szFilename, "\")
If pos1 = 0 Then pos1 = 1 Else pos1 = pos1 + 1
pos2 = InStrRev(szFilename, ".")
If pos2 = 0 Then pos2 = Len(szFilename) + 1

szOutPattern = sDir & Mid$(szFilename, pos1, pos2 - pos1) 'szFilename & "_desam\" & Mid$(szFilename, pos1, pos2 - pos1)
MakeDir sDir 'szFilename & "_desam\"

If IsLIB(szFilename) Then
  '  If MsgBox("Ce fichier est une library. Son désassemblage peut produire un très grand nombre de fichiers." & vbCrLf & "Voulez-vous continuer ?", vbExclamation Or vbYesNo) = vbYes Then
        DysLIBFile szFilename, szOutPattern
   ' End If
ElseIf IsPE(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable PE.", vbInformation
    DysPE szFilename, szOutPattern, True
ElseIf IsNE(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable NE (format non supporté)", vbCritical
    'TODO
ElseIf IsIntelLE(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable LE (VxD).", vbInformation
    DysLEFile szFilename, szOutPattern
ElseIf IsMZ(szFilename) Then
   ' MsgBox "Ce fichier est un exécutable MZ (MS-DOS).", vbInformation
    DysMZ szFilename, szOutPattern, True
ElseIf IsCOFF(szFilename) Then
    'MsgBox "Ce fichier est un fichier objet COFF.", vbInformation
    DysCOFF szFilename, szOutPattern, True
'TODO OMF
Else
   ' MsgBox "Ce fichier est dans un format non supporté. Désolé.", vbCritical
End If
End Sub
