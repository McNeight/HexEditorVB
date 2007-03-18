VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetionErreur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =======================================================
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A complete hexadecimal editor for Windows ©
' (Editeur hexadécimal complet pour Windows ©)
'
' Copyright © 2006-2007 by Alain Descotes.
'
' This file is part of Hex Editor VB.
'
' Hex Editor VB is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' Hex Editor VB is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with Hex Editor VB; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
' =======================================================


Option Explicit

'=======================================================
'//CLASS DE GESTION DES ERREURS
'//ECRIT UN *.LOG AVEC LES ERREURS RENCONTREES
'=======================================================

'=======================================================
'TYPE & ENUMS
'=======================================================
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'=======================================================
'CONSTANTES
'=======================================================
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10
Private Const INVALID_HANDLE_VALUE          As Long = -1


'=======================================================
'APIS
'=======================================================
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long


'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private bMakeSoundIDE As Boolean
Private sLogFile As String

'=======================================================
'EVENEMENTS
'=======================================================
Public Event DisplayError(Error As ErrObject)



Private Sub Class_Initialize()

    'initialise les variables privées
    bMakeSoundIDE = True
    sLogFile = App.Path & "\ErrLog.txt"
    
    'efface l'erreur actuelle
    Err.Clear
    
End Sub

'=======================================================
'ajoute une erreur au log
'=======================================================
Public Sub AddError(ByVal sZone As String, Optional ByVal WriteLog As Boolean = True)
Dim lFile As Long
Dim sTemp As String

    'On Error Resume Next
    
    'ajoute du texte à la console
    Call AddTextToConsole("Erreur rencontrée, écriture dans le log...")
    
    'obtient un N° de fichier dispo
    lFile = FreeFile
    
    'vérifie que le fichier de log existe, le créé le cas échéant
    If Not (FileExists(sLogFile)) Then
        If FolderExists(GetFolderFromPath(sLogFile)) = False Then
            'alors il faut créé les dossiers précédant le path du fichier
            sTemp = sLogFile
            While FolderExists(sTemp) = False
                'alors on remonte d'un dossier
                sTemp = GetPrevFolder(sTemp)
                DoEvents
            Wend
            
            sTemp = GetFolderAfter(sLogFile, sTemp)
            
            While sTemp <> GetFolderFromPath(sLogFile)
                'on créé désormais les dossiers inexistants
                MkDir GetFolderAfter(sLogFile, sTemp)
                sTemp = GetFolderAfter(sLogFile, sTemp)
                DoEvents
            Wend
        
        
        End If
    End If
    
    'le dossier existe maintenant, on créé donc le fichier
    Open sLogFile For Append As lFile
    
    'stemp va contenir le contenu de la ligne à écrite
    sTemp = vbNewLine & "Date=[" & Date$ & "]"
    sTemp = sTemp & vbNewLine & "Time=[" & Time$ & "]"
    sTemp = sTemp & vbNewLine & "Error number=[" & Err.Number & "]"
    sTemp = sTemp & vbNewLine & "Error source=[" & Err.Source & "]"
    sTemp = sTemp & vbNewLine & "Error description=[" & Err.Description & "]"
    sTemp = sTemp & vbNewLine & "Error zone=[" & sZone & "]"
    
    'écrit dans le fichier
    RaiseEvent DisplayError(Err)
    
    If WriteLog Then Print #lFile, sTemp
    If bMakeSoundIDE And App.LogMode <> 1 Then Beep 300, 100    'car pas dans l'IDE
    
    Close lFile 'ferme le fichier

    'ajoute du texte à la console
    Call AddTextToConsole(sTemp)
    
    'ajoute du texte à la console
    Call AddTextToConsole("Erreur écrite dans le log")
End Sub

'=======================================================
'efface le fichier log
'=======================================================
Public Function DeleteLogFile() As Long
    DeleteLogFile = DeleteFile(sLogFile)
End Function




'=======================================================
'propriétés de la classe
'=======================================================
Public Property Get MakeSoundIDE() As Boolean: MakeSoundIDE = bMakeSoundIDE: End Property
Public Property Let MakeSoundIDE(MakeSoundIDE As Boolean): bMakeSoundIDE = MakeSoundIDE: End Property
Public Property Get LogFile() As String: LogFile = sLogFile: End Property
Public Property Let LogFile(LogFile As String): sLogFile = LogFile: End Property

'=======================================================
'renvoie le nombre d'erreurs dans le fichier log
'=======================================================
Public Property Get NumberOfErrorInLogFile() As Long
Dim lFile As Long
Dim Cnt As Long
Dim tmp As String
    
    If FileExists(sLogFile) = False Then
        'alors pas de fichier
        NumberOfErrorInLogFile = 0
        Exit Property
    End If
    
    Cnt = 0
    lFile = FreeFile
    Open sLogFile For Input As lFile
    While Not EOF(lFile)
        Input #lFile, tmp
        If InStr(1, tmp, "Date=[") <> 0 Then Cnt = Cnt + 1    'si contient la string "Date=["
        'alors c'était la première ligne de description d'une erreur
    Wend
    
    Close lFile
    
    NumberOfErrorInLogFile = Cnt
        
End Property

'=======================================================
'renvoie les erreurs sous forme de tableau
'=======================================================
Public Property Get GetErrors() As Variant
Dim tmp() As clsErr
Dim tmpText As String
Dim lFile As Long
Dim x As Long, y As Long

    'ouvre le fichier log et obtient une string qui contient tout le texte
    If FileExists(sLogFile) = False Then
        GetErrors = 0
        Exit Property
    End If
    
    'redimensionne au nombre d'erreur dans le fichier
    ReDim tmp(Me.NumberOfErrorInLogFile)
    If UBound(tmp()) = 0 Then Exit Property
    
    'ouvre le fichier
    lFile = FreeFile: y = 0
    Open sLogFile For Input As lFile
    'x est un compteur qui détermine le type de la ligne qui sera prise dans le fichier
    '1 ==> Date
    '2 ==> Heure
    '3 ==> Numéro d'erreur
    '4 ==> Source
    '5 ==> description
    '6 ==> zone
    For y = 1 To UBound(tmp())
        Input #lFile, tmpText   'la première ligne est vierge
        Set tmp(y) = New clsErr
        For x = 1 To 6
            Input #lFile, tmpText   'obtient la ligne
            Select Case x   'en fonction du numéro de la ligne
                Case 1
                    tmp(y).ErrDate = tmpText
                 Case 2
                    tmp(y).ErrTime = tmpText
                Case 3
                    tmp(y).ErrNumber = Val(Mid$(tmpText, 15, Len(tmpText) - 15))
                Case 4
                    tmp(y).ErrSource = tmpText
                Case 5
                    tmp(y).ErrDescription = tmpText
                Case 6
                    tmp(y).ErrZone = tmpText
                End Select
        Next x
    Next y
        
    Close lFile 'ferme le fichier
    GetErrors = tmp()  'affecte au Variant
    
End Property



'//SUB & FUNCTIONS PRIVEES

'=======================================================
'obtient le dossier fils du dossier sKnownFolder
'=======================================================
Private Function GetFolderAfter(ByVal sPath As String, ByVal sKnownFolder) As String
Dim l As Long
Dim l2 As Long

    l = InStr(Len(sKnownFolder) + 1, sPath, "\", vbBinaryCompare)
    
    If l = 0 Then
        'pas de path
        GetFolderAfter = sKnownFolder
        Exit Function
    End If
    
    l2 = InStr(l, sPath, "\", vbBinaryCompare)
    
    GetFolderAfter = Left$(sPath, l2)
    
End Function

'=======================================================
'test de l'existence d'un répertoire
'=======================================================
Private Function FolderExists(strFolder As String) As Boolean
    FolderExists = CBool(PathIsDirectory(strFolder))
End Function

'=======================================================
'obtient un dossier (le prédécent) depuis un path
'=======================================================
Private Function GetPrevFolder(ByVal sPath As String) As String
Dim l As Long

    If Right$(sPath, 1) = "\" Then sPath = Left$(sPath, Len(sPath) - 1)
    
    l = InStrRev(sPath, "\", , vbBinaryCompare)
    
    If l = 0 Then
        'pas de path
        GetPrevFolder = sPath
        Exit Function
    End If
    
    GetPrevFolder = Left$(sPath, l)
    
End Function

'=======================================================
'obtient un dossier depuis un path
'=======================================================
Private Function GetFolderFromPath(ByVal sPath As String) As String
Dim l As Long

    l = InStrRev(sPath, "\", , vbBinaryCompare)
    
    If l = 0 Then
        'pas de path
        GetFolderFromPath = sPath
        Exit Function
    End If
    
    GetFolderFromPath = Left$(sPath, l)
    
End Function

'=======================================================
'teste l'existence un fichier
'=======================================================
Public Function FileExists(ByVal sFile As String) As Boolean
    FileExists = CBool(PathFileExists(sFile) And Not (FolderExists(sFile)))
End Function
