VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInternet"
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
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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
'CLASSE DE TELECHARGEMENT D'INFORMATIONS SUR INTERNET
'=======================================================


'=======================================================
'APIS
'=======================================================
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'=======================================================
'CONSTANTES
'=======================================================
'Private Const NO_UPDATE As String = "You are using the latest version of Hex Editor VB. No new update is available."
'Private Const CAN_NOT_CHECK As String = "Hex Editor VB is not able to retrieve informations on the Internet. Please check your Internet configuration or try another time."
'Private Const NEW_UPDATE As String = "A new version is available ! Download it now ?"
Private Const SOURCEFORGE_URL As String = "http://sourceforge.net/project/showfiles.php?group_id=186829"
Private Const GENERIC_READ                  As Long = &H80000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_SHARE_WRITE              As Long = &H2
Private Const OPEN_EXISTING                 As Long = 3
Private Const FILE_BEGIN                    As Long = 0


'=======================================================
'Renvoie une string contenant le r�sultat (-1 ==> rat�, 0 ==> rien, 1 ==> MAJ dispo)
'indiquant qu'il existe une nouvelle MAJ ou pas
'=======================================================
Public Function CheckUpdate(ByRef NotesFile As String, ByRef _
    NewVersion As String, ByRef sPageDL As String) As Long
    
Dim Ret As Long
Dim a As Long
Dim c As Long
Dim d As Long
Dim b As Long
Dim sPath As String
Dim S As String
Dim s2 As String

    On Error Resume Next
    
    'g�n�re un nom de fichier temporaire
    sPath = GetTempName("hevb_MAJ")
    
    'on proc�de au t�l�chargement de la page
    Ret = URLDownloadToFile(0, SOURCEFORGE_URL, sPath, 0, 0)
    
    If Ret <> 0 Then
        CheckUpdate = -1 '�chec
        Exit Function
    End If
    
    'on r�cup�re le fichier en m�moire
    S = LoadFileInString(sPath)
    
    'on d�truit le fichier temporaire
    Call Kill(sPath)
    
    '//on extrait les infos de ce fichier
    
    'r�cup�re la position d'une string unique dans la page
    a = InStr(1, S, "<tfoot><tr><td></td></tr></tfoot>", vbBinaryCompare)
    
    'on d�cale de 275 caract�res
    a = a + 275
    
    'on r�cup�re � partir de l� la string jusqu'au premier '<'
    b = InStr(a, S, "<", vbBinaryCompare)
    
    'on r�cup�re notre string
    s2 = Mid$(S, a, b - a)
    
    'on compare avec la string qui d�fini la version de Hex Editor VB
    If s2 = HEX_EDITOR_VB_VERSION Then
        'version actuelle
        CheckUpdate = 0
        Exit Function
    End If
    
    NewVersion = s2
    
    'si on est l�, c'est qu'il existe une nouvelle version !!
    'dans ce cas, il faut r�cup�rer le fichier de notes
    
    'r�cup�re l'URL du package
    c = InStr(a - 100, S, "<td><a href=", vbBinaryCompare)
    d = InStr(c, S, "title=", vbBinaryCompare)
    sPageDL = Replace$("http://sourceforge.net/" & Mid$(S, c + 14, d - c - 16), _
        "amp;", vbNullString)
    
    'obtient un nouveau fichier temporaire
    sPath = GetTempName("hevb_MAJ")
    
    'on r�cup�re le fichier de notes
    a = InStr(b, S, "<td><a href=", vbBinaryCompare)
    b = InStr(a, S, "title=", vbBinaryCompare)
    S = Mid$(S, a + 14, b - a - 16)
    
    NotesFile = Replace$(S, "amp;", vbNullString)
    CheckUpdate = 1
    
End Function

'=======================================================
'r�cup�re les notes de MAJ
'=======================================================
Public Function GetNotes(ByVal URL As String, ByRef Res As String) As Long
    
Dim a As Long
Dim b As Long
Dim S As String
Dim s2 As String
Dim sPath As String
Dim Ret As Long

    On Error Resume Next

    'obtient un nouveau fichier temporaire
    sPath = GetTempName("hevb_MAJ")

    'on proc�de au t�l�chargement du fichier de notes
    Ret = URLDownloadToFile(0, "http://sourceforge.net/" & URL, sPath, 0, 0)
    
    If Ret <> 0 Then
        GetNotes = -1
        Exit Function  'rat�
    End If
    
    'sinon on r�cup�re le fichier dans une nouvelle string
    S = LoadFileInString(sPath)
        
    'on d�truit le fichier temporaire
    Call Kill(sPath)
    
    '//on proc�de � nouveau � l'extraction des informations
    a = InStr(1, S, "HISTORY (ENGLISH)", vbBinaryCompare)
    b = InStr(a, S, "AUTEUR")
    
    If a = 0 Or b = 0 Then
        GetNotes = -1
        Exit Function  'rat�
    End If
    
    Res = Mid$(S, a + 81, b - a - 158)
    GetNotes = Ret
    
End Function

'=======================================================
'G�n�re un nom de fichier temporaire
'=======================================================
Private Function GetTempName(ByVal FileName As String) As String
Dim sBuf As String
Dim S As String

    '//obtient le r�pertoire temporaire
    sBuf = String$(256, vbNullChar)
    
    'obtient le dossier temp
    Call GetTempPath(256, sBuf)
    
    'formate le path
    sBuf = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
    
    '//obtient un path unique
    'buffer
    S = String$(256, vbNullChar)
    
    'obtient le dossier temp
    Call GetTempFileName(sBuf, FileName, 0, S)
    
    'formate le path
    S = Left$(S, InStr(S, vbNullChar) - 1)

    GetTempName = S
    
End Function

'=======================================================
'R�cup�re le contenu d'un fichier dans une string
'=======================================================
Private Function LoadFileInString(ByVal Source As String) As String
Dim hFile As Long
Dim lSize As Long
Dim Ret As Long
Dim S As String
    
    'obtient le handle du fichier
    hFile = CreateFile(Source, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    'obtient sa taille pour cr�er un buffer
    lSize = FileLen(Source)
 
    'initialise le pointeur sur le fichier (au d�but)
    Call SetFilePointer(hFile, 0, 0, FILE_BEGIN)
    
    'cr�� le buffer
    S = String$(lSize, 0)
    
    'charge le fichier
    Ret = ReadFile(hFile, ByVal S, lSize, Ret, ByVal 0&)
     
    'ferme le handle
    Call CloseHandle(hFile)
    
    LoadFileInString = S
    
End Function

'=======================================================
'Fonction permettant d'ouvrir un fichier
'=======================================================
Public Function ShellOpenFile(ByVal File As String, ByVal hWnd As Long, _
    Optional Parameters As String = vbNullString, Optional ByVal Directory _
    As String = vbNullString) As Long
    
    ShellOpenFile = ShellExecute(hWnd, "open", File, Parameters, _
        Directory, 1)
End Function

