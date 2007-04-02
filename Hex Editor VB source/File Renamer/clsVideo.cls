VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =======================================================
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A Windows utility which allows to rename lots of file (part of Hex Editor VB)
'
' Copyright (c) 2006-2007 by Alain Descotes.
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
'CLASSE DE RECUPERATION DES INFOS SUR LES VIDEOS
'=======================================================


'=======================================================
'VARIABLES PUBLIQUES DE LA CLASSE
'=======================================================
Public strFile As String
Public lngWidth As Long
Public lngHeight As Long
Public lngLength As Long
Public lngRate As Long
Public lngScale As Long
Public lngStreams As Long
Public lngMaxBytesPerSec As Long
Public strFileType As String
Public strDuration As String
Public lngModifiedStreams As Long
Public lngSamplesPerSecond As Double
Public lngHour As Long
Public lngMinute As Long
Public lngSecond As Long


'=======================================================
'CONSTANTES
'=======================================================
Private Const OF_SHARE_DENY_WRITE               As Long = &H20

'=======================================================
'TYPES
'=======================================================
Private Type AVIFileInfo
    dwMaxBytesPerSec As Long
    dwFlags As Long
    dwCaps As Long
    dwStreams As Long
    dwSuggestedBufferSize As Long
    dwWidth As Long
    dwHeight As Long
    dwScale As Long
    dwRate As Long
    dwLength As Long
    dwEditCount As Long
    szFileType As String * 64
End Type

'=======================================================
'APIS
'=======================================================
Private Declare Function AVIFileOpen Lib "avifil32" Alias "AVIFileOpenA" (ppfile As Long, ByVal szFile As String, ByVal mode As Long, pclsidHandler As Any) As Long
Private Declare Function AVIFileRelease Lib "avifil32" (ByVal pfile As Long) As Long
Private Declare Function AVIFileInfo Lib "avifil32" Alias "AVIFileInfoA" (ByVal pfile As Long, pfi As AVIFileInfo, ByVal lSize As Long) As Long
Private Declare Sub AVIFileInit Lib "avifil32" ()
Private Declare Sub AVIFileExit Lib "avifil32" ()


'=======================================================
'PROCEDURES/FUNCTIONS PUBLIQUES
'=======================================================

'=======================================================
'récupère les infos
'=======================================================
Public Sub GetVideoInfo()
Dim hFile As Long
Dim AviInfo As AVIFileInfo

    'initialise la librairie AVI
    AVIFileInit
    
    'obtient le handle du fichier
    If AVIFileOpen(hFile, Me.strFile, OF_SHARE_DENY_WRITE, ByVal 0&) = 0 Then
        'c'est donc accessible
        
        If AVIFileInfo(hFile, AviInfo, Len(AviInfo)) = 0 Then
            'c'est bon, on récupère les infos
            With AviInfo
                lngWidth = .dwWidth
                lngHeight = .dwHeight
                lngLength = .dwLength
                lngRate = .dwRate
                lngScale = .dwScale
                lngStreams = .dwStreams
                lngMaxBytesPerSec = .dwMaxBytesPerSec
                strFileType = FormatedString(.szFileType)
                lngModifiedStreams = .dwEditCount
                lngSamplesPerSecond = .dwRate / .dwScale
                strDuration = FormatedDuration()
            End With
        End If
        
        'ferme le handle
        AVIFileRelease hFile
    End If
    
    'décharge la librairie
    AVIFileExit
End Sub

'=======================================================
'formate la durée
'=======================================================
Private Function FormatedDuration() As String
Dim s As String
    
    'secondes=nbre frames/fps
    lngSecond = Round(lngLength / lngSamplesPerSecond)
    
    lngHour = 0: lngMinute = 0
    While lngSecond - 3600 > 0
        lngSecond = lngSecond - 3600
        lngHour = lngHour + 1
    Wend
    While lngSecond - 60 > 0
        lngSecond = lngSecond - 60
        lngMinute = lngMinute + 1
    Wend

    If lngHour Then s = CStr(lngHour) & " h "
    If lngMinute Or lngHour Then s = s & IIf((lngMinute < 10) And lngHour, "0" & CStr(lngMinute), CStr(lngMinute)) & " min "
    If lngSecond Or lngMinute Or lngHour Then s = s & IIf((lngSecond < 10) And (lngMinute Or lngHour), "0" & CStr(lngSecond), CStr(lngSecond)) & " sec"
    FormatedDuration = s
End Function

'=======================================================--------
'formatage de string
'=======================================================--------
Private Function FormatedString(ByVal sString As String) As String

    If InStr(sString, vbNullChar) Then _
    FormatedString = Left$(sString, InStr(sString, vbNullChar) - 1) Else _
    FormatedString = sString
    
    FormatedString = Trim$(FormatedString)
End Function

