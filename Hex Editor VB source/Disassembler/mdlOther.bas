Attribute VB_Name = "mdlOther"
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
'MODULE CONTENANT DES FONCTIONS DIVERSES
'=======================================================


'=======================================================
'récupère le nom de l'utilisateur
'=======================================================
Public Function GetUserName() As String
Dim strS As String
Dim ret As Long

    'créé un buffer
    strS = String$(200, 0)
    
    'récupère le Name
    ret = GetUserNameA(strS, 199)
    If ret <> 0 Then GetUserName = Left$(strS, 199) Else GetUserName = vbNullString
End Function

'=======================================================
'récupère la version de Windows
'=======================================================
Public Function GetWindowsVersion(Optional ByRef sWindowsVersion As String, Optional ByRef lBuildNumber As Long) As WINDOWS_VERSION
Dim OS As OSVERSIONINFO
Dim s As String, l As Long

    'taille de la structure
    OS.dwOSVersionInfoSize = Len(OS)
    
    'récupère l'info sur la version
    If GetVersionEx(OS) = 0 Then
        'échec
        sWindowsVersion = "Cannot retrieve information"
        GetWindowsVersion = UnKnown_OS
        Exit Function
    End If
        
    'numéro de la build
    lBuildNumber = OS.dwBuildNumber
    
    'récupère la version en fonction de Major et Minor
    Select Case OS.dwMajorVersion
        Case 6
            GetWindowsVersion = [Windows Vista]
            sWindowsVersion = "Windows Vista"
            Exit Function
        Case 5
            If OS.dwMinorVersion = 2 Then
                GetWindowsVersion = [Windows Server 2003]
                sWindowsVersion = "Windows Server 2003"
            ElseIf OS.dwMinorVersion = 1 Then
                GetWindowsVersion = [Windows XP]
                sWindowsVersion = "Windows XP"
            ElseIf OS.dwMinorVersion = 0 Then
                GetWindowsVersion = [Windows 2000]
                sWindowsVersion = "Windows 2000"
            End If
            Exit Function
        Case 4
            If OS.dwMinorVersion = 90 Then
                GetWindowsVersion = [Windows Me]
                sWindowsVersion = "Windows ME"
            ElseIf OS.dwMinorVersion = 10 Then
                GetWindowsVersion = [Windows 98]
                sWindowsVersion = "Windows 98"
            ElseIf OS.dwMinorVersion = 0 Then
                GetWindowsVersion = [Windows 95]
                sWindowsVersion = "Windows 95"
            End If
            Exit Function
    End Select
    
    GetWindowsVersion = [UnKnown_OS]
    
End Function

'=======================================================
'obtient le path du fichier temp à créer
'=======================================================
Public Function ObtainTempPath() As String
Dim sBuf As String

    '//obtient le répertoire temporaire
    'buffer
    sBuf = String$(256, vbNullChar)
    'obtient le dossier temp
    GetTempPath 256, sBuf
    'formate le path
    ObtainTempPath = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
    
End Function
