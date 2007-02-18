VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'//CLASS DE GESTION DU FICHIER DE CONFIGURATION
'-------------------------------------------------------

Public sDefaultPath As String   'path par défaut


'-------------------------------------------------------
'APIS
'-------------------------------------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'-------------------------------------------------------
'définit les préférences du fichier *.ini
'-------------------------------------------------------
Private Sub LetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sValeur As String, ByVal sFile As String)
    WritePrivateProfileString sSection, sVariable, sValeur, sFile
End Sub

'-------------------------------------------------------
'permet de récupérer les préférences du fichier *.ini
'-------------------------------------------------------
Private Function GetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sFile As String) As String
Dim lng As Integer
Dim sReturn As String

    'créé le buffer
    sReturn = String$(255, 0)
    
    'obtient la valeur
    lng = GetPrivateProfileString(sSection, sVariable, "", sReturn, 255, sFile)
    
    'formate la string
    GetPref = Left$(sReturn, lng)
End Function

'-------------------------------------------------------
'permet de sauver le fichier *.ini dans le fichier sFile
'-------------------------------------------------------
Public Function SaveIniFile(ByVal IniFile As clsIniPref) As Boolean
Dim strIniFile As String

    On Error GoTo GetError
    
    SaveIniFile = False
    strIniFile = IniFile.IniFilePath

    'sauvegarde les infos dans le fichier
    With IniFile
        LetPref "Appearance", "BackGroundColor", .app_BackGroundColor, .IniFilePath
        LetPref "Appearance", "OffsetForeColor", .app_OffsetForeColor, .IniFilePath
        LetPref "Appearance", "HexaForeColor", .app_HexaForeColor, .IniFilePath
        LetPref "Appearance", "StringsForeColor", .app_StringsForeColor, .IniFilePath
        LetPref "Appearance", "BaseForeColor", .app_BaseForeColor, .IniFilePath
        LetPref "Appearance", "TitleBackGroundColor", .app_TitleBackGroundColor, .IniFilePath
        LetPref "Appearance", "LinesColor", .app_LinesColor, .IniFilePath
        LetPref "Appearance", "SelectionColor", .app_SelectionColor, .IniFilePath
        LetPref "Appearance", "ModifiedItems", .app_ModifiedItems, .IniFilePath
        LetPref "Appearance", "SelectedItems", .app_SelectedItems, .IniFilePath
        LetPref "Appearance", "BookMarkColor", .app_BookMarkColor, .IniFilePath
        LetPref "Appearance", "ModifiedSelectedItems", .app_ModifiedSelectedItems, .IniFilePath
        LetPref "Appearance", "Grid", .app_Grid, .IniFilePath
        LetPref "Integration", "FileContextual", .integ_FileContextual, .IniFilePath
        LetPref "Integration", "FolderContextual", .integ_FolderContextual, .IniFilePath
        LetPref "Integration", "SendTo", .integ_SendTo, .IniFilePath
        LetPref "General", "MaximizeWhenOpen", .general_MaximizeWhenOpen, .IniFilePath
        LetPref "General", "DisplayIcon", .general_DisplayIcon, .IniFilePath
        LetPref "General", "DisplayInfos", .general_DisplayInfos, .IniFilePath
        LetPref "General", "DisplayData", .general_DisplayData, .IniFilePath
        LetPref "General", "OpenSubFiles", .general_OpenSubFiles, .IniFilePath
        LetPref "General", "CloseHomeWhenChosen", .general_CloseHomeWhenChosen, .IniFilePath
        LetPref "General", "ResoX", .general_ResoX, .IniFilePath
        LetPref "General", "ResoY", .general_ResoY, .IniFilePath
        LetPref "General", "AllowMultipleInstances", .general_AllowMultipleInstances, .IniFilePath
        LetPref "General", "DoNotChangeDates", .general_DoNotChangeDates, .IniFilePath
        LetPref "General", "Splash", .general_Splash, .IniFilePath
        LetPref "Environnement", "OS", .env_OS, .IniFilePath  '0=XP,1=Vista
        LetPref "Environnement", "Lang", .env_Lang, .IniFilePath
        LetPref "Historique", "NumberOfSave", .histo_NumberOfSave, .IniFilePath
        LetPref "FileExplorer", "ShowPath", .explo_ShowPath, .IniFilePath
        LetPref "FileExplorer", "ShowHiddenFiles", .explo_ShowHiddenFiles, .IniFilePath
        LetPref "FileExplorer", "ShowHiddenFolders", .explo_ShowHiddenFolders, .IniFilePath
        LetPref "FileExplorer", "ShowSystemFiles", .explo_ShowSystemFiles, .IniFilePath
        LetPref "FileExplorer", "ShowSystemFodlers", .explo_ShowSystemFodlers, .IniFilePath
        LetPref "FileExplorer", "ShowROFiles", .explo_ShowROFiles, .IniFilePath
        LetPref "FileExplorer", "ShowROFolders", .explo_ShowROFolders, .IniFilePath
        LetPref "FileExplorer", "AllowMultipleSelection", .explo_AllowMultipleSelection, .IniFilePath
        LetPref "FileExplorer", "AllowFileSuppression", .explo_AllowFileSuppression, .IniFilePath
        LetPref "FileExplorer", "AllowFolderSuppression", .explo_AllowFolderSuppression, .IniFilePath
        LetPref "FileExplorer", "IconType", .explo_IconType, .IniFilePath
        LetPref "FileExplorer", "DefaultPath", .explo_DefaultPath, .IniFilePath
        LetPref "FileExplorer", "Pattern", .explo_Pattern, .IniFilePath
        LetPref "FileExplorer", "Height", .explo_Height, .IniFilePath
        LetPref "FileExplorer", "HideColumnTitle", .explo_HideColumnTitle, .IniFilePath
        LetPref "Executable", "HasCrashed", .exe_HasCrashed, .IniFilePath
        LetPref "Appearance", "OffsetTitleForeColor", .app_OffsetTitleForeColor, .IniFilePath
    End With
    
    SaveIniFile = True

GetError:
End Function

'-------------------------------------------------------
'permet de récupérer le fichier *ini ou de le créer le cas échéant
'-------------------------------------------------------
Public Function GetIniFile(Optional ByVal sFile As String) As clsIniPref
Dim tCls As clsIniPref
Dim strIniFile As String

    Set tCls = New clsIniPref

    'path par défaut si pas précisé
    If Len(sFile) = 0 Then sFile = Me.sDefaultPath
    
    'obtient les infos sur le fichier
    With tCls
        .app_BackGroundColor = Val(GetPref("Appearance", "BackGroundColor", sFile))
        .app_OffsetForeColor = Val(GetPref("Appearance", "OffsetForeColor", sFile))
        .app_HexaForeColor = Val(GetPref("Appearance", "HexaForeColor", sFile))
        .app_StringsForeColor = Val(GetPref("Appearance", "StringsForeColor", sFile))
        .app_BaseForeColor = Val(GetPref("Appearance", "BaseForeColor", sFile))
        .app_TitleBackGroundColor = Val(GetPref("Appearance", "TitleBackGroundColor", sFile))
        .app_LinesColor = Val(GetPref("Appearance", "LinesColor", sFile))
        .app_SelectionColor = Val(GetPref("Appearance", "SelectionColor", sFile))
        .app_ModifiedItems = Val(GetPref("Appearance", "ModifiedItems", sFile))
        .app_SelectedItems = Val(GetPref("Appearance", "SelectedItems", sFile))
        .app_BookMarkColor = Val(GetPref("Appearance", "BookMarkColor", sFile))
        .app_ModifiedSelectedItems = Val(GetPref("Appearance", "ModifiedSelectedItems", sFile))
        .app_Grid = Val(GetPref("Appearance", "Grid", sFile))
        .app_OffsetTitleForeColor = Val(GetPref("Appearance", "OffsetTitleForeColor", sFile))
        .integ_FileContextual = (Val(GetPref("Integration", "FileContextual", sFile)))
        .integ_FolderContextual = (Val(GetPref("Integration", "FolderContextual", sFile)))
        .integ_SendTo = (Val(GetPref("Integration", "SendTo", sFile)))
        .general_MaximizeWhenOpen = (Val(GetPref("General", "MaximizeWhenOpen", sFile)))
        .general_DisplayIcon = (Val(GetPref("General", "DisplayIcon", sFile)))
        .general_DisplayInfos = (Val(GetPref("General", "DisplayInfos", sFile)))
        .general_DisplayData = (Val(GetPref("General", "DisplayData", sFile)))
        .general_Splash = (Val(GetPref("General", "Splash", sFile)))
        .general_ResoX = Val(GetPref("General", "ResoX", sFile))
        .general_ResoY = Val(GetPref("General", "ResoY", sFile))
        .general_AllowMultipleInstances = (Val(GetPref("General", "AllowMultipleInstances", sFile)))
        .general_DoNotChangeDates = (Val(GetPref("General", "DoNotChangeDates", sFile)))
        .env_OS = Val(GetPref("Environnement", "OS", sFile))
        .env_Lang = GetPref("Environnement", "Lang", sFile)
        .histo_NumberOfSave = Val(GetPref("Historique", "NumberOfSave", sFile))
        .explo_ShowPath = (Val(GetPref("FileExplorer", "ShowPath", sFile)))
        .explo_ShowHiddenFiles = (Val(GetPref("FileExplorer", "ShowHiddenFiles", sFile)))
        .explo_ShowHiddenFolders = (Val(GetPref("FileExplorer", "ShowHiddenFolders", sFile)))
        .explo_ShowSystemFiles = (Val(GetPref("FileExplorer", "ShowSystemFiles", sFile)))
        .explo_ShowSystemFodlers = (Val(GetPref("FileExplorer", "ShowSystemFodlers", sFile)))
        .explo_ShowROFiles = (Val(GetPref("FileExplorer", "ShowROFiles", sFile)))
        .explo_ShowROFolders = (Val(GetPref("FileExplorer", "ShowROFolders", sFile)))
        .explo_AllowMultipleSelection = (Val(GetPref("FileExplorer", "AllowMultipleSelection", sFile)))
        .explo_AllowFileSuppression = (Val(GetPref("FileExplorer", "AllowFileSuppression", sFile)))
        .explo_AllowFolderSuppression = (Val(GetPref("FileExplorer", "AllowFolderSuppression", sFile)))
        .explo_IconType = Val(GetPref("FileExplorer", "IconType", sFile))
        .explo_DefaultPath = GetPref("FileExplorer", "DefaultPath", sFile)
        .explo_Pattern = GetPref("FileExplorer", "Pattern", sFile)
        .explo_Height = Val(GetPref("FileExplorer", "Height", sFile))
        .explo_HideColumnTitle = (Val(GetPref("FileExplorer", "HideColumnTitle", sFile)))
        .exe_HasCrashed = (Val(GetPref("Executable", "HasCrashed", sFile)))
        .general_OpenSubFiles = Val(GetPref("General", "OpenSubFiles", sFile))
        .general_CloseHomeWhenChosen = Val(GetPref("General", "CloseHomeWhenChosen", sFile))
        .IniFilePath = sFile
    End With
        
    Set GetIniFile = tCls
    Set tCls = Nothing

End Function
