VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniPref"
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
'//TYPE PERSO
'CONTIENT LES SECTIONS ET VARIABLES DU FICHIER *INI
'C'EST CE FICHIER QU'IL FAUT EDITER POUR CHANGER LES PREFS
'=======================================================

'=======================================================
'variables publiques qui constituent le type de variable définit par la classe
'=======================================================
Public IniFilePath As String
Public app_BackGroundColor As Long
Public app_OffsetForeColor As Long
Public app_HexaForeColor As Long
Public app_StringsForeColor As Long
Public app_BaseForeColor As Long
Public app_OffsetTitleForeColor As Long
Public app_TitleBackGroundColor As Long
Public app_LinesColor As Long
Public app_OffsetsHex As Byte
Public app_SelectionColor As Long
Public app_ModifiedItems As Long
Public app_SelectedItems As Long
Public app_BookMarkColor As Long
Public app_ModifiedSelectedItems As Long
Public app_Grid As Byte
Public integ_FileContextual As Byte
Public integ_FolderContextual As Byte
Public integ_SendTo As Byte
Public general_MaximizeWhenOpen As Byte
Public general_DisplayIcon As Byte
Public general_DisplayExplore As Byte
Public general_DisplayConsole As Byte
Public general_DisplayInfos As Byte
Public general_DisplayData As Byte
Public general_ResoX As Long
Public general_ResoY As Long
Public general_QuickBackup As Byte
Public general_OpenSubFiles As Byte
Public general_CloseHomeWhenChosen As Byte
Public general_ShowAlert As Byte
Public general_AllowMultipleInstances As Byte
Public general_DoNotChangeDates As Byte
Public general_Splash As Byte
Public env_OS As Byte   '0=XP, 1=Vista
Public env_Lang As String
Public histo_NumberOfSave As Long
Public explo_ShowPath As Byte
Public explo_ShowHiddenFiles As Byte
Public explo_ShowHiddenFolders As Byte
Public explo_ShowSystemFiles As Byte
Public explo_ShowSystemFodlers As Byte
Public explo_ShowROFiles As Byte
Public explo_ShowROFolders As Byte
Public explo_AllowMultipleSelection As Byte
Public explo_AllowFileSuppression As Byte
Public explo_AllowFolderSuppression As Byte
Public explo_IconType As Byte
Public explo_DefaultPath As String
Public explo_Pattern As String
Public explo_Height As Long
Public explo_HideColumnTitle As Byte
Public exe_HasCrashed As Byte
Public console_ForeColor As Long
Public console_BackColor As Long
Public console_Heigth As Long
Public console_Load As Byte
 

'RESTE :
'env_Lang
'env_OS
'general_DoNotChangeDates
'general_AllowMultipleInstances
'general_ResoY/X (fixer bug qui ne change pas la taille)

