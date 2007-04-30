VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
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
'//CLASSE DE GESTION DU REGISTRE
'=======================================================

'=======================================================
'CONSTANTES
'=======================================================
Private Const REG_OPTION_BACKUP_RESTORE         As Long = 4      ' open for backup or restore
Private Const REG_OPTION_VOLATILE               As Long = 1             ' Key is not preserved when system is rebooted
Private Const REG_OPTION_NON_VOLATILE           As Long = 0        ' Key is preserved when system is rebooted
Private Const STANDARD_RIGHTS_ALL               As Long = &H1F0000
Private Const SYNCHRONIZE                       As Long = &H100000
Private Const READ_CONTROL                      As Long = &H20000
Private Const STANDARD_RIGHTS_READ              As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE             As Long = (READ_CONTROL)
Private Const KEY_CREATE_LINK                   As Long = &H20
Private Const KEY_CREATE_SUB_KEY                As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS            As Long = &H8
Private Const KEY_NOTIFY                        As Long = &H10
Private Const KEY_QUERY_VALUE                   As Long = &H1
Private Const KEY_SET_VALUE                     As Long = &H2
Private Const KEY_READ                          As Long = ((STANDARD_RIGHTS_READ Or _
    KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE                         As Long = ((STANDARD_RIGHTS_WRITE Or _
    KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE                       As Long = (KEY_READ)
Private Const KEY_ALL_ACCESS                    As Long = ((STANDARD_RIGHTS_ALL Or _
    KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_NO_MORE_ITEMS               As Long = 259&
Private Const BUFFER_SIZE                       As Long = 255


'=======================================================
'ENUM & TYPES
'=======================================================
'enum de sélection de clé racine
Public Enum KEY_TYPE
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum
Public Enum REGKEY_TYPE
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
End Enum

'=======================================================
'APIS
'=======================================================
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


'=======================================================
'récupère une valeur d'une clé
'exemple d'utilisation :
'    Dim sPath As String
'    Dim hKey As Long
'        sPath = "SYSTEM\CurrentControlSet\Services\LeService"
'        If RegOpenKey(HKEY_LOCAL_MACHINE, sPath, hKey) = 0 Then
'            RetrieveServiceInfo = RegQueryStringValue(hKey, sValeur)
'            RegCloseKey hKey
'        End If
'=======================================================
Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
Dim lValueType As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim lResult As Long
Dim strData As String
Dim lData As Long

    'obtiention des infos sur la clé
    If 0 = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0&, lDataBufSize) Then
        If (lValueType = REG_SZ Or lValueType = REG_MULTI_SZ Or lValueType = REG_EXPAND_SZ) Then
            'valeur REG_SZ ou multi_SZ
            'buffer
            strBuf = Space$(lDataBufSize)
            'contenu de la clé
            lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'formatage de la string
                RegQueryStringValue = Left$(strBuf, InStr(strBuf, vbNullChar) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            'valeur BINARY
            Dim strData2 As Integer
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData2, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData2
            End If
        ElseIf lValueType = REG_DWORD Then
            'valeur dword
            lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, lData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = CStr(lData)
            End If
        End If
    End If
End Function

'=======================================================
'obtient la valeur d'une clé registre
'exemple : ReadValue(HKEY_CURRENT_USER, "Le path de la clé", "Le nom de la valeur de clé")
'=======================================================
Public Function ReadValue(ByVal hKey As KEY_TYPE, ByVal strPath As String, ByVal strValue As String) As String
Dim Ret As Long

    'ouvre la clé
    Call RegOpenKey(hKey, strPath, Ret)
    
    'récupère la valeur de la clé
    ReadValue = RegQueryStringValue(Ret, strValue)
    
    'referme la clé
    Call RegCloseKey(Ret)
End Function

'=======================================================
'sauvegarde une valeur de clé
'exemple : WriteValue HKEY_CURRENT_USER, "folder1\folder2", "Nom de la valeur", "Contenu de la valeur de clé", REG_BINARY
'=======================================================
Public Function WriteValue(ByVal hKey As KEY_TYPE, ByVal strPath As String, ByVal strValue As String, ByVal strData As String, ByVal tType As REGKEY_TYPE) As Long
Dim Ret As Long

    'créé la nouvelle clé
    Call RegCreateKey(hKey, strPath, Ret)
    
    'sauvegarde la clé
    WriteValue = RegSetValueEx(Ret, strValue, 0, tType, ByVal strData, Len(strData))
    
    'la referme
    Call RegCloseKey(Ret)
End Function

'=======================================================
'supprime une valeur de clé
'exemple : DelKeyValue HKEY_CURRENT_USER, "folder1\folder2", "Valeur de clé"
'=======================================================
Public Function DelKeyValue(ByVal hKey As KEY_TYPE, ByVal strPath As String, ByVal strValue As String) As Long
Dim Ret As Long
   
    'ouvre la clé
    Call RegOpenKey(hKey, strPath, Ret)
    
    'supprime la valeur de clé
    DelKeyValue = RegDeleteValue(Ret, strValue)
    
    'referme la clé
    Call RegCloseKey(Ret)
End Function

'=======================================================
'supprime une clé
'exemple : DelKey HKEY_USERS, "folder1\folder2"
'=======================================================
Public Function DelKey(ByVal hKey As KEY_TYPE, ByVal strPath As String) As Long
    DelKey = RegDeleteKey(hKey, strPath)
End Function

'=======================================================
'créé une nouvelle clé
'exemple : CreateKey HKEY_USERS, "folder1\folder2"
'=======================================================
Public Function CreateKey(ByVal hKey As KEY_TYPE, ByVal strPath As String) As Long
Dim result As Long
Dim Ret As Long

    'vérifie l'existence de la clé
    If DoesKeyExist(hKey, strPath) Then Exit Function
    
    'la clé n'existe pas, donc on la créé
    CreateKey = RegCreateKeyEx(hKey, strPath, 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, ByVal 0&, result, Ret)

    'referme la clé
    Call RegCloseKey(result)
End Function

'=======================================================
'teste la présence d'une clé
'=======================================================
Public Function DoesKeyExist(ByVal hKey As KEY_TYPE, ByVal strPath As String) As Boolean
Dim result As Long

    'ouvre la clé
    Call RegOpenKey(hKey, strPath, result)
    
    'affecte le résultat à la fonction
    DoesKeyExist = CBool(result)
    
    'referme la clé
    Call RegCloseKey(result)
End Function

'=======================================================
'énumère les sous clés
'stocke de 1 à ubound
'=======================================================
Public Sub GetSubKeys(ByVal hKey As KEY_TYPE, ByVal strPath As String, _
    ByRef SubKeys() As String)
    
Dim Cnt As Long
Dim sName As String
Dim sData As String
Dim Ret As Long
Dim RetData As Long
Dim result As Long
Dim s() As String

    ReDim s(0)
    Ret = BUFFER_SIZE
    
    'ouvre la clé
    If RegOpenKey(hKey, strPath, result) = 0 Then
    
        'créé un buffer
        sName = Space(BUFFER_SIZE)
        
        'énumère les sous clé
        While RegEnumKeyEx(result, Cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            
            'récupère la clé
            ReDim Preserve s(UBound(s()) + 1)
            s(UBound(s())) = Left$(sName, Ret)
            
            'clé suivante
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
        Wend
        
        'referme la clé
        Call RegCloseKey(result)
    End If
    
    'affecte le tableau
    SubKeys = s
End Sub

'=======================================================
'énumère les valeurs de clé
'=======================================================
Public Sub GetKeyValues(ByVal hKey As KEY_TYPE, ByVal strPath As String, _
    ByRef KeyValues() As String, ByRef KeyValuesData() As String)
    
Dim Cnt As Long
Dim sName As String
Dim sData As String
Dim Ret As Long
Dim RetData As Long
Dim result As Long
Dim s() As String
Dim s2() As String

    ReDim s(0)
    ReDim s2(0)
    
    'ouvre la clé
    If RegOpenKey(hKey, strPath, result) = 0 Then
    
        'buffer
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        
        'énumère les valeurs
        While RegEnumValue(result, Cnt, sName, Ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            
            'récupère les values
            If RetData > 0 Then
                ReDim Preserve s(UBound(s()) + 1)
                ReDim Preserve s2(UBound(s2()) + 1)
                s(UBound(s())) = Left$(sName, Ret)
                s2(UBound(s2())) = Left$(sData, RetData - 1)
            End If
            
            'clé suivante
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        
        'referme la clé
        Call RegCloseKey(hKey)
    End If
    
    'affectation des tableaux
    KeyValues = s
    KeyValuesData = s2
End Sub


