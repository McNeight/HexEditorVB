VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniForm"
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
'//CLASS DE GESTION DU FICHIER DE CONFIGURATION
'=======================================================

'=======================================================
'APIS
'=======================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'=======================================================
'définit les préférences du fichier *.ini
'=======================================================
Private Sub LetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sValeur As String, ByVal sFile As String)
    Call WritePrivateProfileString(sSection, sVariable, sValeur, sFile)
End Sub

'=======================================================
'permet de récupérer les préférences du fichier *.ini
'=======================================================
Private Function GetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sFile As String) As String
Dim lng As Integer
Dim sReturn As String

    'créé le buffer
    sReturn = String$(255, 0)
    
    'obtient la valeur
    lng = GetPrivateProfileString(sSection, sVariable, vbNullString, sReturn, 255, sFile)
    
    'formate la string
    GetPref = Left$(sReturn, lng)
End Function

'=======================================================
'permet de sauver le fichier *.ini dans le fichier strIniFile
'A MODIFIER EN FONCTION DE VOS BESOINS !
'=======================================================
Public Sub SaveFormSettings(ByVal strIniFile As String, Frm As Form)
Dim Obj As Control

    On Error Resume Next
    
        For Each Obj In Frm.Controls
        
            With Obj
            
            'mettre "pref..." en tag pour autoriser la sauvegarde
            'ajouter une string différente après "pref" pour différencier deux controles indexés
            If Left$(LCase$(.Tag), 4) = "pref" Then
                'teste le type de obj, et sauvegarde les propriétés que vous
                'voulez concernant ce composant
                If TypeOf Obj Is TextBox Then
                    LetPref .Name & .Tag, "text", .Text, strIniFile
                ElseIf TypeOf Obj Is CheckBox Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is vkCheck Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is OptionButton Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is vkOptionButton Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is ProgressBar Then
                    LetPref .Name & .Tag, "min", .Min, strIniFile
                    LetPref .Name & .Tag, "max", .Max, strIniFile
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is HScrollBar Then
                    LetPref .Name & .Tag, "min", .Min, strIniFile
                    LetPref .Name & .Tag, "max", .Max, strIniFile
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is VScrollBar Then
                    LetPref .Name & .Tag, "min", .Min, strIniFile
                    LetPref .Name & .Tag, "max", .Max, strIniFile
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is DriveListBox Then
                    LetPref .Name & .Tag, "drive", .Drive, strIniFile
                ElseIf TypeOf Obj Is DirListBox Then
                    LetPref .Name & .Tag, "path", .Path, strIniFile
                ElseIf TypeOf Obj Is Slider Then
                    LetPref .Name & .Tag, "min", .Min, strIniFile
                    LetPref .Name & .Tag, "max", .Max, strIniFile
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is CommonDialog Then
                    LetPref .Name & .Tag, "initdir", .initdire, strIniFile
                    LetPref .Name & .Tag, "filename", .Filename, strIniFile
                ElseIf TypeOf Obj Is FileListBox Then
                    LetPref .Name & .Tag, "path", .Path, strIniFile
                ElseIf TypeOf Obj Is TabStrip Then
                    LetPref .Name & .Tag, "SelectedItemIndex", .SelectedItem.Index, strIniFile
                ElseIf TypeOf Obj Is ComboBox Then
                    LetPref .Name & .Tag, "text", .Text, strIniFile
                ElseIf TypeOf Obj Is vkOptionButton Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is vkCheck Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                ElseIf TypeOf Obj Is vkToggleButton Then
                    LetPref .Name & .Tag, "value", .Value, strIniFile
                End If
            End If
                        
        End With
    Next Obj
    
    If Left$(LCase$(Obj.Tag), 4) = "pref" Then
        'ici, il faut aussi gérer les propriété de la form en elle même
        With Frm
            LetPref .Name, "height", .Height, strIniFile
            LetPref .Name, "width", .Width, strIniFile
            LetPref .Name, "top", .Top, strIniFile
            LetPref .Name, "left", .Left, strIniFile
            LetPref .Name, "WindowState", .WindowState, strIniFile
        End With
    End If
    
End Sub

'=======================================================
'permet de récupérer le fichier *ini ou de le créer le cas échéant
'A MODIFIER EN FONCTION DE VOS BESOINS !
'=======================================================
Public Sub GetFormSettings(ByVal strIniFile As String, Frm As Form)
Dim Obj As Control

    On Error Resume Next

        For Each Obj In Frm.Controls
        
            With Obj
        
            'mettre "pref..." en tag pour autoriser la sauvegarde
            'ajouter une string différente après "pref" pour différencier deux controles indexés
            If Left$(LCase$(.Tag), 4) = "pref" Then
                'teste le type de controle (rajouter les autres controles que vous voulez supporter)
                'et change les propriétés
                If TypeOf Obj Is TextBox Then
                    .Text = GetPref(.Name & .Tag, "text", strIniFile)
                ElseIf TypeOf Obj Is CheckBox Then
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is vkCheck Then
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is OptionButton Then
                    .Value = CBool(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is vkOptionButton Then
                    .Value = CBool(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is ProgressBar Then
                    .Min = CLng(GetPref(.Name & .Tag, "min", strIniFile))
                    .Max = CLng(GetPref(.Name & .Tag, "max", strIniFile))
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is HScrollBar Then
                    .Min = CLng(GetPref(.Name & .Tag, "min", strIniFile))
                    .Max = CLng(GetPref(.Name & .Tag, "max", strIniFile))
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is VScrollBar Then
                    .Min = CLng(GetPref(.Name & .Tag, "min", strIniFile))
                    .Max = CLng(GetPref(.Name & .Tag, "max", strIniFile))
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is DriveListBox Then
                    .Drive = GetPref(.Name & .Tag, "drive", strIniFile)
                ElseIf TypeOf Obj Is DirListBox Then
                    .Path = GetPref(.Name & .Tag, "path", strIniFile)
                ElseIf TypeOf Obj Is Slider Then
                    .Min = CLng(GetPref(.Name & .Tag, "min", strIniFile))
                    .Max = CLng(GetPref(.Name & .Tag, "max", strIniFile))
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is CommonDialog Then
                    .InitDir = GetPref(.Name & .Tag, "initdir", strIniFile)
                    .Filename = GetPref(.Name & .Tag, "filename", strIniFile)
                ElseIf TypeOf Obj Is FileListBox Then
                    .Path = GetPref(.Name & .Tag, "path", strIniFile)
                ElseIf TypeOf Obj Is TabStrip Then
                    .Tabs.Item(CLng(GetPref(.Name & .Tag, "SelectedItemIndex", strIniFile))).Selected = True
                ElseIf TypeOf Obj Is ComboBox Then
                    .Text = GetPref(.Name & .Tag, "text", strIniFile)
                ElseIf TypeOf Obj Is vkOptionButton Then
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is vkCheck Then
                    .Value = CLng(GetPref(.Name & .Tag, "value", strIniFile))
                ElseIf TypeOf Obj Is vkToggleButton Then
                    .Value = CBool(GetPref(.Name & .Tag, "value", strIniFile))
                End If
            End If
                        
        End With
    Next Obj
    
    If Left$(LCase$(Obj.Tag), 4) = "pref" Then
        'pareil avec la form en elle même
        With Frm
            .Height = CLng(GetPref(.Name, "height", strIniFile))
            .Width = CLng(GetPref(.Name, "width", strIniFile))
            .Top = CLng(GetPref(.Name, "top", strIniFile))
            .Left = CLng(GetPref(.Name, "left", strIniFile))
            .WindowState = CLng(GetPref(.Name, "WindowState", strIniFile))
        End With
    End If

End Sub

