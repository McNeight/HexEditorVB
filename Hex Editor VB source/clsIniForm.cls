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
    WritePrivateProfileString sSection, sVariable, sValeur, sFile
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
        
        'mettre "pref..." en tag pour autoriser la sauvegarde
        'ajouter une string différente après "pref" pour différencier deux controles indexés
        If Left$(LCase$(Obj.Tag), 4) = "pref" Then
            'teste le type de obj, et sauvegarde les propriétés que vous
            'voulez concernant ce composant
            If TypeOf Obj Is TextBox Then
                LetPref Obj.Name & Obj.Tag, "text", Obj.Text, strIniFile
            ElseIf TypeOf Obj Is CheckBox Then
                LetPref Obj.Name & Obj.Tag, "value", Obj.Value, strIniFile
            ElseIf TypeOf Obj Is OptionButton Then
                LetPref Obj.Name & Obj.Tag, "value", Obj.Value, strIniFile
            ElseIf TypeOf Obj Is ProgressBar Then
                LetPref Obj.Name & Obj.Tag, "min", Obj.Min, strIniFile
                LetPref Obj.Name & Obj.Tag, "max", Obj.Max, strIniFile
                LetPref Obj.Name & Obj.Tag, "value", Obj.Value, strIniFile
            ElseIf TypeOf Obj Is HScrollBar Then
                LetPref Obj.Name & Obj.Tag, "min", Obj.Min, strIniFile
                LetPref Obj.Name & Obj.Tag, "max", Obj.Max, strIniFile
                LetPref Obj.Name & Obj.Tag, "value", Obj.Value, strIniFile
            ElseIf TypeOf Obj Is VScrollBar Then
                LetPref Obj.Name & Obj.Tag, "min", Obj.Min, strIniFile
                LetPref Obj.Name & Obj.Tag, "max", Obj.Max, strIniFile
                LetPref Obj.Name & Obj.Tag, "value", Obj.Value, strIniFile
            ElseIf TypeOf Obj Is DriveListBox Then
                LetPref Obj.Name & Obj.Tag, "drive", Obj.Drive, strIniFile
            ElseIf TypeOf Obj Is DirListBox Then
                LetPref Obj.Name & Obj.Tag, "path", Obj.Path, strIniFile
            ElseIf TypeOf Obj Is Slider Then
                LetPref Obj.Name & Obj.Tag, "min", Obj.Min, strIniFile
                LetPref Obj.Name & Obj.Tag, "max", Obj.Max, strIniFile
                LetPref Obj.Name & Obj.Tag, "value", Obj.Value, strIniFile
            ElseIf TypeOf Obj Is CommonDialog Then
                LetPref Obj.Name & Obj.Tag, "initdir", Obj.initdire, strIniFile
                LetPref Obj.Name & Obj.Tag, "filename", Obj.Filename, strIniFile
            ElseIf TypeOf Obj Is FileListBox Then
                LetPref Obj.Name & Obj.Tag, "path", Obj.Path, strIniFile
            ElseIf TypeOf Obj Is TabStrip Then
                LetPref Obj.Name & Obj.Tag, "SelectedItemIndex", Obj.SelectedItem.Index, strIniFile
            ElseIf TypeOf Obj Is ComboBox Then
                LetPref Obj.Name & Obj.Tag, "text", Obj.Text, strIniFile
            End If
        End If
                    
    Next Obj
    
    If Left$(LCase$(Obj.Tag), 4) = "pref" Then
        'ici, il faut aussi gérer les propriété de la form en elle même
        LetPref Frm.Name, "height", Frm.Height, strIniFile
        LetPref Frm.Name, "width", Frm.Width, strIniFile
        LetPref Frm.Name, "top", Frm.Top, strIniFile
        LetPref Frm.Name, "left", Frm.Left, strIniFile
        LetPref Frm.Name, "WindowState", Frm.WindowState, strIniFile
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
    
        'mettre "pref..." en tag pour autoriser la sauvegarde
        'ajouter une string différente après "pref" pour différencier deux controles indexés
        If Left$(LCase$(Obj.Tag), 4) = "pref" Then
            'teste le type de controle (rajouter les autres controles que vous voulez supporter)
            'et change les propriétés
            If TypeOf Obj Is TextBox Then
                Obj.Text = GetPref(Obj.Name & Obj.Tag, "text", strIniFile)
            ElseIf TypeOf Obj Is CheckBox Then
                Obj.Value = CLng(GetPref(Obj.Name & Obj.Tag, "value", strIniFile))
            ElseIf TypeOf Obj Is OptionButton Then
                Obj.Value = CBool(GetPref(Obj.Name & Obj.Tag, "value", strIniFile))
            ElseIf TypeOf Obj Is ProgressBar Then
                Obj.Min = CLng(GetPref(Obj.Name & Obj.Tag, "min", strIniFile))
                Obj.Max = CLng(GetPref(Obj.Name & Obj.Tag, "max", strIniFile))
                Obj.Value = CLng(GetPref(Obj.Name & Obj.Tag, "value", strIniFile))
            ElseIf TypeOf Obj Is HScrollBar Then
                Obj.Min = CLng(GetPref(Obj.Name & Obj.Tag, "min", strIniFile))
                Obj.Max = CLng(GetPref(Obj.Name & Obj.Tag, "max", strIniFile))
                Obj.Value = CLng(GetPref(Obj.Name & Obj.Tag, "value", strIniFile))
            ElseIf TypeOf Obj Is VScrollBar Then
                Obj.Min = CLng(GetPref(Obj.Name & Obj.Tag, "min", strIniFile))
                Obj.Max = CLng(GetPref(Obj.Name & Obj.Tag, "max", strIniFile))
                Obj.Value = CLng(GetPref(Obj.Name & Obj.Tag, "value", strIniFile))
            ElseIf TypeOf Obj Is DriveListBox Then
                Obj.Drive = GetPref(Obj.Name & Obj.Tag, "drive", strIniFile)
            ElseIf TypeOf Obj Is DirListBox Then
                Obj.Path = GetPref(Obj.Name & Obj.Tag, "path", strIniFile)
            ElseIf TypeOf Obj Is Slider Then
                Obj.Min = CLng(GetPref(Obj.Name & Obj.Tag, "min", strIniFile))
                Obj.Max = CLng(GetPref(Obj.Name & Obj.Tag, "max", strIniFile))
                Obj.Value = CLng(GetPref(Obj.Name & Obj.Tag, "value", strIniFile))
            ElseIf TypeOf Obj Is CommonDialog Then
                Obj.InitDir = GetPref(Obj.Name & Obj.Tag, "initdir", strIniFile)
                Obj.Filename = GetPref(Obj.Name & Obj.Tag, "filename", strIniFile)
            ElseIf TypeOf Obj Is FileListBox Then
                Obj.Path = GetPref(Obj.Name & Obj.Tag, "path", strIniFile)
            ElseIf TypeOf Obj Is TabStrip Then
                Obj.Tabs.Item(CLng(GetPref(Obj.Name & Obj.Tag, "SelectedItemIndex", strIniFile))).Selected = True
            ElseIf TypeOf Obj Is ComboBox Then
                Obj.Text = GetPref(Obj.Name & Obj.Tag, "text", strIniFile)
            End If
        End If

                    
    Next Obj
    
    If Left$(LCase$(Obj.Tag), 4) = "pref" Then
        'pareil avec la form en elle même
        Frm.Height = CLng(GetPref(Frm.Name, "height", strIniFile))
        Frm.Width = CLng(GetPref(Frm.Name, "width", strIniFile))
        Frm.Top = CLng(GetPref(Frm.Name, "top", strIniFile))
        Frm.Left = CLng(GetPref(Frm.Name, "left", strIniFile))
        Frm.WindowState = CLng(GetPref(Frm.Name, "WindowState", strIniFile))
    End If

End Sub

