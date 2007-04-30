VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLang"
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
'//CONTROLE DE CHANGEMENT DE LANGUE
'=======================================================


'AJOUTER LE TAG "lang_ok" POUR LES LISTBOX/COMBOX POUR EFFECTUER LA LECTURE DES ITEMS


'=======================================================
'APIS
'=======================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private sLanguage As String     'langue utilisée
Private sLangFolder As String    'path du répertoire de langue
Private mParent As Object   'form parente


'=======================================================
'EVENTS
'=======================================================
Public Event LanguageChanged(OldLanguage As String, NewLanguage As String)


'=======================================================
'PROPERTIES
'=======================================================
Public Property Get Language() As String: Language = sLanguage: End Property
Public Property Let Language(Language As String)
If sLanguage <> Language Then
    RaiseEvent LanguageChanged(sLanguage, Language)
    sLanguage = Language
    Call LoadControlsCaption 'met à jour les controles
End If
End Property
Public Property Get LangFolder() As String: LangFolder = sLangFolder: End Property
Public Property Let LangFolder(LangFolder As String): sLangFolder = LangFolder: End Property



'=======================================================
'CLASS SUBS
'=======================================================
Private Sub Class_Initialize()
    With Me
        .LangFolder = App.Path & "\Lang"
        .Language = "French"
    End With
End Sub
Private Sub Class_Terminate()
    Set mParent = Nothing   'libère la mémoire
End Sub




'=======================================================
'active le changement de langue
'=======================================================
Public Sub ActiveLang(Frm As Form)
    Set mParent = Frm
End Sub

'=======================================================
'permet de récupérer une string
'=======================================================
Public Function GetString(ByVal StringName As String) As String
    
Dim lng As Integer
Dim sReturn As String
Dim strFile As String
Dim Section As String
Dim sFile As String

    On Error Resume Next
    
    'créé le buffer
    sReturn = String$(255, 0)
        
    'détermine le fichier *.ini concerné
    sFile = Me.LangFolder & "\" & Me.Language & ".ini"
    
    'détermine la section
    Section = mParent.Name
    
    'obtient la valeur
    lng = GetPrivateProfileString(Section, StringName, vbNullString, sReturn, _
        255, sFile)
    
    'formate la string
    GetString = Left$(sReturn, lng)
End Function

'=======================================================
'définit les préférences du fichier *.ini
'=======================================================
Private Sub LetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sValeur As String, ByVal sFile As String)
    Call WritePrivateProfileString(sSection, sVariable, sValeur, sFile)
End Sub

'=======================================================
'affecte les propriétés Text et Caption pour tous les controles de la form
'=======================================================
Public Sub LoadControlsCaption()
Dim Obj As Control
Dim X As Long
Dim Y As Long
Dim s As String

    On Error Resume Next
    
    'If UserControl.Ambient.UserMode = False Then Exit Sub  'on est dans l'IDE
    
    'vérifie que le fichier existe bien...
    If FileLen(Me.LangFolder & "\" & Me.Language & ".ini") = 0 Then Exit Sub       'fichier inexistant, donc on loade rien

    For Each Obj In mParent.Controls
        'pour chaque controle, on récupère la valeur de son nom et on affecte
        'la string Caption/Texte au controle
        'possibilité d'ajouter facilement les autres controles à gérer (notemment les
        'usercontrols)
        
        s = Obj.Name & "|Index"
        s = s & Trim$(Str$(Obj.Index))  '==> deux concaténation car erreur sur la deuxième
        'ligne pour les composants non indexés
        
        'récupère le ToolTipText
        Obj.ToolTipText = Me.GetString(s & "|ToolTipText")
        
        If TypeOf Obj Is TextBox Then
            Obj.Text = Me.GetString(s & "|Text")
        ElseIf TypeOf Obj Is CheckBox Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is OptionButton Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is ComboBox Then
            
            'vérifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on loade
                
                'clear
                Obj.Clear
                
                'récupère le nombre d'éléments
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Obj.AddItem Me.GetString(s & "|Item" & Trim$(Str$(X)))
                Next X
                
            End If
            
        ElseIf TypeOf Obj Is Label Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is CommandButton Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is Frame Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is ListBox Then
            'alors là faut aussi sauvegarder les Items qu'il y a dedans si tag OK
            
            'vérifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on loade
                
                'clear
                Obj.Clear
                
                'récupère le nombre d'éléments
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Obj.AddItem Me.GetString(s & "|Item" & Trim$(Str$(X)))
                Next X
                
            End If
        ElseIf TypeOf Obj Is Menu Then
            'le menu
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is ListView Then
            'on prend les en-têtes de colonne
            
            'vérifie que le tag est Ok
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on loade
                
                'récupère le nombre d'éléments
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Obj.ColumnHeaders.Item(X).Text = Me.GetString(s & "|Item" & _
                        Trim$(Str$(X)))
                Next X
                
            End If
            
        ElseIf TypeOf Obj Is TabStrip Then
            'on prend les captions des tab
        
            'vérifie que le tag est Ok
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, ol loade
                
                'récupère le nombre d'éléments
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Obj.Tabs.Item(X).Caption = Me.GetString(s & "|Item" & _
                        Trim$(Str$(X)))
                Next X
                
            End If
        End If

    Next Obj

    
    'aussi le caption de la form
    mParent.Caption = Me.GetString(mParent.Name)
End Sub

'=======================================================
'créé le contenu du fichier *.ini
'=======================================================
Public Sub WriteIniFileFormIDEform()
Dim Obj As Control
Dim sFile As String
Dim X As Long
Dim Y As Long
Dim s As String

    On Error Resume Next
    
    sFile = Me.LangFolder & "\" & Me.Language & ".ini"  'le fichier à sauvegarder
    
    'If UserControl.Ambient.UserMode = False Then Exit Sub  'on est dans l'IDE ==> inutile si compilmé en OCX
    
    For Each Obj In mParent.Controls
        'pour chaque control de la form, on écrit la variable
        
        s = Obj.Name & "|Index"
        s = s & Trim$(Str$(Obj.Index))  '==> deux concaténation car erreur sur la deuxième
        'ligne pour les composants non indexés
        
        Call LetPref(mParent.Name, s & "|ToolTipText", Obj.ToolTipText, sFile)
        
        If TypeOf Obj Is TextBox Then
            Call LetPref(mParent.Name, s & "|Text", Obj.Text, sFile)
        ElseIf TypeOf Obj Is CheckBox Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is OptionButton Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is ComboBox Then
            
            'vérifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'récupère le nombre d'éléments
                Call LetPref(mParent.Name, s & "|Count", Obj.ListCount, sFile)
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(X)), Obj.List(X - 1), sFile)
                Next X
            End If
            
        ElseIf TypeOf Obj Is Label Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is CommandButton Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is Frame Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is Menu Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is ListBox Then
                    
            'vérifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'récupère le nombre d'éléments
                Call LetPref(mParent.Name, s & "|Count", Obj.ListCount, sFile)
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(X)), Obj.List(X - 1), sFile)
                Next X
            End If
        ElseIf TypeOf Obj Is ListView Then
            
            'vérifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'récupère le nombre d'éléments
                Call LetPref(mParent.Name, s & "|Count", Obj.ColumnHeaders.Count, sFile)
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(X)), _
                        Obj.ColumnHeaders.Item(X).Text, sFile)
                Next X
                
            End If
        ElseIf TypeOf Obj Is TabStrip Then
            
            'vérifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'récupère le nombre d'éléments
                Call LetPref(mParent.Name, s & "|Count", Obj.Tabs.Count, sFile)
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(X)), Obj.Tabs.Item(X).Caption, sFile)
                Next X
            End If
        End If
    Next Obj

    'sauvegarde le caption de la form
    Call LetPref(mParent.Name, mParent.Name, mParent.Caption, sFile)
    
End Sub


'=======================================================
'permet d'ajouter une string dans le fichier ini
'=======================================================
Public Sub AddSimpleStringToFile(ByVal StringName As String, Value As String)
    '/!\ conflit si identifiant de string porte un nom de controle /!\
    Call LetPref(mParent.Name, StringName, Value, Me.LangFolder & "\" & _
        Me.Language & ".ini")
End Sub
