VERSION 5.00
Begin VB.UserControl ctrlLanguage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   3255
   ScaleWidth      =   3765
End
Attribute VB_Name = "ctrlLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
Private mParent As Form  'form parente
Attribute mParent.VB_VarHelpID = -1


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
    LoadControlsCaption 'met à jour les controles
End If
End Property
Public Property Get LangFolder() As String: LangFolder = sLangFolder: End Property
Public Property Let LangFolder(LangFolder As String): sLangFolder = LangFolder: End Property



'=======================================================
'CONTROL SUBS
'=======================================================
Private Sub UserControl_Initialize()
    Me.LangFolder = App.Path & "\Lang"
    Me.Language = "French"
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.LangFolder = .ReadProperty("LangFolder", App.Path & "\Lang")
        Me.Language = .ReadProperty("Language", "French")
    End With
    Set mParent = UserControl.Parent    'récupère la form parent
    
    LoadControlsCaption 'met à jour les controles
End Sub
Private Sub UserControl_Resize()
'resize impossible
    UserControl.Width = 800
    UserControl.Height = 800
End Sub
Private Sub UserControl_Terminate()
    Set mParent = Nothing   'libère la mémoire
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Language", Me.Language, "French")
        Call .WriteProperty("LangFolder", Me.LangFolder, App.Path & "\Lang")
    End With
End Sub



'=======================================================
'permet de récupérer une string
'=======================================================
Public Function GetString(ByVal StringName As String, Optional ByVal Language As String) As String
Dim lng As Integer
Dim sLang As String
Dim sReturn As String
Dim strFile As String
Dim Section As String
Dim sFile As String

    On Error Resume Next
    
    'créé le buffer
    sReturn = String$(255, 0)
    
    'récupère le nom de la langue
    If IsMissing(Language) Then sLang = Me.Language Else sLang = Me.Language
    
    'détermine le fichier *.ini concerné
    sFile = Me.LangFolder & "\" & sLang & ".ini"
    
    'détermine la section
    Section = mParent.Name
    
    'obtient la valeur
    lng = GetPrivateProfileString(Section, StringName, vbNullString, sReturn, 255, sFile)
    
    'formate la string
    GetString = Left$(sReturn, lng)
End Function

'=======================================================
'définit les préférences du fichier *.ini
'=======================================================
Private Sub LetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sValeur As String, ByVal sFile As String)
    WritePrivateProfileString sSection, sVariable, sValeur, sFile
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
    
    If UserControl.Ambient.UserMode = False Then Exit Sub  'on est dans l'IDE
    
    'vérifie que le fichier existe bien...
    If FileLen(Me.LangFolder & "\" & Me.Language & ".ini") = 0 Then Exit Sub    'fichier inexistant, donc on loade rien
    
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
            If Obj.Tag = "lang_ok" Then
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
            If Obj.Tag = "lang_ok" Then
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
    
    If UserControl.Ambient.UserMode = False Then Exit Sub  'on est dans l'IDE
    
    For Each Obj In mParent.Controls
        'pour chaque control de la form, on écrit la variable
        
        s = Obj.Name & "|Index"
        s = s & Trim$(Str$(Obj.Index))  '==> deux concaténation car erreur sur la deuxième
        'ligne pour les composants non indexés
        
        LetPref mParent.Name, s & "|ToolTipText", Obj.ToolTipText, sFile
        
        If TypeOf Obj Is TextBox Then
            LetPref mParent.Name, s & "|Text", Obj.Text, sFile
        ElseIf TypeOf Obj Is CheckBox Then
            LetPref mParent.Name, s & "|Caption", Obj.Caption, sFile
        ElseIf TypeOf Obj Is OptionButton Then
            LetPref mParent.Name, s & "|Caption", Obj.Caption, sFile
        ElseIf TypeOf Obj Is ComboBox Then
            
            'vérifie que le Tag est OK
            If Obj.Tag = "lang_ok" Then
                'c'est bon, on sauvegarde
                
                'récupère le nombre d'éléments
                LetPref mParent.Name, s & "|Count", Obj.ListCount, sFile
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    LetPref mParent.Name, s & "|Item" & Trim$(Str$(X)), Obj.List(X - 1), sFile
                Next X
            End If
            
        ElseIf TypeOf Obj Is Label Then
            LetPref mParent.Name, s & "|Caption", Obj.Caption, sFile
        ElseIf TypeOf Obj Is CommandButton Then
            LetPref mParent.Name, s & "|Caption", Obj.Caption, sFile
        ElseIf TypeOf Obj Is Frame Then
            LetPref mParent.Name, s & "|Caption", Obj.Caption, sFile
        ElseIf TypeOf Obj Is Menu Then
            LetPref mParent.Name, s & "|Caption", Obj.Caption, sFile
        ElseIf TypeOf Obj Is ListBox Then
                    
            'vérifie que le Tag est OK
            If Obj.Tag = "lang_ok" Then
                'c'est bon, on sauvegarde
                
                'récupère le nombre d'éléments
                LetPref mParent.Name, s & "|Count", Obj.ListCount, sFile
                Y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les éléments
                For X = 1 To Y
                    LetPref mParent.Name, s & "|Item" & Trim$(Str$(X)), Obj.List(X - 1), sFile
                Next X
            End If
        End If
    Next Obj

    'sauvegarde le caption de la form
    LetPref mParent.Name, mParent.Name, mParent.Caption, sFile
End Sub


'=======================================================
'permet d'ajouter une string dans le fichier ini
'=======================================================
Public Sub AddSimpleStringToFile(ByVal StringName As String, Value As String)
    '/!\ conflit si identifiant de string porte un nom de controle /!\
    LetPref mParent.Name, StringName, Value, Me.LangFolder & "\" & Me.Language & ".ini"
End Sub
