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
'//CONTROLE DE CHANGEMENT DE LANGUE
'=======================================================


'AJOUTER LE TAG "lang_ok" POUR LES LISTBOX/COMBOX POUR EFFECTUER LA LECTURE DES ITEMS


'=======================================================
'APIS
'=======================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private sLanguage As String     'langue utilis�e
Private sLangFolder As String    'path du r�pertoire de langue
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
    Call LoadControlsCaption 'met � jour les controles
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
    Set mParent = Nothing   'lib�re la m�moire
End Sub




'=======================================================
'active le changement de langue
'=======================================================
Public Sub ActiveLang(Frm As Form)
    Set mParent = Frm
End Sub

'=======================================================
'permet de r�cup�rer une string
'=======================================================
Public Function GetString(ByVal StringName As String) As String
    
Dim lng As Integer
Dim sReturn As String
Dim strFile As String
Dim Section As String
Dim sFile As String

    On Error Resume Next
    
    'cr�� le buffer
    sReturn = String$(255, 0)
        
    'd�termine le fichier *.ini concern�
    sFile = Me.LangFolder & "\" & Me.Language & ".ini"
    
    'd�termine la section
    Section = mParent.Name
    
    'obtient la valeur
    lng = GetPrivateProfileString(Section, StringName, vbNullString, sReturn, _
        255, sFile)
    
    'formate la string
    GetString = Left$(sReturn, lng)
End Function

'=======================================================
'd�finit les pr�f�rences du fichier *.ini
'=======================================================
Private Sub LetPref(ByVal sSection As String, ByVal sVariable As String, ByVal sValeur As String, ByVal sFile As String)
    Call WritePrivateProfileString(sSection, sVariable, sValeur, sFile)
End Sub

'=======================================================
'affecte les propri�t�s Text et Caption pour tous les controles de la form
'=======================================================
Public Sub LoadControlsCaption()
Dim Obj As Control
Dim x As Long
Dim y As Long
Dim s As String

    On Error Resume Next
    
    'If UserControl.Ambient.UserMode = False Then Exit Sub  'on est dans l'IDE
    
    'v�rifie que le fichier existe bien...
    If FileLen(Me.LangFolder & "\" & Me.Language & ".ini") = 0 Then Exit Sub       'fichier inexistant, donc on loade rien

    For Each Obj In mParent.Controls
        'pour chaque controle, on r�cup�re la valeur de son nom et on affecte
        'la string Caption/Texte au controle
        'possibilit� d'ajouter facilement les autres controles � g�rer (notemment les
        'usercontrols)
        
        s = Obj.Name & "|Index"
        s = s & Trim$(Str$(Obj.Index))  '==> deux concat�nation car erreur sur la deuxi�me
        'ligne pour les composants non index�s
        
        'r�cup�re le ToolTipText
        Obj.ToolTipText = Me.GetString(s & "|ToolTipText")
        
        If TypeOf Obj Is TextBox Then
            Obj.Text = Me.GetString(s & "|Text")
        ElseIf TypeOf Obj Is CheckBox Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is OptionButton Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is ComboBox Then
            
            'v�rifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on loade
                
                'clear
                Obj.Clear
                
                'r�cup�re le nombre d'�l�ments
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Obj.AddItem Me.GetString(s & "|Item" & Trim$(Str$(x)))
                Next x
                
            End If
            
        ElseIf TypeOf Obj Is Label Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is CommandButton Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is Frame Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is Listbox Then
            'alors l� faut aussi sauvegarder les Items qu'il y a dedans si tag OK
            
            'v�rifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on loade
                
                'clear
                Obj.Clear
                
                'r�cup�re le nombre d'�l�ments
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Obj.AddItem Me.GetString(s & "|Item" & Trim$(Str$(x)))
                Next x
                
            End If
        ElseIf TypeOf Obj Is Menu Then
            'le menu
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is ListView Then
            'on prend les en-t�tes de colonne
            
            'v�rifie que le tag est Ok
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on loade
                
                'r�cup�re le nombre d'�l�ments
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Obj.ColumnHeaders.Item(x).Text = Me.GetString(s & "|Item" & _
                        Trim$(Str$(x)))
                Next x
                
            End If
            
        ElseIf TypeOf Obj Is TabStrip Then
            'on prend les captions des tab
        
            'v�rifie que le tag est Ok
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, ol loade
                
                'r�cup�re le nombre d'�l�ments
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Obj.Tabs.Item(x).Caption = Me.GetString(s & "|Item" & _
                        Trim$(Str$(x)))
                Next x
                
            End If
        ElseIf TypeOf Obj Is vkFrame Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is vkOptionButton Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is vkCheck Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is vkLabel Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is vkCommand Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is vkToggleButton Then
            Obj.Caption = Me.GetString(s & "|Caption")
        ElseIf TypeOf Obj Is vkTextBox Then
            Obj.LegendText = Me.GetString(s & "|LegendText")
        End If

    Next Obj

    
    'aussi le caption de la form
    mParent.Caption = Me.GetString(mParent.Name)
End Sub

'=======================================================
'cr�� le contenu du fichier *.ini
'=======================================================
Public Sub WriteIniFileFormIDEform()
Dim Obj As Control
Dim sFile As String
Dim x As Long
Dim y As Long
Dim s As String

    On Error Resume Next
    
    sFile = Me.LangFolder & "\" & Me.Language & ".ini"  'le fichier � sauvegarder
    
    'If UserControl.Ambient.UserMode = False Then Exit Sub  'on est dans l'IDE ==> inutile si compilm� en OCX
    
    For Each Obj In mParent.Controls
        'pour chaque control de la form, on �crit la variable
        
        s = Obj.Name & "|Index"
        s = s & Trim$(Str$(Obj.Index))  '==> deux concat�nation car erreur sur la deuxi�me
        'ligne pour les composants non index�s
        
        Call LetPref(mParent.Name, s & "|ToolTipText", Obj.ToolTipText, sFile)
        
        If TypeOf Obj Is TextBox Then
            Call LetPref(mParent.Name, s & "|Text", Obj.Text, sFile)
        ElseIf TypeOf Obj Is CheckBox Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is OptionButton Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is ComboBox Then
            
            'v�rifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'r�cup�re le nombre d'�l�ments
                Call LetPref(mParent.Name, s & "|Count", Obj.ListCount, sFile)
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(x)), Obj.List(x - 1), sFile)
                Next x
            End If
            
        ElseIf TypeOf Obj Is Label Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is CommandButton Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is Frame Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is Menu Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is Listbox Then
                    
            'v�rifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'r�cup�re le nombre d'�l�ments
                Call LetPref(mParent.Name, s & "|Count", Obj.ListCount, sFile)
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(x)), Obj.List(x - 1), sFile)
                Next x
            End If
        ElseIf TypeOf Obj Is ListView Then
            
            'v�rifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'r�cup�re le nombre d'�l�ments
                Call LetPref(mParent.Name, s & "|Count", Obj.ColumnHeaders.Count, sFile)
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(x)), _
                        Obj.ColumnHeaders.Item(x).Text, sFile)
                Next x
                
            End If
        ElseIf TypeOf Obj Is TabStrip Then
            
            'v�rifie que le Tag est OK
            If InStr(1, Obj.Tag, "lang_ok", vbBinaryCompare) Then
                'c'est bon, on sauvegarde
                
                'r�cup�re le nombre d'�l�ments
                Call LetPref(mParent.Name, s & "|Count", Obj.Tabs.Count, sFile)
                y = Val(Me.GetString(s & "|Count"))
                
                'ajoute les �l�ments
                For x = 1 To y
                    Call LetPref(mParent.Name, s & "|Item" & Trim$(Str$(x)), Obj.Tabs.Item(x).Caption, sFile)
                Next x
            End If
        ElseIf TypeOf Obj Is vkFrame Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is vkOptionButton Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is vkCheck Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is vkLabel Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is vkCommand Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is vkToggleButton Then
            Call LetPref(mParent.Name, s & "|Caption", Obj.Caption, sFile)
        ElseIf TypeOf Obj Is vkTextBox Then
            Call LetPref(mParent.Name, s & "|LegendText", Obj.LegendText, sFile)
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
