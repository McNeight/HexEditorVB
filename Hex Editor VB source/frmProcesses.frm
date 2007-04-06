VERSION 5.00
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Object = "{2245E336-2835-4C1E-B373-2395637023C8}#1.0#0"; "ProcessView_OCX.ocx"
Begin VB.Form frmProcesses 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "S�lection du processus � ouvrir"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
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
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ProcessView_OCX.ProcessView PV 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
      Sorted          =   0   'False
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Rafraichir"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Ouvrir de processus"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Ouvrir de processus"
      Top             =   3600
      Width           =   735
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'FORM QUI INVITE A SELECTIONNER UN PROCESSUS A EDITER
'=======================================================

Private bFirst As Boolean

Private Sub cmdFermer_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim lH As Long
Dim Frm As Form

    'v�rfie que le processus est ouvrable
    lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(PV.SelectedItem.Tag))
    CloseHandle lH
    
    If lH = 0 Then
        'pas possible
        MsgBox "Acc�s impossible � ce processus", vbInformation, "Erreur"
        Exit Sub
    End If
    
    Me.Hide
    
    'possible affiche une nouvelle fen�tre
    Set Frm = New MemPfm
    Call Frm.GetFile(Val(PV.SelectedItem.Tag))
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"

    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entr�es dans les menus

    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call PV.Refresh
End Sub

Private Sub Form_Activate()
Dim ND As Node
    
    On Error Resume Next
    
    'on expand si c'est la premi�re activation de la form
    If bFirst = False Then
        bFirst = True
        
        'on extend tous les noeuds
        With PV
            For Each ND In .Nodes
                ND.Expanded = True
            Next ND
            
            'met en surbrillance le premier
            .Nodes.Item(1).Selected = True
            .SetFocus
        End With
        
    End If
End Sub

Private Sub Form_Load()

    #If MODE_DEBUG Then
        If App.LogMode = 0 Then
            'on cr�� le fichier de langue fran�ais
            Lang.Language = "French"
            Lang.LangFolder = LANG_PATH
            Lang.WriteIniFileFormIDEform
        End If
    #End If
    
    If App.LogMode = 0 Then
        'alors on est dans l'IDE
        Lang.LangFolder = LANG_PATH
    Else
        Lang.LangFolder = App.Path & "\Lang"
    End If
    
    'applique la langue d�sir�e aux controles
    Lang.Language = cPref.env_Lang
    Lang.LoadControlsCaption
    
    bFirst = False
End Sub

Private Sub PV_DblClick()
    If PV.SelectedItem Is Nothing Then Exit Sub
    Call cmdOk_Click
End Sub
