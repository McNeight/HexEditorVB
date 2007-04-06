VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmGestWindows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion des fenêtres"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGestWindows.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCLoseIt 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Fermer la fenêtre sélectionnée"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdShowIt 
      Caption         =   "Afficher"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Afficher la fenêtre au premier plan"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      ToolTipText     =   "Quitter"
      Top             =   3360
      Width           =   1815
   End
   Begin ComctlLib.ListView LV 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Tag             =   "lang_ok"
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nom"
         Object.Width           =   14111
      EndProperty
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuShowIt 
         Caption         =   "&Afficher"
      End
      Begin VB.Menu mnuCloseIt 
         Caption         =   "&Fermer"
      End
   End
End
Attribute VB_Name = "frmGestWindows"
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
'FORM DE GESTIONS DES FENETRES AFFICHEES
'=======================================================

Private Sub cmdCLoseIt_Click()
'ferme les fenêtres sélectionnées
Dim Frm As Form
Dim x As Long

    On Error GoTo ErrGestion
    
    If Not (LV.SelectedItem Is Nothing) Then
        'alors au moins une sélection ==> demande confirmation
        If Not (MsgBox("Confirmer la fermeture ?", vbInformation + vbYesNo, _
        "Attention") = vbYes) Then Exit Sub
    Else
        Exit Sub    'pas de sélection
    End If
    
    LV.Visible = False

    'liste les form et ferme les sélectionnées
    For Each Frm In Forms
        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is MemPfm) Or (TypeOf Frm Is physPfm) Then
            For x = LV.ListItems.Count To 1 Step -1
                If LV.ListItems.Item(x).Selected And LV.ListItems.Item(x).SubItems(1) = _
                Frm.Caption Then
                    SendMessage Frm.hWnd, WM_CLOSE, 0, 0
                    'Unload frm
                    'lNbChildFrm = lNbChildFrm - 1
                    LV.ListItems.Remove x
                End If
            Next x
        End If
    Next Frm
    
    '/!\ NE PAS ENLEVER
    '/!\ BUG NON RESOLU
    '/!\ Après déchargement des form (juste en haut), des form nommées "Form1" (caption par
    'défaut) subsistent
    For Each Frm In Forms
        If Frm.Caption = "Form1" Then SendMessage Frm.hWnd, WM_CLOSE, 0, 0
    Next Frm
    
    LV.Visible = True
    
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmGestWindows.cmdCLoseltClick", True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdShowIt_Click()
'montre la fenêtre sélectionnée
Dim Frm As Form

    'liste les form et affiche celle qui est sélectionnée
    For Each Frm In Forms
        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is MemPfm) Or (TypeOf Frm Is physPfm) Then
            If LV.SelectedItem.SubItems(1) = Frm.Caption Then
                'SendMessage frm.hwnd, WM_SHOWWINDOW, 0, 0   'pas de .Show car form modale affichée
                Frm.ZOrder 0
                If cPref.general_MaximizeWhenOpen Then Frm.WindowState = vbMaximized
            End If
        End If
    Next Frm

End Sub

Private Sub Form_Load()
Dim Frm As Form
        
    #If MODE_DEBUG Then
        If App.LogMode = 0 Then
            'on créé le fichier de langue français
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
    
    'applique la langue désirée aux controles
    Lang.Language = MyLang
    Lang.LoadControlsCaption
    
    LV.ListItems.Clear
    
    'liste les forms et les ajoute au LV
    For Each Frm In Forms
        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is MemPfm) Or (TypeOf Frm Is physPfm) Then
            LV.ListItems.Add Text:=TypeOfForm(Frm)
            LV.ListItems.Item(LV.ListItems.Count).SubItems(1) = Frm.Caption
        End If
    Next Frm
    
End Sub

Private Sub LV_DblClick()
    cmdShowIt_Click
End Sub
