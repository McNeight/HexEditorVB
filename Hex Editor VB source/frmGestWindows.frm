VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   5183
      TabIndex        =   2
      ToolTipText     =   "Quitter"
      Top             =   3413
      Width           =   1815
   End
   Begin VB.CommandButton cmdShowIt 
      Caption         =   "Afficher"
      Height          =   375
      Left            =   143
      TabIndex        =   1
      ToolTipText     =   "Afficher la fenêtre au premier plan"
      Top             =   3413
      Width           =   1455
   End
   Begin VB.CommandButton cmdCLoseIt 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1823
      TabIndex        =   0
      ToolTipText     =   "Fermer la fenêtre sélectionnée"
      Top             =   3413
      Width           =   1455
   End
   Begin ComctlLib.ListView LV 
      Height          =   3255
      Left            =   23
      TabIndex        =   3
      Tag             =   "lang_ok"
      Top             =   53
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
Private Lang As New clsLang

Private Sub cmdCLoseIt_Click()
'ferme les fenêtres sélectionnées
Dim Frm As Form
Dim X As Long

    On Error GoTo ErrGestion
    
    If Not (LV.SelectedItem Is Nothing) Then
        'alors au moins une sélection ==> demande confirmation
        If Not (MsgBox(Lang.GetString("_Conf"), vbInformation + vbYesNo, _
            Lang.GetString("_War")) = vbYes) Then Exit Sub
    Else
        Exit Sub    'pas de sélection
    End If
    
    LV.Visible = False

    'liste les form et ferme les sélectionnées
    For Each Frm In Forms
        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is MemPfm) Or (TypeOf Frm Is physPfm) Then
            For X = LV.ListItems.Count To 1 Step -1
                If LV.ListItems.Item(X).Selected And LV.ListItems.Item(X).SubItems(1) = _
                Frm.Caption Then
                    Call SendMessage(Frm.hWnd, WM_CLOSE, 0, 0)
                    'Unload frm
                    'lNbChildFrm = lNbChildFrm - 1
                    LV.ListItems.Remove X
                End If
            Next X
        End If
    Next Frm
    
    '/!\ NE PAS ENLEVER
    '/!\ BUG NON RESOLU
    '/!\ Après déchargement des form (juste en haut), des form nommées "Form1" (caption par
    'défaut) subsistent
    For Each Frm In Forms
        If Frm.Caption = "Form1" Then Call SendMessage(Frm.hWnd, WM_CLOSE, 0, 0)
    Next Frm
    
    LV.Visible = True
    
    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
    
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
        
    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
                .Language = "French"
                .LangFolder = LANG_PATH
                .WriteIniFileFormIDEform
            End If
        #End If
        
        If App.LogMode = 0 Then
            'alors on est dans l'IDE
            .LangFolder = LANG_PATH
        Else
            .LangFolder = App.Path & "\Lang"
        End If
        
        'applique la langue désirée aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        .LoadControlsCaption
    End With
    
    LV.ListItems.Clear
    
    'liste les forms et les ajoute au LV
    For Each Frm In Forms
        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is _
            MemPfm) Or (TypeOf Frm Is physPfm) Then
            
            LV.ListItems.Add Text:=TypeOfForm(Frm)
            LV.ListItems.Item(LV.ListItems.Count).SubItems(1) = Frm.Caption
        End If
    Next Frm
    
End Sub

Private Sub LV_DblClick()
    Call cmdShowIt_Click
End Sub
