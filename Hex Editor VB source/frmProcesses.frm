VERSION 5.00
Object = "{2245E336-2835-4C1E-B373-2395637023C8}#1.0#0"; "ProcessView_OCX.ocx"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmProcesses 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sélection du processus à ouvrir"
   ClientHeight    =   4065
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
   HelpContextID   =   44
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCommand cmdRefresh 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Rafrichir la liste des processus"
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Rafrachir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkCommand cmdOK 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Ouvrir de processus"
      Top             =   3600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vkUserContolsXP.vkCommand cmdFermer 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Fermer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ProcessView_OCX.ProcessView PV 
      Height          =   3495
      Left            =   15
      TabIndex        =   0
      Top             =   23
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
      Sorted          =   0   'False
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
'FORM QUI INVITE A SELECTIONNER UN PROCESSUS A EDITER
'=======================================================

Private Lang As New clsLang
Private bFirst As Boolean

Private Sub cmdFermer_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim lH As Long
Dim Frm As Form

    'vérfie que le processus est ouvrable
    lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(PV.SelectedItem.Tag))
    Call CloseHandle(lH)
    
    If lH = 0 Then
        'pas possible
        MsgBox Lang.GetString("_AccessDen"), vbInformation, Lang.GetString("_Error")
        Exit Sub
    End If
    
    Me.Hide
    
    'possible affiche une nouvelle fenêtre
    Set Frm = New MemPfm
    Call Frm.GetFile(Val(PV.SelectedItem.Tag))
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"

    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus

    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call PV.Refresh
End Sub

Private Sub Form_Activate()
Dim ND As Node
    
    On Error Resume Next
    
    'on expand si c'est la première activation de la form
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
    
    bFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub

Private Sub PV_DblClick()
    If PV.SelectedItem Is Nothing Then Exit Sub
    Call cmdOk_Click
End Sub
