VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProcesses 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sélection du processus à ouvrir"
   ClientHeight    =   4095
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
   ScaleHeight     =   4095
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   2055
      TabIndex        =   2
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "Ouvrir de processus"
      Top             =   3480
      Width           =   1215
   End
   Begin ComctlLib.ListView LV 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
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
         Text            =   "PID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Processus"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'FORM QUI INVITE A SELECTIONNER UN PROCESSUS A EDITER
'-------------------------------------------------------

Private Sub cmdFermer_Click()
'fermer
    Unload Me
End Sub

Private Sub cmdOk_Click()
'ok
Dim lH As Long
Dim Frm As Form


    'vérfie que le processus est ouvrable
    lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(LV.SelectedItem.Text))
    CloseHandle lH
    
    If lH = 0 Then
        'pas possible
        MsgBox "Accès impossible à ce processus", vbInformation, "Erreur"
        Exit Sub
    End If
    
    Me.Hide
    
    'possible affiche une nouvelle fenêtre
    Set Frm = New MemPfm
    Call Frm.GetFile(Val(LV.SelectedItem.Text), LV.SelectedItem.SubItems(1))
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"

    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus

    Unload Me
End Sub

Private Sub Form_Load()
'fait la liste des processus en mémoire
Dim p() As ProcessItem
Dim x As Long
Dim clsProc As clsProcess   'appel à une classe de gestion de processus

    Set clsProc = New clsProcess

    LV.ListItems.Clear

    'énumération
    clsProc.EnumerateProcesses p()
    
    'affiche la liste
    For x = 0 To UBound(p) - 1
        LV.ListItems.Add Text:=p(x).th32ProcessID
        LV.ListItems.Item(x + 1).SubItems(1) = p(x).szExeFile
    Next x
    
    LV.ListItems.Item(LV.ListItems.Count).Selected = True   'met le dernier process en surbrillance
    
End Sub

Private Sub LV_DblClick()
    If LV.SelectedItem Is Nothing Then Exit Sub
    cmdOk_Click
End Sub
