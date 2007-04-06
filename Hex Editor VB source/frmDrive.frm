VERSION 5.00
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Object = "{3AF19019-2368-4F9C-BBFC-FD02C59BD0EC}#1.0#0"; "DriveView_OCX.ocx"
Begin VB.Form frmDrive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S�lection d'un disque"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDrive.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Actualiser"
      Height          =   375
      Left            =   1238
      TabIndex        =   3
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   3240
      Width           =   975
   End
   Begin DriveView_OCX.DriveView DV 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5106
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ouvrir..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Ouvrir ce lecteur"
      Top             =   3240
      Width           =   975
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
End
Attribute VB_Name = "frmDrive"
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
'FORM PERMETTANT LE CHOIX DU DRIVE A OUVRIR
'=======================================================


'=======================================================
'ouvre diskFrm
'=======================================================
Private Sub cmdOk_Click()
Dim Frm As Form
Dim sDrive As String
Dim cDr As clsDiskInfos

    If DV.SelectedItem Is Nothing Then Exit Sub
    
    Set cDr = New clsDiskInfos
    
    'on check si c'est un disque logique ou un disque physique
    If Left$(DV.SelectedItem.Key, 3) = "log" Then
    
        'disque logique
    
        'v�rifie que le drive est accessible
        If DV.IsSelectedDriveAccessible = False Then
            Set cDr = Nothing   'inaccessible, alors on sort de cette proc�dure
            Exit Sub
        End If
        
        'affiche une nouvelle fen�tre
        Set Frm = New diskPfm
        
        Call Frm.GetDrive(DV.SelectedItem.Text)  'renseigne sur le path s�lectionn�
        
        Unload Me   'quitte cette form
        
        Frm.Show    'affiche la nouvelle
        lNbChildFrm = lNbChildFrm + 1
        frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
        
    Else
    
        'disque physique
        
        'v�rifie que le drive est accessible
        If DV.IsSelectedDriveAccessible = False Then
            Set cDr = Nothing   'inaccessible, alors on sort de cette proc�dure
            Exit Sub
        End If
        
        'affiche une nouvelle fen�tre
        Set Frm = New physPfm
        
        Call Frm.GetDrive(Val(Mid$(DV.SelectedItem.Text, 3, 1))) 'renseigne sur le path s�lectionn�
        
        Unload Me   'quitte cette form
        
        Frm.Show    'affiche la nouvelle
        lNbChildFrm = lNbChildFrm + 1
        frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
        
        
    End If
    
 
    'lib�re la classe
    Set cDr = Nothing
End Sub

Private Sub cmdNO_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
'on refresh la liste des drives
    Call DV.Refresh
    Call MarkUnaccessibleDrives(Me.DV)
End Sub

Private Sub DV_DblClick()
    If DV.SelectedItem Is Nothing Then Exit Sub
    If DV.SelectedItem.Children <> 0 Then Exit Sub
    cmdOk_Click
End Sub

Private Sub Form_Activate()
    Call MarkUnaccessibleDrives(Me.DV)  'marque les drives inaccessibles
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
End Sub
