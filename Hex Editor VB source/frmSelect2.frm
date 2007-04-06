VERSION 5.00
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmSelect2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "     Sélectionner une zone"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelect2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1485
      TabIndex        =   3
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Sélectionner"
      Height          =   375
      Left            =   165
      TabIndex        =   2
      ToolTipText     =   "Procéder à la restriction"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtSize 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Taille de la sélection"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtFrom 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Offset de départ"
      Top             =   120
      Width           =   1095
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Label Label1 
      Caption         =   "Taille de la sélection"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "A partir du byte"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelect2"
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
'FORM POUR SELECTIONNER UNE ZONE PARTICULIERE
'=======================================================

Private Sub cmdOk_Click()
Dim lFrom As Currency
Dim lTo As Currency

    On Error GoTo ErrGestion
    
    'récupère les valeurs numériques
    lFrom = FormatedVal_(txtFrom.Text)
    lTo = lFrom + FormatedVal_(txtSize.Text) - 2 '-2 pour régler à la bonne taille
        
    'vérifie que la plage est OK
    If lFrom < frmContent.ActiveForm.HW.FirstOffset Or lTo > frmContent.ActiveForm.HW.MaxOffset Then
        Unload Me
        Exit Sub
    End If
    If lFrom > frmContent.ActiveForm.HW.MaxOffset Or lTo < frmContent.ActiveForm.HW.FirstOffset Then
        Unload Me
        Exit Sub
    End If
    
    'fait la sélection désirée
    frmContent.ActiveForm.HW.SelectZone 16 - (By16(lFrom) - lFrom), By16(lFrom) - 16, 17 - (By16(lTo) - lTo), By16(lTo) - 16
    
    'refresh le label qui contient la taille de la sélection
    frmContent.ActiveForm.Sb.Panels(4).Text = "Sélection=[" & CStr(frmContent.ActiveForm.HW.NumberOfSelectedItems) & " bytes]"
    frmContent.ActiveForm.Label2(9) = frmContent.ActiveForm.Sb.Panels(4).Text
    frmContent.ActiveForm.HW.Refresh
    Unload Me
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmSelect2.cmdOkClick", True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

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
    Lang.Language = cPref.env_Lang
    Lang.LoadControlsCaption
    
    If frmContent.ActiveForm Is Nothing Then Unload Me
    
    'affiche l'élément actuellement sélectionné dans l'activeform
    txtFrom.Text = CStr(frmContent.ActiveForm.HW.Item.Offset + frmContent.ActiveForm.HW.Item.Col)
End Sub
