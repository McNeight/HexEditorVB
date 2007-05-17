VERSION 5.00
Begin VB.Form frmSelect2 
   BackColor       =   &H00F9E5D9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sélectionner une zone"
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
   HelpContextID   =   44
   Icon            =   "frmSelect2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFrom 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1545
      TabIndex        =   3
      ToolTipText     =   "Offset de départ"
      Top             =   105
      Width           =   1095
   End
   Begin VB.TextBox txtSize 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1665
      TabIndex        =   2
      ToolTipText     =   "Taille de la sélection"
      Top             =   465
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Sélectionner"
      Height          =   375
      Left            =   150
      TabIndex        =   1
      ToolTipText     =   "Procéder à la restriction"
      Top             =   945
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1470
      TabIndex        =   0
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A partir du byte"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Taille de la sélection"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   465
      Width           =   1455
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
Private Lang As New clsLang

Private Sub cmdOk_Click()
Dim lFrom As Currency
Dim lTo As Currency

    On Error GoTo ErrGestion
    
    'récupère les valeurs numériques
    lFrom = FormatedVal_(txtFrom.Text)
    lTo = lFrom + FormatedVal_(txtSize.Text) - 2 '-2 pour régler à la bonne taille
        
    'vérifie que la plage est OK
    If lFrom < frmContent.ActiveForm.HW.FirstOffset Or lTo > _
        frmContent.ActiveForm.HW.MaxOffset Then
        Unload Me
        Exit Sub
    End If
    If lFrom > frmContent.ActiveForm.HW.MaxOffset Or lTo < _
        frmContent.ActiveForm.HW.FirstOffset Then
        Unload Me
        Exit Sub
    End If
    
    With frmContent.ActiveForm
        'fait la sélection désirée
        .HW.SelectZone 16 - (By16(lFrom) - lFrom), By16(lFrom) - 16, 17 - (By16(lTo) - lTo), By16(lTo) - 16
        
        'refresh le label qui contient la taille de la sélection
        .Sb.Panels(4).Text = Lang.GetString("_Sel") & CStr(.HW.NumberOfSelectedItems) & " bytes]"
        .Label2(9) = .Sb.Panels(4).Text
        .HW.Refresh
    End With
    
    Unload Me
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmSelect2.cmdOkClick", True
End Sub

Private Sub cmdQuit_Click()
    Unload Me
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
        Call .ActiveLang(Me): Call .ActiveLang(Me): .Language = cPref.env_Lang
        .LoadControlsCaption
        Call .ActiveLang(Me)
    End With
    
    If frmContent.ActiveForm Is Nothing Then Unload Me
    
    'affiche l'élément actuellement sélectionné dans l'activeform
    txtFrom.Text = CStr(frmContent.ActiveForm.HW.Item.Offset + frmContent.ActiveForm.HW.Item.Col)
End Sub
