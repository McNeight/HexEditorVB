VERSION 5.00
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmConvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversions"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPlan 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Tag             =   "pref"
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdQuitter 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   5055
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   4815
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   4815
         Begin VB.TextBox txtI 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Tag             =   "pref1"
            ToolTipText     =   "Valeur dans la base d'arrivée"
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtBase 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   4200
            TabIndex        =   5
            Tag             =   "pref"
            Text            =   "12"
            ToolTipText     =   "Base personnelle"
            Top             =   120
            Width           =   615
         End
         Begin VB.ComboBox cbO 
            Height          =   315
            ItemData        =   "frmConvert.frx":058A
            Left            =   2520
            List            =   "frmConvert.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "pref lang_ok"
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "dans la base"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   4815
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   4815
         Begin VB.TextBox txtI 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   0
            Tag             =   "pref0"
            ToolTipText     =   "Valeur dans la base de départ"
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtBase 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   4200
            TabIndex        =   2
            Tag             =   "pref"
            Text            =   "12"
            ToolTipText     =   "Base personnelle"
            Top             =   120
            Width           =   615
         End
         Begin VB.ComboBox cbI 
            Height          =   315
            ItemData        =   "frmConvert.frx":058E
            Left            =   2520
            List            =   "frmConvert.frx":0590
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Tag             =   "pref lang_ok"
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "dans la base"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   11
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
End
Attribute VB_Name = "frmConvert"
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
'FORM DE CONVERSIONS ENTRE BASES
'=======================================================

Private clsPref As clsIniForm
Private cConv As clsConvert

Private Sub cbI_Click()
    Call DisplayResult
    txtBase(0).Enabled = cbI.Text = "Autre"
End Sub

Private Sub cbO_Click()
    Call DisplayResult
    txtBase(1).Enabled = cbO.Text = "Autre"
End Sub

Private Sub chkPlan_Click()
    If chkPlan.Value = 1 Then Call SetFormForeBackGround(Me, SetFormForeGround) Else Call SetFormForeBackGround(Me, SetFormBackGround)
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If chkPlan.Value Then Call SetFormForeBackGround(Me, SetFormForeGround) Else _
        Call SetFormForeBackGround(Me, SetFormBackGround)
End Sub

'=======================================================
'affiche la conversion
'=======================================================
Private Sub DisplayResult()
Dim s2 As String

    With cConv
        .CurrentString = txtI(0).Text
        
        Select Case cbI.Text
            Case "Décimale"
                .CurrentBase = 10
            Case "Octale"
                .CurrentBase = 8
            Case "Héxadécimale"
                .CurrentBase = 16
            Case "Binaire"
                .CurrentBase = 2
            Case "Autre"
                .CurrentBase = Val(txtBase(0).Text)
            Case Else
                'ANSI ASCII
        End Select
        
        Select Case cbO.Text
            Case "Décimale"
                s2 = .Convert(10)
            Case "Octale"
                s2 = .Convert(8)
            Case "Héxadécimale"
                s2 = .Convert(16)
            Case "Binaire"
                s2 = .Convert(2)
            Case "Autre"
                s2 = .Convert(Val(txtBase(1).Text))
            Case Else
                'ANSI ASCII
        End Select
        
        If .ConversionFailed Then
            txtI(1).ForeColor = RED_COLOR
            txtI(1).Text = "Echec"
        Else
            txtI(1).ForeColor = vbBlack
            txtI(1).Text = s2
        End If
    End With
    
End Sub

Private Sub Form_Load()

    'loading des preferences
    Set clsPref = New clsIniForm
    Set cConv = New clsConvert
    
    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
                .Language = "French"
                .LangFolder = LANG_PATH
                Call .WriteIniFileFormIDEform
            End If
        #End If
        
        If App.LogMode = 0 Then
            'alors on est dans l'IDE
            .LangFolder = LANG_PATH
        Else
            .LangFolder = App.Path & "\Lang"
        End If
        
        'applique la langue désirée aux controles
        .Language = cPref.env_Lang
        Call .LoadControlsCaption
    End With

    Call clsPref.GetFormSettings(App.Path & "\Preferences\Conversion.ini", Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\Conversion.ini", Me)
    Set clsPref = Nothing
    Set cConv = Nothing
End Sub

Private Sub txtBase_Change(Index As Integer)
    Call DisplayResult
End Sub

Private Sub txtI_Change(Index As Integer)
    Call DisplayResult
End Sub
