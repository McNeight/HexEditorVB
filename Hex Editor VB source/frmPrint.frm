VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impression"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Options d'impression"
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   113
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   3855
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   3855
         Begin VB.CheckBox Check1 
            Caption         =   "Valeurs ASCII"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Tag             =   "pref0"
            ToolTipText     =   "Ajouter les valeurs ASCII"
            Top             =   120
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Valeurs hexa"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Tag             =   "pref1"
            ToolTipText     =   "Ajouter les valeurs hexa"
            Top             =   480
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Offsets"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   15
            Tag             =   "pref2"
            ToolTipText     =   "Afficher les offsets"
            Top             =   840
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Informations sur le fichier"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   14
            Tag             =   "pref3"
            ToolTipText     =   "Afficher les informations sur le fichier en première page"
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Titre"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   13
            Tag             =   "pref4"
            ToolTipText     =   "Ajouter un titre"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtTitle 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   12
            Tag             =   "pref"
            Text            =   "Titre du document"
            ToolTipText     =   "Titre du document"
            Top             =   1560
            Width           =   2895
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Imprimer tout"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Tag             =   "pref0"
            ToolTipText     =   "Tout imprimer"
            Top             =   2040
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Imprimer sélection"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Tag             =   "pref1"
            ToolTipText     =   "Imprimer uniquement la sélection"
            Top             =   2280
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Imprimer de "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Tag             =   "pref2"
            ToolTipText     =   "Imprimer une plage d'offsets"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtFrom 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Tag             =   "pref"
            Text            =   "00000000"
            ToolTipText     =   "Offset inférieur"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtTo 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2760
            TabIndex        =   7
            Tag             =   "pref"
            Text            =   "10000000"
            ToolTipText     =   "Offset supérieur"
            Top             =   2520
            Width           =   975
         End
         Begin VB.TextBox txtFontSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Tag             =   "pref"
            Text            =   "10"
            ToolTipText     =   "Taille de la police d'impression (Courier New)"
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "à"
            Height          =   255
            Left            =   2520
            TabIndex        =   19
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Taille de la police :"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2880
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Lancer l'impression"
      Top             =   3833
      Width           =   975
   End
   Begin VB.CommandButton cmdPrintOpt 
      Caption         =   "Options de l'imprimante"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Configurer l'imprimante/impression"
      Top             =   3833
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   4313
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aperçu avant impression..."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Visualiser un aperçu avant d'imprimer"
      Top             =   4313
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrint"
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
'FORM D'IMPRESSION
'=======================================================

Private Lang As New clsLang
Private clsPref As clsIniForm

Private Sub cmdPrint_Click()
'impression


    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_PrintOK"))
End Sub

Private Sub cmdPrintOpt_Click()
'options de l'imprimante
Dim pt As Printer

    If frmContent.ActiveForm Is Nothing Then Exit Sub

    'affiche la boite de dialogue de choix de l'imprimante
    'récupère les propriétés du printer choisi
    Call GetPrinter(pt)
    
    If pt Is Nothing Then Exit Sub
    
    If TypeOfActiveForm = "Pfm" Then Call PrintFile(CCur(Val(txtFrom.Text)), _
        CCur(Val(txtTo.Text)), CBool(Check1(1).Value), CBool(Check1(0).Value), _
        CBool(Check1(2).Value), CBool(Check1(3).Value), _
        CLng(Val(txtFontSize.Text)), pt, IIf(Check1(4).Value, txtTitle.Text, _
        vbNullString))
        'lance l'impression du fichier en cours de visualisation
    
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
'aperçu avant impression

End Sub

Private Sub Form_Load()

    Set clsPref = New clsIniForm
    
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
    
    'loading des preferences
    Call clsPref.GetFormSettings(App.Path & "\Preferences\PrintFile.ini", Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\PrintFile.ini", Me)
    Set clsPref = Nothing
End Sub
