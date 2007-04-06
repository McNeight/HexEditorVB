VERSION 5.00
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporter"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkClip 
      Caption         =   "Copier dans le clipboard"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Permet de copier dans le clipboard plutôt que de créer un fichier"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Ne pas sauvegarder"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Lancer la sauvegarde"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Lancer la sauvegarde"
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format d'export"
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   4215
      Begin VB.CheckBox chkOffset 
         Caption         =   "Ajouter les offsets"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Permer d'ajouter les offsets au fichier (en hexa ou décimal, selon les préférences)"
         Top             =   840
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Index           =   1
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3975
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtOpt 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2400
            TabIndex        =   6
            Text            =   "texte de l'option"
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkString 
            Caption         =   "Ajouter les valeurs ASCII"
            Height          =   195
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "Permet l'ajout des valeurs ASCII au fichier"
            Top             =   840
            Width           =   2175
         End
         Begin VB.ComboBox cbFormat 
            Height          =   315
            ItemData        =   "frmExport.frx":000C
            Left            =   0
            List            =   "frmExport.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Format d'exportation"
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label lbl 
            Caption         =   "option"
            Height          =   255
            Left            =   2400
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fichier à sauvegarder"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3975
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   3480
            TabIndex        =   1
            ToolTipText     =   "Choix du fichier à sauvegarder"
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   0
            ToolTipText     =   "Emplacement du fichier à sauvegarder"
            Top             =   120
            Width           =   3255
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
Attribute VB_Name = "frmExport"
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
'FORM D'EXPORT MULTIFORMATS
'=======================================================

Private bEntireFile As Boolean

Private Sub cbFormat_Click()
'affiche l'option du format

    lbl.Visible = False
    txtOpt.Visible = False
    chkString.Enabled = True
    chkOffset.Enabled = True
            
    Select Case cbFormat.Text
        Case "RTF"
        
        Case "Texte"
            Me.Caption = "Exporter en texte (fichier 5 fois plus grand)"
        Case "Source C"
            Me.Caption = "Exporter en code C (fichier 6 fois plus grand)"
            chkString.Enabled = False
            chkOffset.Enabled = False
        Case "Source VB"
            Me.Caption = "Exporter en code VB (fichier 2 fois plus grand)"
            chkString.Enabled = False
            chkOffset.Enabled = False
            lbl.Caption = "Car. séparateur"
            txtOpt.Text = vbNullString
            lbl.Visible = True
            lbl.Enabled = True
            txtOpt.Enabled = True
            txtOpt.ToolTipText = "Caractère de séparation des valeurs hexadécimales"
            txtOpt.Visible = True
        Case "Source JAVA"
            Me.Caption = "Exporter en code JAVA (fichier 6 fois plus grand)"
            chkString.Enabled = False
            chkOffset.Enabled = False
        Case "HTML"
            Me.Caption = "Exporter en HTML (fichier 13 fois plus grand)"
            lbl.Caption = "Taille (1-10)"
            txtOpt.Text = "3"
            lbl.Visible = True
            lbl.Enabled = True
            txtOpt.Enabled = True
            txtOpt.ToolTipText = "Taille du texte"
            txtOpt.Visible = True
        Case "Else"
            Me.Caption = "Exporter"
    End Select
End Sub

Private Sub chkClip_Click()
    txtFile.Enabled = Not (CBool(chkClip.Value))
    cmdBrowse.Enabled = Not (CBool(chkClip.Value))
    Frame1(0).Enabled = Not (CBool(chkClip.Value))
End Sub

Private Sub cmdBrowse_Click()
'browse for file
Dim sFile As String

    sFile = cFile.ShowSave("Sélectionner le fichier à créer", Me.hWnd, "Tous|*.*", App.Path)
    
    If cFile.FileExists(sFile) Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà. Le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If
    
    txtFile.Text = sFile
        
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'lance la sauvegarde
Dim x As Long

    'ajoute du texte à la console
    Call AddTextToConsole("Exportation en cours...")
    Frame1(0).Enabled = False
    txtFile.Enabled = False
    chkOffset.Enabled = False
    chkString.Enabled = False
    lbl.Enabled = False
    txtOpt.Enabled = False
    cbFormat.Enabled = False
    cmdBrowse.Enabled = False
    Frame1(1).Enabled = False
    cmdSave.Enabled = False
    cmdQuit.Enabled = False
    DoEvents

    Select Case cbFormat.Text
        Case "HTML"
            
            x = Int(Abs(Val(txtOpt.Text)))
            If x < 1 Or x > 10 Then
                MsgBox "Taille non valide", vbCritical, "Attention"
                GoTo ResumeMe
            End If
            
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsHTML(txtFile.Text, CBool(chkOffset.Value), CBool(chkString.Value), _
                    frmContent.ActiveForm.Caption, -1, , x, CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsHTML(txtFile.Text, CBool(chkOffset.Value), CBool(chkString.Value), _
                    "az", 1, 1, x, CBool(chkClip.Value))
            End If
            
        Case "RTF"
            
            
            
        Case "Texte"
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsTEXT(txtFile.Text, CBool(chkOffset.Value), CBool(chkString.Value), _
                    frmContent.ActiveForm.Caption, -1, , CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsTEXT(txtFile.Text, CBool(chkOffset.Value), CBool(chkString.Value), _
                    "az", 1, 1, CBool(chkClip.Value))
            End If
            
        Case "Source C"
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsC(txtFile.Text, frmContent.ActiveForm.Caption, -1, , CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsC(txtFile.Text, frmContent.ActiveForm.Caption, 1, 1, CBool(chkClip.Value))
            End If
            
        Case "Source VB"
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsVB(txtFile.Text, frmContent.ActiveForm.Caption, -1, , txtOpt.Text, CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsVB(txtFile.Text, frmContent.ActiveForm.Caption, 1, 1, txtOpt.Text, CBool(chkClip.Value))
            End If
            
        Case "Source JAVA"
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsJAVA(txtFile.Text, frmContent.ActiveForm.Caption, -1, , CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsJAVA(txtFile.Text, frmContent.ActiveForm.Caption, 1, 1, CBool(chkClip.Value))
            End If
            
    End Select
    
ResumeMe:
    Frame1(0).Enabled = Not (CBool(chkClip.Value))
    txtFile.Enabled = Not (CBool(chkClip.Value))
    chkOffset.Enabled = (cbFormat.Text = "HTML" Or cbFormat.Text = "Texte")
    chkString.Enabled = (cbFormat.Text = "HTML" Or cbFormat.Text = "Texte")
    lbl.Enabled = (cbFormat.Text = "HTML" Or cbFormat.Text = "Source VB")
    txtOpt.Enabled = (cbFormat.Text = "HTML" Or cbFormat.Text = "Source VB")
    cbFormat.Enabled = True
    cmdBrowse.Enabled = Not (CBool(chkClip.Value))
    Frame1(1).Enabled = True
    cmdSave.Enabled = True
    cmdQuit.Enabled = True
    DoEvents
    'ajoute du texte à la console
    Call AddTextToConsole("Exportation terminée")
End Sub

'=======================================================
'à appeler si on veut sauver un fichier entier
'=======================================================
Public Sub IsEntireFile()
    bEntireFile = True
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
End Sub
