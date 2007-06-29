VERSION 5.00
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
   HelpContextID   =   15
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Fichier à sauvegarder"
      Height          =   855
      Index           =   0
      Left            =   113
      TabIndex        =   10
      Top             =   105
      Width           =   4215
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3975
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   13
            ToolTipText     =   "Emplacement du fichier à sauvegarder"
            Top             =   120
            Width           =   3255
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   3480
            TabIndex        =   12
            ToolTipText     =   "Choix du fichier à sauvegarder"
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format d'export"
      Height          =   1575
      Index           =   1
      Left            =   113
      TabIndex        =   3
      Top             =   1425
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
         Begin VB.ComboBox cbFormat 
            Height          =   315
            ItemData        =   "frmExport.frx":000C
            Left            =   0
            List            =   "frmExport.frx":0022
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "pref lang_ok"
            ToolTipText     =   "Format d'exportation"
            Top             =   120
            Width           =   3975
         End
         Begin VB.CheckBox chkString 
            Caption         =   "Ajouter les valeurs ASCII"
            Height          =   195
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Permet l'ajout des valeurs ASCII au fichier"
            Top             =   840
            Width           =   2175
         End
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
         Begin VB.Label lbl 
            Caption         =   "option"
            Height          =   255
            Left            =   2400
            TabIndex        =   9
            Top             =   480
            Visible         =   0   'False
            Width           =   1335
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Lancer la sauvegarde"
      Height          =   375
      Left            =   113
      TabIndex        =   2
      ToolTipText     =   "Lancer la sauvegarde"
      Top             =   3105
      Width           =   2175
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   2873
      TabIndex        =   1
      ToolTipText     =   "Ne pas sauvegarder"
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CheckBox chkClip 
      Caption         =   "Copier dans le clipboard"
      Height          =   195
      Left            =   113
      TabIndex        =   0
      ToolTipText     =   "Permet de copier dans le clipboard plutôt que de créer un fichier"
      Top             =   1065
      Width           =   2535
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

Private Lang As New clsLang
Private bEntireFile As Boolean

Private Sub cbFormat_Click()
'affiche l'option du format

    lbl.Visible = False
    txtOpt.Visible = False
    chkString.Enabled = True
    chkOffset.Enabled = True
    
    With Lang
        Select Case cbFormat.Text
            Case .GetString("_RTF!")
            
            Case .GetString("_Text!")
                Me.Caption = .GetString("_ExportTxt")
            Case .GetString("_SourceC!")
                Me.Caption = .GetString("_ExportC")
                chkString.Enabled = False
                chkOffset.Enabled = False
            Case .GetString("_VB!")
                Me.Caption = .GetString("_ExportVB")
                chkString.Enabled = False
                chkOffset.Enabled = False
                lbl.Caption = .GetString("_CarSep")
                txtOpt.Text = vbNullString
                lbl.Visible = True
                lbl.Enabled = True
                txtOpt.Enabled = True
                txtOpt.ToolTipText = .GetString("_CarSepTool")
                txtOpt.Visible = True
            Case .GetString("_JAVA!")
                Me.Caption = .GetString("_ExportJAVA")
                chkString.Enabled = False
                chkOffset.Enabled = False
            Case .GetString("_HTML!")
                Me.Caption = .GetString("_ExportHTML")
                lbl.Caption = .GetString("_Size")
                txtOpt.Text = "3"
                lbl.Visible = True
                lbl.Enabled = True
                txtOpt.Enabled = True
                txtOpt.ToolTipText = .GetString("_SizeTool")
                txtOpt.Visible = True
            Case Else
                Me.Caption = .GetString("_ElseExport")
        End Select
    End With
End Sub

Private Sub chkClip_Click()
    With chkClip
        txtFile.Enabled = Not (CBool(.Value))
        cmdBrowse.Enabled = Not (CBool(.Value))
        Frame1(0).Enabled = Not (CBool(.Value))
    End With
End Sub

Private Sub cmdBrowse_Click()
'browse for file
Dim sFile As String

    sFile = cFile.ShowSave(Lang.GetString("_FileToCreate"), Me.hWnd, _
        Lang.GetString("_All") & "|*.*", App.Path)
    
    If cFile.FileExists(sFile) Then
        'fichier déjà existant
        If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
    
    txtFile.Text = sFile
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'lance la sauvegarde
Dim X As Long

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_Exporting"))
    
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
        Case Lang.GetString("_HTML!")
            
            X = Int(Abs(Val(txtOpt.Text)))
            If X < 1 Or X > 10 Then
                MsgBox Lang.GetString("_SizeNoOk"), vbCritical, _
                    Lang.GetString("_War")
                GoTo ResumeMe
            End If
            
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsHTML(txtFile.Text, CBool(chkOffset.Value), _
                    CBool(chkString.Value), frmContent.ActiveForm.Caption, _
                    -1, , X, CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsHTML(txtFile.Text, CBool(chkOffset.Value), _
                    CBool(chkString.Value), "az", 1, 1, X, CBool(chkClip.Value))
            End If
            
        Case Lang.GetString("_RTF!")
            
            
            
        Case Lang.GetString("_Text!")
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsTEXT(txtFile.Text, CBool(chkOffset.Value), _
                    CBool(chkString.Value), frmContent.ActiveForm.Caption, _
                    -1, , CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsTEXT(txtFile.Text, CBool(chkOffset.Value), _
                CBool(chkString.Value), "az", 1, 1, CBool(chkClip.Value))
            End If
            
        Case Lang.GetString("_SourceC!")
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsC(txtFile.Text, frmContent.ActiveForm.Caption, -1, _
                    , CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsC(txtFile.Text, frmContent.ActiveForm.Caption, 1, _
                    1, CBool(chkClip.Value))
            End If
            
        Case Lang.GetString("_VB!")
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsVB(txtFile.Text, frmContent.ActiveForm.Caption, _
                    -1, , txtOpt.Text, CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsVB(txtFile.Text, frmContent.ActiveForm.Caption, _
                    1, 1, txtOpt.Text, CBool(chkClip.Value))
            End If
            
        Case Lang.GetString("_JAVA!")
            If bEntireFile Then
                'sauvegarde d'un fichier entier
                Call SaveAsJAVA(txtFile.Text, frmContent.ActiveForm.Caption, _
                    -1, , CBool(chkClip.Value))
            Else
                'sauvegarde d'une plage d'offset
                Call SaveAsJAVA(txtFile.Text, frmContent.ActiveForm.Caption, _
                    1, 1, CBool(chkClip.Value))
            End If
            
    End Select
    
ResumeMe:
    Frame1(0).Enabled = Not (CBool(chkClip.Value))
    txtFile.Enabled = Not (CBool(chkClip.Value))
    With Lang
        chkOffset.Enabled = (cbFormat.Text = .GetString("_HTML!") Or cbFormat.Text = .GetString("_Text!"))
        chkString.Enabled = (cbFormat.Text = .GetString("_HTML!") Or cbFormat.Text = .GetString("_Text!"))
        lbl.Enabled = (cbFormat.Text = .GetString("_HTML!") Or cbFormat.Text = .GetString("_VB!"))
        txtOpt.Enabled = (cbFormat.Text = .GetString("_HTML!") Or cbFormat.Text = .GetString("_VB!"))
    End With
    cbFormat.Enabled = True
    cmdBrowse.Enabled = Not (CBool(chkClip.Value))
    Frame1(1).Enabled = True
    cmdSave.Enabled = True
    cmdQuit.Enabled = True
    
    DoEvents
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ExpOk"))
End Sub

'=======================================================
'à appeler si on veut sauver un fichier entier
'=======================================================
Public Sub IsEntireFile()
    bEntireFile = True
End Sub

Private Sub Form_Load()

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
            Lang.LangFolder = LANG_PATH
        Else
            Lang.LangFolder = App.Path & "\Lang"
        End If
        
        'applique la langue désirée aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        Call .LoadControlsCaption
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub
