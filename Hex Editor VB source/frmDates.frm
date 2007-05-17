VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDates 
   BackColor       =   &H00F9E5D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changement de dates"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   30
   Icon            =   "frmDates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   833
      TabIndex        =   25
      ToolTipText     =   "Emplacement du fichier"
      Top             =   173
      Width           =   3615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4673
      TabIndex        =   24
      ToolTipText     =   "Sélection du fichier"
      Top             =   173
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de création actuelle"
      Height          =   1095
      Index           =   3
      Left            =   3593
      TabIndex        =   21
      Top             =   653
      Width           =   3015
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de création"
      Height          =   855
      Index           =   0
      Left            =   113
      TabIndex        =   17
      Top             =   653
      Width           =   3375
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
         Format          =   64356355
         CurrentDate     =   39133.9583333333
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3135
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   20
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernier accès"
      Height          =   855
      Index           =   1
      Left            =   113
      TabIndex        =   13
      Top             =   1853
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3015
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   15
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   120
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   64356355
            CurrentDate     =   39133
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernière modification"
      Height          =   855
      Index           =   2
      Left            =   113
      TabIndex        =   9
      Top             =   3053
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3015
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   11
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   120
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   12
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   64356355
            CurrentDate     =   39133
         End
      End
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   113
      TabIndex        =   8
      ToolTipText     =   "Appliquer les changements"
      Top             =   4133
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1553
      TabIndex        =   7
      ToolTipText     =   "Fermer la fenêtre"
      Top             =   4133
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernière modification actuelle"
      Height          =   1215
      Index           =   5
      Left            =   3593
      TabIndex        =   4
      Top             =   3053
      Width           =   3015
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernier accès actuelle"
      Height          =   1095
      Index           =   4
      Left            =   3593
      TabIndex        =   1
      Top             =   1853
      Width           =   3015
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Rafraichir"
      Height          =   255
      Left            =   5273
      TabIndex        =   0
      ToolTipText     =   "Rafraichir les dates"
      Top             =   173
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fichier :"
      Height          =   255
      Index           =   0
      Left            =   113
      TabIndex        =   26
      Top             =   173
      Width           =   615
   End
End
Attribute VB_Name = "frmDates"
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
'FORM PERMETTANT DE CHANGER LES DATES D'UN FICHIER
'=======================================================
Private Lang As New clsLang
Private mFile As FileSystemLibrary.File

'=======================================================
'attribution des dates spécifiées au Fichier (txtFile.text)
'=======================================================
Private Sub AttribDates()

    Call cFile.SetFileDates(txtFile.Text, DT(0).Value, DT(1).Value, DT(2).Value)
    
    MsgBox Lang.GetString("_ChangeOk"), vbInformation, _
        Lang.GetString("_DateChange")
    
End Sub

Private Sub cmdAppliquer_Click()
    Call AttribDates
    Call ChangeDates
End Sub

Private Sub cmdBrowse_Click()
'ouvre un fichier

    txtFile.Text = cFile.ShowOpen(Lang.GetString("_SelFile"), Me.hWnd, _
        Lang.GetString("_All") & " |*.*")
    
    Call ChangeDates
End Sub

Private Sub cmdDefaut_Click(Index As Integer)
'affiche dans les textboxes les dates par défaut

    On Error Resume Next
    
    DT(Index).Value = lblDate(Index).Caption & " " & lblHour(Index).Caption
    
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call ChangeDates
End Sub

'=======================================================
'récupère les dates et change les controles
'=======================================================
Private Sub ChangeDates()
    Set mFile = cFile.GetFile(txtFile.Text) 'récupère le fichier
    
    'affichage des infos de dates
    With mFile
        lblDate(0).Caption = Left$(.DateCreated, 10)
        lblDate(1).Caption = Left$(.DateLastAccessed, 10)
        lblDate(2).Caption = Left$(.DateLastModified, 10)
        
        'affichage des infos d'heures
        lblHour(0).Caption = Right$(.DateCreated, 8)
        lblHour(1).Caption = Right$(.DateLastAccessed, 8)
        lblHour(2).Caption = Right$(.DateLastModified, 8)
    End With
    
    'heures par défaut
    cmdDefaut_Click (0): cmdDefaut_Click (1): cmdDefaut_Click (2)
End Sub

'=======================================================
'récupère le fichier
'=======================================================
Public Sub GetFile(ByVal tFile As FileSystemLibrary.File)
    
    'active la gestion des langues
    Call Lang.ActiveLang(Me)
    
    Set mFile = tFile
    Call ChangeDates
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
    
End Sub
