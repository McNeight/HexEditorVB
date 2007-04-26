VERSION 5.00
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDates 
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
   Icon            =   "frmDates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Rafraichir"
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Rafraichir les dates"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernier accès actuelle"
      Height          =   1095
      Index           =   4
      Left            =   3600
      TabIndex        =   22
      Top             =   1920
      Width           =   3015
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernière modification actuelle"
      Height          =   1215
      Index           =   5
      Left            =   3600
      TabIndex        =   19
      Top             =   3120
      Width           =   3015
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      ToolTipText     =   "Fermer la fenêtre"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Appliquer les changements"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernière modification"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3015
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   8
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   120
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   63569923
            CurrentDate     =   39133
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernier accès"
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3015
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   6
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   120
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   63569923
            CurrentDate     =   39133
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de création"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   3375
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
         Format          =   63569923
         CurrentDate     =   39133.9583333333
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   3135
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   4
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de création actuelle"
      Height          =   1095
      Index           =   3
      Left            =   3600
      TabIndex        =   12
      Top             =   720
      Width           =   3015
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "Sélection du fichier"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Emplacement du fichier"
      Top             =   240
      Width           =   3615
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Label Label1 
      Caption         =   "Fichier :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   240
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

Private mFile As filesystemlibrary.File

'=======================================================
'attribution des dates spécifiées au Fichier (text1.text)
'=======================================================
Private Sub AttribDates()
Dim lngHandle As Long
Dim udtFileTime As FILETIME
Dim ucreationFileTime As FILETIME
Dim ucaccessFileTime As FILETIME
Dim udtLocalTime As FILETIME
Dim ucreationFileTimelocal As FILETIME
Dim ucaccessFileTimelocal As FILETIME
Dim udtSystemTime As SYSTEMTIME
Dim ucreationSystemTime As SYSTEMTIME
Dim ucaccessSystemTime As SYSTEMTIME

    On Error GoTo ErrGestion

    
    'stocke les données de date & d'heure à la nouvelle heure (last modification)
    With udtSystemTime
        .wYear = DT(2).Year
        .wMonth = DT(2).Month
        .wDay = DT(2).Day
        '.wDayOfWeek = Weekday(Left$(DT(2).Value, 10)) - 1
        .wHour = DT(2).Hour
        .wMinute = DT(2).Minute
        .wSecond = DT(2).Second
        .wMilliseconds = 0
    End With
    
    'idem pour la date & heure de création
    With ucreationSystemTime
        .wYear = DT(0).Year
        .wMonth = DT(0).Month
        .wDay = DT(0).Day
        '.wDayOfWeek = Weekday(Left$(DT(0).Value, 10)) - 1
        .wHour = DT(0).Hour
        .wMinute = DT(0).Minute
        .wSecond = DT(0).Second
        .wMilliseconds = 0
    End With
        
    'idem pour la date et heure de dernier accès
    With ucaccessSystemTime
        .wYear = DT(1).Year
        .wMonth = DT(1).Month
        .wDay = DT(1).Day
        '.wDayOfWeek = Weekday(Left$(DT(1).Value, 10)) - 1
        .wHour = DT(1).Hour
        .wMinute = DT(1).Minute
        .wSecond = DT(1).Second
        .wMilliseconds = 0
    End With
    
    'convertir SystemTime ==> FileTime pour les 3 heures/dates
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    SystemTimeToFileTime ucreationSystemTime, ucreationFileTimelocal
    SystemTimeToFileTime ucaccessSystemTime, ucaccessFileTimelocal
    
    'convertir l'heure locale vers l'heure universelle
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    LocalFileTimeToFileTime ucreationFileTimelocal, ucreationFileTime
    LocalFileTimeToFileTime ucaccessFileTimelocal, ucaccessFileTime
    
    'obtient le handle du fichier
    lngHandle = CreateFile(txtFile.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    'applique les changements
    SetFileTime lngHandle, ucreationFileTime, ucaccessFileTime, udtFileTime
    
    'ferme le handle
    CloseHandle lngHandle
    
    MsgBox Lang.GetString("_ChangeOk"), vbInformation, Lang.GetString("_DateChange")
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmDates.AttribDates", True
    
    MsgBox Lang.GetString("_ErrorH") & vbNewLine & Err.Description, vbInformation, Lang.GetString("_DateChange")
End Sub

Private Sub cmdAppliquer_Click()
    AttribDates
    Call ChangeDates
End Sub

Private Sub cmdBrowse_Click()
'ouvre un fichier

    txtFile.Text = cFile.ShowOpen(Lang.GetString("_SelFile"), Me.hWnd, Lang.GetString("_All") & " |*.*")
    
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
    lblDate(0).Caption = Left$(mFile.DateCreated, 10)
    lblDate(1).Caption = Left$(mFile.DateLastAccessed, 10)
    lblDate(2).Caption = Left$(mFile.DateLastModified, 10)
    
    'affichage des infos d'heures
    lblHour(0).Caption = Right$(mFile.DateCreated, 8)
    lblHour(1).Caption = Right$(mFile.DateLastAccessed, 8)
    lblHour(2).Caption = Right$(mFile.DateLastModified, 8)
    
    'heures par défaut
    cmdDefaut_Click (0): cmdDefaut_Click (1): cmdDefaut_Click (2)
End Sub

'=======================================================
'récupère le fichier
'=======================================================
Public Sub GetFile(ByVal tFile As filesystemlibrary.File)
    Set mFile = tFile
    Call ChangeDates
End Sub

Private Sub Form_Load()
    #If MODE_DEBUG Then
        If App.LogMode = 0 And CREATE_FRENCH_FILE Then
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
