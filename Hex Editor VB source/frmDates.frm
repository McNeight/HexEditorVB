VERSION 5.00
Begin VB.Form frmDates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changement de dates"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8505
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Rafraichir"
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      ToolTipText     =   "Rafraichir les dates"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernier accès actuelle"
      Height          =   1455
      Index           =   4
      Left            =   5400
      TabIndex        =   18
      Top             =   2160
      Width           =   3015
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernière modification actuelle"
      Height          =   1455
      Index           =   5
      Left            =   5400
      TabIndex        =   12
      Top             =   3600
      Width           =   3015
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   4785
      TabIndex        =   8
      ToolTipText     =   "Fermer la fenêtre"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "Appliquer"
      Height          =   375
      Left            =   2505
      TabIndex        =   7
      ToolTipText     =   "Appliquer les changements"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernière modification"
      Height          =   1455
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   5175
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   2
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4935
         TabIndex        =   11
         Top             =   240
         Width           =   4935
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   2
            Left            =   600
            MaxLength       =   2
            TabIndex        =   53
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   52
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   2
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   51
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   2
            Left            =   600
            MaxLength       =   2
            TabIndex        =   50
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   49
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Index           =   2
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   48
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   375
            Index           =   2
            Left            =   3840
            TabIndex        =   17
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "jour"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "mois"
            Height          =   195
            Index           =   2
            Left            =   1200
            TabIndex        =   58
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "année"
            Height          =   195
            Index           =   2
            Left            =   2400
            TabIndex        =   57
            Top             =   120
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "heure"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "minute"
            Height          =   195
            Index           =   2
            Left            =   1155
            TabIndex        =   55
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "seconde"
            Height          =   195
            Index           =   2
            Left            =   2280
            TabIndex        =   54
            Top             =   600
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de dernier accès"
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5175
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   1
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4935
         TabIndex        =   10
         Top             =   240
         Width           =   4935
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   600
            MaxLength       =   2
            TabIndex        =   41
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   40
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   39
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   1
            Left            =   600
            MaxLength       =   2
            TabIndex        =   38
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   37
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Index           =   1
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   36
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   16
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "jour"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "mois"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   46
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "année"
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   45
            Top             =   120
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "heure"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "minute"
            Height          =   195
            Index           =   1
            Left            =   1155
            TabIndex        =   43
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "seconde"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   42
            Top             =   600
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de création"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5175
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   4935
         TabIndex        =   9
         Top             =   240
         Width           =   4935
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   2
            TabIndex        =   29
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   28
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   0
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   27
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   2
            TabIndex        =   26
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   25
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Index           =   0
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   24
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdDefaut 
            Caption         =   "Par défaut"
            Height          =   375
            Index           =   0
            Left            =   3840
            TabIndex        =   15
            ToolTipText     =   "Dates par défaut (actuelles) du fichier"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "jour"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "mois"
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   34
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "année"
            Height          =   195
            Index           =   0
            Left            =   2400
            TabIndex        =   33
            Top             =   120
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "heure"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "minute"
            Height          =   195
            Index           =   0
            Left            =   1155
            TabIndex        =   31
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "seconde"
            Height          =   195
            Index           =   0
            Left            =   2280
            TabIndex        =   30
            Top             =   600
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date de création actuelle"
      Height          =   1455
      Index           =   3
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   3015
      Begin VB.Label lblHour 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   6120
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
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Fichier :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmDates"
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
'FORM PERMETTANT DE CHANGER LES DATES D'UN FICHIER
'-------------------------------------------------------

'-------------------------------------------------------
'attribution des dates spécifiées au Fichier (text1.text)
'-------------------------------------------------------
Private Sub AttribDates()
Dim modif_Date As Date, lngHandle As Long, creation_Date As Date, access_Date As Date
Dim udtFileTime As FILETIME, ucreationFileTime As FILETIME, ucaccessFileTime As FILETIME
Dim udtLocalTime As FILETIME, ucreationFileTimelocal As FILETIME, ucaccessFileTimelocal As FILETIME
Dim udtSystemTime As SYSTEMTIME, ucreationSystemTime As SYSTEMTIME, ucaccessSystemTime As SYSTEMTIME

    On Error GoTo ErrGestion
    
    'obtient les dates à partir des textboxes (utilisé pour le DayOfWeek)
    modif_Date = Text2(2).Text & "/" & Text3(2).Text & "/" & Text4(2).Text
    creation_Date = Text2(0).Text & "/" & Text3(0).Text & "/" & Text4(0).Text
    access_Date = Text2(1).Text & "/" & Text3(1).Text & "/" & Text4(1).Text
    
    'stocke les données de date & d'heure à la nouvelle heure (last modification)
    With udtSystemTime
        .wYear = Year(modif_Date)
        .wMonth = Month(modif_Date)
        .wDay = Day(modif_Date)
        .wDayOfWeek = Weekday(modif_Date) - 1
        .wHour = FormatedVal(Text5(2).Text)
        .wMinute = FormatedVal(Text6(2).Text)
        .wSecond = FormatedVal(Text7(2).Text)
        .wMilliseconds = 0
    End With
    
    'idem pour la date & heure de création
    With ucreationSystemTime
        .wYear = Year(creation_Date)
        .wMonth = Month(creation_Date)
        .wDay = Day(creation_Date)
        .wDayOfWeek = Weekday(creation_Date) - 1
        .wHour = FormatedVal(Text5(0).Text)
        .wMinute = FormatedVal(Text6(0).Text)
        .wSecond = FormatedVal(Text7(0).Text)
        .wMilliseconds = 0
    End With
        
    'idem pour la date et heure de dernier accès
    With ucaccessSystemTime
        .wYear = Year(access_Date)
        .wMonth = Month(access_Date)
        .wDay = Day(access_Date)
        .wDayOfWeek = Weekday(access_Date) - 1
        .wHour = FormatedVal(Text5(1).Text)
        .wMinute = FormatedVal(Text6(1).Text)
        .wSecond = FormatedVal(Text7(1).Text)
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
    
    'avertir du succès de la modification
    MsgBox "Le changement a fonctionné", vbInformation, "Changements des dates"
    
    Exit Sub

ErrGestion:
    clsERREUR.AddError "frmDates.AttribDates", True
    
    MsgBox "Une erreur est survenue." & vbNewLine & Err.Description, vbInformation, "Changements des dates"
End Sub

Private Sub cmdAppliquer_Click()
    AttribDates
    txtFile_Change
End Sub

Private Sub cmdBrowse_Click()
'ouvre un fichier

    On Error GoTo CancelErr
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "Sélection d'un fichier"
        .Filter = "Tous |*.*"
        .ShowOpen
        txtFile.Text = .Filename
    End With
    
    Exit Sub
    
CancelErr:
End Sub

Private Sub cmdDefaut_Click(Index As Integer)
'affiche dans les textboxes les dates par défaut

    On Error Resume Next
    
    Text2(Index).Text = Day(lblDate(Index).Caption)
    Text3(Index).Text = Month(lblDate(Index).Caption)
    Text4(Index).Text = Year(lblDate(Index).Caption)
    Text5(Index).Text = Hour(lblHour(Index).Caption)
    Text6(Index).Text = Minute(lblHour(Index).Caption)
    Text7(Index).Text = Second(lblHour(Index).Caption)
        
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    txtFile_Change
End Sub

Private Sub txtFile_Change()
'si le fichier existe, alors on obtient les dates et on les affiches dans les captions, et les textboxes
Dim lHandle As Long
Dim CREATION_D As FILETIME, CREATION__D As SYSTEMTIME
Dim ACCESS_D As FILETIME, ACCESS__D As SYSTEMTIME
Dim MODIF_D As FILETIME, MODIF__D As SYSTEMTIME
   
    If cFile.FileExists(txtFile.Text) Then
        'récupère le handle du fichier
        lHandle = CreateFile(txtFile.Text, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
        
        'récupère les dates
        GetFileTime lHandle, CREATION_D, ACCESS_D, MODIF_D
        
        'conversion de dates
        'temps local
        FileTimeToLocalFileTime CREATION_D, CREATION_D
        FileTimeToLocalFileTime ACCESS_D, ACCESS_D
        FileTimeToLocalFileTime MODIF_D, MODIF_D
        'temps système
        FileTimeToSystemTime CREATION_D, CREATION__D
        FileTimeToSystemTime ACCESS_D, ACCESS__D
        FileTimeToSystemTime MODIF_D, MODIF__D
        
        'affichage des infos de dates
        lblDate(0).Caption = Left$(FileTimeToString(CREATION_D, False), 10)
        lblDate(1).Caption = Left$(FileTimeToString(ACCESS_D, False), 10)
        lblDate(2).Caption = Left$(FileTimeToString(MODIF_D, False), 10)
        
        'affichage des infos d'heures
        lblHour(0).Caption = CREATION__D.wHour & ":" & CREATION__D.wMinute & ":" & CREATION__D.wSecond
        lblHour(1).Caption = ACCESS__D.wHour & ":" & ACCESS__D.wMinute & ":" & ACCESS__D.wSecond
        lblHour(2).Caption = MODIF__D.wHour & ":" & MODIF__D.wMinute & ":" & MODIF__D.wSecond
    End If
    
    'heures par défaut
    cmdDefaut_Click (0): cmdDefaut_Click (1): cmdDefaut_Click (2)

    CloseHandle lHandle
End Sub
