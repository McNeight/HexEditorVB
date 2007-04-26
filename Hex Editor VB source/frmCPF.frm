VERSION 5.00
Object = "{EF4A8ABF-4214-4B3F-8F82-ACF6D11FA80D}#1.0#0"; "BGraphe_OCX.ocx"
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmCPF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparaison de fichiers"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCPF.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.CommandButton cmdExamineDifferences 
      Caption         =   "Examiner les différences"
      Height          =   495
      Left            =   1099
      TabIndex        =   5
      ToolTipText     =   "Lancer une analyse détaillée"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCLose 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   5697
      TabIndex        =   7
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveReport 
      Caption         =   "Sauvegarder le rapport..."
      Height          =   495
      Left            =   3297
      TabIndex        =   6
      ToolTipText     =   "Sauvegarder le rapport au format texte"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Résultats"
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   8055
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   120
         ScaleHeight     =   4215
         ScaleWidth      =   7815
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   7815
         Begin BGraphe_OCX.BGraphe BG2 
            Height          =   3015
            Left            =   3960
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5318
            BarreColor1     =   0
            BarreColor2     =   16711680
         End
         Begin BGraphe_OCX.BGraphe BG1 
            Height          =   3015
            Left            =   0
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5318
            BarreColor1     =   0
            BarreColor2     =   16711680
         End
         Begin VB.Label lblTailles 
            Caption         =   "Tailles :"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   6015
         End
         Begin VB.Label lblMatch 
            Caption         =   "Pourcentage de correspondance :"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   240
            Width           =   6015
         End
         Begin VB.Label lblDiffer 
            Caption         =   "Différences :"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   480
            Width           =   6015
         End
         Begin VB.Label Label1 
            Caption         =   "Fichier 1"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Fichier 2"
            Height          =   255
            Left            =   3960
            TabIndex        =   16
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Label lblF2 
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label lblF1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3720
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "Lancer l'analyse"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      ToolTipText     =   "Analyser"
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fichier 2"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   4455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtFile2 
            Height          =   285
            Left            =   0
            TabIndex        =   3
            ToolTipText     =   "Emplacement du fichier 2"
            Top             =   0
            Width           =   3375
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   4
            ToolTipText     =   "Sélectionner le fichier 2"
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fichier 1"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtFile1 
            Height          =   285
            Left            =   0
            TabIndex        =   1
            ToolTipText     =   "Emplacement du fichier 1"
            Top             =   0
            Width           =   3375
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   2
            ToolTipText     =   "Sélectionner le fichier 1"
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin ProgressBar_OCX.pgrBar PGB 
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Progression de l'analyse"
      Top             =   960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmCPF.frx":058A
      FrontPicture    =   "frmCPF.frx":05A6
   End
End
Attribute VB_Name = "frmCPF"
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
'FORM DE COMPARAISON DE FICHIERS
'=======================================================

'=======================================================
'VARIABLES
'=======================================================
Private F1(255) As Long
Private F2(255) As Long


Private Sub cmdBrowse_Click(Index As Integer)
'browse for file

    On Error GoTo ErrGes
    
    With frmContent.CMD
        .CancelError = True
        .Filter = Lang.GetString("_All") & "  |*.*"
   
        If Index = 0 Then
            'fichier 1
            .DialogTitle = Lang.GetString("_SelectFile1")
            .ShowOpen
            txtFile1.Text = .Filename
        Else
            'fichier 2
            .DialogTitle = Lang.GetString("_SelectFile2")
            .ShowOpen
            txtFile2.Text = .Filename
        End If
    End With

ErrGes:
    
End Sub

Private Sub cmdCLose_Click()
'quitte
    Unload Me
End Sub

Private Sub cmdExamineDifferences_Click()
    
    frmMerge.Show vbModal
End Sub

Private Sub cmdGo_Click()
'lance l'analyse
    
    If cFile.FileExists(txtFile1.Text) = False Or cFile.FileExists(txtFile2.Text) = False Then
        'un des deux fichiers n'existe pas
        MsgBox Lang.GetString("_FileMiss"), vbInformation, Lang.GetString("_Failed")
        Exit Sub
    End If
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_AnalyseProc"))
    
    LaunchAnalys    'lance l'analyse
    DisplayResults  'affiche les résultats
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_AnalyseFin"))
    
End Sub

Private Sub cmdSaveReport_Click()
'sauvegarde le rapport



    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_RapFin"))
End Sub

'=======================================================
'lance l'analyse des deux fichiers
'=======================================================
Private Sub LaunchAnalys()
Dim lLength1 As Long, lLength2 As Long
Dim x As Long
Dim y As Long
Dim b As Byte
Dim l As Long
Dim tOver As OVERLAPPED
Dim strBuffer As String
Dim curByte As Currency
Dim lngFile As Long
Dim curByteOld As Currency
    
    On Error GoTo ErrGestion
    
    'vide les listes
    For x = 0 To 255
        F1(x) = 0: F2(x) = 0
    Next x
    
    'prépare la progressbar
    lLength1 = cFile.GetFileSizes(txtFile1.Text).FileSize: lLength2 = cFile.GetFileSizes(txtFile2.Text).FileSize
    PGB.Min = 0: PGB.Max = lLength1 + lLength2: PGB.Value = 0
    x = 0
    
    'obtient le handle du fichier
    lngFile = CreateFile(txtFile1.Text, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)
    
    'vérifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    'créé un buffer de 50Ko
    strBuffer = String$(51200, 0) 'buffer de 50K
    
    curByte = 0
    Do Until curByte > lLength1  'tant que le fichier n'est pas fini
    
        x = x + 1
    
        'prépare le type OVERLAPPED - obtient 2 long à la place du Currency
        GetLargeInteger curByte, tOver.Offset, tOver.OffsetHigh
        
        'obtient la string sur le buffer
        ReadFileEx lngFile, ByVal strBuffer, 51200, tOver, AddressOf CallBackFunction
        
        If curByte + 51200 <= lLength1 Then
            'alors on prend bien 51200 car
            l = 51200
        Else
            'on prend que les derniers car
            l = lLength1 - curByte
        End If
        
        For y = 1 To l
            b = Asc(Mid$(strBuffer, y, 1))
            'ajoute une occurence
            F1(b) = F1(b) + 1
        Next y
        
        If (x Mod 10) = 0 Then
            'rend la main
            DoEvents
            PGB.Value = curByte
        End If
        
        curByte = curByte + 51200
    
    Loop
    
    'Close lFile
    CloseHandle lngFile
      

 
    x = 0: curByteOld = curByte
    
    'obtient le handle du fichier
    lngFile = CreateFile(txtFile2.Text, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)
    
    'vérifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    'créé un buffer de 50Ko
    strBuffer = String$(51200, 0) 'buffer de 50K
    
    curByte = 0
    Do Until curByte > lLength2  'tant que le fichier n'est pas fini
    
        x = x + 1
    
        'prépare le type OVERLAPPED - obtient 2 long à la place du Currency
        GetLargeInteger curByte, tOver.Offset, tOver.OffsetHigh
        
        'obtient la string sur le buffer
        ReadFileEx lngFile, ByVal strBuffer, 51200, tOver, AddressOf CallBackFunction
        
        If curByte + 51200 <= lLength2 Then
            'alors on prend bien 51200 car
            l = 51200
        Else
            'on prend que les derniers car
            l = lLength2 - curByte
        End If
        
        For y = 1 To l
            b = Asc(Mid$(strBuffer, y, 1))
            'ajoute une occurence
            F2(b) = F2(b) + 1
        Next y
        
        If (x Mod 10) = 0 Then
            'rend la main
            DoEvents
            PGB.Value = curByte + curByteOld
        End If
        
        curByte = curByte + 51200
    
    Loop
    
    CloseHandle lngFile
    
    PGB.Value = PGB.Max
    DisplayResults   'affiche les résultats

    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmCPF.LaunchAnalysis", True
    
End Sub

'=======================================================
'affiche les résultats (graphes & labels)
'=======================================================
Private Sub DisplayResults()
Dim x As Long

    'remplit les graphes
    BG1.ClearGraphe
    BG1.ClearValues
    BG2.ClearGraphe
    BG2.ClearValues
    
    For x = 0 To 255
        BG1.AddValue x, F1(x)
        BG2.AddValue x, F2(x)
    Next x

    'trace les graphes
    BG1.TraceGraph
    BG2.TraceGraph
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
