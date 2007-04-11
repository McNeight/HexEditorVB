VERSION 5.00
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form frmISO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Création de fichier ISO"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmISO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar PGB 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmISO.frx":000C
      FrontPicture    =   "frmISO.frx":0028
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      ToolTipText     =   "Ferme cette fenêtre"
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fichier ISO résultant"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   120
         ScaleHeight     =   330
         ScaleWidth      =   3015
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            TabIndex        =   4
            Tag             =   "Choix du fichier ISO résultant"
            Top             =   0
            Width           =   2295
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   255
            Left            =   2520
            TabIndex        =   3
            ToolTipText     =   "Sélectionne un fichier résultant"
            Top             =   0
            Width           =   375
         End
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      ToolTipText     =   "Démarre la création du fichier ISO"
      Top             =   240
      Width           =   615
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
End
Attribute VB_Name = "frmISO"
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
'FORM DE CREATION DE FICHIER ISO
'=======================================================

Private tDrive As clsDrive
Private clsPref As clsIniForm
Dim sFile As String


Private Sub cmdBrowse_Click()
'browse for file

    On Error GoTo ErrGes
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_BrowseForFile")
        .Filter = "Iso File |*.iso|"
        .Filename = vbNullString
        .ShowSave
        txtFile.Text = .Filename
    End With
    
ErrGes:
End Sub

Private Sub cmdGo_Click()
'lance la création du fichier ISO
Dim lngSec As Long
Dim s As String
Dim x As Long
Dim lFile As Long
Dim lBPC As Long
Dim lTLS As Long
Dim hDisk As Long

    'récupère le fichier
    sFile = txtFile.Text
    
    'vérifie si le fichier existe pas déjà
    If cFile.FileExists(sFile) Then
        'message de confirmation
        x = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War"))
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'vérifie que le dossier existe bien
    If cFile.FolderExists(cFile.GetParentDirectory(sFile)) = False Then
        'inexistant
        MsgBox Lang.GetString("_DirDoesNot"), vbCritical, Lang.GetString("_Error")
        Exit Sub
    End If
    
    txtFile.Enabled = False
    cmdBrowse.Enabled = False
    cmdGO.Enabled = False
    cmdQuit.Enabled = False
    
    PGB.Max = tDrive.TotalPhysicalSectors - 1
    PGB.Min = 0
    PGB.Value = 0
    
    'créé un fichier vide
    Call cFile.CreateEmptyFile(sFile, True)
    
    'récupère les handles
    lFile = GetFileHandle(sFile)
    hDisk = GetDiskHandle(tDrive.VolumeLetter & ":\")
    
    lBPC = tDrive.BytesPerSector
    lTLS = tDrive.TotalPhysicalSectors - 1
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ISOBegin"))
    
    'pour chaque secteur
    For lngSec = 0 To lTLS
        
        'récupère le contenu du secteur
        Call DirectReadSHandle(hDisk, lngSec, lBPC, lBPC, s)
        
        'sauve dans le fichier
        Call WriteBytesToFileEndHandle(lFile, s)
        
        If (lngSec Mod 500) = 0 Then
            DoEvents   'rend la main de tps en tps
            PGB.Value = lngSec
        End If
            
    Next lngSec
    
    PGB.Value = PGB.Max
    
    txtFile.Enabled = True
    cmdBrowse.Enabled = True
    cmdGO.Enabled = True
    cmdQuit.Enabled = True

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ISOEnd"))

    'referme les handles
    CloseHandle lFile
    CloseHandle hDisk

End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'loading des preferences
    Set clsPref = New clsIniForm
        
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

'=======================================================
'récupère la description du CD/DVD
'=======================================================
Public Sub GetDrive(ByVal tDriveRec As clsDrive)
    Set tDrive = tDriveRec
End Sub
