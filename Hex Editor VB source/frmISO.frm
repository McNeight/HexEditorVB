VERSION 5.00
Object = "{3AF19019-2368-4F9C-BBFC-FD02C59BD0EC}#1.0#0"; "DriveView_OCX.ocx"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmISO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Création de fichier ISO"
   ClientHeight    =   3255
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
   HelpContextID   =   39
   Icon            =   "frmISO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   685
      Left            =   120
      TabIndex        =   4
      Top             =   2010
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1217
      Caption         =   "Fichier ISO résultant"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtFile 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Tag             =   "Choix du fichier ISO résultant"
         Top             =   320
         Width           =   2295
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         ToolTipText     =   "Sélectionne un fichier résultant"
         Top             =   320
         Width           =   375
      End
   End
   Begin vkUserContolsXP.vkBar PGB 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      Value           =   1
      BackPicture     =   "frmISO.frx":000C
      FrontPicture    =   "frmISO.frx":0028
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3608
      TabIndex        =   2
      ToolTipText     =   "Démarre la création du fichier ISO"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   3368
      TabIndex        =   1
      ToolTipText     =   "Ferme cette fenêtre"
      Top             =   2760
      Width           =   855
   End
   Begin DriveView_OCX.DriveView DV 
      Height          =   1815
      Left            =   128
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3201
      DisplayPhysicalDrives=   0   'False
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

Private Lang As New clsLang
Private tDrive As DriveView_OCX.clsDrive
Private clsPref As clsIniForm
Dim sFile As String


Private Sub cmdBrowse_Click()
'browse for file

    On Error GoTo ErrGes
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_BrowseForFile")
        .Filter = "Iso File |*.iso|"
        .FileName = vbNullString
        .ShowSave
        txtFile.Text = .FileName
    End With
    
ErrGes:
End Sub

Private Sub cmdGo_Click()
'lance la création du fichier ISO
Dim lngSec As Long
Dim S As String
Dim X As Long
Dim lFile As Long
Dim lBPC As Long
Dim lTLS As Long
Dim hDisk As Long

    'récupère le fichier
    sFile = txtFile.Text
    
    'vérifie si le fichier existe pas déjà
    If cFile.FileExists(sFile) Then
        'message de confirmation
        X = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If
    
    'vérifie que le dossier existe bien
    If cFile.FolderExists(cFile.GetParentFolderName(sFile)) = False Then
        'inexistant
        MsgBox Lang.GetString("_DirDoesNot"), vbCritical, Lang.GetString("_Error")
        Exit Sub
    End If
    
    txtFile.Enabled = False
    cmdBrowse.Enabled = False
    cmdGO.Enabled = False
    cmdQuit.Enabled = False
    
    With PGB
        .Max = tDrive.TotalPhysicalSectors - 1
        .Min = 0
        .Value = 0
    End With
    
    'créé un fichier vide
    Call cFile.CreateEmptyFile(sFile, True)
    
    'récupère les handles
    lFile = GetFileHandleWrite(sFile)
    hDisk = GetDiskHandleRead(tDrive.VolumeLetter & ":\")
    
    lBPC = tDrive.BytesPerSector
    lTLS = tDrive.TotalPhysicalSectors - 1
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ISOBegin"))
    
    'pour chaque secteur
    For lngSec = 0 To lTLS
        
        'récupère le contenu du secteur
        Call DirectReadSHandle(hDisk, lngSec, lBPC, lBPC, S)
        
        'sauve dans le fichier
        Call WriteBytesToFileEndHandle(lFile, S)
        
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
    Call CloseHandle(lFile)
    Call CloseHandle(hDisk)

End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub DV_NodeClick(ByVal Node As ComctlLib.INode)
Dim S As String

    If DV.IsSelectedDriveAccessible Then
    
        'récupère le type de fichier du drive
        S = DV.GetSelectedDrive.FileSystemName
        
        If S = "UDF" Or S = "CDFS" Then
            'on a choisi le drive
            Set tDrive = DV.GetSelectedDrive
            cmdGO.Enabled = True
        Else
            cmdGO.Enabled = False
        End If
    Else
        cmdGO.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Call MarkNonCDDrives(DV)
End Sub

Private Sub Form_Load()
'loading des preferences
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
        
        DV.LogicalDrivesString = .GetString("_LogicalString")
    
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Lang = Nothing
End Sub
