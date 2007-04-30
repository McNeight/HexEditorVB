VERSION 5.00
Begin VB.Form frmSaveProcess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sauvegarder le contenu mémoire du processus"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaveProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstList 
      Height          =   2985
      Left            =   128
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   488
      Width           =   3855
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Tout enregistrer (au minimum 2Go sont requis)"
      Height          =   375
      Left            =   128
      TabIndex        =   9
      Tag             =   "pref"
      ToolTipText     =   "Enregistre toute la mémoire (/!\ 2Go sont requis)"
      Top             =   3608
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contenu de l'enregistrement"
      Height          =   1575
      Left            =   4208
      TabIndex        =   4
      Top             =   128
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3015
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox chkHexa 
            Caption         =   "Sauvegarder les valeurs hexa"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Tag             =   "pref"
            ToolTipText     =   "Sauvegarder les valeurs hexa"
            Top             =   120
            Width           =   2895
         End
         Begin VB.CheckBox chkASCII 
            Caption         =   "Sauvegarder les valeurs ASCII"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Tag             =   "pref"
            ToolTipText     =   "Sauvegarder les valeurs ASCII réelles uniquement si coché seul"
            Top             =   480
            Width           =   2895
         End
         Begin VB.CheckBox chkOffset 
            Caption         =   "Sauvegarder les offsets"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Tag             =   "pref"
            ToolTipText     =   "La sauvegarde des offsets nécessite la sauvegarde de strings formatées"
            Top             =   840
            Width           =   2895
         End
      End
   End
   Begin VB.TextBox txtPath 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4208
      TabIndex        =   3
      ToolTipText     =   "Chemin du fichier résultat"
      Top             =   3008
      Width           =   2775
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   7088
      TabIndex        =   2
      ToolTipText     =   "Sélectionner l'emplacement du fichier à sauvegarder"
      Top             =   3008
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder"
      Height          =   375
      Left            =   4208
      TabIndex        =   1
      ToolTipText     =   "Sauvegarder dans le fichier sélectionné"
      Top             =   3608
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   5888
      TabIndex        =   0
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   3608
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Zones mémoire à enregistrer"
      Height          =   255
      Left            =   128
      TabIndex        =   13
      Top             =   128
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Chemin du fichier"
      Height          =   255
      Left            =   4208
      TabIndex        =   12
      Top             =   2648
      Width           =   3255
   End
   Begin VB.Label lblSize 
      Caption         =   "Taille du fichier résultant=[0]"
      Height          =   615
      Left            =   4208
      TabIndex        =   11
      ToolTipText     =   "Taille estimée du fichier qui sera créé"
      Top             =   1808
      Width           =   3255
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Tout cocher"
      End
      Begin VB.Menu mnuDeselectAll 
         Caption         =   "&Tout décocher"
      End
   End
End
Attribute VB_Name = "frmSaveProcess"
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
'FORM DE SAUVEGARDE DU CONTENU MEMOIRE D'UN PROCESSUS
'=======================================================

Private Lang As New clsLang
Private clsPref As clsIniForm
Private LS() As Long

Private Sub chkAll_Click()
    lstList.Enabled = Not (CBool(chkAll.Value))
    If chkAll.Value Then
        'alors on coche tout
        Call mnuSelectAll_Click
    End If
End Sub

Private Sub cmdBrowse_Click()
'browse
Dim X As Long
    
    On Error GoTo CancelPushed
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SelFileToSave")
        .Filter = Lang.GetString("_ExeFile") & " |*.exe|" & Lang.GetString("_TxtFile") & "|*.txt|Tous|*.*"
        .Filename = vbNullString
        .ShowSave
        txtPath.Text = .Filename
    End With
    
    If cFile.FileExists(txtPath.Text) Then
        'message de confirmation
        X = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + _
            vbYesNo, Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If
    
CancelPushed:
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'effectue la sauvegarde


    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_SaveOk"))
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
    Call clsPref.GetFormSettings(App.Path & "\Preferences\SaveProcess.ini", Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\SaveProcess.ini", Me)
    Set clsPref = Nothing
End Sub

'=======================================================
'recalcule la taille totale
'=======================================================
Private Sub RecalcSize()
'alors on recalcule la taille du fichier résultat
Dim lSize As Long
Dim X As Long
Dim Y As Long
Dim s As String
    
    lSize = 0
    For X = 0 To lstList.ListCount - 1
        s = Left$(lstList.List(X), Len(lstList.List(X)) - 1)  'garde l'item sans le ']' final
        Y = InStrRev(s, "[", , vbBinaryCompare)
        s = Mid$(s, Y + 1, Len(s) - Y) 'contient la taille
        
        If lstList.Selected(X) Then
            'ajoute la taille
            lSize = lSize + Val(s)
        End If
    Next X
    
    lblSize.Caption = Lang.GetString("_SizeRes") & Trim$(Str$(lSize)) & "]" & _
        vbNewLine & FormatedSize(lSize)
End Sub

Private Sub lstList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'affiche le popup menu sur le listbox
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp
    
    Call RecalcSize  'recalcule la taille
End Sub

Private Sub mnuDeselectAll_Click()
'décoche toutes les cases
Dim X As Long
    
    lstList.Visible = False
    
    For X = lstList.ListCount - 1 To 0 Step -1
        lstList.Selected(X) = False
    Next X
    
    lblSize.Caption = Lang.GetString("_SizeRes") & "0]"
    lstList.Visible = True
End Sub

Private Sub mnuSelectAll_Click()
'coche toutes les cases
Dim X As Long
    
    lstList.Visible = False
    
    For X = lstList.ListCount - 1 To 0 Step -1
        ValidateRect lstList.hWnd, 0&
        lstList.Selected(X) = True
    Next X
    
    lstList.Visible = True
    
    Call RecalcSize  'recalcule la taille
End Sub

'=======================================================
'obtient le processus concerné par l'enregistrement
'=======================================================
Public Sub GetProcess(ByVal lPID As Long, sFile As String)
Dim clsProc As clsMemoryRW
Dim LB() As Long
Dim X As Long

    txtPath.Text = sFile
    
    '//ajoute au listbox toutes les zones mémoire
    'liste les zones
    Set clsProc = New clsMemoryRW
    Call clsProc.RetrieveMemRegions(lPID, LB(), LS())
    
    Call lstList.Clear
    lstList.Visible = False
    
    'les ajoute
    For X = 1 To UBound(LS())
        lstList.AddItem "Offset=[" & CStr(LB(X)) & "], " & Lang.GetString("_Size") & "=[" & CStr(LS(X)) & "]"
    Next X
    
    lstList.Visible = True
        
End Sub
