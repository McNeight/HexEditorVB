VERSION 5.00
Begin VB.Form frmSaveProcess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sauvegarder le contenu m�moire du processus"
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      ToolTipText     =   "Fermer cette fen�tre"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "Sauvegarder dans le fichier s�lectionn�"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      ToolTipText     =   "S�lectionner l'emplacement du fichier � sauvegarder"
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Chemin du fichier r�sultat"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contenu de l'enregistrement"
      Height          =   1575
      Left            =   4200
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3015
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox chkOffset 
            Caption         =   "Sauvegarder les offsets"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Tag             =   "pref"
            ToolTipText     =   "La sauvegarde des offsets n�cessite la sauvegarde de strings format�es"
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox chkASCII 
            Caption         =   "Sauvegarder les valeurs ASCII"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Tag             =   "pref"
            ToolTipText     =   "Sauvegarder les valeurs ASCII r�elles uniquement si coch� seul"
            Top             =   480
            Width           =   2895
         End
         Begin VB.CheckBox chkHexa 
            Caption         =   "Sauvegarder les valeurs hexa"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Tag             =   "pref"
            ToolTipText     =   "Sauvegarder les valeurs hexa"
            Top             =   120
            Width           =   2895
         End
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Tout enregistrer (au minimum 2Go sont requis)"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Tag             =   "pref"
      ToolTipText     =   "Enregistre toute la m�moire (/!\ 2Go sont requis)"
      Top             =   3600
      Width           =   3855
   End
   Begin VB.ListBox lstList 
      Height          =   2985
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lblSize 
      Caption         =   "Taille du fichier r�sultant=[0]"
      Height          =   615
      Left            =   4200
      TabIndex        =   13
      ToolTipText     =   "Taille estim�e du fichier qui sera cr��"
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Chemin du fichier"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Zones m�moire � enregistrer"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Tout cocher"
      End
      Begin VB.Menu mnuDeselectAll 
         Caption         =   "&Tout d�cocher"
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
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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
Dim x As Long
    
    On Error GoTo CancelPushed
    
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "S�lection du fichier � enregistrer"
        .Filter = "Executable |*.exe|Fichier texte|*.txt|Tous|*.*"
        .ShowSave
        txtPath.Text = .Filename
    End With
    
    If cFile.FileExists(txtPath.Text) Then
        'message de confirmation
        x = MsgBox("Le fichier existe d�j�, le remplacer ?", vbInformation + vbYesNo, "Attention")
        If Not (x = vbYes) Then Exit Sub
    End If
    
CancelPushed:
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'effectue la sauvegarde


    'ajoute du texte � la console
    Call AddTextToConsole("Sauvegarde termin�e")
End Sub

Private Sub Form_Load()
    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\SaveProcess.ini", Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\SaveProcess.ini", Me
    Set clsPref = Nothing
End Sub

'=======================================================
'recalcule la taille totale
'=======================================================
Private Sub RecalcSize()
'alors on recalcule la taille du fichier r�sultat
Dim lSize As Long
Dim x As Long
Dim y As Long
Dim s As String
    
    lSize = 0
    For x = 0 To lstList.ListCount - 1
        s = Left$(lstList.List(x), Len(lstList.List(x)) - 1)  'garde l'item sans le ']' final
        y = InStrRev(s, "[", , vbBinaryCompare)
        s = Mid$(s, y + 1, Len(s) - y) 'contient la taille
        
        If lstList.Selected(x) Then
            'ajoute la taille
            lSize = lSize + Val(s)
        End If
    Next x
    lblSize.Caption = "Taille du fichier r�sultant=[" & Trim$(Str$(lSize)) & "]" & vbNewLine & FormatedSize(lSize)
End Sub

Private Sub lstList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu sur le listbox
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp
    
    RecalcSize  'recalcule la taille
End Sub

Private Sub mnuDeselectAll_Click()
'd�coche toutes les cases
Dim x As Long
    
    lstList.Visible = False
    
    For x = lstList.ListCount - 1 To 0 Step -1
        lstList.Selected(x) = False
    Next x
    
    lblSize.Caption = "Taille du fichier r�sultant=[0]"
    lstList.Visible = True
End Sub

Private Sub mnuSelectAll_Click()
'coche toutes les cases
Dim x As Long
    
    lstList.Visible = False
    
    For x = lstList.ListCount - 1 To 0 Step -1
        ValidateRect lstList.hWnd, 0&
        lstList.Selected(x) = True
    Next x
    
    lstList.Visible = True
    
    RecalcSize  'recalcule la taille
End Sub

'=======================================================
'obtient le processus concern� par l'enregistrement
'=======================================================
Public Sub GetProcess(ByVal lPID As Long, sFile As String)
Dim clsProc As clsMemoryRW
Dim LB() As Long
Dim x As Long

    txtPath.Text = sFile
    
    '//ajoute au listbox toutes les zones m�moire
    'liste les zones
    Set clsProc = New clsMemoryRW
    clsProc.RetrieveMemRegions lPID, LB(), LS()
    
    lstList.Clear
    lstList.Visible = False
    
    'les ajoute
    For x = 1 To UBound(LS())
        lstList.AddItem "Offset=[" & CStr(LB(x)) & "], taille=[" & CStr(LS(x)) & "]"
    Next x
    
    lstList.Visible = True
        
End Sub
