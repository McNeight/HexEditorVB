VERSION 5.00
Object = "{BEF0F0EF-04C8-45BD-A6A9-68C01A66CB51}#2.0#0"; "vkUserControlsXP.ocx"
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
   HelpContextID   =   19
   Icon            =   "frmSaveProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkCheck chkAll 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Enregistre toute la mémoire (/!\ 2Go sont requis)"
      Top             =   3600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      BackColor       =   16777215
      BackStyle       =   0
      Caption         =   "Tout enregistrer (au minimum 2Go sont requis)"
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
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1695
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      Caption         =   "Contenu de l'enregistrement"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCheck chkOffset 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "La sauvegarde des offsets nécessite la sauvegarde de strings formatées"
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Sauvegarder les offsets"
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
      Begin vkUserContolsXP.vkCheck chkASCII 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Sauvegarder les valeurs ASCII réelles uniquement si coché seul"
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Sauvegarder les valeurs ASCII"
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
      Begin vkUserContolsXP.vkCheck chkHexa 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Sauvegarder les valeurs hexa"
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Sauvegarder les valeurs hexa"
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
   Begin vkUserContolsXP.vkListBox lstList 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   0   'False
      Sorted          =   0
      StyleCheckBox   =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zones mémoire à enregistrer"
      Height          =   255
      Left            =   128
      TabIndex        =   6
      Top             =   128
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Chemin du fichier"
      Height          =   255
      Left            =   4208
      TabIndex        =   5
      Top             =   2648
      Width           =   3255
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "Taille du fichier résultant=[0]"
      Height          =   615
      Left            =   4215
      TabIndex        =   4
      ToolTipText     =   "Taille estimée du fichier qui sera créé"
      Top             =   1920
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
Dim x As Long
    
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
        x = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + _
            vbYesNo, Lang.GetString("_War"))
        If Not (x = vbYes) Then Exit Sub
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
Dim x As Long
Dim y As Long
Dim s As String
    
    lSize = 0
    For x = 1 To lstList.ListCount
        s = Left$(lstList.List(x), Len(lstList.List(x)) - 1)  'garde l'item sans le ']' final
        y = InStrRev(s, "[", , vbBinaryCompare)
        s = Mid$(s, y + 1, Len(s) - y) 'contient la taille
        
        If lstList.Selected(x) Then
            'ajoute la taille
            lSize = lSize + Val(s)
        End If
    Next x
    
    lblSize.Caption = Lang.GetString("_SizeRes") & Trim$(Str$(lSize)) & "]" & _
        vbNewLine & FormatedSize(lSize)
End Sub

Private Sub lstList_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
'affiche le popup menu sur le listbox
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp
    
    Call RecalcSize  'recalcule la taille
End Sub

Private Sub mnuDeselectAll_Click()
'décoche toutes les cases
Dim x As Long
    
    Call lstList.UnCheckAll
    
    lblSize.Caption = Lang.GetString("_SizeRes") & "0]"
End Sub

Private Sub mnuSelectAll_Click()
'coche toutes les cases
Dim x As Long
    
    Call lstList.CheckAll
    
    Call RecalcSize  'recalcule la taille
End Sub

'=======================================================
'obtient le processus concerné par l'enregistrement
'=======================================================
Public Sub GetProcess(ByVal lPID As Long, sFile As String)
Dim clsProc As clsMemoryRW
Dim LB() As Long
Dim x As Long

    txtPath.Text = sFile
    
    '//ajoute au listbox toutes les zones mémoire
    'liste les zones
    Set clsProc = New clsMemoryRW
    Call clsProc.RetrieveMemRegions(lPID, LB(), LS())
    
    Call lstList.Clear
    
    'les ajoute
    With lstList
        .UnRefreshControl = True
        For x = 1 To UBound(LS())
            Call .AddItem("Offset=[" & CStr(LB(x)) & "], " & Lang.GetString("_Size") & "=[" & CStr(LS(x)) & "]")
        Next x
        .UnRefreshControl = False
        Call .Refresh
    End With
            
End Sub
