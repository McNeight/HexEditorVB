VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmDisAsm 
   BackColor       =   &H8000000C&
   Caption         =   "D�sassembleur d'ex�cutables"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   7425
   Icon            =   "frmDisAsm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5385
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14993
            MinWidth        =   14993
            Text            =   "Status = Ready"
            TextSave        =   "Status = Ready"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Menu rmnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuDisAsm 
         Caption         =   "&D�sassembler un fichier..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu rnuSave 
         Caption         =   "&Enregistrer"
         Begin VB.Menu mnuSaveASM 
            Caption         =   "&Liste des instructions ASM..."
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu rmnuDisplay 
      Caption         =   "&Affichage"
      Begin VB.Menu mnuShowASM 
         Caption         =   "&Code ASM"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuShowLog 
         Caption         =   "&Log"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuShowInfos 
         Caption         =   "&Informations sur le fichier"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuShowImp 
         Caption         =   "&Imports"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuShowExports 
         Caption         =   "&Exports"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuShowDat 
         Caption         =   "&Donn�es"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu rmnuWindow 
      Caption         =   "&Fen�tres"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&En cascade"
      End
      Begin VB.Menu mnuMH 
         Caption         =   "Mosa�que &horizontale"
      End
      Begin VB.Menu mnuMV 
         Caption         =   "Mosa�que &verticale"
      End
      Begin VB.Menu mnuReorganize 
         Caption         =   "&R�organiser les icones"
      End
   End
   Begin VB.Menu rmnuHelp 
      Caption         =   "&Aide"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Aide..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&A propos"
      End
   End
End
Attribute VB_Name = "frmDisAsm"
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
'PROJET DE DESASSEMBLAGE D'EXECUTABLES
'=======================================================

Private Sub mnuDisAsm_Click()
'ouvre un fichier
Dim s As String
Dim b As Boolean
Dim s2 As String

    'choix du fichier
    s = cFile.ShowOpen("Choix du fichier � d�sassembler", Me.hWnd, _
        "Fichiers d�sassemblables|*.exe;*.dll|Tous|*.*", App.Path, , b)
    If b Then Exit Sub
    If cFile.FileExists(s) = False Then Exit Sub
    
    'r�cup�re le path temporaire et cr�� un nom de dossier
    s2 = ObtainTempPath & "\" & cFile.GetFileFromPath(s)
    
    'on lance la proc�dure de d�sassemblage
    Sb.Panels(1).Text = "D�sassemblage en cours..."
    
    Call DisassembleWin32Executable(s, s2)
    
    Sb.Panels(1).Text = "Status = Ready"
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    EndProgram
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub
Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuMH_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuMV_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuQuit_Click()
    Call EndProgram
End Sub
Private Sub mnuReorganize_Click()
    Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuShowASM_Click()
    Me.mnuShowASM.Checked = Not (Me.mnuShowASM.Checked)
End Sub
Private Sub mnuShowDat_Click()
    Me.mnuShowDat.Checked = Not (Me.mnuShowDat.Checked)
End Sub
Private Sub mnuShowExports_Click()
    Me.mnuShowExports.Checked = Not (Me.mnuShowExports.Checked)
End Sub
Private Sub mnuShowImp_Click()
    Me.mnuShowImp.Checked = Not (Me.mnuShowImp.Checked)
End Sub
Private Sub mnuShowInfos_Click()
    Me.mnuShowInfos.Checked = Not (Me.mnuShowInfos.Checked)
End Sub
Private Sub mnuShowLog_Click()
    Me.mnuShowLog.Checked = Not (Me.mnuShowLog.Checked)
End Sub
