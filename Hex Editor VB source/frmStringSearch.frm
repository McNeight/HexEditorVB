VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{6ADE9E73-F694-428F-BF86-06ADD29476A5}#1.0#0"; "ProgressBar_OCX.ocx"
Begin VB.Form frmStringSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche de chaînes de caractères"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStringSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin ProgressBar_OCX.pgrBar PGB 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Progression de la recherche"
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      BackColorTop    =   13027014
      BackColorBottom =   15724527
      Value           =   1
      BackPicture     =   "frmStringSearch.frx":058A
      FrontPicture    =   "frmStringSearch.frx":05A6
   End
   Begin ComctlLib.ListView LV 
      Height          =   3735
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Offset"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Chaîne trouvée"
         Object.Width           =   7673
      EndProperty
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder les résultats"
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "Sauvegarder les résultats (format texte)"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Lancer la recherche"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      ToolTipText     =   "Lancer la recherche"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options de recherche"
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   4095
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   4095
         Begin VB.CheckBox chkAccent 
            Caption         =   "Rechercher des caractères accentués"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Tag             =   "pref"
            ToolTipText     =   "Rechercher des caractères avec des accents ("
            Top             =   960
            Width           =   3735
         End
         Begin VB.CheckBox chkAddSignet 
            Caption         =   "Ajouter un signet pour les chaines trouvées"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Tag             =   "pref"
            ToolTipText     =   "Ajouter un signet à chaque offset où une string est trouvée"
            Top             =   1200
            Width           =   3975
         End
         Begin VB.CheckBox chkSigns 
            Caption         =   "Rechercher des signes"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Tag             =   "pref"
            ToolTipText     =   "Inclure les signes dans la recherche"
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox chkMaj 
            Caption         =   "Rechercher des majuscules"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Tag             =   "pref"
            ToolTipText     =   "Inclure les majuscules dans la recherche"
            Top             =   480
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkMin 
            Caption         =   "Rechercher des minuscules"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Tag             =   "pref"
            ToolTipText     =   "Inclure les minuscules dans la recherche"
            Top             =   240
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3240
            TabIndex        =   8
            Tag             =   "pref"
            Text            =   "5"
            ToolTipText     =   "Taille minimale (au dessous de cette taille, les suites de caractères ne sont pas considérées comme des strings)"
            Top             =   1460
            Width           =   735
         End
         Begin VB.CheckBox chkNumb3r 
            Caption         =   "Rechercher des chiffres"
            Height          =   255
            Left            =   0
            TabIndex        =   1
            Tag             =   "pref"
            ToolTipText     =   "Inclure les chiffres dans la recherche"
            Top             =   0
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Taille minimale de la chaîne de caractères :"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   1460
            Width           =   3135
         End
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Résultats de la recherche"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   6375
   End
End
Attribute VB_Name = "frmStringSearch"
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
'FORM DE RECHERCHE DE STRINGS DANS LE FICHIER/MEMOIRE
'=======================================================

Private clsPref As clsIniForm

Private Sub cmdGo_Click()
'lance la recherche
Dim tRes() As SearchResult
Dim lngRes() As Long
Dim strRes() As String
Dim i As Long
Dim bAddSign As Boolean

    'On Error GoTo ErrGestion

    txtSize.Text = FormatedVal(txtSize.Text)
    If Val(txtSize.Text) = 0 Then
        'pas de longueur ==> pas de recherche ;)
        MsgBox "Veuillez spécifier une longueur de chaine non nulle", vbInformation, "Pas de recherche possible"
        Exit Sub
    End If

    If frmContent.ActiveForm Is Nothing Then Exit Sub

    'ajoute du texte à la console
    Call AddTextToConsole("Recherche de strings en cours...")
    
    cmdSave.Enabled = False
    chkNumb3r.Enabled = False
    chkMin.Enabled = False
    chkMaj.Enabled = False
    chkSigns.Enabled = False
    Label1.Enabled = False
    txtSize.Enabled = False
    cmdQuit.Enabled = False
    chkAccent.Enabled = False
    chkAddSignet.Enabled = False
    
    If TypeOfActiveForm = "Pfm" Then
        'alors c'est un fichier classique
        
        'lance la recherche
        SearchStringInFile frmContent.ActiveForm.Caption, Val(txtSize.Text), CBool(chkSigns.Value), CBool(chkMaj.Value), CBool(chkMin.Value), CBool(chkNumb3r.Value), CBool(chkAccent.Value), tRes(), Me.pgb
        
    ElseIf TypeOfActiveForm = "Mem" Then
        'alors c'est dans la mémoire
        
        'lance la recherche
        cMem.SearchEntireStringMemory Val(frmContent.ActiveForm.Tag), Val(txtSize.Text), CBool(chkSigns.Value), CBool(chkMaj.Value), CBool(chkMin.Value), CBool(chkNumb3r.Value), CBool(chkAccent.Value), lngRes(), strRes(), Me.pgb
        
        'sauvegarde dans la variable tRes
        ReDim tRes(UBound(lngRes()))
        For i = 1 To UBound(lngRes())
            tRes(i).curOffset = CCur(lngRes(i))
            tRes(i).strString = strRes(i)
        Next i
        
    Else
        'alors c'est dans le disque

        'lance la recherche
        SearchStringInFile frmContent.ActiveForm.Caption, Val(txtSize.Text), CBool(chkSigns.Value), CBool(chkMaj.Value), CBool(chkMin.Value), CBool(chkNumb3r.Value), CBool(chkAccent.Value), tRes(), Me.pgb
        
    End If
    
    'ajoute du texte à la console
    Call AddTextToConsole("Affichage des résultats...")
    
    LV.ListItems.Clear
    
    bAddSign = chkAddSignet.Value
    
    'affiche les résultats
    For i = 1 To UBound(tRes)
        With LV
            If bAddSign Then
                'ajoute un signet

                frmContent.ActiveForm.HW.AddSignet By16(tRes(i).curOffset)
                frmContent.ActiveForm.HW.TraceSignets
                
                'ajoute le signet à la listview
                frmContent.ActiveForm.lstSignets.ListItems.Add Text:=Trim$(Str$((By16(tRes(i).curOffset))))
                frmContent.ActiveForm.lstSignets.ListItems.Item(frmContent.ActiveForm.lstSignets.ListItems.Count).SubItems(1) = tRes(i).strString
            End If
            .ListItems.Add Text:=CStr(tRes(i).curOffset)
            .ListItems.Item(i).SubItems(1) = tRes(i).strString
        End With
    Next i
    
    Label2.Caption = "Résultats de la recherche " & CStr(UBound(tRes()))
    
    cmdSave.Enabled = True
    chkNumb3r.Enabled = True
    chkMin.Enabled = True
    chkMaj.Enabled = True
    chkSigns.Enabled = True
    Label1.Enabled = True
    txtSize.Enabled = True
    cmdQuit.Enabled = True
    chkAddSignet.Enabled = True
    chkAccent.Enabled = True

    'ajoute du texte à la console
    Call AddTextToConsole("Recherche terminée")
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmStringSearch.cmdGOClick", True
End Sub

Private Sub cmdQuit_Click()
'ferme
    Unload Me
End Sub

Private Sub cmdSave_Click()
'sauvegarde les résultats
Dim lFile As Long
Dim sFile As String
Dim x As Long

    On Error GoTo CancelPushed
    
    'affiche la boite de dialogue Sauvegarder
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "Enregistrer les résultats"
        .Filter = "Tous|*.*"
        .ShowSave
        sFile = .Filename
    End With
    
    If cFile.FileExists(sFile) Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà. Le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If

    Label2.Caption = "Saving file..."
    
    lFile = FreeFile 'obtient un numéro disponible
    Open sFile For Output As lFile  'ouvre le fichier
    
    Print #lFile, "Recherche de [" & txtSize.Text & "] caractères consécutifs" & vbNewLine & "[fichier]=" & sFile & vbNewLine & "[date]=" & Date$ & "  " & Time$ & vbNewLine & "[match]=" & LV.ListItems.Count
    
    For x = 1 To LV.ListItems.Count 'sauvegarde chaque élément du ListView
        Print #lFile, "[offset]=" & CStr(LV.ListItems.Item(x)) & "  [string]=" & LV.ListItems.Item(x).SubItems(1)
        DoEvents
    Next x
    
    Close lFile
    
CancelPushed:
    Label2.Caption = "Résultats de la recherche"
End Sub

Private Sub Form_Load()
    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\StringSearch.ini", Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\StringSearch.ini", Me
    Set clsPref = Nothing
End Sub

Private Sub LV_ItemClick(ByVal Item As ComctlLib.ListItem)
'se rend à l'offset

    If (frmContent.ActiveForm Is Nothing) Then Exit Sub
    
    frmContent.ActiveForm.HW.FirstOffset = By16(Val(Item.Text))
    frmContent.ActiveForm.VS.Value = By16(Val(Item.Text)) / 16 - 1
    Call frmContent.ActiveForm.VS_Change(frmContent.ActiveForm.VS.Value)

End Sub
