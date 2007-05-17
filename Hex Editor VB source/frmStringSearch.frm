VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5B5F5394-748F-414C-9FDD-08F3427C6A09}#3.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmStringSearch 
   BackColor       =   &H00F9E5D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche de chaînes de caractères"
   ClientHeight    =   6960
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
   HelpContextID   =   21
   Icon            =   "frmStringSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkBar PGB 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      Value           =   1
      BackPicture     =   "frmStringSearch.frx":058A
      FrontPicture    =   "frmStringSearch.frx":05A6
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
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      Caption         =   "Options de recherche"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin vkUserContolsXP.vkCheck chkAddSignet 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Ajouter un signet pour les chaines trouvées"
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
      Begin vkUserContolsXP.vkCheck chkAccent 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Rechercher des caractères accentués"
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
      Begin vkUserContolsXP.vkCheck chkSigns 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Rechercher des signes"
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
      Begin vkUserContolsXP.vkCheck chkMaj 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Rechercher des majuscules"
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
      Begin vkUserContolsXP.vkCheck chkMin 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Rechercher des minuscules"
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
      Begin vkUserContolsXP.vkCheck chkNumb3r 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         BackColor       =   16777215
         BackStyle       =   0
         Caption         =   "Rechercher des chiffres"
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
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Tag             =   "pref"
         Text            =   "5"
         ToolTipText     =   "Taille minimale (au dessous de cette taille, les suites de caractères ne sont pas considérées comme des strings)"
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Taille minimale de la chaîne de caractères :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Lancer la recherche"
      Height          =   375
      Left            =   4785
      TabIndex        =   3
      ToolTipText     =   "Lancer la recherche"
      Top             =   218
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sauvegarder les résultats"
      Height          =   495
      Left            =   4785
      TabIndex        =   2
      ToolTipText     =   "Sauvegarder les résultats (format texte)"
      Top             =   938
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   4785
      TabIndex        =   1
      ToolTipText     =   "Fermer cette fenêtre"
      Top             =   1778
      Width           =   1575
   End
   Begin ComctlLib.ListView LV 
      Height          =   3735
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "lang_ok"
      Top             =   3105
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Résultats de la recherche"
      Height          =   255
      Left            =   105
      TabIndex        =   4
      Top             =   2865
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

Private Lang As New clsLang
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
        MsgBox Lang.GetString("_PleaseNoNull"), vbInformation, Lang.GetString("_NoSPoss")
        Exit Sub
    End If

    If frmContent.ActiveForm Is Nothing Then Exit Sub

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_SearchCour"))
    
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
        Call SearchStringInFile(frmContent.ActiveForm.Caption, Val(txtSize.Text), _
            CBool(chkSigns.Value), CBool(chkMaj.Value), CBool(chkMin.Value), _
            CBool(chkNumb3r.Value), CBool(chkAccent.Value), tRes(), Me.PGB)
        
    ElseIf TypeOfActiveForm = "Mem" Then
        'alors c'est dans la mémoire
        
        'lance la recherche
        Call cMem.SearchEntireStringMemory(Val(frmContent.ActiveForm.Tag), _
            Val(txtSize.Text), CBool(chkSigns.Value), CBool(chkMaj.Value), _
            CBool(chkMin.Value), CBool(chkNumb3r.Value), CBool(chkAccent.Value), _
            lngRes(), strRes(), Me.PGB)
        
        'sauvegarde dans la variable tRes
        ReDim tRes(UBound(lngRes()))
        For i = 1 To UBound(lngRes())
            tRes(i).curOffset = CCur(lngRes(i))
            tRes(i).strString = strRes(i)
        Next i
        
    ElseIf TypeOfActiveForm = "Disk" Then
        'alors c'est dans le disque

        'lance la recherche
        Call SearchStringInFile(frmContent.ActiveForm.Caption, Val(txtSize.Text), _
            CBool(chkSigns.Value), CBool(chkMaj.Value), CBool(chkMin.Value), _
            CBool(chkNumb3r.Value), CBool(chkAccent.Value), tRes(), Me.PGB)
        
    Else
        'disque physique
        
        'lance la recherche
        Call SearchStringInFile(frmContent.ActiveForm.Caption, Val(txtSize.Text), _
            CBool(chkSigns.Value), CBool(chkMaj.Value), CBool(chkMin.Value), _
            CBool(chkNumb3r.Value), CBool(chkAccent.Value), tRes(), Me.PGB)
        
    End If
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ShowRes"))
    
    Call LV.ListItems.Clear
    
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
    
    Label2.Caption = Lang.GetString("_SearchResult") & " " & CStr(UBound(tRes()))
    
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
    Call AddTextToConsole(Lang.GetString("_SearchFin"))
    
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
        .DialogTitle = Lang.GetString("_SaveRes")
        .Filter = Lang.GetString("_All") & "|*.*"
        .Filename = vbNullString
        .ShowSave
        sFile = .Filename
    End With
    
    If cFile.FileExists(sFile) Then
        'fichier déjà existant
        If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + _
            vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If

    Label2.Caption = "Saving file..."
    
    lFile = FreeFile 'obtient un numéro disponible
    Open sFile For Output As lFile  'ouvre le fichier
    
    Print #lFile, Lang.GetString("_SearchOf") & txtSize.Text & _
        Lang.GetString("_ConsecChar") & vbNewLine & Lang.GetString("_FileIs") & _
        sFile & vbNewLine & Lang.GetString("_DateIs") & Date$ & "  " & Time$ & _
        vbNewLine & "[match]=" & LV.ListItems.Count
    
    For x = 1 To LV.ListItems.Count 'sauvegarde chaque élément du ListView
        Print #lFile, "[offset]=" & CStr(LV.ListItems.Item(x)) & "  [string]=" & _
        LV.ListItems.Item(x).SubItems(1)
        DoEvents
    Next x
    
    Close lFile
    
CancelPushed:
    Label2.Caption = Lang.GetString("_SearchResult")
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
    Call clsPref.GetFormSettings(App.Path & "\Preferences\StringSearch.ini", Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\StringSearch.ini", Me)
    Set clsPref = Nothing
End Sub

Private Sub LV_ItemClick(ByVal Item As ComctlLib.ListItem)
'se rend à l'offset

    If (frmContent.ActiveForm Is Nothing) Then Exit Sub
    
    With frmContent.ActiveForm
        .HW.FirstOffset = By16(Val(Item.Text))
        .VS.Value = By16(Val(Item.Text)) / 16 - 1
        Call .VS_Change(.VS.Value)
    End With

End Sub
