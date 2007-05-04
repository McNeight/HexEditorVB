VERSION 5.00
Begin VB.Form frmAdvancedConversion 
   Caption         =   "Conversion avancée"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   28
   Icon            =   "frmAdvancedConversion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   7695
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7695
      Begin VB.Frame Frame3 
         Caption         =   "Définition du séparateur"
         Enabled         =   0   'False
         Height          =   975
         Left            =   3840
         TabIndex        =   22
         Top             =   840
         Width           =   3735
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   3495
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   270
            Width           =   3495
            Begin VB.TextBox txtSepH 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   2280
               TabIndex        =   12
               Tag             =   "pref"
               Text            =   "00"
               ToolTipText     =   "Valeur hexa faisant office de séparateur"
               Top             =   360
               Width           =   615
            End
            Begin VB.TextBox txtSepS 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   2280
               TabIndex        =   10
               Tag             =   "pref"
               Text            =   "-"
               ToolTipText     =   "String faisant office de séparateur"
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton optSep 
               Caption         =   "Définir par valeur hexa"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   11
               Tag             =   "pref"
               ToolTipText     =   "Définir le séparateur par une valeur hexa (utile en cas de séparateur non écrivable au clavier)"
               Top             =   360
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton optSep 
               Caption         =   "Définir par string"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   9
               Tag             =   "pref"
               ToolTipText     =   "Définir le sépareteur par une string"
               Top             =   0
               Width           =   1695
            End
         End
      End
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Tag             =   "pref"
         Text            =   "2"
         ToolTipText     =   "Taille des paquets"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtBaseO 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Tag             =   "pref"
         ToolTipText     =   "Base personnelle"
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optUseSeparator 
         Caption         =   "Utiliser un séparateur"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Tag             =   "pref"
         ToolTipText     =   "Utiliser un séparateur entre les paquets"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton optUseFixedSize 
         Caption         =   "Utiliser une taille de paquet fixe"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "pref"
         ToolTipText     =   "Utiliser une taille de paquets fixes"
         Top             =   960
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Fermer"
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         ToolTipText     =   "Fermer cette fenêtre"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrad 
         Caption         =   "Traduire"
         Height          =   375
         Left            =   4200
         TabIndex        =   0
         ToolTipText     =   "Lancer la conversion"
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtBaseI 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Tag             =   "pref"
         ToolTipText     =   "Base personnelle"
         Top             =   60
         Width           =   735
      End
      Begin VB.ComboBox cbO 
         Height          =   315
         ItemData        =   "frmAdvancedConversion.frx":058A
         Left            =   1200
         List            =   "frmAdvancedConversion.frx":05A0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "pref lang_ok"
         ToolTipText     =   "Base d'arrivée"
         Top             =   420
         Width           =   1815
      End
      Begin VB.ComboBox cbI 
         Height          =   315
         ItemData        =   "frmAdvancedConversion.frx":05E0
         Left            =   1200
         List            =   "frmAdvancedConversion.frx":05F6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "pref lang_ok"
         ToolTipText     =   "Base de départ"
         Top             =   60
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "vers"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "convertir de "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sortie"
      Height          =   2535
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   7695
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   7455
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
         Begin VB.TextBox txtO 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   120
            Width           =   6015
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entrée"
      Height          =   2295
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   7455
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   7455
         Begin VB.TextBox txtI 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   0
            Width           =   6015
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "&Effacer"
      End
      Begin VB.Menu mnuTiretPopUp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Couper"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copier"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Coller"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Sélectionner tout"
      End
      Begin VB.Menu mnuTiretPopUp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Enregistrer..."
      End
      Begin VB.Menu mnuLoadFile 
         Caption         =   "&Charger depuis un fichier..."
      End
   End
   Begin VB.Menu mnuPopUp2 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete2 
         Caption         =   "&Effacer"
      End
      Begin VB.Menu mnuTiretPopUp12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut2 
         Caption         =   "&Couper"
      End
      Begin VB.Menu mnuCopy2 
         Caption         =   "&Copier"
      End
      Begin VB.Menu mnuPaste2 
         Caption         =   "&Coller"
      End
      Begin VB.Menu mnuSelectAll2 
         Caption         =   "&Sélectionner tout"
      End
      Begin VB.Menu mnuTiretPopUp22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave2 
         Caption         =   "&Enregistrer..."
      End
      Begin VB.Menu mnuLoadFile2 
         Caption         =   "&Charger depuis un fichier..."
      End
   End
End
Attribute VB_Name = "frmAdvancedConversion"
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
'FORM DE CONVERSION AVANCEE
'=======================================================

Private clsPref As clsIniForm
Private Lang As New clsLang
Private cConv As clsConvert

Private Sub cbI_Click()
    txtBaseI.Enabled = (cbI.Text = Lang.GetString("_Other!"))
End Sub

Private Sub cbO_Click()
    txtBaseO.Enabled = (cbO.Text = Lang.GetString("_Other!"))
End Sub

Private Sub cmdCLose_Click()
    Call Unload(Me)
End Sub

Private Sub cmdTrad_Click()
'lance la conversion
    Call LaunchExtraConversion
End Sub

Private Sub Form_Load()

    'loading des preferences
    Set clsPref = New clsIniForm
    Set cConv = New clsConvert
    
    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on créé le fichier de langue français
                .Language = "French"
                .LangFolder = LANG_PATH
                Call .WriteIniFileFormIDEform
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
        Call .LoadControlsCaption
    End With

    
    Call clsPref.GetFormSettings(App.Path & "\Preferences\AdvancedConversion.ini", Me)
    optSep(1).Value = Not (optSep(0).Value)
End Sub

Private Sub Form_Resize()
'resisze les composants de la form

    On Error Resume Next

    With Picture3
        .Width = Me.Width
        .Left = 0
        .Height = 1900
    End With
    With Picture1
        .Top = 240
        .Width = Me.Width - 480
        .Left = 80
    End With
    With Picture2
        .Top = 240
        .Left = 80
        .Width = Me.Width - 480
    End With
    With Frame1
        .Width = Me.Width - 320
        .Left = 120
        .Top = 120
        .Height = (Me.Height - 2400) / 2
    End With
    With Frame2
        .Width = Me.Width - 320
        .Left = 120
        .Top = 2090 + Frame1.Height
        .Height = Frame1.Height - 250
    End With
    Picture3.Top = 200 + Frame1.Height
    Picture1.Height = Frame1.Height - 300
    Picture2.Height = Frame2.Height - 300
    With txtI
        .Left = 70
        .Width = Picture1.Width - 160
        .Top = 30
        .Height = Picture1.Height - 80
    End With
    With txtO
        .Left = 70
        .Width = Picture2.Width - 160
        .Top = 30
        .Height = Picture2.Height - 80
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde les prefs
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\AdvancedConversion.ini", Me)
    
    Set clsPref = Nothing
    Set cConv = Nothing
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp
End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp2
End Sub

Private Sub optSep_Click(Index As Integer)
    txtSepS.Enabled = optSep(0).Value
    txtSepH.Enabled = Not (txtSepS.Enabled)
End Sub
Private Sub optUseFixedSize_Click()
    Frame3.Enabled = optUseSeparator.Value
    If Frame3.Enabled Then
        txtSepS.Enabled = optSep(0).Value
        txtSepH.Enabled = Not (txtSepS.Enabled)
        optSep(0).Enabled = True
        optSep(1).Enabled = True
    Else
        txtSepS.Enabled = False
        txtSepH.Enabled = False
        optSep(0).Enabled = False
        optSep(1).Enabled = False
    End If
End Sub
Private Sub optUseSeparator_Click()
    Frame3.Enabled = optUseSeparator.Value
    If Frame3.Enabled Then
        txtSepS.Enabled = optSep(0).Value
        txtSepH.Enabled = Not (txtSepS.Enabled)
        optSep(0).Enabled = True
        optSep(1).Enabled = True
    Else
        txtSepS.Enabled = False
        txtSepH.Enabled = False
        optSep(0).Enabled = False
        optSep(1).Enabled = False
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp2
End Sub

'=======================================================
'procède à la conversion
'=======================================================
Private Sub LaunchExtraConversion()
Dim sO As String
Dim s As String
Dim LS As Long
Dim lmax As Long
Dim sSep As String
Dim x As Long
Dim sA() As String

    If cbI.ListIndex < 0 Or cbO.ListIndex < 0 Then Exit Sub 'pas de base sélectionnée
    If (cbI.ListIndex = 5 And FormatedVal(txtBaseI.Text) = 0) Or _
        (cbO.ListIndex = 5 And FormatedVal(txtBaseO.Text) = 0) Then Exit Sub 'pas de base perso définie
    
    txtO.Text = vbNullString
    sO = vbNullString
    
    If optUseFixedSize.Value Then
        'alors on fait une conversion par taille de paquet fixe
        LS = FormatedVal(txtSize.Text)
        If LS = 0 Then
            'taille nulle
            MsgBox Lang.GetString("_SizeNoOk"), vbCritical, Lang.GetString("_War")
            Exit Sub
        End If
        
        Me.Caption = Lang.GetString("_ConvCour")
        
        lmax = Len(txtI.Text)
        For x = 1 To lmax Step LS
        
            If (x Mod 1000) = 0 Then DoEvents
            
            'on extrait le(s) caractère(s)
            s = Mid$(txtI.Text, x, LS)
            
            'on récupère la valeur formatée et on ajoute au buffer final
            sO = sO & GetCv(s)
        Next x
        
        'on affiche çà
        txtO.Text = sO
        
        Me.Caption = Lang.GetString("_AdConv")
        
    Else
        'alors on fait une conversion par séparateur
        If (optSep(0).Value And Len(txtSepS.Text) = 0) Or (optSep(1).Value And _
            Len(txtSepH) = 0) Then
            
            'impossible car pas de spérateur
            MsgBox Lang.GetString("_NoGoodSep"), vbCritical, Lang.GetString("_War")
            
            Exit Sub
        End If
        
        Me.Caption = Lang.GetString("_ConvCour")
        
        'définit le caractère séparant
        If optSep(0).Value Then sSep = txtSepS.Text Else sSep = _
            Str2Hex(txtSepH.Text)
        
        'récupère toutes les valeurs séparément
        sA() = Split(txtI.Text, sSep, , vbBinaryCompare)
        
        For x = 0 To UBound(sA())
            If (x Mod 1000) = 0 Then DoEvents
            sO = sO & GetCv(sA(x)) & sSep
        Next x
        
        'on affiche en virant le dernier séparateur
        txtO.Text = Left$(sO, Len(sO) - Len(sSep))

        Me.Caption = Lang.GetString("_AdConv")
        
    End If
    
End Sub

'=======================================================
'renvoie une valeur formatée en fonction des choix de base
'=======================================================
Private Function GetCv(ByVal sIn As String) As String
Dim s2 As String

    cConv.CurrentString = sIn
    
    With Lang
        Select Case cbI.Text
            Case .GetString("_Decimal!")
                cConv.CurrentBase = 10
            Case .GetString("_Octal!")
                cConv.CurrentBase = 8
            Case .GetString("_Hexa!")
                cConv.CurrentBase = 16
            Case .GetString("_Binary!")
                cConv.CurrentBase = 2
            Case .GetString("_Other!")
                cConv.CurrentBase = Val(txtBaseI.Text)
            Case Else
                'ANSI ASCII
        End Select
        
        Select Case cbO.Text
            Case .GetString("_Decimal!")
                s2 = cConv.Convert(10)
            Case .GetString("_Octal!")
                s2 = cConv.Convert(8)
            Case .GetString("_Hexa!")
                s2 = cConv.Convert(16)
            Case .GetString("_Binary!")
                s2 = cConv.Convert(2)
            Case .GetString("_Other!")
                s2 = cConv.Convert(Val(txtBaseI.Text))
            Case Else
                'ANSI ASCII
        End Select
    End With
    
    'Byte2FormatedString
    
    GetCv = s2
End Function
