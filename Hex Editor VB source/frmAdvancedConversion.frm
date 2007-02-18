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
   Icon            =   "frmAdvancedConversion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   7695
      TabIndex        =   6
      Top             =   2520
      Width           =   7695
      Begin VB.Frame Frame3 
         Caption         =   "Définition du séparateur"
         Enabled         =   0   'False
         Height          =   975
         Left            =   3840
         TabIndex        =   18
         Top             =   840
         Width           =   3735
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   3495
            TabIndex        =   19
            Top             =   270
            Width           =   3495
            Begin VB.TextBox txtSepH 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   2280
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   20
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
         TabIndex        =   17
         Tag             =   "pref"
         Text            =   "2"
         ToolTipText     =   "Taille des paquets"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtBaseO 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Tag             =   "pref"
         ToolTipText     =   "Base personnelle"
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optUseSeparator 
         Caption         =   "Utiliser un séparateur"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Tag             =   "pref"
         ToolTipText     =   "Utiliser un séparateur entre les paquets"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton optUseFixedSize 
         Caption         =   "Utiliser une taille de paquet fixe"
         Height          =   255
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   13
         ToolTipText     =   "Fermer cette fenêtre"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdTrad 
         Caption         =   "Traduire"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         ToolTipText     =   "Lancer la conversion"
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txtBaseI 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Tag             =   "pref"
         ToolTipText     =   "Base personnelle"
         Top             =   60
         Width           =   735
      End
      Begin VB.ComboBox cbO 
         Height          =   315
         ItemData        =   "frmAdvancedConversion.frx":08CA
         Left            =   1200
         List            =   "frmAdvancedConversion.frx":08E0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "pref"
         ToolTipText     =   "Base d'arrivée"
         Top             =   420
         Width           =   1815
      End
      Begin VB.ComboBox cbI 
         Height          =   315
         ItemData        =   "frmAdvancedConversion.frx":0920
         Left            =   1200
         List            =   "frmAdvancedConversion.frx":0936
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "pref"
         ToolTipText     =   "Base de départ"
         Top             =   60
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "vers"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "convertir de "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sortie"
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   7695
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   7455
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   120
            Width           =   6015
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entrée"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   7455
         TabIndex        =   1
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
            TabIndex        =   2
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
' -----------------------------------------------
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' -----------------------------------------------
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
' -----------------------------------------------


Option Explicit

'-------------------------------------------------------
'FORM DE CONVERSION AVANCEE
'-------------------------------------------------------

Private clsPref As clsIniForm

Private Sub cmdCLose_Click()
    Unload Me
End Sub

Private Sub cmdTrad_Click()
'lance la conversion
    Call LaunchExtraConversion
End Sub

Private Sub Form_Load()
    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\AdvancedConversion.ini", Me
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
    clsPref.SaveFormSettings App.Path & "\Preferences\AdvancedConversion.ini", Me
    Set clsPref = Nothing
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopup
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
    If Button = 2 Then Me.PopupMenu Me.mnuPopup
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche le popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopUp2
End Sub

'-------------------------------------------------------
'procède à la conversion
'-------------------------------------------------------
Private Sub LaunchExtraConversion()
Dim sO As String
Dim s As String
Dim LS As Long
Dim lMax As Long
Dim sSep As String
Dim x As Long
Dim sA() As String

    If cbI.ListIndex < 0 Or cbO.ListIndex < 0 Then Exit Sub 'pas de base sélectionnée
    If (cbI.ListIndex = 5 And FormatedVal(txtBaseI.Text) = 0) Or (cbO.ListIndex = 5 And FormatedVal(txtBaseO.Text) = 0) Then Exit Sub 'pas de base perso définie
    
    txtO.Text = vbNullString
    sO = vbNullString
    
    If optUseFixedSize.Value Then
        'alors on fait une conversion par taille de paquet fixe
        LS = FormatedVal(txtSize.Text)
        If LS = 0 Then
            'taille nulle
            MsgBox "La taille des paquets n'est pas convenable.", vbCritical, "Attention"
            Exit Sub
        End If
        
        Me.Caption = "Conversion..."
        
        lMax = Len(txtI.Text)
        For x = 1 To lMax Step LS
        
            If (x Mod 1000) = 0 Then DoEvents
            
            'on extrait le(s) caractère(s)
            s = Mid$(txtI.Text, x, LS)
            
            'on récupère la valeur formatée et on ajoute au buffer final
            sO = sO & GetCv(s)
        Next x
        
        'on affiche çà
        txtO.Text = sO
        
        Me.Caption = "Conversion avancée"
        
    Else
        'alors on fait une conversion par séparateur
        If (optSep(0).Value And Len(txtSepS.Text) = 0) Or (optSep(1).Value And Len(txtSepH) = 0) Then
            'impossible car pas de spérateur
            MsgBox "Le séparateur n'est pas convenable.", vbCritical, "Attention"
            Exit Sub
        End If
        
        Me.Caption = "Conversion..."
        
        'définit le caractère séparant
        If optSep(0).Value Then sSep = txtSepS.Text Else sSep = Str2Hex(txtSepH.Text)
        
        'récupère toutes les valeurs séparément
        sA() = Split(txtI.Text, sSep, , vbBinaryCompare)
        
        For x = 0 To UBound(sA())
        
            If (x Mod 1000) = 0 Then DoEvents
            sO = sO & GetCv(sA(x)) & sSep
        Next x
        
        'on affiche en virant le dernier séparateur
        txtO.Text = Left$(sO, Len(sO) - Len(sSep))

        Me.Caption = "Conversion avancée"
        
    End If
    
End Sub

'-------------------------------------------------------
'renvoie une aleur formatée en fonction des choix de base
'-------------------------------------------------------
Private Function GetCv(ByVal sIn As String) As String
Dim s2 As String

    'On Error Resume Next    'évite les dépassement de capacité si l'user rentre n'importe quoi

    Select Case cbI.Text
        Case "Décimale"
            Select Case cbO.Text
                Case "Décimale"
                    s2 = sIn
                Case "Octale"
                    s2 = Oct$(FormatedVal(sIn))
                Case "Héxadécimale"
                    s2 = Hex$(FormatedVal(sIn))
                Case "Binaire"
                    s2 = Dec2Bin(FormatedVal(sIn))
                Case "ANSI ASCII"
                    s2 = Byte2FormatedString(FormatedVal(sIn))
                Case "Autre"
                    
            End Select
        Case "Octale"
            Select Case cbO.Text
                Case "Décimale"
                    s2 = Oct2Dec(sIn)
                Case "Octale"
                    s2 = sIn
                Case "Héxadécimale"
                    s2 = Hex$(Oct2Dec(sIn))
                Case "Binaire"
                    s2 = Dec2Bin(Oct2Dec(sIn))
                Case "ANSI ASCII"
                    s2 = Byte2FormatedString(Oct2Dec(sIn))
                Case "Autre"
                
            End Select
        Case "Héxadécimale"
            Select Case cbO.Text
                Case "Décimale"
                    s2 = Hex2Dec(sIn)
                Case "Octale"
                    s2 = Hex2Oct(sIn)
                Case "Héxadécimale"
                    s2 = sIn
                Case "Binaire"
                    s2 = Dec2Bin(Hex2Dec(sIn))
                Case "ANSI ASCII"
                    s2 = Hex2Str(sIn)
                Case "Autre"
                
            End Select
        Case "Binaire"
            Select Case cbO.Text
                Case "Décimale"
                    s2 = Bin2Dec(sIn)
                Case "Octale"
                    s2 = Oct$(Bin2Dec(sIn))
                Case "Héxadécimale"
                    s2 = Hex$(Bin2Dec(sIn))
                Case "Binaire"
                    s2 = sIn
                Case "ANSI ASCII"
                    s2 = Byte2FormatedString(Bin2Dec(sIn))
                Case "Autre"
                
            End Select
        Case "ANSI ASCII"
            Select Case cbO.Text
                Case "Décimale"
                    s2 = Str2Dec(sIn)
                Case "Octale"
                    s2 = Str2Oct(sIn)
                Case "Héxadécimale"
                    s2 = Str2Hex(sIn)
                Case "Binaire"
                    s2 = Dec2Bin(Str2Dec(sIn))
                Case "ANSI ASCII"
                    s2 = sIn
                Case "Autre"
                
            End Select
        Case "Autre"
            Select Case cbO.Text
                Case "Décimale"
                    
                Case "Octale"
                    
                Case "Héxadécimale"
                    
                Case "Binaire"
                    
                Case "ANSI ASCII"
                    
                Case "Autre"
                    
            End Select
    End Select
    
    GetCv = s2
End Function
