VERSION 5.00
Object = "{2ED9CD5C-C64E-4F0C-B719-F9D0F542DD03}#1.0#0"; "BGraphe_OCX.ocx"
Object = "{6ADE9E73-F694-428F-BF86-06ADD29476A5}#1.0#0"; "ProgressBar_OCX.ocx"
Begin VB.Form frmAnalys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistiques"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnalys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Fichier"
      Height          =   1335
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   9855
         TabIndex        =   7
         Top             =   240
         Width           =   9855
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Fichier=[path]"
            Top             =   0
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "Fichier=[path]"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "Fichier=[path]"
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "Fichier=[path]"
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "Fichier=[path]"
            Top             =   0
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "Fichier=[path]"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   1
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "Fichier=[path]"
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "Fichier=[path]"
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtFile 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Fichier=[path]"
            Top             =   0
            Width           =   2895
         End
      End
   End
   Begin VB.CommandButton cmdSaveStats 
      Caption         =   "Sauvegarder statistiques..."
      Height          =   495
      Left            =   1245
      TabIndex        =   4
      ToolTipText     =   "Sauvegarder les statistiques au format texte"
      Top             =   7320
      Width           =   2655
   End
   Begin VB.CommandButton cmdSaveBMP 
      Caption         =   "Sauvegarder BMP..."
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      ToolTipText     =   "Sauvegarder la bitmap des occurences"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   7365
      TabIndex        =   2
      ToolTipText     =   "Quitter cette fenêtre"
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Occurences"
      Height          =   5895
      Left            =   50
      TabIndex        =   1
      Top             =   1320
      Width           =   10095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5535
         ScaleWidth      =   9855
         TabIndex        =   5
         Top             =   240
         Width           =   9855
         Begin ProgressBar_OCX.pgrBar PGB 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Progression de l'analyse"
            Top             =   5160
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   661
            BackColorTop    =   13027014
            BackColorBottom =   15724527
            Value           =   1
            BackPicture     =   "frmAnalys.frx":08CA
            FrontPicture    =   "frmAnalys.frx":08E6
         End
         Begin BGraphe_OCX.BGraphe BG 
            Height          =   5055
            Left            =   0
            TabIndex        =   18
            ToolTipText     =   "Occurences de distribution des bytes"
            Top             =   0
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   8916
            BarreColor1     =   0
            BarreColor2     =   16711680
         End
         Begin VB.CommandButton cmdAnalyse 
            Caption         =   "Analyser"
            Height          =   375
            Left            =   4560
            TabIndex        =   6
            ToolTipText     =   "Lance l'analyse"
            Top             =   5160
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Byte [65] = [A] : 45845"
            Height          =   255
            Left            =   5880
            TabIndex        =   17
            Top             =   5280
            Width           =   3855
         End
      End
   End
End
Attribute VB_Name = "frmAnalys"
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
'FORM POUR L'ANALYSE DES FICHIERS
'-------------------------------------------------------

Private sFile As String


'-------------------------------------------------------
'obtient le fichier à analyser
'stocke ses propriétés dans les composants d'affichage
'-------------------------------------------------------
Public Sub GetFile(ByVal sFil As String)
Dim sDescription As String
Dim sCopyright As String
Dim sVersion As String
Dim lSize As Long
Dim lAttribute As Long
Dim dAccess As String
Dim dCreation As String
Dim dModification As String
Dim lPages As Long
Dim cF As clsFile

    sFile = sFil
    Me.Caption = sFil
    
    'affiche les infos sur le fichier dans les textboxes

    'nom du fichier
    txtFile.Text = "[" & Me.Caption & "]"
    
    'récupère les infos sur le fichier
    Set cF = cFile.GetFile(sFil)
    
    'récupère les infos sur les fichiers *.exe, *.dll...
    sDescription = cF.EXEFileDescription
    sVersion = cF.EXEFileVersion
    sCopyright = cF.EXELegalCopyright
    
    sVersion = IIf(sVersion = vbNullString, "--", sVersion)
    sCopyright = IIf(sCopyright = vbNullString, "--", sCopyright)
    sDescription = IIf(sDescription = vbNullString, "--", sDescription)
    
    'récupère les dates
    dCreation = cF.CreationDate
    dAccess = cF.LastAccessDate
    dModification = cF.LastModificationDate
    
    'la taille
    lSize = cF.FileSize
    
    'attribut
    lAttribute = cF.FileAttributes
    
    'affiche tout çà
    TextBox(0).Text = "Taille=[" & CStr(lSize) & " Octets  -  " & CStr(Round(lSize / 1024, 3)) & " Ko" & "]"
    TextBox(1).Text = "Attribut=[" & CStr(lAttribute) & "]"
    TextBox(2).Text = "Création=[" & dCreation & "]"
    TextBox(3).Text = "Accès=[" & dAccess & "]"
    TextBox(4).Text = "Modification=[" & dModification & "]"
    TextBox(5).Text = "Version=[" & sVersion & "]"
    TextBox(6).Text = "Description=[" & sDescription & "]"
    TextBox(7).Text = "Copyright=[" & sCopyright & "]"
  
End Sub

Private Sub BG_MouseMove(bByteX As Byte, lOccurence As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.Caption = "Byte=[" & CStr(bByteX) & "] = [" & Byte2FormatedString(bByteX) & "]  :   " & CStr(lOccurence)
End Sub

Public Sub cmdAnalyse_Click()
'lance l'analyse du fichier sFile
Dim lngLen As Long
Dim x As Long
Dim y As Long
Dim b As Byte
Dim l As Long
Dim F(255) As Long
Dim tOver As OVERLAPPED
Dim strBuffer As String
Dim curByte As Currency
Dim lngFile As Long

    On Error GoTo ErrGestion
    
    BG.ClearGraphe
    BG.ClearValues
    
    'prépare la progressbar
    lngLen = cFile.GetFileSize(sFile)
    pgb.Min = 0: pgb.Max = lngLen: pgb.Value = 0
    
    'obtient le handle du fichier
    lngFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, 0&, 0&)
    
    'vérifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    'créé un buffer de 50Ko
    strBuffer = String$(51200, 0) 'buffer de 50K
    
    curByte = 0
    Do Until curByte > lngLen  'tant que le fichier n'est pas fini
    
        x = x + 1
    
        'prépare le type OVERLAPPED - obtient 2 long à la place du Currency
        GetLargeInteger curByte, tOver.Offset, tOver.OffsetHigh
        
        'obtient la string sur le buffer
        ReadFileEx lngFile, ByVal strBuffer, 51200, tOver, AddressOf CallBackFunction
        
        If curByte + 51200 <= lngLen Then
            'alors on prend bien 51200 car
            l = 51200
        Else
            'on prend que les derniers car
            l = lngLen - curByte
        End If
        
        For y = 1 To l
            b = Asc(Mid$(strBuffer, y, 1))
            'ajoute une occurence
            F(b) = F(b) + 1
        Next y
        
        If (x Mod 10) = 0 Then
            'rend la main
            DoEvents
            pgb.Value = curByte
        End If
        
        curByte = curByte + 51200
    
    Loop

    CloseHandle lngFile
    
    'remplit le BG
    For x = 0 To 255
        BG.AddValue x, F(x)
    Next x
        
    pgb.Value = pgb.Max
    BG.TraceGraph
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmAnalysis.cmdAnalyseClick", True
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub cmdSaveBMP_Click()
'sauvegarder en bmp
Dim s As String
Dim x As Long

    On Error GoTo Err
    
    'affiche la boite de dialogue "sauvegarder"
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "Sauvegarder une image bitmap"
        .Filter = "Bitmap Image|*.bmp|"
        .ShowSave
        s = .Filename
    End With
    
    'formate le nom (add terminaison)
    If LCase(Right$(s, 4)) <> ".bmp" Then s = s & ".bmp"
    
    If cFile.FileExists(s) Then
        'message de confirmation
        x = MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention")
        If Not (x = vbYes) Then Exit Sub
    End If

    'sauvegarde
    Call BG.SaveBMP(s, cPref.general_ResoX, cPref.general_ResoY)
    
Err:
End Sub

Private Sub cmdSaveStats_Click()
'sauvegarde les stats dans un fichier *.log
Dim s As String
Dim x As Long
Dim s2 As String

    On Error GoTo Err
    
    'affiche la boite de dialogue "sauvegarder"
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "Sauvegarder les statistiques"
        .Filter = "Fichier log|*.log|"
        .ShowSave
        s = .Filename
    End With
    
    'formate le nom (add terminaison)
    If LCase(Right$(s, 4)) <> ".log" Then s = s & ".log"
    
    If cFile.FileExists(s) Then
        'message de confirmation
        x = MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention")
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'créé le fichier
    cFile.CreateEmptyFile s
    
    s2 = vbNullString
    'créé la string
    For x = 0 To 255
        s2 = s2 & "Byte=[" & Trim$(Str$(x)) & "] --> occurence=[" & Trim$(Str$(BG.GetValue(x))) & "]" & vbNewLine
    Next x
    
    'sauvegarde le fichier
    cFile.SaveDATAinFile s, Left$(s2, Len(s2) - 2), True
Err:
End Sub
