VERSION 5.00
Object = "{EF4A8ABF-4214-4B3F-8F82-ACF6D11FA80D}#1.0#0"; "BGraphe_OCX.ocx"
Object = "{BC0A7EAB-09F8-454A-AB7D-447C47D14F18}#1.0#0"; "ProgressBar_OCX.ocx"
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
         TabStop         =   0   'False
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
      TabIndex        =   2
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
      TabIndex        =   4
      ToolTipText     =   "Quitter cette fen�tre"
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Occurences"
      Height          =   5895
      Left            =   50
      TabIndex        =   5
      Top             =   1320
      Width           =   10095
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5535
         ScaleWidth      =   9855
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   9855
         Begin ProgressBar_OCX.pgrBar PGB 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Progression de l'analyse"
            Top             =   5160
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   661
            BackColorTop    =   13027014
            BackColorBottom =   15724527
            Value           =   1
            BackPicture     =   "frmAnalys.frx":058A
            FrontPicture    =   "frmAnalys.frx":05A6
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
            TabIndex        =   1
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
'FORM POUR L'ANALYSE DES FICHIERS
'=======================================================

Private sFile As String
Private Lang As New clsLang


'=======================================================
'obtient le fichier � analyser
'stocke ses propri�t�s dans les composants d'affichage
'=======================================================
Public Sub GetFile(ByVal sFil As String)
Dim sDescription As String
Dim sCopyright As String
Dim sVersion As String
Dim lPages As Long
Dim cF As FileSystemLibrary.File

    'active la gestion des langues
    Call Lang.ActiveLang(Me)
    
    sFile = sFil
    Me.Caption = sFil
    
    'affiche les infos sur le fichier dans les textboxes

    'nom du fichier
    txtFile.Text = "[" & Me.Caption & "]"
    
    'r�cup�re les infos sur le fichier
    Set cF = cFile.GetFile(sFil)
    
    'r�cup�re les infos sur les fichiers *.exe, *.dll...
    With cF
        sDescription = .FileVersionInfos.FileDescription
        sVersion = .FileVersionInfos.FileVersion
        sCopyright = .FileVersionInfos.Copyright
        
        sVersion = IIf(sVersion = vbNullString, "--", sVersion)
        sCopyright = IIf(sCopyright = vbNullString, "--", sCopyright)
        sDescription = IIf(sDescription = vbNullString, "--", sDescription)
        
        'affiche tout ��
        TextBox(0).Text = Lang.GetString("_Size") & CStr(.FileSize) & _
            " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        TextBox(1).Text = Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        TextBox(2).Text = Lang.GetString("_Creat") & .DateCreated & "]"
        TextBox(3).Text = Lang.GetString("_Acces") & .DateLastAccessed & "]"
        TextBox(4).Text = Lang.GetString("_Modif") & .DateLastModified & "]"
        TextBox(5).Text = Lang.GetString("_Version") & sVersion & "]"
        TextBox(6).Text = Lang.GetString("_Description") & sDescription & "]"
        TextBox(7).Text = Lang.GetString("_CopyR") & sCopyright & "]"
    End With
  
End Sub

Private Sub BG_MouseMove(bByteX As Byte, lOccurence As Long, Button As Integer, _
    Shift As Integer, X As Single, Y As Single)
    
    Label1.Caption = "Byte=[" & CStr(bByteX) & "] = [" & _
        Byte2FormatedString(bByteX) & "]  :   " & CStr(lOccurence)
        
End Sub

Public Sub cmdAnalyse_Click()
'lance l'analyse du fichier sFile
Dim lngLen As Long
Dim X As Long
Dim Y As Long
Dim b As Byte
Dim l As Long
Dim F(255) As Long
Dim tOver As OVERLAPPED
Dim strBuffer As String
Dim curByte As Currency
Dim lngFile As Long

    On Error GoTo ErrGestion
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_AnalCou"))
    
    Call BG.ClearGraphe
    Call BG.ClearValues
    
    'pr�pare la progressbar
    lngLen = cFile.GetFileSize(sFile)
    With PGB
        .Min = 0: .Max = lngLen: .Value = 0
    End With
    
    'obtient le handle du fichier
    lngFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, 0&, _
        OPEN_EXISTING, 0&, 0&)
    
    'v�rifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    'cr�� un buffer de 50Ko
    strBuffer = String$(51200, 0) 'buffer de 50K
    
    curByte = 0
    Do Until curByte > lngLen  'tant que le fichier n'est pas fini
    
        X = X + 1
    
        'pr�pare le type OVERLAPPED - obtient 2 long � la place du Currency
        Call GetLargeInteger(curByte, tOver.Offset, tOver.OffsetHigh)
        
        'obtient la string sur le buffer
        Call ReadFileEx(lngFile, ByVal strBuffer, 51200, tOver, _
            AddressOf CallBackFunction)
        
        If curByte + 51200 <= lngLen Then
            'alors on prend bien 51200 car
            l = 51200
        Else
            'on prend que les derniers car
            l = lngLen - curByte
        End If
        
        For Y = 1 To l
            b = Asc(Mid$(strBuffer, Y, 1))
            'ajoute une occurence
            F(b) = F(b) + 1
        Next Y
        
        If (X Mod 10) = 0 Then
            'rend la main
            DoEvents
            PGB.Value = curByte
        End If
        
        curByte = curByte + 51200
    
    Loop

    Call CloseHandle(lngFile)
    
    'remplit le BG
    For X = 0 To 255
        Call BG.AddValue(X, F(X))
    Next X
        
    PGB.Value = PGB.Max
    Call BG.TraceGraph
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_AnalTer"))
    
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
Dim X As Long

    On Error GoTo Err
    
    'affiche la boite de dialogue "sauvegarder"
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_BitmapImg")
        .Filter = "Bitmap Image|*.bmp|"
        .Filename = vbNullString
        .ShowSave
        s = .Filename
    End With
    
    'formate le nom (add terminaison)
    If LCase(Right$(s, 4)) <> ".bmp" Then s = s & ".bmp"
    
    If cFile.FileExists(s) Then
        'message de confirmation
        X = MsgBox(Lang.GetString("_FileAlreadyEx"), vbInformation + vbYesNo, _
            Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If

    'sauvegarde
    Call BG.SaveBMP(s, cPref.general_ResoX, cPref.general_ResoY)

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_ImgSaved"))
    
Err:
End Sub

Private Sub cmdSaveStats_Click()
'sauvegarde les stats dans un fichier *.log
Dim s As String
Dim X As Long
Dim s2 As String

    On Error GoTo Err
    
    'affiche la boite de dialogue "sauvegarder"
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SaveStat")
        .Filter = Lang.GetString("_LogFile") & " |*.log|"
        .Filename = vbNullString
        .ShowSave
        s = .Filename
    End With
    
    'formate le nom (add terminaison)
    If LCase(Right$(s, 4)) <> ".log" Then s = s & ".log"
    
    If cFile.FileExists(s) Then
        'message de confirmation
        X = MsgBox(Lang.GetString("_FileAlreadyEx"), vbInformation + vbYesNo, _
            Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If
    
    'cr�� le fichier
    Call cFile.CreateEmptyFile(s)
    
    s2 = vbNullString
    'cr�� la string
    For X = 0 To 255
        s2 = s2 & "Byte=[" & Trim$(Str$(X)) & "] --> " & Lang.GetString("_Occ") _
            & Trim$(Str$(BG.GetValue(X))) & "]" & vbNewLine
    Next X
    
    'sauvegarde le fichier
    Call cFile.SaveDataInFile(s, Left$(s2, Len(s2) - 2), True)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_StatSaved"))
Err:
End Sub

Private Sub Form_Load()

    With Lang
        #If MODE_DEBUG Then
            If App.LogMode = 0 And CREATE_FRENCH_FILE Then
                'on cr�� le fichier de langue fran�ais
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
        
        'applique la langue d�sir�e aux controles
        Call .ActiveLang(Me): .Language = cPref.env_Lang
        .LoadControlsCaption
    End With
    
End Sub
