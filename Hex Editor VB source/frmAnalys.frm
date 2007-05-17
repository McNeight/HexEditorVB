VERSION 5.00
Object = "{EF4A8ABF-4214-4B3F-8F82-ACF6D11FA80D}#1.0#0"; "BGraphe_OCX.ocx"
Object = "{5B5F5394-748F-414C-9FDD-08F3427C6A09}#3.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form frmAnalys 
   BackColor       =   &H00F9E5D9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistiques"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   27
   Icon            =   "frmAnalys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin vkUserContolsXP.vkFrame vkFrame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      Caption         =   "Occurences"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdAnalyse 
         Caption         =   "Analyser"
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         ToolTipText     =   "Lance l'analyse"
         Top             =   5400
         Width           =   975
      End
      Begin BGraphe_OCX.BGraphe BG 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Occurences de distribution des bytes"
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   8705
         BarreColor1     =   0
         BarreColor2     =   16711680
      End
      Begin vkUserContolsXP.vkBar PGB 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   5400
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Value           =   1
         BackPicture     =   "frmAnalys.frx":058A
         FrontPicture    =   "frmAnalys.frx":05A6
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Byte [65] = [A] : 45845"
         Height          =   255
         Left            =   5880
         TabIndex        =   17
         Top             =   5520
         Width           =   3855
      End
   End
   Begin vkUserContolsXP.vkFrame vkFrame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2566
      BackColor2      =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtFile 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Fichier=[path]"
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Fichier=[path]"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Fichier=[path]"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Fichier=[path]"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Fichier=[path]"
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Fichier=[path]"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Fichier=[path]"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Fichier=[path]"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Fichier=[path]"
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdSaveStats 
      Caption         =   "Sauvegarder statistiques..."
      Height          =   495
      Left            =   1245
      TabIndex        =   0
      ToolTipText     =   "Sauvegarder les statistiques au format texte"
      Top             =   7680
      Width           =   2655
   End
   Begin VB.CommandButton cmdSaveBMP 
      Caption         =   "Sauvegarder BMP..."
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "Sauvegarder la bitmap des occurences"
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Quitter"
      Height          =   495
      Left            =   7365
      TabIndex        =   2
      ToolTipText     =   "Quitter cette fenêtre"
      Top             =   7680
      Width           =   1575
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
'FORM POUR L'ANALYSE DES FICHIERS
'=======================================================

Private sFile As String
Private Lang As New clsLang


'=======================================================
'obtient le fichier à analyser
'stocke ses propriétés dans les composants d'affichage
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
    
    'récupère les infos sur le fichier
    Set cF = cFile.GetFile(sFil)
    
    'récupère les infos sur les fichiers *.exe, *.dll...
    With cF
        sDescription = .FileVersionInfos.FileDescription
        sVersion = .FileVersionInfos.FileVersion
        sCopyright = .FileVersionInfos.Copyright
        
        sVersion = IIf(sVersion = vbNullString, "--", sVersion)
        sCopyright = IIf(sCopyright = vbNullString, "--", sCopyright)
        sDescription = IIf(sDescription = vbNullString, "--", sDescription)
        
        'affiche tout çà
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
    Shift As Integer, x As Single, y As Single)
    
    Label1.Caption = "Byte=[" & CStr(bByteX) & "] = [" & _
        Byte2FormatedString(bByteX) & "]  :   " & CStr(lOccurence)
        
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
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_AnalCou"))
    
    Call BG.ClearGraphe
    Call BG.ClearValues
    
    'prépare la progressbar
    lngLen = cFile.GetFileSize(sFile)
    With PGB
        .Min = 0: .Max = lngLen: .Value = 0
    End With
    
    'obtient le handle du fichier
    lngFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, 0&, _
        OPEN_EXISTING, 0&, 0&)
    
    'vérifie que le handle est valide
    If lngFile = INVALID_HANDLE_VALUE Then Exit Sub
    
    'créé un buffer de 50Ko
    strBuffer = String$(51200, 0) 'buffer de 50K
    
    curByte = 0
    Do Until curByte > lngLen  'tant que le fichier n'est pas fini
    
        x = x + 1
    
        'prépare le type OVERLAPPED - obtient 2 long à la place du Currency
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
        
        For y = 1 To l
            b = Asc(Mid$(strBuffer, y, 1))
            'ajoute une occurence
            F(b) = F(b) + 1
        Next y
        
        If (x Mod 10) = 0 Then
            'rend la main
            DoEvents
            PGB.Value = curByte
        End If
        
        curByte = curByte + 51200
    
    Loop

    Call CloseHandle(lngFile)
    
    'remplit le BG
    For x = 0 To 255
        Call BG.AddValue(x, F(x))
    Next x
        
    PGB.Value = PGB.Max
    Call BG.TraceGraph
    
    'ajoute du texte à la console
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
Dim x As Long

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
        x = MsgBox(Lang.GetString("_FileAlreadyEx"), vbInformation + vbYesNo, _
            Lang.GetString("_War"))
        If Not (x = vbYes) Then Exit Sub
    End If

    'sauvegarde
    Call BG.SaveBMP(s, cPref.general_ResoX, cPref.general_ResoY)

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ImgSaved"))
    
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
        x = MsgBox(Lang.GetString("_FileAlreadyEx"), vbInformation + vbYesNo, _
            Lang.GetString("_War"))
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'créé le fichier
    Call cFile.CreateEmptyFile(s)
    
    s2 = vbNullString
    'créé la string
    For x = 0 To 255
        s2 = s2 & "Byte=[" & Trim$(Str$(x)) & "] --> " & Lang.GetString("_Occ") _
            & Trim$(Str$(BG.GetValue(x))) & "]" & vbNewLine
    Next x
    
    'sauvegarde le fichier
    Call cFile.SaveDataInFile(s, Left$(s2, Len(s2) - 2), True)
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_StatSaved"))
Err:
End Sub

Private Sub Form_Load()

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
    
End Sub
