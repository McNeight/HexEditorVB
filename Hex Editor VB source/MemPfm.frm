VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C60799F1-7AA3-45BA-AFBF-5BEAB08BC66C}#1.0#0"; "HexViewer_OCX.ocx"
Object = "{5B5F5394-748F-414C-9FDD-08F3427C6A09}#3.0#0"; "vkUserControlsXP.ocx"
Begin VB.Form MemPfm 
   BackColor       =   &H00F9E5D9&
   Caption         =   "Ouverture d'un processus..."
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   7
   Icon            =   "MemPfm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   8730
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Fichier=[Modifié]"
            TextSave        =   "Fichier=[Modifié]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Page=[0/0]"
            TextSave        =   "Page=[0/0]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Offset=[0]"
            TextSave        =   "Offset=[0]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Sélection=[0 Bytes]"
            TextSave        =   "Sélection=[0 Bytes]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin vkUserContolsXP.vkFrame FrameInfos 
      Height          =   6975
      Left            =   4680
      TabIndex        =   15
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   12303
      Caption         =   "Informations"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView lstHisto 
         Height          =   1575
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "lang_ok"
         Top             =   4920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2778
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
            Text            =   "Action"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Rang"
            Object.Width           =   706
         EndProperty
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "MemPfm.frx":08CA
         Top             =   840
         Width           =   2895
      End
      Begin ComctlLib.TabStrip TB2 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "lang_ok"
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Fichier cible"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Processus"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdMAJ 
         Caption         =   "Mettre à jour"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         ToolTipText     =   "Mettre à jour les informations"
         Top             =   6600
         Width           =   1695
      End
      Begin ComctlLib.TabStrip TB 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "lang_ok"
         Top             =   4515
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Historique"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Signets"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ListView lstSignets 
         Height          =   1575
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "lang_ok"
         Top             =   4920
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Commentaire"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.TextBox txtProc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "MemPfm.frx":08D9
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Statistiques"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pages=[pages]"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sélection=[selection]"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Offset=[offset]"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Offset Maximum=[offset max]"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Historique=[nombre]"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   21
         Top             =   4200
         Width           =   2895
      End
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.TabStrip MemTB 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Liste des plages disponibles"
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin HexViewer_OCX.HexViewer HW 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4471
      strTag1         =   "0"
      strTag2         =   "0"
   End
   Begin vkUserContolsXP.vkFrame FrameIcon 
      Height          =   3015
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   5318
      BackColor2      =   16777215
      Caption         =   "Icones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ComctlLib.ListView lvIcon 
         Height          =   2535
         Left            =   60
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   4471
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   327682
         Icons           =   "IMG"
         SmallIcons      =   "IMG"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin vkUserContolsXP.vkFrame FrameData 
      Height          =   1455
      Left            =   2400
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2566
      Caption         =   "Valeur"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   0
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Octal :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Hexa :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "ASCII :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
   End
   Begin vkUserContolsXP.vkVScroll VS 
      Height          =   2895
      Left            =   3360
      TabIndex        =   29
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5106
      Value           =   0
      MouseInterval   =   1
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   3240
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "MemPfm"
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
'FORM D'EDITION DE LA MEMOIRE D'UN PROCESSUS
'=======================================================

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private Lang As New clsLang
Private lBgAdress As Long   'offset de départ de page
Private lEdAdress As Long   'offset de fin de page
Private NumberPerPage As Long   'nombre de lignes visibles par Page
Private pRs As Long, pr As Long, pc As Long, pCs As Long 'sauvegarde de la sélection
Private lLength As Long 'taille du fichier
Private ChangeListO() As Long
Private ChangeListC() As Long
Private ChangeListS() As String
Private ChangeListDim As Long
Private sProcess As String
Private lHandle As Long 'handle du processus ouvert
Private lFile As Long   'n° d'ouverture du fichier
Private bOkToOpen As Boolean
Private lMinAdrr As Long
Private lMaxAdrr As Long
Private lBA() As Long
Private lRS() As Long
Private clsProc As clsProcess
Private pProcess As ProcessItem
Private mouseUped As Boolean
Private bFirstChange As Boolean
Private bytFirstChange As Byte

Public cUndo As clsUndoItem 'infos générales sur 'historique
Private cHisto() As clsUndoSubItem  'historique pour le Undo/Redo


Private Sub cmdGoTaskMgr_Click()

    'va dans le gestionnaire de processus
    frmProcess.Show
    
    'permet au curseur de se mettre directement sur le processus édité
    Call SendKeys(cFile.GetFileName(Me.Caption))
    
End Sub

'=======================================================
'effectue la MAJ des informations sur le processus
'=======================================================
Private Sub cmdMAJ_Click()
Dim sDescription As String
Dim sCopyright As String
Dim sVersion As String
Dim lPages As Long
Dim cF As FileSystemLibrary.File
Dim s As String
    
    s = pProcess.szImagePath
    
    'récupère les infos sur le fichier
    Set cF = cFile.GetFile(pProcess.szImagePath)
    
    'récupère les infos sur les fichiers *.exe, *.dll...
    With cF
        sVersion = .FileVersionInfos.FileVersion
        sCopyright = .FileVersionInfos.Copyright
        sDescription = .FileVersionInfos.FileDescription
    
        sVersion = IIf(sVersion = vbNullString, "--", sVersion)
        sCopyright = IIf(sCopyright = vbNullString, "--", sCopyright)
        sDescription = IIf(sDescription = vbNullString, "--", sDescription)
        
        'affiche tout çà
        s = s & vbNewLine & Lang.GetString("_SizeIs") & CStr(.FileSize) & _
            " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        s = s & vbNewLine & Lang.GetString("_AttrIs") & CStr(.Attributes) & "]"
        s = s & vbNewLine & Lang.GetString("_CreaIs") & .DateCreated & "]"
        s = s & vbNewLine & Lang.GetString("_AccessIs") & .DateLastAccessed & "]"
        s = s & vbNewLine & Lang.GetString("_ModifIs") & .DateLastModified & "]"
        s = s & vbNewLine & Lang.GetString("_Version") & sVersion & "]"
        s = s & vbNewLine & Lang.GetString("_DescrIs") & sDescription & "]"
        s = s & vbNewLine & "Copyright=[" & sCopyright & "]"
       
        Label2(8).Caption = Me.Sb.Panels(2).Text
        Label2(9).Caption = Lang.GetString("_SelIs") & CStr(HW.NumberOfSelectedItems) & " bytes]"
        Label2(10).Caption = Me.Sb.Panels(3).Text
        Label2(11).Caption = Lang.GetString("_MaxOff") & CStr(16 * Int(lLength / 16)) & "]"
        'Label2(12).Caption = "[" & sDescription & "]"
    End With
    
    txtFile.Text = s
    
    'affiche les informations concernant le processus
    Call cmdRefreshMemInfo_Click
    
End Sub

Private Sub cmdRefreshMemInfo_Click()
'refresh process info
Dim p As ProcessItem
Dim s As String
    
    Set p = clsProc.GetProcess(pProcess.th32ProcessID, True, False, True)   'réobtient les infos
    
    With p
        s = "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & Lang.GetString("_ParentProc") & .procParentProcess.szExeFile & " - " & pProcess.procParentProcess.th32ProcessID & "]"
        s = s & vbNewLine & Lang.GetString("_UsedMem") & .procMemory.WorkingSetSize & "]"
        s = s & vbNewLine & Lang.GetString("_SWAP") & .procMemory.PagefileUsage & "]"
        s = s & vbNewLine & Lang.GetString("_Prior") & PriorityFromLong(.pcPriClassBase) & "-" & .pcPriClassBase & "]"
        s = s & vbNewLine & "Threads=[" & .cntThreads & "]"
    End With
    
    txtProc.Text = s
End Sub

'=======================================================
'permet de lancer le Resize depuis uen autre form
'=======================================================
Public Sub ResizeMe()
    Form_Resize
End Sub

Private Sub Form_Activate()

    If HW.Visible Then HW.SetFocus
    'Call VS_Change(VS.Value)
    ReDim ChangeListO(1) As Long
    ReDim ChangeListC(1) As Long
    ReDim ChangeListS(1) As String
    ChangeListDim = 1
    
    HW.Refresh
    
    bOkToOpen = False 'pas prêt à l'ouverture
    
    Call UpdateWindow(Me.hWnd)     'refresh de la form

End Sub

Private Sub Form_Load()

    'instancie la classe Undo
    Set cUndo = New clsUndoItem
    
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
    
    'subclasse la form pour éviter de resizer trop
    #If USE_FORM_SUBCLASSING Then
        Call LoadResizing(Me.hWnd, 9000, 6000)
    #End If
    
    'subclasse également lvIcon pour éviter le drag & drop
    Call HookLVDragAndDrop(lvIcon.hWnd)
    
    'affecte les valeurs générales (type) à l'historique
    cUndo.tEditType = edtProcess
    Set cUndo.Frm = Me
    Set cUndo.lvHisto = Me.lstHisto
    ReDim cHisto(0)
    Set cHisto(0) = New clsUndoSubItem
    
    'affiche ou non les éléments en fonction des paramètres d'affichage de frmcontent
    With frmContent
        Me.HW.Visible = .mnuTab.Checked
        Me.MemTB.Visible = .mnuTab.Checked
        Me.VS.Visible = .mnuTab.Checked
        Me.FrameData.Visible = .mnuEditTools.Checked
        Me.FrameInfos.Visible = .mnuInformations.Checked
        Me.FrameIcon.Visible = .mnuShowIcons.Checked
    End With
    
    'change les couleurs du HW
    With cPref
        HW.BackColor = .app_BackGroundColor
        HW.OffsetForeColor = .app_OffsetForeColor
        HW.HexForeColor = .app_HexaForeColor
        HW.StringForeColor = .app_StringsForeColor
        HW.OffsetTitleForeColor = .app_OffsetTitleForeColor
        HW.BaseTitleForeColor = .app_BaseForeColor
        HW.TitleBackGround = .app_TitleBackGroundColor
        HW.LineColor = .app_LinesColor
        HW.SelectionColor = .app_SelectionColor
        HW.ModifiedItemColor = .app_ModifiedItems
        HW.ModifiedSelectedItemColor = .app_ModifiedSelectedItems
        HW.SignetColor = .app_BookMarkColor
        HW.Grid = .app_Grid
        HW.UseHexOffset = CBool(.app_OffsetsHex)

        'en grand dans la MDIform
        If .general_MaximizeWhenOpen Then Me.WindowState = vbMaximized
    End With
    
    frmContent.Sb.Panels(1).Text = "Status=[Opening process]"
    frmContent.Sb.Refresh
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'FileContent = vbNullString
    lNbChildFrm = lNbChildFrm - 1
    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
    
    Close lFile 'ferme le fichier
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'redimensionne le MemTB
    With MemTB
        .Width = 9620
        .Left = IIf(FrameInfos.Visible, FrameInfos.Width, 0) + 50
        .Top = 0
    End With
    
    'redimensionne/bouge le frameInfo
    With FrameInfos
        .Top = 10
        .Height = Me.Height - 700
        .Left = 20
        cmdMAJ.Top = .Height - 330
        lstHisto.Height = .Height - 5350
        lstSignets.Height = .Height - 5350
        'Me.pctContain_cmdMAJ.Height = .Height - 350
    End With
    
    'met le Grid à la taille de la fenêtre
    With HW
        .Width = 9620
        .Height = Me.Height - 400 - Sb.Height - MemTB.Height
        .Left = IIf(FrameInfos.Visible, FrameInfos.Width, 0) + 50
        .Top = MemTB.Height
    End With
    
    'bouge le frameData
    FrameData.Top = 100
    FrameData.Left = IIf(HW.Visible, HW.Width + HW.Left, _
        IIf(FrameInfos.Visible, FrameInfos.Width, 0)) + 500
    
    'bouge le frameIcon
    FrameIcon.Top = 1700
    FrameIcon.Left = IIf(HW.Visible, HW.Width + HW.Left, _
        IIf(FrameInfos.Visible, FrameInfos.Width, 0)) + 500
    
    'calcule le nombre de lignes du Grid à afficher
    'NumberPerPage = Int(Me.Height / 250) - 1
    NumberPerPage = Int(HW.Height / 250) - 1
    
    HW.NumberPerPage = NumberPerPage
    HW.Refresh
   
    With VS
        Call VS_Change(.Value)
        .Top = MemTB.Height
        .Height = Me.Height - 430 - Sb.Height - MemTB.Height
        .Left = IIf(Me.Width < 13100, Me.Width - 350, HW.Left + HW.Width)
    End With
            
End Sub

'=======================================================
'affiche toutes les valeurs hexa visualisable par le HW
'obtenues par lecture dans la mémoire du processus
'=======================================================
Private Sub OpenFile(ByVal lBg As Long, ByVal lEd As Long)
Dim tmpText As String
Dim a As Long
Dim s As String
Dim b As Long
Dim c As Long
Dim lLength As Long
Dim e As Byte
Dim s2 As String
Dim mbi As MEMORY_BASIC_INFORMATION
Dim lLenMBI As Long
Dim si As SYSTEM_INFO
Dim lpMem As Long, Ret As Long, lPos As Long, sBuffer As String
Dim lWritten As Long
Dim sSearchString As String
Dim CalcAddress As Long
Dim sReplaceString As String

    On Error GoTo ErrGestion
    
    'initialise les variables
    c = 1

    HW.ChangeValues 'permet d'empêcher de voir des valeurs hexa vers la fin du fichier
    
    For a = lBg To lEd Step 16
        
        c = c + 1

        'obtient la string correspondant à l'adresse mémoire lue
        sBuffer = cMem.ReadBytesH(lHandle, a - 1, 16)
        
        'ajoute une string formatée
        HW.AddStringValue c - 1, Formated16String(sBuffer)
        
        'ajoute les valeurs hexa à partir de la VRAIE string
        For e = 1 To 16
            s2 = Str2Hex(Mid$(sBuffer, e, 1))
            If Len(s2) = 1 Then s2 = "0" & s2  'formate la string en rajoutant "0" devant si nécessaire
            HW.AddHexValue c - 1, e, s2
        Next e

    Next a

    'Done.
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
    
    'HW.Refresh

    Exit Sub
ErrGestion:
    clsERREUR.AddError "MemPfm.OpenFile", True
End Sub

'=======================================================
'renvoie si l'offset contient une modification
'=======================================================
Private Function IsOffsetModified(ByVal lOffset As Long, ByRef lPlace As Long) As Boolean
Dim x As Long
    
    IsOffsetModified = False
    
    For x = ChangeListDim To 2 Step -1      'ordre décroissant pour pouvoir détecter la dernière modification
    'dans le cas où il y a eu plusieurs modifs dans le même offset
        If ChangeListO(x) = lOffset + 1 Then
            'quelque chose de modifié dans cet ligne
            lPlace = x
            IsOffsetModified = True
            Exit Function
        End If
    Next x
    
End Function

'=======================================================
'renvoie si la case a été modifiée ou non (permet l'affichage en couleur dans HW)
'=======================================================
Private Function IsModified(ByVal lCol As Long, ByVal lOffset As Long) As Boolean
Dim x As Long
    
    IsModified = False
    
    For x = 2 To ChangeListDim
        If ChangeListO(x) = lOffset + 1 Then
            'quelque chose de modifié dans cet ligne
            If ChangeListC(x) = lCol Then
                IsModified = True
                Exit Function
            End If
        End If
    Next x
End Function

'=======================================================
'obtient le nom du processus et l'ouvre ==> procedure qui est appelée
'pour intialiser l'ouverture mémoire
'=======================================================
Public Sub GetFile(ByVal lPID As Long)
Dim l As Long
Dim si As SYSTEM_INFO

    On Error GoTo 5
    
    'active la gestion des langues
    Call Lang.ActiveLang(Me)
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_OpProc") & " " & Trim$(Str$(lPID)) & " ...")
    
    'obtient les infos sur le processus
    Set clsProc = New clsProcess
    Set pProcess = clsProc.GetProcess(lPID, True, True, True)
    
    'permettra de renseigner plus tard sur le processus (à partir d'une autre form)
    Me.Tag = CStr(lPID)
    
    Me.Caption = pProcess.szImagePath
    
    'ajoute l'icone du processus
    Set Me.Icon = CreateIcon(pProcess.szImagePath)
    
    'on récupère les zones mémoires utilisées par le processus
    Call cMem.RetrieveMemRegions(lPID, lBA(), lRS())
    
    If UBound(lBA) = 0 Then
        'pas de données du processus dans l'espace virtuel
        MsgBox Lang.GetString("_NoDataVirtual"), vbInformation, Lang.GetString("_CanNotOpen")
        Unload Me
    End If
    
    'il existe des régions non nulles, alors on créé le nombre de tab à MemTB nécessaire
    '/!\ LIMITE à 255 /!\
    For l = 1 To UBound(lBA)
        MemTB.Tabs.Add pvIndex:=l, pvCaption:="0x" & (lBA(l))
    Next l
    MemTB.Tabs.Remove (MemTB.Tabs.Count) 'enlève le dernier, qui est vide

    MemTB.Tabs.Add pvIndex:=1, pvCaption:="Tout"
5
    On Error GoTo ErrGestion
    
    'récupère le handle du processus
    lHandle = OpenProcess(PROCESS_ALL_ACCESS, False, lPID)
 
    'obtient les adresses mémoire lisibles
    Call GetSystemInfo(si)
    
    lMinAdrr = si.lpMinimumApplicationAddress
    lMaxAdrr = si.lpMaximumApplicationAddress
    
    VS.Min = By16(lBA(1) / 16)  'adresse min de lecture
    VS.Max = By16((lBA(1) + lRS(1)) / 16) 'adresse max de lecture
    VS.Value = By16(lBA(1) / 16)
    
    VS.SmallChange = 1
    VS.LargeChange = NumberPerPage - 1
    
    Call MemTB_Click 'refresh
    
    'stocke dans les tag les valeurs Max et Min des offsets
    HW.curTag1 = HW.FirstOffset
    HW.curTag2 = HW.MaxOffset
    HW.FileSize = 2147483648# '2Go de mémoire virtuelle
    
    
    'affiche aussi les icones du fichier
    Call LoadIconesToLV(pProcess.szImagePath, lvIcon, Me.pct, Me.IMG)
 
    Call cmdMAJ_Click    'refresh les infos

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ProcOpened"))
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "MemPfm.GetFile", True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cUndo = Nothing
    
    #If USE_FORM_SUBCLASSING Then
        'alors enlève le subclassing
        Call RestoreResizing(Me.hWnd)
    #End If
    
    'enleve le hook sur lvIcon également
    Call UnHookLVDragAndDrop(lvIcon.hWnd)
    
    'Call frmContent.MDIForm_Resize 'évite le bug d'affichage
End Sub

Private Sub HW_GotFocus()
    HW.Refresh
End Sub

Private Sub HW_KeyDown(KeyCode As Integer, Shift As Integer)
'gère les touches qui changent le VS, gère le changement de valeur

    On Error GoTo ErrGestion
    
    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    With HW
        If KeyCode = vbKeyUp Then
            'alors monte
            If .FirstOffset = VS.Min * 16 And .Item.Line = 1 Then Exit Sub 'tout au début déjà
            'on remonte d'une ligne alors
            .Item.Line = .Item.Line - 1
            If .Item.Line = 0 Then
                'alors on remonte le firstoffset
                .Item.Line = 1
                .FirstOffset = .FirstOffset - 16
                VS.Value = VS.Value - 1
                Call VS_Change(VS.Value)
            End If
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
        
        If KeyCode = vbKeyDown Then
            'alors descend
            If .FirstOffset + .Item.Line * 16 - 16 = By16(.MaxOffset) Then Exit Sub  'tout en bas
            'on descend d'une ligne alors
            .Item.Line = .Item.Line + 1
            If .Item.Line = .NumberPerPage Then
                'alors on descend le firstoffset
                .Item.Line = .NumberPerPage - 1
                .FirstOffset = .FirstOffset + 16
                VS.Value = VS.Value + 1
                Call VS_Change(VS.Value)
            End If
            'change le VS
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
    End With
    
    With VS
        If KeyCode = vbKeyEnd Then
            'alors aller tout à la fin
            .Value = .Max
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyHome Then
            'alors tout au début
            .Value = .Min
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyPageUp Then
            'alors monter de NumberPerPage
            .Value = IIf((.Value - NumberPerPage) > .Min, .Value - NumberPerPage, .Min)
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyPageDown Then
            'alors descendre de NumberPerPage
            .Value = IIf((.Value + NumberPerPage) < .Max, .Value + NumberPerPage, .Max)
            Call VS_Change(.Value)
        End If
    End With
    
    With HW
        If KeyCode = vbKeyLeft Then
            'alors va à gauche
            If .FirstOffset = VS.Min * 16 And .Item.Col = 1 And .Item.Line = 1 Then Exit Sub 'tout au début déjà
            If .Item.Col = 1 Then
                'tout à gauche ==> on remonte d'une ligne alors
                .Item.Col = 16: .Item.Line = .Item.Line - 1
                If .Item.Line = 0 Then
                    'alors on remonte le firstoffset
                    .Item.Line = 1
                    .FirstOffset = .FirstOffset - 16
                    VS.Value = VS.Value - 1
                    Call VS_Change(VS.Value)
                End If
            Else
                'va à gauche
                .Item.Col = .Item.Col - 1
            End If
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
             
        If KeyCode = vbKeyRight Then
            'alors va à droite
            If .FirstOffset + .Item.Line * 16 - 16 = By16(.MaxOffset) And .Item.Col = 16 Then Exit Sub  'tout à la fin déjà
            If .Item.Col = 16 Then
                'tout à droite ==> on descend d'une ligne alors
                .Item.Col = 1: .Item.Line = .Item.Line + 1
                If .Item.Line = .NumberPerPage Then
                    'alors on descend le firstoffset
                    .Item.Line = .NumberPerPage - 1
                    .FirstOffset = .FirstOffset + 16
                    VS.Value = VS.Value + 1
                    Call VS_Change(VS.Value)
                End If
            Else
                'va à droite
                .Item.Col = .Item.Col + 1
            End If
            'change le VS
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
        
        'réenregistre le numéro de l'offset actuel dans hw.item
        .Item.Offset = .Item.Line * 16 - 16
        'affecte les autres valeurs dans Item
        '.Item.tType = tHex
        .Item.Value = .Value(.Item.Line, .Item.Col)
    End With
    
    DoEvents
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "MemPfm.KeyDown", True
End Sub

Private Sub lvIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu frmContent.mnuPopupIcon
    End If
End Sub

Private Sub HW_KeyPress(KeyAscii As Integer)
'change les valeurs dans le tableau
Dim s As String
Dim sKey As Long
Dim bytHex As Byte
Dim Valu As Byte
Dim x As Byte

    On Error GoTo ErrGestion

    If HW.Item.tType = tHex Then  'si l'on est dans la zone hexa
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Then
            'alors on a ajouté 0,1,...,9,A,B,....,F
            'on change directement dans le tableau à afficher
            
            'détermine Valu en fonction de KeyAscii
            '65 --> "A"
            '48 --> "0"
            '49 --> "1"
            If (KeyAscii >= 48 And KeyAscii <= 57) Then
                Valu = KeyAscii - 48
            ElseIf (KeyAscii >= 65 And KeyAscii <= 70) Then
                Valu = KeyAscii - 55
            ElseIf (KeyAscii >= 97 And KeyAscii <= 102) Then
                Valu = KeyAscii - 87
            End If
            
            If bFirstChange Then
                'alors c'est la seconde valeur hexa
                bytHex = 16 * bytFirstChange + Valu
                bFirstChange = False
            Else
                bytFirstChange = Valu
                bFirstChange = True
                Exit Sub
            End If
            
            'le nouveau byte est donc désormais bytHex
            'l'ancien byte est s
            s = cMem.ReadBytes(pProcess.th32ProcessID, HW.FirstOffset + 16 * (HW.Item.Line - 1) + HW.Item.Col - 1, 1)
            
            'applique le changement
            Call AddAChange(bytHex)
            
            'ajoute l'historique
            Me.AddHistoFrm actByteWritten, Chr_(bytHex), s, HW.FirstOffset + 16 * (HW.Item.Line - 1) + HW.Item.Col - 1, , HW.Item.Col, , pProcess.th32ProcessID
            
           'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
        End If
    ElseIf HW.Item.tType = tString Then
        'alors voici la zone STRING
        'on ne tappe QU'UNE SEULE VALEUR
        
            'le nouveau byte est donc désormais bytHex
            'l'ancien byte est s
            s = cMem.ReadBytes(pProcess.th32ProcessID, HW.FirstOffset + 16 * (HW.Item.Line - 1) + HW.Item.Col - 1, 1)

            'applique le changement
            Call AddAChange(KeyAscii)

            'ajoute l'historique
            Me.AddHistoFrm actByteWritten, Chr_(KeyAscii), s, HW.FirstOffset + 16 * (HW.Item.Line - 1) + HW.Item.Col - 1, , HW.Item.Col, , pProcess.th32ProcessID
            
           'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
    End If

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.KeyPress", True
End Sub

Private Sub HW_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, Item As ItemElement)
Dim s As String
Dim r As Long
Dim l As Currency

    'popup menu
    If Button = 2 Then
        frmContent.mnuDeleteSelection.Enabled = False
        frmContent.mnuCut.Enabled = False
        Me.PopupMenu frmContent.rmnuEdit ', X + GD.Left, Y + GD.Top
    End If
    
    'calcule l'offset (hexa ou décimal)
    l = Item.Line * 16 + HW.FirstOffset - 16 + Item.Col - 1
    If cPref.app_OffsetsHex Then
        clsConv.CurrentString = Trim$(Str$(l))
        s = clsConv.Convert(16, 10)
    Else
        s = CStr(l)
    End If
    Me.Sb.Panels(3).Text = "Offset=[" & s & "]"
    
    Label2(10).Caption = Me.Sb.Panels(3).Text
    
    If Button = 1 Then
        'alors on a sélectionné un objet
        
        If Item.Col < 1 Or Item.Col > 16 Then Exit Sub
        
        If Item.Value = vbNullString Then
            'pas de valeur
            txtValue(0).Text = vbNullString
            txtValue(1).Text = vbNullString
            txtValue(2).Text = vbNullString
            txtValue(3).Text = vbNullString
            FrameData.Caption = "No data"
            Exit Sub
        End If
        
        'affiche la donnée sélectionnée dans frmData
        If Item.tType = tHex Then
            'valeur hexa sélectionnée
            With Item
                txtValue(0).Text = .Value
                txtValue(1).Text = Hex2Dec(.Value)
                txtValue(2).Text = Hex2Str(.Value)
                txtValue(3).Text = Hex2Oct(.Value)
                FrameData.Caption = "Data=[" & .Value & "]"
            End With
        End If
        If Item.tType = tString Then
            'valeur strind sélectionnée
            s = HW.Value(Item.Line, Item.Col)
            txtValue(0).Text = s
            txtValue(1).Text = Hex2Dec(s)
            txtValue(2).Text = Hex2Str(s)
            txtValue(3).Text = Hex2Oct(s)
            FrameData.Caption = "Data=[" & s & "]"
        End If
        
    ElseIf Button = 4 And Shift = 0 Then
        'click avec la molette, et pas de Shift or Control
        'on ajoute (ou enlève) un signet

        If HW.IsSignet(Item.Offset) = False Then
            'on l'ajoute
            HW.AddSignet Item.Offset
            Me.lstSignets.ListItems.Add Text:=CStr(Item.Offset)
            Call HW.TraceSignets
        Else
        
            'alors on l'enlève
            While HW.IsSignet(HW.Item.Offset)
                'on supprime
                HW.RemoveSignet Val(HW.Item.Offset)
            Wend
            
            'enlève du listview
            For r = lstSignets.ListItems.Count To 1 Step -1
                If lstSignets.ListItems.Item(r).Text = CStr(HW.Item.Offset) Then
                    lstSignets.ListItems.Remove r
                End If
            Next r
            
            Call Refresh
        End If
    ElseIf Button = 4 And Shift = 2 Then
        'click molette + control
        'sélectionne une zone définie
        frmSelect.GetEditFunction 0 'selection mode
        frmSelect.Show vbModal
    End If
    
End Sub

Private Sub HW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Sb.Panels(4).Text = Lang.GetString("_SelIs") & _
        CStr(HW.NumberOfSelectedItems) & " bytes]"
    Label2(9) = Me.Sb.Panels(4).Text
End Sub

Private Sub HW_MouseWheel(ByVal lSens As Long)

    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    With VS
        If lSens > 0 Then
            'alors on descend
            .Value = IIf((.Value - .Min) > 0, .Value - 3, .Min)
            Call VS_Change(.Value)
        Else
            'alors on monte
            .Value = IIf((.Value + 3) <= .Max, .Value + 3, .Max)
            Call VS_Change(.Value)
        End If
    End With
    
    DoEvents
End Sub

Private Sub HW_UserMakeFirstOffsetChangeByMovingMouse()
    VS.Value = HW.FirstOffset / 16
End Sub

Private Sub lstSignets_ItemClick(ByVal Item As ComctlLib.ListItem)
'va au signet
    If mouseUped Then
        Me.HW.FirstOffset = Val(Item.Text)
        Me.HW.Refresh
        Me.VS.Value = Me.HW.FirstOffset / 16
        mouseUped = False   'évite de devoir bouger le HW si l'on sélectionne pleins d'items
        'par exemple avec Shift
    End If
End Sub

Private Sub lstSignets_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tLst As ListItem
Dim s As String
Dim r As Long

    If Button = 2 Then
        'alors clic droit ==> on affiche la boite de dialogue "commentaire" sur le comment
        'qui a été sélectionné
        Set tLst = lstSignets.HitTest(x, y)
        If tLst Is Nothing Then Exit Sub
        s = InputBox(Lang.GetString("_AddCommentFor") & " " & tLst.Text, _
            Lang.GetString("_AddCom"))
        If StrPtr(s) <> 0 Then
            'ajoute le commentaire
            tLst.SubItems(1) = s
        End If
    End If
    
    If Button = 4 Then
        'mouse du milieu ==> on supprime le signet
        Set tLst = lstSignets.HitTest(x, y)
        If tLst Is Nothing Then Exit Sub
        
        r = MsgBox(Lang.GetString("_DelSig") & " " & tLst.Text & " ?", _
            vbInformation + vbYesNo, Lang.GetString("_War"))
        If r <> vbYes Then Exit Sub
        
        'on supprime
        HW.RemoveSignet Val(tLst.Text)
        
        'on enlève du listview
        lstSignets.ListItems.Remove tLst.Index
    End If
        
End Sub

Private Sub MemTB_Click()
'alors on change HW et VS en fonction du Tab
Dim l As Long

    l = MemTB.SelectedItem.Index
   
    With VS
        If l = 1 Then
            'affiche tout
            .Min = lMinAdrr / 16
            .Max = lMaxAdrr / 16
            .Value = .Min
            Call VS_Change(.Value)
        Else
            'affiche qu'une partie
            'change les valeurs du VS
            .Min = By16(lBA(l - 1) / 16)
            .Max = By16((lBA(l - 1) + lRS(l - 1)) / 16)
            .Value = .Min
            Call VS_Change(.Value)
        End If
    End With
    
    HW.MaxOffset = VS.Max * 16
    
    HW.Refresh

End Sub

Private Sub MemTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche un popupmenu créé dynamiquement
Dim l As Long

    On Error Resume Next

    If Button = 2 Then
        l = AfficherMenu
        MemTB.Tabs.Item(l).Selected = True
        Call MemTB_Click
    End If
    
End Sub

Private Sub TB_Click()
'change l'affichage de historique/signets
    If TB.SelectedItem.Index = 1 Then
        lstHisto.Visible = True
        lstSignets.Visible = False
    Else
        lstHisto.Visible = False
        lstSignets.Visible = True
    End If
End Sub

Private Sub TB2_Click()
'change les données affichées
    If TB2.SelectedItem.Index = 1 Then
        'alors c'est le Tab qui concerne le fichier
        txtFile.Visible = True
        txtProc.Visible = False
    Else
        'alors c'est le tab qui concerne les infos sur le process
        'alors c'est le Tab qui concerne le fichier
        txtProc.Visible = True
        txtFile.Visible = False
    End If
End Sub

Private Sub lstSignets_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'permet de ne pas changer le HW dans le cas de multiples sélections
    mouseUped = True
End Sub

Private Sub lstSignets_KeyDown(KeyCode As Integer, Shift As Integer)
'vire les signets si touche suppr
Dim r As Long

    mouseUped = True
    
    If KeyCode = vbKeyDelete Then
        'touche suppr
        If lstSignets.SelectedItem.Selected Then
            'alors on supprime quelque chose
            r = MsgBox(Lang.GetString("_DelSign"), vbInformation + vbYesNo, _
                Lang.GetString("_War"))
            If r <> vbYes Then Exit Sub
        
            For r = lstSignets.ListItems.Count To 1 Step -1
                If lstSignets.ListItems.Item(r).Selected Then _
                    lstSignets.ListItems.Remove r
            Next r
        End If
    End If
        
End Sub

'=======================================================
'change la valeur du VS
'public, car cette sub est aussi appelée pour le refresh
'=======================================================
Public Sub VS_Change(Value As Currency)
Dim lPages As Long

    'réaffiche la Grid
    Call OpenFile(16 * Value + 1, VS.Value * 16 + NumberPerPage * 16)
    
    'calcule le nbre de pages
    lPages = lLength / (NumberPerPage * 16) + 1
    Me.Sb.Panels(2).Text = "Page=[" & CStr(1 + Int(VS.Value / NumberPerPage)) & _
        "/" & CStr(lPages) & "]"
    Label2(8).Caption = Me.Sb.Panels(2).Text
    
    HW.FirstOffset = VS.Value * 16
    HW.Refresh
'frmContent.Caption = "min=" & VS.Min & "   val=" & VS.Value & "   max=" & VS.Max
End Sub

'=======================================================
'ajoute une valeur changée dans la liste des valeurs changées
'=======================================================
Public Sub AddChange(ByVal lOffset As Long, ByVal lCol As Long, ByVal sString As String)
   
    'redimensionne le tableau
    ChangeListDim = ChangeListDim + 1
    ReDim Preserve ChangeListO(ChangeListDim) As Long
    ReDim Preserve ChangeListS(ChangeListDim) As String
    ReDim Preserve ChangeListC(ChangeListDim) As Long
    
    'ajoute les nouvelles valeurs
    ChangeListO(ChangeListDim) = lOffset + 1
    ChangeListS(ChangeListDim) = sString
    ChangeListC(ChangeListDim) = lCol
    
    Call Me.VS_Change(VS.Value)
End Sub

'=======================================================
'permet la sauvegarde du contenu mémoire dans un fichier
'=======================================================
Public Function GetNewFile(ByVal sFile2 As String) As String
    
    Call frmSaveProcess.GetProcess(pProcess.th32ProcessID, sFile2) 'renseigne sur le processus à sauvegarder
    frmSaveProcess.Show vbModal

End Function

'=======================================================
'affiche un menu dynamique (popup menu) contenant la liste
'de toutes les zones mémoire disponibles
'=======================================================
Public Function AfficherMenu() As Long
Dim pMenuInfo As MENUITEMINFO 'définit les info de l'item de menu ajouté
Dim pPositionCurseur As POINTAPI 'stocke la position actuelle du curseur
Dim lHandleMenu As Long 'stocke le handle du menu
Dim lHandleSousMenu() As Long 'stocke les handles des sous-menus
Dim x As Long

    On Error GoTo erreur0

    ReDim lHandleSousMenu(MemTB.Tabs.Count)  'nombre de sous-menus à afficher
    
    'on définit le handle du menu popup
    Let lHandleMenu = CreatePopupMenu
    
    'insère MemTB.Tabs.Count sous-menus
    For x = MemTB.Tabs.Count To 1 Step -1
        With pMenuInfo
            .cbSize = Len(pMenuInfo)
            .fType = MFT_STRING
            .fState = MFS_ENABLED
            .dwTypeData = MemTB.Tabs.Item(x).Caption
            .cch = Len(pMenuInfo.dwTypeData)
            .wID = x
            .fMask = MIIM_ID Or MIIM_TYPE Or MIIM_STATE
            .hSubMenu = lHandleMenu
        End With
        Call InsertMenuItem(lHandleMenu, 0, True, pMenuInfo)
        lHandleSousMenu(x) = CreatePopupMenu
    Next x
    
    'on affiche le menu crée
    Call GetCursorPos(pPositionCurseur)
    AfficherMenu = TrackPopupMenuEx(lHandleMenu, TPM_LEFTALIGN Or _
        TPM_RIGHTBUTTON Or TPM_RETURNCMD, pPositionCurseur.x, _
        pPositionCurseur.y, Me.hWnd, ByVal 0&)
    
    Call DestroyMenu(lHandleMenu)
    For x = 1 To MemTB.Tabs.Count
        Call DestroyMenu(lHandleSousMenu(0))
    Next x

Exit Function

erreur0:
    AfficherMenu = -1
End Function

'=======================================================
'fonction ayant uniquement pour but d'exister, on l'appelle à partir d'une
'autre fonction pour tester si frmcontent.activeform est
'une form d'édition mémoire, edition de fichier ou de disque
'=======================================================
Public Function Useless() As String
    Useless = "Mem"
End Function

'=======================================================
'effectue un changement dans la mémoire du processus
'=======================================================
Public Sub AddAChange(ByVal sNewByte As Long)
Dim s As String
Dim x As Long

    'écrit le nouveau byte dans la mémoire
    Call cMem.WriteBytesH(lHandle, HW.FirstOffset + 16 * (HW.Item.Line - 1) + _
        HW.Item.Col - 1, Chr_(sNewByte))
    
    'refresh
    Call VS_Change(VS.Value)
End Sub

'=======================================================
'changement des valeurs dans le FrameData
'=======================================================
Private Sub txtValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'change la valeur de l'item sélectionné du HW de frmcontent.activeform (i_tem)
Dim I_tem As ItemElement

    Set I_tem = HW.Item
    
    If KeyCode = 13 Then
        'alors appui sur "enter"
    
        Select Case Index
            Case 0
                'alors on change les autres champs que le champ "Hexa"
                txtValue(1).Text = Hex2Dec(txtValue(0).Text)
                txtValue(2).Text = Hex2Str(txtValue(0).Text)
                txtValue(3).Text = Hex2Oct(txtValue(0).Text)
            Case 1
                'alors on change les autres champs que le champ "decimal"
                txtValue(0).Text = Hex$(Val(txtValue(1).Text))
                txtValue(2).Text = Byte2FormatedString(Val(txtValue(1).Text))
                txtValue(3).Text = Oct$(Val(txtValue(1).Text))
            Case 2
                'alors on change les autres champs que le champ "ASCII"
                txtValue(0).Text = Str2Hex(txtValue(2).Text)
                txtValue(1).Text = Str2Dec(txtValue(2).Text)
                txtValue(3).Text = Str2Oct(txtValue(2).Text)
            Case 3
                'alors on change les autres champs que le champ "octal"
                txtValue(0).Text = Hex$(Oct2Dec(Val(txtValue(3).Text)))
                txtValue(1).Text = Oct2Dec(Val(txtValue(3).Text))
                txtValue(2).Text = Chr_(Oct2Dec(Val(txtValue(3).Text)))
        End Select

        AddAChange (Hex2Dec(txtValue(0).Text))
        
    End If
        
End Sub

'=======================================================
'ajout d 'un élément à l'historique
'=======================================================
Public Sub AddHistoFrm(ByVal tUndo As UNDO_TYPE, Optional ByVal sData1 As String, _
    Optional ByVal sData2 As String, Optional ByVal curData1 As Currency, _
    Optional ByVal curData2 As Currency, Optional ByVal bytData1 As Byte, _
    Optional ByVal bytData2 As Byte, Optional ByVal lngData1 As Long)
    
    Call AddHisto(-1, cUndo, cHisto(), tUndo, sData1, sData2, curData1, _
        curData2, bytData1, bytData2, lngData1)
    lstHisto.ListItems.Item(lstHisto.ListItems.Count).Selected = True
End Sub

'=======================================================
'ajout d 'un élément à l'historique
'=======================================================
Public Sub UndoM()
    'On Error Resume Next
    
    If lstHisto.SelectedItem Is Nothing Then Exit Sub
    
    With lstHisto
        If .SelectedItem.Index > 1 Then
            .ListItems.Item(.SelectedItem.Index - 1).Selected = True
            cUndo.lRang = .SelectedItem.Index + 1
        Else
            .ListItems.Item(.SelectedItem.Index).Selected = False
            cUndo.lRang = 1
        End If
    End With
    
    Call UndoMe(cUndo, cHisto())
    
    DoEvents    '/!\ DO NOT REMOVE ! (permet d'effectuer les changements d'enabled et de ne pas pouvoir appuyer sur ctrl+Z quand on est au tout début)
    
    Call ModifyHistoEnabled 'vérifie que c'est Ok pour les enabled
End Sub

'=======================================================
'ajout d 'un élément à l'historique
'=======================================================
Public Sub RedoM()
    On Error Resume Next
    With lstHisto
        If .SelectedItem.Index = 1 And .ListItems.Item(1).Selected = False Then
            'rien de sélectionné
            .ListItems.Item(1).Selected = True
            cUndo.lRang = .SelectedItem.Index
        ElseIf .SelectedItem.Index < .ListItems.Count Then
            .ListItems.Item(.SelectedItem.Index + 1).Selected = True
            cUndo.lRang = .SelectedItem.Index
        Else
            .ListItems.Item(.SelectedItem.Index).Selected = False
            cUndo.lRang = .ListItems.Count
        End If
    End With
    
    Call RedoMe(cUndo, cHisto())
    
    DoEvents    '/!\ DO NOT REMOVE !
    
    Call ModifyHistoEnabled 'vérifie que c'est Ok pour les enabled
End Sub

'=======================================================
'refresh simple du HW
'=======================================================
Public Sub RefreshHW()
    Call VS_Change(VS.Value)
End Sub

Private Sub VS_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, x As Long, y As Long)
'alors une action a été effectuée (lance le popup menu)
    If Button = vbRightButton Then Me.PopupMenu frmContent.rmnuPos
End Sub

Private Sub VS_Scroll()
    DoEvents     '/!\ DO NOT REMOVE !
    Call VS_Change(VS.Value)
End Sub
