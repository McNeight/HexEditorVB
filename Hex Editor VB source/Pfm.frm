VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C60799F1-7AA3-45BA-AFBF-5BEAB08BC66C}#1.0#0"; "HexViewer_OCX.ocx"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.Form Pfm 
   AutoRedraw      =   -1  'True
   Caption         =   "Ouverture d'un fichier..."
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   6
   Icon            =   "Pfm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   7635
   Visible         =   0   'False
   Begin vkUserContolsXP.vkVScroll VS 
      Height          =   2895
      Left            =   2880
      TabIndex        =   25
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5106
      Value           =   0
      MouseInterval   =   1
   End
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7860
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Fichier=[Modifi�]"
            TextSave        =   "Fichier=[Modifi�]"
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
            Text            =   "S�lection=[0 Bytes]"
            TextSave        =   "S�lection=[0 Bytes]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin vkUserContolsXP.vkFrame FrameInfos 
      Height          =   6975
      Left            =   3720
      TabIndex        =   14
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
      Begin vkUserContolsXP.vkCommand cmdMAJ 
         Height          =   255
         Left            =   720
         TabIndex        =   27
         ToolTipText     =   "Mettre � jour les informations"
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "Mettre � jour"
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
      Begin ComctlLib.ListView lstSignets 
         Height          =   1575
         Left            =   120
         TabIndex        =   24
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
      Begin ComctlLib.TabStrip TB 
         Height          =   375
         Left            =   120
         TabIndex        =   16
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
      Begin ComctlLib.ListView lstHisto 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
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
      Begin vkUserContolsXP.vkTextBox txtFile 
         Height          =   2175
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3836
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         LegendText      =   "Informations sur le fichier"
         LegendForeColor =   12937777
         LegendType      =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Historique=[nombre]"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   23
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Offset Maximum=[offset max]"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Offset=[offset]"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S�lection=[selection]"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pages=[pages]"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   3240
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
         TabIndex        =   18
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fichier"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
   End
   Begin vkUserContolsXP.vkFrame FrameIcon 
      Height          =   3015
      Left            =   360
      TabIndex        =   3
      Top             =   3120
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
         TabIndex        =   4
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
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin HexViewer_OCX.HexViewer HW 
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4471
      strTag1         =   "0"
      strTag2         =   "0"
   End
   Begin vkUserContolsXP.vkFrame FrameData 
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   6240
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
         Index           =   2
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   9
         Top             =   840
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
         Index           =   0
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "ASCII :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   840
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
         Caption         =   "Hexa :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Octal :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   2520
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Pfm"
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
'FORM D'EDITION DU CONTENU D'UN FICHIER
'=======================================================

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private Lang As New clsLang
Private lBgAdress As Currency   'offset de d�part de page
Private lEdAdress As Currency   'offset de fin de page
Private NumberPerPage As Long   'nombre de lignes visibles par Page
Private lLength As Currency 'taille du fichier
Private ChangeListO() As Currency
Private ChangeListC() As Currency
Private ChangeListS() As String
Private ChangeListDim As Long
Private lFile As Long   'n� d'ouverture du fichier
Private bOkToOpen As Boolean
Private mouseUped As Boolean
Private bFirstChange As Boolean
Private bytFirstChange As Byte
Private bHasMaximized As Long
Private lngFormStyle As Long

Public cUndo As clsUndoItem 'infos g�n�rales sur 'historique
Private cHisto() As clsUndoSubItem  'historique pour le Undo/Redo

Public TheFile As FileSystemLibrary.File

Public Sub cmdMAJ_Click()
'MAJ des infos
Dim lPages As Long
Dim cFic As FileSystemLibrary.File
Dim S As String

    On Error Resume Next
    
    'nom du fichier
    S = "[" & Me.Caption & "]"
    
    'r�cup�re les infos sur le fichier
    Set cFic = cFile.GetFile(Me.Caption)
    
    'affiche tout ��
    With cFic
        S = S & vbNewLine & Lang.GetString("_SizeIs") & CStr(.FileSize) & " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        S = S & vbNewLine & Lang.GetString("_AttrIs") & CStr(.Attributes) & "]"
        S = S & vbNewLine & Lang.GetString("_CreaIs") & .DateCreated & "]"
        S = S & vbNewLine & Lang.GetString("_AccessIs") & .DateLastAccessed & "]"
        S = S & vbNewLine & Lang.GetString("_ModifIs") & .DateLastModified & "]"
        S = S & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        S = S & vbNewLine & Lang.GetString("_DescrIs") & .FileVersionInfos.FileDescription & "]"
        S = S & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
    End With
    
    txtFile.Text = S
    
    Label2(8).Caption = Me.Sb.Panels(2).Text
    Label2(9).Caption = Lang.GetString("_SelIs") & CStr(HW.NumberOfSelectedItems) & " bytes]"
    Label2(10).Caption = Me.Sb.Panels(3).Text
    Label2(11).Caption = Lang.GetString("_MaxOff") & CStr(16 * Int(lLength / 16)) & "]"

    'Label2(12).Caption = "[" & sDescription & "]"
    
    Set cFic = Nothing

End Sub

Private Sub Form_Activate()
    
    cmdMAJ_Click
    If HW.Visible Then HW.SetFocus
    Call VS_Change(VS.Value)
    ReDim ChangeListO(1) As Currency
    ReDim ChangeListC(1) As Currency
    ReDim ChangeListS(1) As String
    ChangeListDim = 1
    
    HW.Refresh

    Call UpdateWindow(Me.hWnd)     'refresh de la form
End Sub

Private Sub Form_Load()
        
    'instancie la classe Undo
    Set cUndo = New clsUndoItem
    
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
    
    lngFormStyle = GetWindowLong(Me.hWnd, GWL_STYLE)

    'subclasse la form pour �viter de resizer trop
    #If USE_FORM_SUBCLASSING Then
        Call LoadResizing(Me.hWnd, 9000, 6000)
    #End If
    
    'subclasse �galement lvIcon pour �viter le drag & drop
    Call HookLVDragAndDrop(lvIcon.hWnd)

    
    'affecte les valeurs g�n�rales (type) � l'historique
    cUndo.tEditType = edtFile
    Set cUndo.Frm = Me
    Set cUndo.lvHisto = Me.lstHisto
    ReDim cHisto(0)
    Set cHisto(0) = New clsUndoSubItem
    
    'affiche ou non les �l�ments en fonction des param�tres d'affichage de frmcontent
    With frmContent
        Me.HW.Visible = .mnuTab.Checked
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
    End With
    
    bOkToOpen = False 'pas pr�t � l'ouverture
    
    With cPref
        'en grand dans la MDIform
        If .general_MaximizeWhenOpen Then Me.WindowState = vbMaximized
    End With
        
    frmContent.Sb.Panels(1).Text = "Status=[Opening File]"
    frmContent.Sb.Refresh
    
    bFirstChange = False 'pas de KeyAscii d�j� appuy� pour modifier une valeur

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'FileContent = vbNullString
    lNbChildFrm = lNbChildFrm - 1
    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
    
    Call CloseHandle(lFile)  'ferme le handle sur le fichier
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'r�cup�re le style de la form
    
    'If Me.WindowState = vbMaximized Then
        'alors si on est sous Vista, il faut supprimer la bordure de la fen�tre pour
        '�viter un bug d'affichage

        

          '  If lngFormStyle And WS_BORDER Then
                'lngFormStyle = lngFormStyle And (Not WS_BORDER)
          '  End If
         '   If lngFormStyle And WS_THICKFRAME Then
                'lngFormStyle = lngFormStyle And (Not WS_THICKFRAME)
          '  End If
            
           ' Call SetWindowLong(Me.hWnd, GWL_STYLE, 0)    'vire le style border

        
 '   Else
            
         '   If (lngFormStyle And WS_BORDER) = 0 Then
               ' lngFormStyle = lngFormStyle And WS_BORDER
         '   End If
          '  If (lngFormStyle And WS_THICKFRAME) = 0 Then
                'lngFormStyle = lngFormStyle And WS_THICKFRAME
         '   End If
            
      '      Call SetWindowLong(Me.hWnd, GWL_STYLE, 118791072)    'rajoute le style border


 '   End If
    

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
    
    'met le Grid � la taille de la fen�tre
    With HW
        .Width = 9620
        .Height = Me.Height - 400 - Sb.Height
        .Left = IIf(FrameInfos.Visible, FrameInfos.Width, 0) + 50
        .Top = 0
    End With
    
    'bouge le frameData
    FrameData.Top = 100
    FrameData.Left = IIf(HW.Visible, HW.Width + HW.Left, _
        IIf(FrameInfos.Visible, FrameInfos.Width, 0)) + 500
    
    'bouge le frameIcon
    FrameIcon.Top = 1700
    FrameIcon.Left = IIf(HW.Visible, HW.Width + HW.Left, _
        IIf(FrameInfos.Visible, FrameInfos.Width, 0)) + 500
    
    'calcule le nombre de lignes du Grid � afficher
    'NumberPerPage = Int(Me.Height / 250) - 1
    NumberPerPage = Int(HW.Height / 250) - 1
    
    HW.NumberPerPage = NumberPerPage
    HW.Refresh
   
    With VS
        Call VS_Change(.Value)
        .Top = 0
        .Height = Me.Height - 430 - Sb.Height
        .Left = IIf(Me.Width < 13100, Me.Width - 350, HW.Left + HW.Width)
    End With
    
    
    'on cr�� un gradient de couleur sur le fond de la Form
'    Dim lC As RGB_COLOR
'    Dim RC As RGB_COLOR
'
'    'gris clair
'    With lC
'        .r = 156
'        .G = 186
'        .b = 234
'    End With
'
'    'blanc
'    With RC
'        .r = 197
'        .G = 215
'        .b = 248
'    End With
'
'    'trace le gradient
'    Call FillGradient(Me, lC, RC)
            
End Sub

'=======================================================
'permet de lancer le Resize depuis uen autre form
'=======================================================
Public Sub ResizeMe()
    Call Form_Resize
End Sub

'=======================================================
'affiche dans le HW les valeurs hexa qui correspondent � la partie
'du fichier qui est visualis�e
'=======================================================
Private Sub OpenFile(ByVal lBg As Currency, ByVal lEd As Currency)
Dim tmpText As String
Dim a As Long
Dim S As String
Dim b As Long
Dim c As Long
Dim s2 As String
Dim lLength As Long
Dim e As Byte
Dim sTemp() As String
Dim lOff As Currency
Dim lPlace As Long
Dim Ret As Long
        
    If bOkToOpen = False Then Exit Sub  'pas pr�t � ouvrir
    
    Call HW.ChangeValues 'permet d'emp�cher de voir des valeurs hexa vers la fin du fichier
    
    'cr�� un buffer qui contiendra les valeurs
    tmpText = String$(By16(lEd - lBg) + 16, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, (lBg - 1) / 10000, 0&, FILE_BEGIN)    'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE
    
    'prend un morceau du fichier
    'tmpText = String$(16, 0)  'buffer
    Ret = ReadFile(lFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'obtient le texte � partir du fichier (depuis l'octet lBg, et pour une intervalle contenant tout les bytes affich�s � l'�cran
    'Get lFile, lBg, tmpText
        
    ReDim sTemp(HW.NumberPerPage)
    
    'divise tmpText en parties de 16 bytes et stocke dans sTemp
    For e = 1 To HW.NumberPerPage

        'stocke la VRAIE string de 16
        sTemp(e) = Mid$(tmpText, e * 16 - 15, 16)
        
        S = vbNullString
        For a = 1 To 16
        
            'calcule l'offset �quivalent
            lOff = (e - 1) * 16 + VS.Value * 16
            
            If IsOffsetModified(lOff, lPlace) Then
                'l'offset est modifi�

                'la string existe dans la liste des string modifi�es
                S = ChangeListS(lPlace)   'ne formate PAS la string
                
                'ajoute les valeurs hexa � partir de la VRAIE NOUVELLE string
                For c = 1 To 16
                    s2 = Hex$(Asc(Mid$(S, c, 1)))
                    If Len(s2) = 1 Then s2 = "0" & s2
                    HW.AddHexValue e, c, s2, IsModified(c, lOff)
                Next c
                
                S = Formated16String(S) 'formate la string POUR L'AFFICHAGE
                Exit For
            Else

                'formate une string de 16 pour l'affichage
                S = S & Byte2FormatedString(Asc(Mid$(sTemp(e), a, 1)))
                
                'ajoute les valeurs hexa � partir de la VRAIE string
                s2 = Hex$(Asc(Mid$(sTemp(e), a, 1)))
                If Len(s2) = 1 Then s2 = "0" & s2
                HW.AddHexValue e, a, s2, False
            End If
            
        Next a
        
        'ajoute la string FORMATEE pour l'affichage
        HW.AddStringValue e, S
        
    Next e

    'Done.
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
    
    'HW.Refresh  'affiche les r�sultats
End Sub

'=======================================================
'renvoie si l'offset contient une modification
'=======================================================
Public Function IsOffsetModified(ByVal lOffset As Currency, ByRef lPlace As Long) As Boolean
Dim X As Long
    
    IsOffsetModified = False
    
    For X = ChangeListDim To 2 Step -1      'ordre d�croissant pour pouvoir d�tecter la derni�re modification
    'dans le cas o� il y a eu plusieurs modifs dans le m�me offset
        If ChangeListO(X) = lOffset + 1 Then
            'quelque chose de modifi� dans cet ligne
            lPlace = X
            IsOffsetModified = True
            Exit Function
        End If
    Next X
    
End Function


'=======================================================
'efface les �lements de la s�lection
'=======================================================
Public Sub DeleteZone()
Dim tempFile As String
Dim X As Long, S As String
Dim tmpText As String
Dim Y As Long
Dim a As Long
Dim lNewPos As Long
Dim e As Long
Dim lLen As Long
Dim lFile2 As Long
Dim lPlace As Long
Dim tempFileLen As Long
Dim OfF As Currency
Dim OfL As Currency
    
    On Error GoTo ErrGestion
    
    'obtient un path temporaire
    Call ObtainTempPathFile(vbNullString, tempFile, cFile.GetFileExtension(Me.Caption))
    
    '//suppression d'une zone
    'cr�� le fichier normalement jusqu'� la zone � enlever
    'affiche le message d'attente
    frmContent.Sb.Panels(1).Text = "Status=[Creating backup " & Me.Caption & "]"
    
    'd�termine les offsets de d�limitation
    OfF = Me.HW.FirstSelectionItem.Offset + Me.HW.FirstSelectionItem.Col
    OfL = Me.HW.SecondSelectionItem.Offset + Me.HW.SecondSelectionItem.Col
    
    lFile2 = FreeFile 'obtient une ouverture dispo
    Open tempFile For Binary Access Write As lFile2   'ouvre le fichier sFile2 pour l'enregistrement
       
    'cr�� un buffer de longueur divisible par 16
    'recoupera � la fin pour la longueur exacte du fichier
    lLen = By16(lLength)
    
    tempFileLen = lLength   'contient la taille que fera le fichier temporaire (apr�s suppression)
    'pour l'instant c'est la taille normale
    
    lNewPos = 0 'pas de d�calage pour le moment
    
    tmpText = String$(16, 0)
       
        For a = 1 To lLen Step 16  'remplit par intervalles de 16
                    
            Get lFile, a, tmpText   'prend un morceau du fichier source (16bytes) � partir de l'octet A

            If Not (((a - 1) >= 16 * Int(OfF / 16)) And ((a - 1) <= By16(OfL))) Then
                'ALORS n'appartient pas � la zone enlev�e
                If IsOffsetModified(a - 1, lPlace) Then
                    'l'offset est modifi�
    
                    'la string existe dans la liste des string modifi�es
                    'ne formate PAS la string
                    
                    'v�rifie que l'on est pas dans la derni�re ligne, si oui, ne prend que la longueur n�cessaire
                    If a + 16 > tempFileLen Then
                        S = Mid$(ChangeListS(lPlace), 1, tempFileLen - a + 1)
                    Else
                        S = ChangeListS(lPlace)
                    End If
                Else
                    
                    'pas de modif, prend � partir du fichier source
                    'v�rfie que l'on est pas dans la derni�re ligne
                    
                    If a + 16 > tempFileLen Then
                        S = Mid$(tmpText, 1, lLength - a + 1)
                    Else
                        S = tmpText
                    End If
                End If
                
                Put lFile2, a + lNewPos, S  'ajoute au fichier r�sultant
                
            Else
                'ALORS appartient � la zone enlev�e
                lNewPos = -OfL + OfF    'd�calage
                tempFileLen = lLength - OfL + OfF  'change la taille du fichier temp
            End If
            
            
            If (a Mod 160017) = 0 Then
                'rend un peu la main
                frmContent.Sb.Panels(1).Text = "Status=[Creating backup " & Me.Caption & "]" & "[" & Round(100 * a / lLength, 2) & " %]"
                DoEvents
            End If

        Next a
    
    Close lFile 'ferme l'ANCIEN FICHIER
    lFile = lFile2  'change le num�ro d'enregistrement ==> fichier temp
    
    Me.Caption = tempFile 'le nouveau nom de fichier
    
    'affiche le message de fin de sauvegarde
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.DeleteZone", True
End Sub

'=======================================================
'renvoie si la case a �t� modifi�e ou non
'=======================================================
Private Function IsModified(ByVal lCol As Long, ByVal lOffset As Currency) As Boolean
Dim X As Long
    
    IsModified = False
    
    For X = 2 To ChangeListDim
        If ChangeListO(X) = lOffset + 1 Then
            'quelque chose de modifi� dans cet ligne
            If ChangeListC(X) = lCol Then
                IsModified = True
                Exit Function
            End If
        End If
    Next X
End Function

'=======================================================
'obtient le nom du fichier � ouvrir, et l'ouvre
'=======================================================
Public Sub GetFile(ByVal sFile As String)
Dim l As Currency

    'On Error GoTo ErrGestion

    'active la gestion des langues
    Call Lang.ActiveLang(Me)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_OpFileCour") & " " & sFile & " ...")
    
    'r�cup�re les infos fichier
    Set TheFile = cFile.GetFile(sFile)
    
    'r�cup�re le contenu du fichier
    lLength = TheFile.FileSize
    
    HW.MaxOffset = lLength 'offset maximal
    HW.FileSize = lLength
    
    Me.Caption = sFile
    
    'ajoute l'icone � notre form
    Set Me.Icon = CreateIcon(sFile)
    
    'obtient un handle vers le fichier
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If NumberPerPage = 0 Then Call Form_Resize  'besoin de cette valeur
        
    'r�gle la taille de VS
    With VS
        .Min = 0
        l = Int(lLength / 16)
        .Max = l
        .Value = 0
        .SmallChange = 1
        .LargeChange = NumberPerPage - 1
    End With
    
    'stocke dans les tag les valeurs Max et Min des offsets
    With HW
        .curTag1 = .FirstOffset
        .curTag2 = .MaxOffset
    End With
    
    bOkToOpen = True 'pr�t � l'ouverture

    'affichage
    Call OpenFile(16 * VS.Value + 1, VS.Value * 16 + NumberPerPage * 16)
    
    'affiche aussi les icones du fichier
    Call LoadIconesToLV(sFile, lvIcon, Me.pct, Me.IMG)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_File") & " " & sFile & " " & Lang.GetString("_Opened"))
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.GetFile", True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cUndo = Nothing
    Set Lang = Nothing
    
    #If USE_FORM_SUBCLASSING Then
        'alors enl�ve le subclassing
        Call RestoreResizing(Me.hWnd)
    #End If
    
    'enleve le hook sur lvIcon �galement
    Call UnHookLVDragAndDrop(lvIcon.hWnd)
    
    'Call frmContent.MDIForm_Resize '�vite le bug d'affichage
End Sub

Private Sub HW_GotFocus()
    HW.Refresh
End Sub

Private Sub HW_KeyDown(KeyCode As Integer, Shift As Integer)
'g�re les touches qui changent le VS, g�re le changement de valeur
    
    On Error GoTo ErrGestion
    
    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    With HW
        If KeyCode = vbKeyUp Then
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors monte
            If .FirstOffset = 0 And .Item.Line = 1 Then Exit Sub  'tout au d�but d�j�
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
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors descend
            
            'on v�rifie que l'on ne d�passe pas la fin du fichier
            If (.FirstOffset + 16 * .Item.Line + .Item.Col) > lLength Then _
                Exit Sub    'd�passe du fichier
                
                
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
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors aller tout � la fin
            .Value = .Max
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyHome Then
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors tout au d�but
            .Value = .Min
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyPageUp Then
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors monter de NumberPerPage
            VS.Value = IIf((.Value - NumberPerPage) > .Min, .Value - NumberPerPage, .Min)
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyPageDown Then
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors descendre de NumberPerPage
            .Value = IIf((.Value + NumberPerPage) < .Max, .Value + NumberPerPage, .Max)
            Call VS_Change(.Value)
        End If
    End With
    
    With HW
        If KeyCode = vbKeyLeft Then
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors va � gauche
            If .FirstOffset = 0 And .Item.Col = 1 And .Item.Line = 1 Then Exit Sub 'tout au d�but d�j�
            If .Item.Col = 1 Then
                'tout � gauche ==> on remonte d'une ligne alors
                .Item.Col = 16: .Item.Line = .Item.Line - 1
                If .Item.Line = 0 Then
                    'alors on remonte le firstoffset
                    .Item.Line = 1
                    .FirstOffset = .FirstOffset - 16
                    VS.Value = VS.Value - 1
                    Call VS_Change(VS.Value)
                End If
            Else
                'va � gauche
                .Item.Col = .Item.Col - 1
            End If
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
             
        If KeyCode = vbKeyRight Then
            bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
            'pour l'�dition dynamique au clavier
            
            'alors on va � droite
            
            'on v�rifie que l'on ne d�passe pas la fin du fichier
            If (.FirstOffset + 16 * (.Item.Line - 1) + .Item.Col) >= lLength Then _
                Exit Sub    'd�passe du fichier
            
            If .FirstOffset + .Item.Line * 16 - 16 = By16(.MaxOffset) And .Item.Col = 16 Then Exit Sub  'tout � la fin d�j�
            If .Item.Col = 16 Then
                'tout � droite ==> on descend d'une ligne alors
                .Item.Col = 1: .Item.Line = .Item.Line + 1
                If .Item.Line = .NumberPerPage Then
                    'alors on descend le firstoffset
                    .Item.Line = .NumberPerPage - 1
                    .FirstOffset = .FirstOffset + 16
                    VS.Value = VS.Value + 1
                    Call VS_Change(VS.Value)
                End If
            Else
                'va � droite
                .Item.Col = .Item.Col + 1
            End If
            'change le VS
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
        
        'r�enregistre le num�ro de l'offset actuel dans hw.item
        .Item.Offset = .Item.Line * 16 - 16
        'affecte les autres valeurs dans Item
        'HW.Item.tType = tHex
        .Item.Value = .Value(.Item.Line, .Item.Col)
    End With
    
    DoEvents

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.KeyDown", True
End Sub

Private Sub HW_KeyPress(KeyAscii As Integer)
'change les valeurs dans le tableau
Dim S As String
Dim sKey As Long
Dim bytHex As Byte
Dim Valu As Byte
Dim X As Byte
Dim s2 As String
    
    If HW.Item.tType = tHex Then  'si l'on est dans la zone hexa
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Then
            'alors on a ajout� 0,1,...,9,A,B,....,F
            'on change directement dans le tableau � afficher
            
            'd�termine Valu en fonction de KeyAscii
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
            
            'le nouveau byte est donc d�sormais bytHex

            's est la nouvelle string
            S = vbNullString
            For X = 1 To 16
                S = S & Hex2Str_(HW.Value(HW.Item.Line, CLng(X)))
            Next X
            s2 = S
            
            'calcule la nouvelle string (partie de gauche ancienne + nouveau byte + partie de droite ancienne)
            S = Mid$(S, 1, HW.Item.Col - 1) & Chr_(bytHex) & Mid$(S, HW.Item.Col + 1, 16 - HW.Item.Col)   'avant & nouvelle & apr�s
            
            'applique le changement
            Call Me.AddChange((HW.Item.Line - 1) * 16 + HW.FirstOffset, HW.Item.Col, S)
            'ajoute l'historique
            Call Me.AddHistoFrm(actByteWritten, s2, S, (HW.Item.Line - 1) * _
                16 + HW.FirstOffset, , HW.Item.Col)
           'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
            
        End If
    ElseIf HW.Item.tType = tString Then
        'alors voici la zone STRING
        'on ne tappe QU'UNE SEULE VALEUR
        
            'le nouveau byte est donc d�sormais KeyAscii

            's est la nouvelle string
            S = vbNullString
            For X = 1 To 16
                S = S & Hex2Str_(HW.Value(HW.Item.Line, CLng(X)))
            Next X
            s2 = S
            
            'calcule la nouvelle string (partie de gauche ancienne + nouveau byte + partie de droite ancienne)
            S = Mid$(S, 1, HW.Item.Col - 1) & Chr_(KeyAscii) & Mid$(S, HW.Item.Col + 1, 16 - HW.Item.Col)   'avant & nouvelle & apr�s
            
            'applique le changement
            Call Me.AddChange((HW.Item.Line - 1) * 16 + HW.FirstOffset, HW.Item.Col, S)
            'ajoute l'historique
            Call Me.AddHistoFrm(actByteWritten, s2, S, (HW.Item.Line - 1) * _
                16 + HW.FirstOffset, , HW.Item.Col)
            'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
            
    End If
End Sub

Public Sub HW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Item As HexViewer_OCX.ItemElement)
Dim S As String
Dim r As Long
Dim l As Currency

    'popup menu
    If Button = 2 Then
        frmContent.mnuDeleteSelection.Enabled = True
        frmContent.mnuCut.Enabled = True
        Me.PopupMenu frmContent.rmnuEdit ', X + GD.Left, Y + GD.Top
    End If
    
    'calcule l'offset (hexa ou d�cimal)
    l = Item.Line * 16 + HW.FirstOffset - 16 + Item.Col - 1
    If cPref.app_OffsetsHex Then
        clsConv.CurrentString = Trim$(Str$(l))
        S = clsConv.Convert(16, 10)
    Else
        S = CStr(l)
    End If
    Me.Sb.Panels(3).Text = "Offset=[" & S & "]"
    
    Label2(10).Caption = Me.Sb.Panels(3).Text
    
    If Button = 1 Then
        'alors on a s�lectionn� un objet
        
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'�dition dynamique au clavier
        
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
        
        'affiche la donn�e s�lectionn�e dans frmData
        If Item.tType = tHex Then
            'valeur hexa s�lectionn�e
            With Item
                txtValue(0).Text = .Value
                txtValue(1).Text = Hex2Dec(.Value)
                txtValue(2).Text = Hex2Str(.Value)
                txtValue(3).Text = Hex2Oct(.Value)
                FrameData.Caption = "Data=[" & .Value & "]"
            End With
        End If
        If Item.tType = tString Then
            'valeur strind s�lectionn�e
            S = HW.Value(Item.Line, Item.Col)
            txtValue(0).Text = S
            txtValue(1).Text = Hex2Dec(S)
            txtValue(2).Text = Hex2Str(S)
            txtValue(3).Text = Hex2Oct(S)
            FrameData.Caption = "Data=[" & S & "]"
        End If
        
    ElseIf Button = 4 And Shift = 0 Then
        'click avec la molette, et pas de Shift or Control
        'on ajoute (ou enl�ve) un signet

        If HW.IsSignet(Item.Offset) = False Then
            'on l'ajoute
            HW.AddSignet Item.Offset
            Me.lstSignets.ListItems.Add Text:=CStr(Item.Offset)
            Call HW.TraceSignets
        Else
        
            'alors on l'enl�ve
            While HW.IsSignet(HW.Item.Offset)
                'on supprime
                HW.RemoveSignet Val(HW.Item.Offset)
            Wend
            
            'enl�ve du listview
            For r = lstSignets.ListItems.Count To 1 Step -1
                If lstSignets.ListItems.Item(r).Text = CStr(HW.Item.Offset) Then
                    lstSignets.ListItems.Remove r
                End If
            Next r
            
            Call Refresh
        End If
    ElseIf Button = 4 And Shift = 2 Then
        'click molette + control
        's�lectionne une zone d�finie
        Call frmSelect.GetEditFunction(0)  'selection mode
        frmSelect.Show vbModal
    End If
    
End Sub

Private Sub HW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Sb.Panels(4).Text = "S�lection=[" & CStr(HW.NumberOfSelectedItems) & " bytes]"
    Label2(9) = Me.Sb.Panels(4).Text
End Sub

Private Sub HW_MouseWheel(ByVal lSens As Long)

    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
    'pour l'�dition dynamique au clavier
    
    With VS
        If lSens > 0 Then
            'alors on descend
            .Value = IIf((.Value - 3) >= 0, .Value - 3, 0)
            Call VS_Change(.Value)
        Else
            'alors on monte
            .Value = IIf((.Value + 3) <= .Max, .Value + 3, .Max)
            Call VS_Change(.Value)
        End If
    End With
End Sub

Private Sub HW_UserMakeFirstOffsetChangeByMovingMouse()
    VS.Value = HW.FirstOffset / 16
End Sub

Private Sub lstHisto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'on supprime l'historique si button=2
'on va � l'�l�ment si button=4
Dim Item As ListItem
        
    If Button = 2 Then
        'r�cup�re l'item
        Set Item = lstHisto.HitTest(X, Y)
        If Item Is Nothing Then Exit Sub
        
        'supprime
        Call DelHisto(Val(Item.SubItems(1)), cHisto(), cUndo)
    ElseIf Button = 4 Then
        'r�cup�re l'item
        Set Item = lstHisto.HitTest(X, Y)
        If Item Is Nothing Then Exit Sub
        
        'active l'�l�ment
        cUndo.lRang = Item.Index
        Item.Selected = True
        Call RedoMe(cUndo, cHisto())
        DoEvents    '/!\ DO NOT REMOVE !
        Call ModifyHistoEnabled 'v�rifie que c'est Ok pour les enabled
    End If
    
End Sub

Private Sub lstSignets_ItemClick(ByVal Item As ComctlLib.ListItem)
'va au signet
    If mouseUped Then
        Me.HW.FirstOffset = Val(Item.Text)
        Me.HW.Refresh
        Me.VS.Value = Me.HW.FirstOffset / 16
        mouseUped = False   '�vite de devoir bouger le HW si l'on s�lectionne pleins d'items
        'par exemple avec Shift
    End If
End Sub

Private Sub lstSignets_KeyDown(KeyCode As Integer, Shift As Integer)
'vire les signets si touche suppr
Dim r As Long

    mouseUped = True
    
    If KeyCode = vbKeyDelete Then
        'touche suppr
        If lstSignets.SelectedItem.Selected Then
            'alors on supprime quelque chose
            r = MsgBox(Lang.GetString("_DelSign"), vbInformation + vbYesNo, Lang.GetString("_War"))
            If r <> vbYes Then Exit Sub
        
            For r = lstSignets.ListItems.Count To 1 Step -1
                If lstSignets.ListItems.Item(r).Selected Then lstSignets.ListItems.Remove r
            Next r
        End If
    End If
        
End Sub

Private Sub lstSignets_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tLst As ListItem
Dim S As String
Dim r As Long

    If Button = 2 Then
        'alors clic droit ==> on affiche la boite de dialogue "commentaire" sur le comment
        'qui a �t� s�lectionn�
        Set tLst = lstSignets.HitTest(X, Y)
        If tLst Is Nothing Then Exit Sub
        S = InputBox(Lang.GetString("_AddCommentFor") & " " & tLst.Text, Lang.GetString("_AddCom"))
        If StrPtr(S) <> 0 Then
            'ajoute le commentaire
            tLst.SubItems(1) = S
        End If
    End If
    
    If Button = 4 Then
        'mouse du milieu ==> on supprime le signet
        Set tLst = lstSignets.HitTest(X, Y)
        If tLst Is Nothing Then Exit Sub
        
        r = MsgBox(Lang.GetString("_DelSig") & " " & tLst.Text & " ?", vbInformation + vbYesNo, Lang.GetString("_War"))
        If r <> vbYes Then Exit Sub
        
        'on supprime
        HW.RemoveSignet Val(tLst.Text)
        
        'on enl�ve du listview
        lstSignets.ListItems.Remove tLst.Index
    End If
        
End Sub

Private Sub lstSignets_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'permet de ne pas changer le HW dans le cas de multiples s�lections
    mouseUped = True
End Sub

Private Sub lvIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu frmContent.mnuPopupIcon
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

Private Sub txtValue_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0   'emp�che le 'BEEP'
End Sub

'=======================================================
'change la valeur du VS
'public, car cette sub est aussi appel�e pour le refresh
'=======================================================
Public Sub VS_Change(Value As Currency)
Dim lPages As Long

    'On Error GoTo ErrGestion
    
    If NumberPerPage = 0 Then Exit Sub
    
    'r�affiche la Grid
    Call OpenFile(16 * Value + 1, VS.Value * 16 + NumberPerPage * 16)
    
    'calcule le nbre de pages
    lPages = lLength / (NumberPerPage * 16) + 1
    Me.Sb.Panels(2).Text = "Page=[" & CStr(1 + Int(VS.Value / NumberPerPage)) & _
        "/" & CStr(lPages) & "]"
    Label2(8).Caption = Me.Sb.Panels(2).Text
    
    HW.FirstOffset = VS.Value * 16
    HW.Refresh

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.VS_Change", True
End Sub

'=======================================================
'ajoute une valeur chang�e dans la liste des valeurs chang�es
'=======================================================
Public Sub AddChange(ByVal lOffset As Currency, ByVal lCol As Long, ByVal sString As String)
   
    'redimensionne le tableau
    ChangeListDim = ChangeListDim + 1
    ReDim Preserve ChangeListO(ChangeListDim) As Currency
    ReDim Preserve ChangeListS(ChangeListDim) As String
    ReDim Preserve ChangeListC(ChangeListDim) As Currency
    
    'ajoute les nouvelles valeurs
    ChangeListO(ChangeListDim) = lOffset + 1
    ChangeListS(ChangeListDim) = sString
    ChangeListC(ChangeListDim) = lCol
    
    Call Me.VS_Change(VS.Value)
End Sub

'=======================================================
'proc�de � la sauvegarde du fichier avec changements � l'emplacement sFile2
'=======================================================
Public Function GetNewFile(ByVal sFile2 As String) As String
Dim X As Long, S As String
Dim tmpText As String
Dim dblAdv As Double
Dim Y As Long
Dim a As Long
Dim e As Long
Dim lLen As Long, lFile2 As Long
Dim lPlace As Long
Dim Ret As Long

    On Error GoTo ErrGestion
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SavingFile") & " " & sFile2 & " ...")

    'affiche le message d'attente
    frmContent.Sb.Panels(1).Text = "Status=[Saving " & Me.Caption & "]"
       
    'obtient un handle vers le fichier � �crire
    'ouverture en ECRITURE, avec overwrite si d�j� existant (car d�j� demand� confirmation avant)
    lFile2 = CreateFile(sFile2, GENERIC_WRITE, FILE_SHARE_READ Or _
        FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0)

    'cr�� un buffer de longueur divisible par 16
    'recoupera � la fin pour la longueur exacte du fichier
    lLen = By16(lLength)
    tmpText = String$(16, 0)
       
        For a = 1 To lLen Step 16  'remplit par intervalles de 16
            
            'bouge le pointeur sur lr fichier au bon emplacement
            Ret = SetFilePointerEx(lFile, (a - 1) / 10000, 0&, FILE_BEGIN)
            'a divis� par 10^4 pour obtenir un nombre d�cimal de Currency
            
            'prend un morceau du fichier
            tmpText = String$(16, 0)  'buffer
            Ret = ReadFile(lFile, ByVal tmpText, 16, Ret, ByVal 0&)
       
            'Get lFile, a, tmpText   'prend un morceau du fichier source (16bytes) � partir de l'octet A
                      
            If IsOffsetModified(a - 1, lPlace) Then
                'l'offset est modifi�

                'la string existe dans la liste des string modifi�es
                'ne formate PAS la string
                
                'v�rifie que l'on est pas dans la derni�re ligne, si oui, ne prend que la longueur n�cessaire
                If a + 16 > lLength Then
                    S = Mid$(ChangeListS(lPlace), 1, lLength - a + 1)
                Else
                    S = ChangeListS(lPlace)
                End If
            Else
                
                'pas de modif, prend � partir du fichier source
                'v�rfie que l'on est pas dans la derni�re ligne
                
                If a + 16 > lLength Then
                    S = Mid$(tmpText, 1, lLength - a + 1)
                Else
                    S = tmpText
                End If
            End If
            
            'proc��de � l'�criture dans le fichier
            'bouge le pointeur
            Ret = SetFilePointerEx(lFile2, 0&, 0&, FILE_END) 'FILE_END ==> �crit � la fin du fichier
            '�criture dans le fichier
            Call WriteFile(lFile2, ByVal S, Len(S), Ret, ByVal 0&)
            
            If (a Mod 160017) = 0 Then
                'rend un peu la main
                dblAdv = Round((100 * a) / lLength, 3)
                frmContent.Sb.Panels(1).Text = "Status=[Saving " & Me.Caption & "]" & "[" & CStr(dblAdv) & " %]"
                DoEvents
            End If

        Next a
    
    'ferme le handle du fichier �crit
    Call CloseHandle(lFile2)
    
    'affiche le message de fin de sauvegarde
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_FileSaved"))
    
    Exit Function
ErrGestion:
    clsERREUR.AddError "Pfm.GetNewFile", True
End Function

'=======================================================
'fonction ayant uniquement pour but d'exister, on l'appelle � partir d'une
'autre fonction pour tester si frmcontent.activeform est
'une form d'�dition m�moire, edition de fichier ou de disque
'=======================================================
Public Function Useless() As String
    Useless = "Pfm"
End Function

'=======================================================
'changement des valeurs dans le FrameData
'=======================================================
Private Sub txtValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'change la valeur de l'item s�lectionn� du HW de frmcontent.activeform (i_tem)
Dim I_tem As ItemElement

    Set I_tem = HW.Item
    
    If KeyCode = 13 Then
        'alors appui sur "enter"
    
        Select Case Index
            Case 0
                'alors on change les autres champs que le champ "Hexa"
                With txtValue(0)
                    txtValue(1).Text = Hex2Dec(.Text)
                    txtValue(2).Text = Hex2Str(.Text)
                    txtValue(3).Text = Hex2Oct(.Text)
                End With
            Case 1
                'alors on change les autres champs que le champ "decimal"
                With txtValue(1)
                    txtValue(0).Text = Hex$(Val(.Text))
                    txtValue(2).Text = Byte2FormatedString(Val(.Text))
                    txtValue(3).Text = Oct$(Val(.Text))
                End With
            Case 2
                'alors on change les autres champs que le champ "ASCII"
                With txtValue(2)
                    txtValue(0).Text = Str2Hex(.Text)
                    txtValue(1).Text = Str2Dec(.Text)
                    txtValue(3).Text = Str2Oct(.Text)
                End With
            Case 3
                'alors on change les autres champs que le champ "octal"
                With txtValue(3)
                    txtValue(0).Text = Hex$(Oct2Dec(Val(.Text)))
                    txtValue(1).Text = Oct2Dec(Val(.Text))
                    txtValue(2).Text = Chr_(Oct2Dec(Val(.Text)))
                End With
        End Select

        With frmContent.ActiveForm.HW
            .AddHexValue I_tem.Line, I_tem.Col, txtValue(0).Text
            .AddOneStringValue I_tem.Line, I_tem.Col, txtValue(2).Text
            ModifyData
        End With
    End If
        
End Sub

'=======================================================
'des donn�es ont �t�s modifi�es ==> on sauvegarde ces changements
'=======================================================
Private Sub ModifyData()
Dim S As String
Dim X As Long
Dim I_tem As ItemElement

    Set I_tem = HW.Item

    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    'd�finit s (nouvelle string)
    S = vbNullString
    For X = 1 To 16
        S = S & Hex2Str_(frmContent.ActiveForm.HW.Value(I_tem.Line, X))
    Next X

    frmContent.ActiveForm.AddChange frmContent.ActiveForm.HW.FirstOffset + 16 * (I_tem.Line - 1), I_tem.Col, S
    
End Sub

'=======================================================
'ajout d 'un �l�ment � l'historique
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
'ajout d 'un �l�ment � l'historique
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
    DoEvents    '/!\ DO NOT REMOVE ! (permet d'effectuer les changements d'enabled et de ne pas pouvoir appuyer sur ctrl+Z quand on est au tout d�but)
    Call ModifyHistoEnabled 'v�rifie que c'est Ok pour les enabled
End Sub

'=======================================================
'ajout d 'un �l�ment � l'historique
'=======================================================
Public Sub RedoM()
    On Error Resume Next
    With lstHisto
        If .SelectedItem.Index = 1 And .ListItems.Item(1).Selected = False Then
            'rien de s�lectionn�
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
    Call ModifyHistoEnabled 'v�rifie que c'est Ok pour les enabled
End Sub

'=======================================================
'refresh simple du HW
'=======================================================
Public Sub RefreshHW()
    Call VS_Change(VS.Value)
End Sub

Private Sub VS_MouseDown(Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
'alors une action a �t� effectu�e (lance le popup menu)
    If Button = vbRightButton Then Me.PopupMenu frmContent.rmnuPos
End Sub

Private Sub VS_Scroll()
    DoEvents     '/!\ DO NOT REMOVE !
    Call VS_Change(VS.Value)
End Sub
