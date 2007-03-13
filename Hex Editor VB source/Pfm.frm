VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{276EF1C1-20F1-4D85-BE7B-06C736C9DCE9}#1.1#0"; "ExtendedVScrollbar_OCX.ocx"
Object = "{4C7ED4AA-BF37-4FCA-80A9-C4E4272ADA0B}#1.1#0"; "HexViewer_OCX.ocx"
Begin VB.Form Pfm 
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
   Icon            =   "Pfm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   7635
   Visible         =   0   'False
   Begin VB.Frame FrameData 
      Caption         =   "Valeur"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   600
      TabIndex        =   32
      Top             =   3000
      Width           =   1695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   50
         ScaleHeight     =   1095
         ScaleWidth      =   1605
         TabIndex        =   33
         Top             =   240
         Width           =   1600
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   3
            Left            =   960
            TabIndex        =   6
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   0
            Left            =   960
            MaxLength       =   2
            TabIndex        =   3
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   1
            Left            =   960
            MaxLength       =   3
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   2
            Left            =   960
            MaxLength       =   1
            TabIndex        =   5
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblValue 
            Caption         =   "Octal :"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   37
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Hexa :"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Decimal :"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "ASCII :"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2760
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame FrameIcon 
      Caption         =   "Icones"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   360
      TabIndex        =   29
      Top             =   4440
      Width           =   1695
      Begin ComctlLib.ListView lvIcon 
         Height          =   2535
         Left            =   65
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1550
         _ExtentX        =   2752
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
         BackColor       =   -2147483633
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.Frame FrameInfos 
      Caption         =   "Informations"
      ForeColor       =   &H00FF0000&
      Height          =   6975
      Left            =   3720
      TabIndex        =   8
      Top             =   240
      Width           =   3135
      Begin VB.PictureBox pctContain_cmdMAJ 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   120
         ScaleHeight     =   6615
         ScaleWidth      =   2955
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   2950
         Begin VB.TextBox txtFile 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   1
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   2
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   3
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   4
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   5
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   6
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   7
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   2160
            Width           =   2895
         End
         Begin VB.CommandButton cmdMAJ 
            Caption         =   "Mettre à jour"
            Height          =   255
            Left            =   600
            TabIndex        =   2
            ToolTipText     =   "Mettre à jour les informations"
            Top             =   6240
            Width           =   1695
         End
         Begin ComctlLib.ListView lstSignets 
            Height          =   1575
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   4560
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
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   4160
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
            Left            =   0
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   4560
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fichier"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   2895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Statistiques"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   26
            Top             =   2520
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Pages=[pages]"
            Height          =   200
            Index           =   8
            Left            =   0
            TabIndex        =   25
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Sélection=[selection]"
            Height          =   200
            Index           =   9
            Left            =   0
            TabIndex        =   24
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Offset=[offset]"
            Height          =   200
            Index           =   10
            Left            =   0
            TabIndex        =   23
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Offset Maximum=[offset max]"
            Height          =   200
            Index           =   11
            Left            =   0
            TabIndex        =   22
            Top             =   3600
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Historique=[nombre]"
            Height          =   200
            Index           =   12
            Left            =   0
            TabIndex        =   21
            Top             =   3840
            Width           =   2895
         End
      End
   End
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
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
   Begin ExtVS.ExtendedVScrollBar VS 
      Height          =   2895
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5106
      Min             =   0
      Value           =   0
      LargeChange     =   100
      SmallChange     =   100
   End
   Begin HexViewer_OCX.HexViewer HW 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4471
      strTag1         =   "0"
      strTag2         =   "0"
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   2520
      Top             =   5400
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
'FORM D'EDITION DU CONTENU D'UN FICHIER
'=======================================================

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private lBgAdress As Currency   'offset de départ de page
Private lEdAdress As Currency   'offset de fin de page
Private NumberPerPage As Long   'nombre de lignes visibles par Page
Private lLenght As Currency 'taille du fichier
Private ChangeListO() As Currency
Private ChangeListC() As Currency
Private ChangeListS() As String
Private ChangeListDim As Long
Private lFile As Long   'n° d'ouverture du fichier
Private bOkToOpen As Boolean
Private mouseUped As Boolean
Private bFirstChange As Boolean
Private bytFirstChange As Byte
Private bHasMaximized As Long
Private lngFormStyle As Long

Public cUndo As clsUndoItem 'infos générales sur 'historique
Private cHisto() As clsUndoSubItem  'historique pour le Undo/Redo

Public TheFile As clsFile

Public Sub cmdMAJ_Click()
'MAJ des infos
Dim lPages As Long
Dim cFic As clsFile

    On Error Resume Next
    
    'nom du fichier
    txtFile.Text = "[" & Me.Caption & "]"
    
    'récupère les infos sur le fichier
    Set cFic = cFile.GetFile(Me.Caption)
    
    'affiche tout çà
    TextBox(0).Text = "Taille=[" & CStr(cFic.FileSize) & " Octets  -  " & CStr(Round(cFic.FileSize / 1024, 3)) & " Ko" & "]"
    TextBox(1).Text = "Attribut=[" & CStr(cFic.FileAttributes) & "]"
    TextBox(2).Text = "Création=[" & cFic.CreationDate & "]"
    TextBox(3).Text = "Accès=[" & cFic.LastAccessDate & "]"
    TextBox(4).Text = "Modification=[" & cFic.LastModificationDate & "]"
    TextBox(5).Text = "Version=[" & cFic.EXEFileVersion & "]"
    TextBox(6).Text = "Description=[" & cFic.EXEFileDescription & "]"
    TextBox(7).Text = "Copyright=[" & cFic.EXELegalCopyright & "]"
   
    Label2(8).Caption = Me.Sb.Panels(2).Text
    Label2(9).Caption = "Sélection=[" & CStr(HW.NumberOfSelectedItems) & " bytes]"
    Label2(10).Caption = Me.Sb.Panels(3).Text
    Label2(11).Caption = "Offset maximum=[" & CStr(16 * Int(lLenght / 16)) & "]"
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

    UpdateWindow Me.hWnd    'refresh de la form
End Sub

Private Sub Form_Load()

    lngFormStyle = GetWindowLong(Me.hWnd, GWL_STYLE)

    'subclasse la form pour éviter de resizer trop
    #If USE_FORM_SUBCLASSING Then
        Call LoadResizing(Me.hWnd, 9000, 6000)
    #End If
    
    'subclasse également lvIcon pour éviter le drag & drop
    Call HookLVDragAndDrop(lvIcon.hWnd)
    
    'instancie la classe Undo
    Set cUndo = New clsUndoItem
    
    'affecte les valeurs générales (type) à l'historique
    cUndo.tEditType = edtFile
    Set cUndo.Frm = Me
    Set cUndo.lvHisto = Me.lstHisto
    ReDim cHisto(0)
    Set cHisto(0) = New clsUndoSubItem
    
    'affiche ou non les éléments en fonction des paramètres d'affichage de frmcontent
    Me.HW.Visible = frmContent.mnuTab.Checked
    Me.VS.Visible = frmContent.mnuTab.Checked
    Me.FrameData.Visible = frmContent.mnuEditTools.Checked
    Me.FrameInfos.Visible = frmContent.mnuInformations.Checked
    Me.FrameIcon.Visible = frmContent.mnuShowIcons.Checked
    
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
    
    bOkToOpen = False 'pas prêt à l'ouverture
    
    With cPref
        'en grand dans la MDIform
        If .general_MaximizeWhenOpen Then Me.WindowState = vbMaximized
    End With
        
    frmContent.Sb.Panels(1).Text = "Status=[Opening File]"
    frmContent.Sb.Refresh
    
    bFirstChange = False 'pas de KeyAscii déjà appuyé pour modifier une valeur
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'FileContent = vbNullString
    lNbChildFrm = lNbChildFrm - 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
    
    CloseHandle lFile 'ferme le handle sur le fichier
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'récupère le style de la form
    
    'If Me.WindowState = vbMaximized Then
        'alors si on est sous Vista, il faut supprimer la bordure de la fenêtre pour
        'éviter un bug d'affichage

        

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
    FrameInfos.Top = 0
    FrameInfos.Height = Me.Height - 700
    FrameInfos.Left = 20
    cmdMAJ.Top = FrameInfos.Height - 650
    lstHisto.Height = FrameInfos.Height - 5300
    lstSignets.Height = FrameInfos.Height - 5300
    Me.pctContain_cmdMAJ.Height = FrameInfos.Height - 350
    
    'met le Grid à la taille de la fenêtre
    HW.Width = 9620
    HW.Height = Me.Height - 400 - Sb.Height
    HW.Left = IIf(FrameInfos.Visible, FrameInfos.Width, 0) + 50
    HW.Top = 0
    
    'bouge le frameData
    FrameData.Top = 100
    FrameData.Left = IIf(HW.Visible, HW.Width + HW.Left, IIf(FrameInfos.Visible, FrameInfos.Width, 0)) + 500
    
    'bouge le frameIcon
    FrameIcon.Top = 1700
    FrameIcon.Left = IIf(HW.Visible, HW.Width + HW.Left, IIf(FrameInfos.Visible, FrameInfos.Width, 0)) + 500
    
    'calcule le nombre de lignes du Grid à afficher
    'NumberPerPage = Int(Me.Height / 250) - 1
    NumberPerPage = Int(HW.Height / 250) - 1
    
    HW.NumberPerPage = NumberPerPage
    HW.Refresh

   
    Call VS_Change(VS.Value)
    
    VS.Top = 0
    VS.Height = Me.Height - 430 - Sb.Height
    VS.Left = IIf(Me.Width < 13100, Me.Width - 350, HW.Left + HW.Width)
            
End Sub

'=======================================================
'permet de lancer le Resize depuis uen autre form
'=======================================================
Public Sub ResizeMe()
    Form_Resize
End Sub

'=======================================================
'affiche dans le HW les valeurs hexa qui correspondent à la partie
'du fichier qui est visualisée
'=======================================================
Private Sub OpenFile(ByVal lBg As Currency, ByVal lEd As Currency)
Dim tmpText As String
Dim a As Long
Dim s As String
Dim b As Long
Dim c As Long
Dim s2 As String
Dim lLenght As Long
Dim e As Byte
Dim sTemp() As String
Dim lOff As Currency
Dim lPlace As Long
Dim Ret As Long
    
    'On Error GoTo ErrGestion
    
    If bOkToOpen = False Then Exit Sub  'pas prêt à ouvrir
    
    HW.ChangeValues 'permet d'empêcher de voir des valeurs hexa vers la fin du fichier
    
    'créé un buffer qui contiendra les valeurs
    tmpText = String$(By16(lEd - lBg) + 16, 0)

    'bouge le pointeur sur lr fichier au bon emplacement
    Ret = SetFilePointerEx(lFile, (lBg - 1) / 10000, 0&, FILE_BEGIN)    'divise par 10000 pour
    'pouvoir renvoyer une currency DECIMALE
    
    'prend un morceau du fichier
    'tmpText = String$(16, 0)  'buffer
    Ret = ReadFile(lFile, ByVal tmpText, Len(tmpText), Ret, ByVal 0&)
    
    'obtient le texte à partir du fichier (depuis l'octet lBg, et pour une intervalle contenant tout les bytes affichés à l'écran
    'Get lFile, lBg, tmpText
        
    ReDim sTemp(HW.NumberPerPage)
    
    'divise tmpText en parties de 16 bytes et stocke dans sTemp
    For e = 1 To HW.NumberPerPage

        'stocke la VRAIE string de 16
        sTemp(e) = Mid$(tmpText, e * 16 - 15, 16)
        
        s = vbNullString
        For a = 1 To 16
        
            'calcule l'offset équivalent
            lOff = (e - 1) * 16 + VS.Value * 16
            
            If IsOffsetModified(lOff, lPlace) Then
                'l'offset est modifié

                'la string existe dans la liste des string modifiées
                s = ChangeListS(lPlace)   'ne formate PAS la string
                
                'ajoute les valeurs hexa à partir de la VRAIE NOUVELLE string
                For c = 1 To 16
                    s2 = Hex$(Asc(Mid$(s, c, 1)))
                    If Len(s2) = 1 Then s2 = "0" & s2
                    HW.AddHexValue e, c, s2, IsModified(c, lOff)
                Next c
                
                s = Formated16String(s) 'formate la string POUR L'AFFICHAGE
                Exit For
            Else

                'formate une string de 16 pour l'affichage
                s = s & Byte2FormatedString(Asc(Mid$(sTemp(e), a, 1)))
                
                'ajoute les valeurs hexa à partir de la VRAIE string
                s2 = Hex$(Asc(Mid$(sTemp(e), a, 1)))
                If Len(s2) = 1 Then s2 = "0" & s2
                HW.AddHexValue e, a, s2, False
            End If
            
        Next a
        
        'ajoute la string FORMATEE pour l'affichage
        HW.AddStringValue e, s
        
    Next e

    'Done.
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
    
    'HW.Refresh  'affiche les résultats
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.OpenFile", True
End Sub

'=======================================================
'renvoie si l'offset contient une modification
'=======================================================
Public Function IsOffsetModified(ByVal lOffset As Currency, ByRef lPlace As Long) As Boolean
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
'efface les élements de la sélection
'=======================================================
Public Sub DeleteZone()
Dim tempFile As String
Dim x As Long, s As String
Dim tmpText As String
Dim y As Long
Dim a As Long
Dim lNewPos As Long
Dim e As Long
Dim lLen As Long, lFile2 As Long
Dim lPlace As Long
Dim tempFileLen As Long
Dim OfF As Currency, OfL As Currency
    
    On Error GoTo ErrGestion
    
    'obtient un path temporaire
    ObtainTempPathFile vbNullString, tempFile, cFile.GetFileExtension(Me.Caption)
    
    '//suppression d'une zone
    'créé le fichier normalement jusqu'à la zone à enlever
    'affiche le message d'attente
    frmContent.Sb.Panels(1).Text = "Status=[Creating backup " & Me.Caption & "]"
    
    'détermine les offsets de délimitation
    OfF = Me.HW.FirstSelectionItem.Offset + Me.HW.FirstSelectionItem.Col
    OfL = Me.HW.SecondSelectionItem.Offset + Me.HW.SecondSelectionItem.Col
    
    lFile2 = FreeFile 'obtient une ouverture dispo
    Open tempFile For Binary Access Write As lFile2   'ouvre le fichier sFile2 pour l'enregistrement
       
    'créé un buffer de longueur divisible par 16
    'recoupera à la fin pour la longueur exacte du fichier
    lLen = By16(lLenght)
    
    tempFileLen = lLenght   'contient la taille que fera le fichier temporaire (après suppression)
    'pour l'instant c'est la taille normale
    
    lNewPos = 0 'pas de décalage pour le moment
    
    tmpText = String$(16, 0)
       
        For a = 1 To lLen Step 16  'remplit par intervalles de 16
                    
            Get lFile, a, tmpText   'prend un morceau du fichier source (16bytes) à partir de l'octet A

            If Not (((a - 1) >= 16 * Int(OfF / 16)) And ((a - 1) <= By16(OfL))) Then
                'ALORS n'appartient pas à la zone enlevée
                If IsOffsetModified(a - 1, lPlace) Then
                    'l'offset est modifié
    
                    'la string existe dans la liste des string modifiées
                    'ne formate PAS la string
                    
                    'vérifie que l'on est pas dans la dernière ligne, si oui, ne prend que la longueur nécessaire
                    If a + 16 > tempFileLen Then
                        s = Mid$(ChangeListS(lPlace), 1, tempFileLen - a + 1)
                    Else
                        s = ChangeListS(lPlace)
                    End If
                Else
                    
                    'pas de modif, prend à partir du fichier source
                    'vérfie que l'on est pas dans la dernière ligne
                    
                    If a + 16 > tempFileLen Then
                        s = Mid$(tmpText, 1, lLenght - a + 1)
                    Else
                        s = tmpText
                    End If
                End If
                
                Put lFile2, a + lNewPos, s  'ajoute au fichier résultant
                
            Else
                'ALORS appartient à la zone enlevée
                lNewPos = -OfL + OfF    'décalage
                tempFileLen = lLenght - OfL + OfF  'change la taille du fichier temp
            End If
            
            
            If (a Mod 160017) = 0 Then
                'rend un peu la main
                frmContent.Sb.Panels(1).Text = "Status=[Creating backup " & Me.Caption & "]" & "[" & Round(100 * a / lLenght, 2) & " %]"
                DoEvents
            End If

        Next a
    
    Close lFile 'ferme l'ANCIEN FICHIER
    lFile = lFile2  'change le numéro d'enregistrement ==> fichier temp
    
    Me.Caption = tempFile 'le nouveau nom de fichier
    
    'affiche le message de fin de sauvegarde
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.DeleteZone", True
End Sub

'=======================================================
'renvoie si la case a été modifiée ou non
'=======================================================
Private Function IsModified(ByVal lCol As Long, ByVal lOffset As Currency) As Boolean
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
'obtient le nom du fichier à ouvrir, et l'ouvre
'=======================================================
Public Sub GetFile(ByVal sFile As String)
Dim l As Currency

    On Error GoTo ErrGestion

    'récupère les infos fichier
    Set TheFile = cFile.GetFile(sFile)
    
    'récupère le contenu du fichier
    lLenght = TheFile.FileSize
    
    HW.MaxOffset = lLenght 'offset maximal
    
    Me.Caption = sFile
    
    'ajoute l'icone à notre form
    Set Me.Icon = CreateIcon(sFile)
    
    'obtient un handle vers le fichier
    lFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    If NumberPerPage = 0 Then Call Form_Resize  'besoin de cette valeur
        
    'règle la taille de VS
    VS.Min = 0
    l = Int(lLenght / 16)
    VS.Max = l
    VS.Value = 0
    VS.SmallChange = 1
    VS.LargeChange = NumberPerPage - 1
    
    'stocke dans les tag les valeurs Max et Min des offsets
    HW.curTag1 = HW.FirstOffset
    HW.curTag2 = HW.MaxOffset
    
    bOkToOpen = True 'prêt à l'ouverture

    'affichage
    OpenFile 16 * VS.Value + 1, VS.Value * 16 + NumberPerPage * 16
    
    'affiche aussi les icones du fichier
    LoadIconesToLV sFile, lvIcon, Me.pct, Me.IMG

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.GetFile", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cUndo = Nothing
    #If USE_FORM_SUBCLASSING Then
        'alors enlève le subclassing
        Call RestoreResizing(Me.hWnd)
    #End If
    
    'enleve le hook sur lvIcon également
    Call UnHookLVDragAndDrop(lvIcon.hWnd)
End Sub

Private Sub HW_GotFocus()
    HW.Refresh
End Sub

Private Sub HW_KeyDown(KeyCode As Integer, Shift As Integer)
'gère les touches qui changent le VS, gère le changement de valeur
    
    On Error GoTo ErrGestion
    
    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    If KeyCode = vbKeyUp Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors monte
        If HW.FirstOffset = 0 And HW.Item.Line = 1 Then Exit Sub  'tout au début déjà
        'on remonte d'une ligne alors
        HW.Item.Line = HW.Item.Line - 1
        If HW.Item.Line = 0 Then
            'alors on remonte le firstoffset
            HW.Item.Line = 1
            HW.FirstOffset = HW.FirstOffset - 16
            VS.Value = VS.Value - 1
            Call VS_Change(VS.Value)
        End If
        HW.ColorItem tHex, HW.Item.Line, HW.Item.Col, HW.Value(HW.Item.Line, HW.Item.Col), HW.SelectionColor, True
        HW.AddSelection HW.Item.Line, HW.Item.Col
    End If
    
    If KeyCode = vbKeyDown Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors descend
        If HW.FirstOffset + HW.Item.Line * 16 - 16 = By16(HW.MaxOffset) Then Exit Sub  'tout en bas
        'on descend d'une ligne alors
        HW.Item.Line = HW.Item.Line + 1
        If HW.Item.Line = HW.NumberPerPage Then
            'alors on descend le firstoffset
            HW.Item.Line = HW.NumberPerPage - 1
            HW.FirstOffset = HW.FirstOffset + 16
            VS.Value = VS.Value + 1
            Call VS_Change(VS.Value)
        End If
        'change le VS
        HW.ColorItem tHex, HW.Item.Line, HW.Item.Col, HW.Value(HW.Item.Line, HW.Item.Col), HW.SelectionColor, True
        HW.AddSelection HW.Item.Line, HW.Item.Col
   End If
    
    If KeyCode = vbKeyEnd Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors aller tout à la fin
        VS.Value = VS.Max
        Call VS_Change(VS.Value)
    End If
    If KeyCode = vbKeyHome Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors tout au début
        VS.Value = VS.Min
        Call VS_Change(VS.Value)
    End If
    If KeyCode = vbKeyPageUp Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors monter de NumberPerPage
        VS.Value = IIf((VS.Value - NumberPerPage) > VS.Min, VS.Value - NumberPerPage, VS.Min)
        Call VS_Change(VS.Value)
    End If
    If KeyCode = vbKeyPageDown Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors descendre de NumberPerPage
        VS.Value = IIf((VS.Value + NumberPerPage) < VS.Max, VS.Value + NumberPerPage, VS.Max)
        Call VS_Change(VS.Value)
    End If
    
    If KeyCode = vbKeyLeft Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors va à gauche
        If HW.FirstOffset = 0 And HW.Item.Col = 1 And HW.Item.Line = 1 Then Exit Sub 'tout au début déjà
        If HW.Item.Col = 1 Then
            'tout à gauche ==> on remonte d'une ligne alors
            HW.Item.Col = 16: HW.Item.Line = HW.Item.Line - 1
            If HW.Item.Line = 0 Then
                'alors on remonte le firstoffset
                HW.Item.Line = 1
                HW.FirstOffset = HW.FirstOffset - 16
                VS.Value = VS.Value - 1
                Call VS_Change(VS.Value)
            End If
        Else
            'va à gauche
            HW.Item.Col = HW.Item.Col - 1
        End If
        HW.ColorItem tHex, HW.Item.Line, HW.Item.Col, HW.Value(HW.Item.Line, HW.Item.Col), HW.SelectionColor, True
        HW.AddSelection HW.Item.Line, HW.Item.Col
    End If
         
    If KeyCode = vbKeyRight Then
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
        'alors va à droite
        If HW.FirstOffset + HW.Item.Line * 16 - 16 = By16(HW.MaxOffset) And HW.Item.Col = 16 Then Exit Sub  'tout à la fin déjà
        If HW.Item.Col = 16 Then
            'tout à droite ==> on descend d'une ligne alors
            HW.Item.Col = 1: HW.Item.Line = HW.Item.Line + 1
            If HW.Item.Line = HW.NumberPerPage Then
                'alors on descend le firstoffset
                HW.Item.Line = HW.NumberPerPage - 1
                HW.FirstOffset = HW.FirstOffset + 16
                VS.Value = VS.Value + 1
                Call VS_Change(VS.Value)
            End If
        Else
            'va à droite
            HW.Item.Col = HW.Item.Col + 1
        End If
        'change le VS
        HW.ColorItem tHex, HW.Item.Line, HW.Item.Col, HW.Value(HW.Item.Line, HW.Item.Col), HW.SelectionColor, True
        HW.AddSelection HW.Item.Line, HW.Item.Col
    End If
    
    'réenregistre le numéro de l'offset actuel dans hw.item
    HW.Item.Offset = HW.Item.Line * 16 - 16
    'affecte les autres valeurs dans Item
    'HW.Item.tType = tHex
    HW.Item.Value = HW.Value(HW.Item.Line, HW.Item.Col)
    
    DoEvents

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.KeyDown", True
End Sub

Private Sub HW_KeyPress(KeyAscii As Integer)
'change les valeurs dans le tableau
Dim s As String
Dim sKey As Long
Dim bytHex As Byte
Dim Valu As Byte
Dim x As Byte
Dim s2 As String

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

            's est la nouvelle string
            s = vbNullString
            For x = 1 To 16
                s = s & Hex2Str_(HW.Value(HW.Item.Line, CLng(x)))
            Next x
            s2 = s
            
            'calcule la nouvelle string (partie de gauche ancienne + nouveau byte + partie de droite ancienne)
            s = Mid$(s, 1, HW.Item.Col - 1) & Chr$(bytHex) & Mid$(s, HW.Item.Col + 1, 16 - HW.Item.Col)   'avant & nouvelle & après
            
            'applique le changement
            Call Me.AddChange((HW.Item.Line - 1) * 16 + HW.FirstOffset, HW.Item.Col, s)
            'ajoute l'historique
            Me.AddHistoFrm actByteWritten, s2, s, (HW.Item.Line - 1) * 16 + HW.FirstOffset, , HW.Item.Col
           'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
            
        End If
    ElseIf HW.Item.tType = tString Then
        'alors voici la zone STRING
        'on ne tappe QU'UNE SEULE VALEUR
        
            'le nouveau byte est donc désormais KeyAscii

            's est la nouvelle string
            s = vbNullString
            For x = 1 To 16
                s = s & Hex2Str_(HW.Value(HW.Item.Line, CLng(x)))
            Next x
            s2 = s
            
            'calcule la nouvelle string (partie de gauche ancienne + nouveau byte + partie de droite ancienne)
            s = Mid$(s, 1, HW.Item.Col - 1) & Chr$(KeyAscii) & Mid$(s, HW.Item.Col + 1, 16 - HW.Item.Col)   'avant & nouvelle & après
            
            'applique le changement
            Call Me.AddChange((HW.Item.Line - 1) * 16 + HW.FirstOffset, HW.Item.Col, s)
            'ajoute l'historique
            Me.AddHistoFrm actByteWritten, s2, s, (HW.Item.Line - 1) * 16 + HW.FirstOffset, , HW.Item.Col
            'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
            
    End If

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.KeyPress", True
End Sub

Public Sub HW_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, Item As ItemElement)
Dim s As String
Dim r As Long

    'popup menu
    If Button = 2 Then
        frmContent.mnuDeleteSelection.Enabled = True
        frmContent.mnuCut.Enabled = True
        Me.PopupMenu frmContent.rmnuEdit ', X + GD.Left, Y + GD.Top
    End If
    Me.Sb.Panels(3).Text = "Offset=[" & CStr(Item.Line * 16 + HW.FirstOffset - 16) & "]"
    Label2(10).Caption = Me.Sb.Panels(3).Text
    
    
    If Button = 1 Then
        'alors on a sélectionné un objet
        
        bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
        'pour l'édition dynamique au clavier
        
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
            txtValue(0).Text = Item.Value
            txtValue(1).Text = Hex2Dec(Item.Value)
            txtValue(2).Text = Hex2Str(Item.Value)
            txtValue(3).Text = Hex2Oct(Item.Value)
            FrameData.Caption = "Data=[" & Item.Value & "]"
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
            HW.TraceSignets
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
            
            Refresh
        End If
    ElseIf Button = 4 And Shift = 2 Then
        'click molette + control
        'sélectionne une zone définie
        frmSelect.GetEditFunction 0 'selection mode
        frmSelect.Show vbModal
    End If
    
End Sub

Private Sub HW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Sb.Panels(4).Text = "Sélection=[" & CStr(HW.NumberOfSelectedItems) & " bytes]"
    Label2(9) = Me.Sb.Panels(4).Text
End Sub

Private Sub HW_MouseWheel(ByVal lSens As Long)

    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    bFirstChange = False 'ALORS IL FAUDRA RETAPER LA PREMIERE PARTIE DE LA STRING HEXA
    'pour l'édition dynamique au clavier
        
    If lSens > 0 Then
        'alors on descend
        VS.Value = IIf((VS.Value - 3) >= 0, VS.Value - 3, 0)
        Call VS_Change(VS.Value)
    Else
        'alors on monte
        VS.Value = IIf((VS.Value + 3) <= VS.Max, VS.Value + 3, VS.Max)
        Call VS_Change(VS.Value)
    End If
End Sub

Private Sub HW_UserMakeFirstOffsetChangeByMovingMouse()
    VS.Value = HW.FirstOffset / 16
End Sub

Private Sub lstHisto_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'on supprime l'historique si button=2
'on va à l'élément si button=4
Dim Item As ListItem
        
    If Button = 2 Then
        'récupère l'item
        Set Item = lstHisto.HitTest(x, y)
        If Item Is Nothing Then Exit Sub
        
        'supprime
        DelHisto Val(Item.SubItems(1)), cHisto(), cUndo
    ElseIf Button = 4 Then
        'récupère l'item
        Set Item = lstHisto.HitTest(x, y)
        If Item Is Nothing Then Exit Sub
        
        'active l'élément
        cUndo.lRang = Item.Index
        Item.Selected = True
        RedoMe cUndo, cHisto()
        DoEvents    '/!\ DO NOT REMOVE !
        Call ModifyHistoEnabled 'vérifie que c'est Ok pour les enabled
    End If
    
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

Private Sub lstSignets_KeyDown(KeyCode As Integer, Shift As Integer)
'vire les signets si touche suppr
Dim r As Long

    mouseUped = True
    
    If KeyCode = vbKeyDelete Then
        'touche suppr
        If lstSignets.SelectedItem.Selected Then
            'alors on supprime quelque chose
            r = MsgBox("Supprimer les signets ?", vbInformation + vbYesNo, "Attention")
            If r <> vbYes Then Exit Sub
        
            For r = lstSignets.ListItems.Count To 1 Step -1
                If lstSignets.ListItems.Item(r).Selected Then lstSignets.ListItems.Remove r
            Next r
        End If
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
        s = InputBox("Ajouter un commentaire pour le signet " & tLst.Text, "Ajout d'un commentaire")
        If StrPtr(s) <> 0 Then
            'ajoute le commentaire
            tLst.SubItems(1) = s
        End If
    End If
    
    If Button = 4 Then
        'mouse du milieu ==> on supprime le signet
        Set tLst = lstSignets.HitTest(x, y)
        If tLst Is Nothing Then Exit Sub
        
        r = MsgBox("Supprimer le signet " & tLst.Text & " ?", vbInformation + vbYesNo, "Attention")
        If r <> vbYes Then Exit Sub
        
        'on supprime
        HW.RemoveSignet Val(tLst.Text)
        
        'on enlève du listview
        lstSignets.ListItems.Remove tLst.Index
    End If
        
End Sub

Private Sub lstSignets_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'permet de ne pas changer le HW dans le cas de multiples sélections
    mouseUped = True
End Sub

Private Sub lvIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
    If KeyAscii = Asc(vbCr) Then KeyAscii = 0   'empêche le 'BEEP'
End Sub

'=======================================================
'change la valeur du VS
'public, car cette sub est aussi appelée pour le refresh
'=======================================================
Public Sub VS_Change(Value As Currency)
Dim lPages As Long

    'On Error GoTo ErrGestion
    
    If NumberPerPage = 0 Then Exit Sub
    
    'réaffiche la Grid
    OpenFile 16 * Value + 1, VS.Value * 16 + NumberPerPage * 16
    
    'calcule le nbre de pages
    lPages = lLenght / (NumberPerPage * 16) + 1
    Me.Sb.Panels(2).Text = "Page=[" & CStr(1 + Int(VS.Value / NumberPerPage)) & "/" & CStr(lPages) & "]"
    Label2(8).Caption = Me.Sb.Panels(2).Text
    
    HW.FirstOffset = VS.Value * 16
    
    HW.Refresh

    Exit Sub
ErrGestion:
    clsERREUR.AddError "Pfm.VS_Change", True
End Sub

'=======================================================
'ajoute une valeur changée dans la liste des valeurs changées
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
'procède à la sauvegarde du fichier avec changements à l'emplacement sFile2
'=======================================================
Public Function GetNewFile(ByVal sFile2 As String) As String
Dim x As Long, s As String
Dim tmpText As String
Dim dblAdv As Double
Dim y As Long
Dim a As Long
Dim e As Long
Dim lLen As Long, lFile2 As Long
Dim lPlace As Long
Dim Ret As Long

    On Error GoTo ErrGestion

    'affiche le message d'attente
    frmContent.Sb.Panels(1).Text = "Status=[Saving " & Me.Caption & "]"
       
    'obtient un handle vers le fichier à écrire
    'ouverture en ECRITURE, avec overwrite si déjà existant (car déjà demandé confirmation avant)
    lFile2 = CreateFile(sFile2, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0)

    'créé un buffer de longueur divisible par 16
    'recoupera à la fin pour la longueur exacte du fichier
    lLen = By16(lLenght)
    tmpText = String$(16, 0)
       
        For a = 1 To lLen Step 16  'remplit par intervalles de 16
            
            'bouge le pointeur sur lr fichier au bon emplacement
            Ret = SetFilePointerEx(lFile, (a - 1) / 10000, 0&, FILE_BEGIN)
            'a divisé par 10^4 pour obtenir un nombre décimal de Currency
            
            'prend un morceau du fichier
            tmpText = String$(16, 0)  'buffer
            Ret = ReadFile(lFile, ByVal tmpText, 16, Ret, ByVal 0&)
       
            'Get lFile, a, tmpText   'prend un morceau du fichier source (16bytes) à partir de l'octet A
                      
            If IsOffsetModified(a - 1, lPlace) Then
                'l'offset est modifié

                'la string existe dans la liste des string modifiées
                'ne formate PAS la string
                
                'vérifie que l'on est pas dans la dernière ligne, si oui, ne prend que la longueur nécessaire
                If a + 16 > lLenght Then
                    s = Mid$(ChangeListS(lPlace), 1, lLenght - a + 1)
                Else
                    s = ChangeListS(lPlace)
                End If
            Else
                
                'pas de modif, prend à partir du fichier source
                'vérfie que l'on est pas dans la dernière ligne
                
                If a + 16 > lLenght Then
                    s = Mid$(tmpText, 1, lLenght - a + 1)
                Else
                    s = tmpText
                End If
            End If
            
            'procéède à l'écriture dans le fichier
            'bouge le pointeur
            Ret = SetFilePointerEx(lFile2, 0&, 0&, FILE_END) 'FILE_END ==> écrit à la fin du fichier
            'écriture dans le fichier
            WriteFile lFile2, ByVal s, Len(s), Ret, ByVal 0&
            
            If (a Mod 160017) = 0 Then
                'rend un peu la main
                dblAdv = Round((100 * a) / lLenght, 3)
                frmContent.Sb.Panels(1).Text = "Status=[Saving " & Me.Caption & "]" & "[" & CStr(dblAdv) & " %]"
                DoEvents
            End If

        Next a
    
    'ferme le handle du fichier écrit
    CloseHandle lFile2
    
    'affiche le message de fin de sauvegarde
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"

    Exit Function
ErrGestion:
    clsERREUR.AddError "Pfm.GetNewFile", True
End Function

'=======================================================
'fonction ayant uniquement pour but d'exister, on l'appelle à partir d'une
'autre fonction pour tester si frmcontent.activeform est
'une form d'édition mémoire, edition de fichier ou de disque
'=======================================================
Public Function Useless() As String
    Useless = "Pfm"
End Function

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
                txtValue(2).Text = Chr$(Oct2Dec(Val(txtValue(3).Text)))
        End Select

        With frmContent.ActiveForm.HW
            .AddHexValue I_tem.Line, I_tem.Col, txtValue(0).Text
            .AddOneStringValue I_tem.Line, I_tem.Col, txtValue(2).Text
            ModifyData
        End With
    End If
        
End Sub

'=======================================================
'des données ont étés modifiées ==> on sauvegarde ces changements
'=======================================================
Private Sub ModifyData()
Dim s As String
Dim x As Long
Dim I_tem As ItemElement

    Set I_tem = HW.Item

    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    'définit s (nouvelle string)
    s = vbNullString
    For x = 1 To 16
        s = s & Hex2Str_(frmContent.ActiveForm.HW.Value(I_tem.Line, x))
    Next x

    frmContent.ActiveForm.AddChange frmContent.ActiveForm.HW.FirstOffset + 16 * (I_tem.Line - 1), I_tem.Col, s
    
End Sub

'=======================================================
'ajout d 'un élément à l'historique
'=======================================================
Public Sub AddHistoFrm(ByVal tUndo As UNDO_TYPE, Optional ByVal sData1 As String, _
    Optional ByVal sData2 As String, Optional ByVal curData1 As Currency, _
    Optional ByVal curData2 As Currency, Optional ByVal bytData1 As Byte, _
    Optional ByVal bytData2 As Byte, Optional ByVal lngData1 As Long)
    
    AddHisto -1, cUndo, cHisto(), tUndo, sData1, sData2, curData1, curData2, bytData1, bytData2, lngData1
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
    UndoMe cUndo, cHisto()
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
    RedoMe cUndo, cHisto()
    DoEvents    '/!\ DO NOT REMOVE !
    Call ModifyHistoEnabled 'vérifie que c'est Ok pour les enabled
End Sub

'=======================================================
'refresh simple du HW
'=======================================================
Public Sub RefreshHW()
    Call VS_Change(VS.Value)
End Sub

Private Sub VS_MouseAction(ByVal lngMouseAction As ExtVS.MOUSE_ACTION)
'alors une action a été effectuée (lance le popup menu)
    If lngMouseAction = RIGHT_CLICK Then
        Me.PopupMenu frmContent.rmnuPos
    End If
End Sub
