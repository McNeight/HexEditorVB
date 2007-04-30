VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C9771C4C-85A3-44E9-A790-1B18202DA173}#1.0#0"; "FileView_OCX.ocx"
Object = "{67F3B6F5-143C-4724-BF0B-20B81F5D8E04}#1.0#0"; "ExtendedVScrollbar_OCX.ocx"
Object = "{C60799F1-7AA3-45BA-AFBF-5BEAB08BC66C}#1.0#0"; "HexViewer_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.Form diskPfm 
   Caption         =   "Ouverture d'un disque..."
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "diskPfm.frx":0000
   LinkTopic       =   "diskPfm"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   10155
   Visible         =   0   'False
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   8160
      Width           =   10155
      _ExtentX        =   17912
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
   Begin VB.TextBox pctPath 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1575
   End
   Begin HexViewer_OCX.HexViewer HW 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4260
      strTag1         =   "0"
      strTag2         =   "0"
   End
   Begin ExtVS.ExtendedVScrollBar VS 
      Height          =   3615
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   6376
      Min             =   0
      Value           =   0
      LargeChange     =   100
      SmallChange     =   100
   End
   Begin VB.Frame FrameFrag 
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   2520
      TabIndex        =   29
      Top             =   4920
      Width           =   2055
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2175
         Index           =   0
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   1815
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton cmdNext 
            Height          =   615
            Left            =   1320
            Picture         =   "diskPfm.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton cmdPrev 
            Height          =   615
            Left            =   1320
            Picture         =   "diskPfm.frx":1F0C
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   660
            Width           =   495
         End
         Begin VB.ListBox lstFrag 
            Height          =   1620
            ItemData        =   "diskPfm.frx":388E
            Left            =   0
            List            =   "diskPfm.frx":3890
            TabIndex        =   9
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblFrag 
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   1815
         End
      End
   End
   Begin VB.Frame FrameInfo2 
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   7080
      TabIndex        =   44
      Top             =   2520
      Width           =   2175
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   50
         ScaleHeight     =   5055
         ScaleWidth      =   2085
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   2085
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   22
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "Fichier=[path]"
            Top             =   3840
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   21
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "Fichier=[path]"
            Top             =   3600
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   20
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "Taille=[taille]"
            Top             =   3360
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   27
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   57
            Text            =   "Taille=[taille]"
            Top             =   2760
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   26
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "Fichier=[path]"
            Top             =   3000
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   19
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "Fichier=[path]"
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   18
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "Fichier=[path]"
            Top             =   1560
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   17
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "Taille=[taille]"
            Top             =   0
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   16
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   52
            Text            =   "Fichier=[path]"
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   7
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "Fichier=[path]"
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   6
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "Fichier=[path]"
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   5
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "Fichier=[path]"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   4
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "Fichier=[path]"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   3
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "Fichier=[path]"
            Top             =   2400
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   2
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "Fichier=[path]"
            Top             =   2160
            Width           =   2895
         End
      End
   End
   Begin FileView_OCX.FileView FV 
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      ShowEntirePath  =   0   'False
      AllowMultiSelect=   0   'False
      HideColumnHeaders=   -1  'True
      ShowFiles       =   0   'False
      ShowDrives      =   0   'False
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
   Begin VB.Frame FrameData 
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   7200
      TabIndex        =   23
      Top             =   840
      Width           =   1695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   50
         ScaleHeight     =   1095
         ScaleWidth      =   1605
         TabIndex        =   24
         TabStop         =   0   'False
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
            TabIndex        =   28
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Hexa :"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Decimal :"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "ASCII :"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.Frame FrameInfos 
      ForeColor       =   &H00FF0000&
      Height          =   6975
      Left            =   3720
      TabIndex        =   13
      Top             =   240
      Width           =   3135
      Begin ComctlLib.ListView lstHisto 
         Height          =   1575
         Left            =   120
         TabIndex        =   62
         TabStop         =   0   'False
         Tag             =   "lang_ok"
         Top             =   4800
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
      Begin VB.PictureBox pctContain_cmdMAJ 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   120
         ScaleHeight     =   6615
         ScaleWidth      =   2955
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   2950
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   1
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   2400
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   13
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   12
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   11
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   960
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   10
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   9
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   480
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   8
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "Taille=[taille]"
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   14
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.TextBox TextBox 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   15
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Text            =   "Fichier=[path]"
            Top             =   1920
            Width           =   2895
         End
         Begin VB.CommandButton cmdMAJ 
            Height          =   255
            Left            =   600
            TabIndex        =   2
            Top             =   6240
            Width           =   1695
         End
         Begin ComctlLib.ListView lstSignets 
            Height          =   1575
            Left            =   0
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "lang_ok"
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
            TabIndex        =   16
            TabStop         =   0   'False
            Tag             =   "lang_ok"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Disque"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   43
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
            TabIndex        =   22
            Top             =   2640
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Pages=[pages]"
            Height          =   200
            Index           =   8
            Left            =   0
            TabIndex        =   21
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Sélection=[selection]"
            Height          =   200
            Index           =   9
            Left            =   0
            TabIndex        =   20
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Offset=[offset]"
            Height          =   200
            Index           =   10
            Left            =   0
            TabIndex        =   19
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Offset Maximum=[offset max]"
            Height          =   200
            Index           =   11
            Left            =   0
            TabIndex        =   18
            Top             =   3600
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Historique=[nombre]"
            Height          =   200
            Index           =   12
            Left            =   0
            TabIndex        =   17
            Top             =   3840
            Width           =   2895
         End
      End
   End
   Begin FileView_OCX.FileView FV2 
      Height          =   1575
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      ShowEntirePath  =   0   'False
      AllowMultiSelect=   0   'False
      ShowDirectories =   0   'False
      ShowDrives      =   0   'False
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
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Label lblGOTO 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Aller à ..."
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   7320
      Width           =   855
   End
End
Attribute VB_Name = "diskPfm"
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
'FORM D'OUVERTURE D'UN DISQUE EN ACCES DIRECT
'=======================================================

'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private lBgAdress As Long   'offset de départ de page
Private lEdAdress As Long   'offset de fin de page
Private NumberPerPage As Long   'nombre de lignes visibles par Page
Private pRs As Long, pr As Long, pc As Long, pCs As Long 'sauvegarde de la sélection
Private lLength As Currency 'taille du fichier
Private lBytesPerSector As Long
Private ChangeListO() As Long
Private ChangeListC() As Long
Private ChangeListS() As String
Private ChangeListDim As Long
Private lFile As Long   'n° d'ouverture du fichier
Private bOkToOpen As Boolean
Private strDrive As String
Private cDrive As FileSystemLibrary.Drive
Private mouseUped As Boolean
Private bFirstChange As Boolean
Private bytFirstChange As Byte
Private hDrive As Long  'handle du disque ouvert
Private N0Clust As String
Private N0LogSec As String
Private N0PhysSec As String

Public cUndo As clsUndoItem 'infos générales sur 'historique
Private cHisto() As clsUndoSubItem  'historique pour le Undo/Redo


Private Sub cmdMAJ_Click()
'MAJ
Dim lPages As Long

    On Error Resume Next
    
    Label2(8).Caption = Me.Sb.Panels(2).Text
    Label2(9).Caption = Lang.GetString("_Selec") & _
        CStr(HW.NumberOfSelectedItems) & " " & Lang.GetString("_Bytes")
    Label2(10).Caption = Me.Sb.Panels(3).Text
    Label2(11).Caption = Lang.GetString("_OffMax") & _
        CStr(16 * Int(lLength / 16)) & "]"
    'Label2(12).Caption = "[" & sDescription & "]"

End Sub

Private Sub cmdNext_Click()
'change le cluster vu

    On Local Error Resume Next
    
    If lstFrag.ListCount = 0 Then Exit Sub
    If lstFrag.ListIndex = lstFrag.ListCount - 1 Then
        lstFrag.Selected(1) = True
    Else
        lstFrag.Selected(lstFrag.ListIndex + 1) = True
    End If
End Sub

Private Sub cmdPrev_Click()
'change le cluster vu

    On Local Error Resume Next

    If lstFrag.ListCount = 0 Then Exit Sub
    If lstFrag.ListIndex = 0 Then
        lstFrag.Selected(lstFrag.ListCount - 1) = True
    Else
        lstFrag.Selected(lstFrag.ListIndex - 1) = True
    End If
End Sub

Private Sub Form_Activate()
'    cmdMAJ_Click
    If HW.Visible Then HW.SetFocus
    Call VS_Change(VS.Value)
    ReDim ChangeListO(1) As Long
    ReDim ChangeListC(1) As Long
    ReDim ChangeListS(1) As String
 '   ChangeListDim = 1
    
    bOkToOpen = False 'pas prêt à l'ouverture
    
    Call UpdateWindow(Me.hWnd)      'refresh de la form
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cUndo = Nothing
    #If USE_FORM_SUBCLASSING Then
        'alors enlève le subclassing
        Call RestoreResizing(Me.hWnd)
    #End If
    
    'Call frmContent.MDIForm_Resize 'évite le bug d'affichage
    
    'on referme le handle du disque ouvert
    Call CloseHandle(hDrive)
End Sub

Private Sub FV2_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim curOffset As Currency

    'affiche le path tronqué dans la picturebox prévue à cet effet
    Call DisplayPath
    
    'affiche les infos sur la fragmentation du fichier
    Call DisplayClustersInfos(FV.Path & "\" & Item.Text)
    
    'click sur un offset ==> affiche cet offset dans le HW
    If lstFrag.ListCount = 0 Then Exit Sub

    curOffset = Val(lstFrag.List(0)) 'numéro du cluster
    curOffset = curOffset * cDrive.BytesPerCluster 'offset (byte)
    curOffset = ByN(curOffset, 16) / 16   'valeur à mettre dans le Scroll
    
    '//A CHANGER 'TODO
    curOffset = curOffset + IIf(cDrive.VolumeLetter = "L", 16800, 0)
        
    VS.Value = curOffset    'applique le refresh dans le VS_Change
    Call VS_Change(VS.Value)
End Sub

Private Sub lblGOTO_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'affiche un popup pour sauter à un autre emplacement du disque
    lblGOTO.BorderStyle = 1
    If Button = 1 Then Me.PopupMenu frmContent.mnuPopupDisk, , _
        lblGOTO.Left - 2050, lblGOTO.Top
    lblGOTO.BorderStyle = 0
End Sub

Private Sub lstFrag_Click()
Dim curOffset As Currency

'click sur un offset ==> affiche cet offset dans le HW

    curOffset = Val(lstFrag.List(lstFrag.ListIndex)) 'numéro du cluster
    curOffset = curOffset * cDrive.BytesPerCluster 'offset (byte)
    curOffset = ByN(curOffset, 16) / 16   'valeur à mettre dans le Scroll
    
    '//A CHANGER 'TODO
    curOffset = curOffset + IIf(cDrive.VolumeLetter = "L", 16800, 0)
    
    If curOffset = 0 Then Exit Sub
    
    VS.Value = curOffset    'applique le refresh dans le VS_Change
    Call VS_Change(VS.Value)
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

Private Sub lstSignets_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'permet de ne pas changer le HW dans le cas de multiples sélections
    mouseUped = True
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
        .Language = cPref.env_Lang
        Call .LoadControlsCaption
    End With
    
    'subclasse la form pour éviter de resizer trop
    #If USE_FORM_SUBCLASSING Then
        Call LoadResizing(Me.hWnd, 9000, 6000)
    #End If
    
    'affecte les valeurs générales (type) à l'historique
    cUndo.tEditType = edtDisk
    Set cUndo.Frm = Me
    Set cUndo.lvHisto = Me.lstHisto
    ReDim cHisto(0)
    Set cHisto(0) = New clsUndoSubItem
    
    'affiche ou non les éléments en fonction des paramètres d'affichage de frmcontent
    With frmContent
        Me.HW.Visible = .mnuTab.Checked
        Me.VS.Visible = .mnuTab.Checked
        Me.FV.Visible = .mnuExploreDisk.Checked
        Me.FV2.Visible = .mnuExploreDisk.Checked
        Me.pctPath.Visible = .mnuExploreDisk.Checked
        Me.FrameFrag.Visible = .mnuExploreDisk.Checked
        Me.FrameData.Visible = .mnuEditTools.Checked
        Me.FrameInfo2.Visible = .mnuInformations.Checked
        Me.FrameInfos.Visible = .mnuInformations.Checked
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
    End With
    
    HW.UseHexOffset = True  'on utilise obligatoirement l'affichage des offsets en hexa
    'car les valeurs sont trop grandes pour le Long, et pas affichable par currency (nmbre trop grand en pixels)
    
    With cPref
        'en grand dans la MDIform
        If .general_MaximizeWhenOpen Then Me.WindowState = vbMaximized
    End With
    
    With Lang
        N0Clust = .GetString("_N0Clust")
        N0LogSec = .GetString("_N0LogSec")
        N0PhysSec = .GetString("_N0PhysSec")
    End With
    
    frmContent.Sb.Panels(1).Text = "Status=[Opening disk]"
    frmContent.Sb.Refresh
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'FileContent = vbNullString
    lNbChildFrm = lNbChildFrm - 1
    frmContent.Sb.Panels(2).Text = Lang.GetString("_Opening") & CStr(lNbChildFrm) & "]"
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    'redimensionne le FV
    With FV
        .Top = 0
        .Left = 50
        .Width = 3000
        .Height = 2535
        .RefreshListViewOnly
        .ShowDrives = False   'empêche de pouvoir changer de Drive
    End With
    
    'positionne le pctPath
    With pctPath
        .Height = 255
        .Top = FV.Height
        .Left = 50
        .Width = Me.Width - 1200
    End With
    With lblGOTO
        .Height = 255
        .Top = pctPath.Top
        .Left = pctPath.Width + 100
    End With
    
    'redimensionne et initialise le FV2
    With FV2
        .Top = 0
        .Left = 100 + FV.Width
        .Width = Me.Width - 2355 - FV.Width
        .Height = 2535
        .RefreshListViewOnly
    End With
    
    'redimensionne le Frame Frag
    With FrameFrag
        .Top = 0
        .Left = IIf(FV.Visible, FV.Width + FV2.Width, 0) + 200
        .Width = 2055
        .Height = 2535
    End With
    
    'redimensionne/bouge le frameInfo
    With FrameInfos
        .Top = IIf(FV.Visible, FV.Height + 300, 0)
        .Height = Me.Height - 700 - .Top
        .Left = 20
        cmdMAJ.Top = .Height - 650
        lstHisto.Height = .Height - 5300
        lstSignets.Height = .Height - 5300
        Me.pctContain_cmdMAJ.Height = .Height - 350
    End With
    
    'met le Grid à la taille de la fenêtre
    With HW
        .Width = 9620
        .Height = Me.Height - 400 - Sb.Height - FrameInfos.Top
        .Left = IIf(FrameInfos.Visible, FrameInfos.Width, 0) + 50
        .Top = FrameInfos.Top
    End With
    
    'bouge le frameData
    With FrameData
        .Top = 100 + HW.Top
        .Left = IIf(HW.Visible, HW.Width + HW.Left, IIf(FrameInfos.Visible, _
            FrameInfos.Width, 0)) + 500
    End With
    
    'bouge le FrameInfo2
    With FrameInfo2
        .Top = 200 + HW.Top + FrameData.Height
        .Left = FrameData.Left - 200
        .Width = Me.Width - HW.Left - HW.Width - 500
        .Height = HW.Height - 1700
        Picture3.Width = .Width - 100
        Picture3.Height = .Height - 300
    End With
    
    With VS
        .Top = FrameInfos.Top
        .Height = Me.Height - 430 - Sb.Height - FrameInfos.Top
        .Left = IIf(Me.Width < 13100, Me.Width - 350, HW.Left + HW.Width)
    End With
    
    'calcule le nombre de lignes du Grid à afficher
    'NumberPerPage = Int(Me.Height / 250) - 1
    NumberPerPage = Int(HW.Height / 250) - 1
    
    With HW
        .NumberPerPage = NumberPerPage
        .Refresh
    End With

    Call DisplayPath
    
    Call VS_Change(VS.Value)
End Sub

'=======================================================
'permet de lancer le Resize depuis une autre form
'=======================================================
Public Sub ResizeMe()
    Call Form_Resize
End Sub

'=======================================================
'affiche le path du fichier sélectionné dans la textbox
'=======================================================
Private Sub DisplayPath()
Dim s As String
Dim l As Long

    'récupère le texte à afficher
    s = FV.Path & "\"
    s = Replace$(s, "\\", "\")  'vire le double slash
    If FV2.ListIndex <> -1 Then s = s & FV2.ListItems.Item(FV2.ListIndex)
    
    'enlève la partie après le vbNullChar de la string
    l = InStr(1, s, vbNullChar)
    If l > 0 Then
        s = Left$(s, l)
    End If
    
    'affiche la string dans la picturebox
    pctPath.Text = cFile.GetFolderName(s)
End Sub

'=======================================================
'affiche dans le HW les valeurs hexa qui correspondent à la partie
'du fichier qui est visualisée
'=======================================================
Private Sub OpenDrive()
Dim r() As Byte
Dim RB() As Byte
Dim RA() As Byte
Dim RD() As Byte
Dim RT() As Byte
Dim Offset As Currency
Dim Sector As Currency
Dim x As Long
Dim s As String
Dim y As Long
Dim h As Long
Dim lDisplayableBytes As Long
Dim Sb As Currency
Dim sA As Currency
Dim offsetFinSectorBef As Currency
Dim offsetFinSectorVis As Currency
Dim offsetFinSectorAft As Currency
Dim lDecal As Long

    'On Error GoTo ErrGestion
    
    'initialise tous les tableaux de bytes
    ReDim r(0): ReDim RB(0): ReDim RA(0): ReDim RD(0): ReDim RT(0)
    
    lDisplayableBytes = HW.NumberPerPage * 16 'nombres d'éléments affichables
    Offset = HW.FirstOffset 'premier offset visualisé
    Sector = Int(Offset / lBytesPerSector) 'secteur visualisé
    
    'détermination des 3 secteurs dont on veut les bytes
    'un secteur AVANT, celui visualisé en partie, et celui APRES
    'de manière à pouvoir visualiser de manière continue les secteurs
    Sb = Sector - 1: sA = Sector + 1
    If Sb >= 0 Then
        Call DirectReadHandle(hDrive, Sb, lBytesPerSector, lBytesPerSector, _
            RB())       'obtient le secteur d'avant celui visualisé
    End If
    If sA <= cDrive.TotalPhysicalSectors Then
        'obtient le secteur d'apres
        Call DirectReadHandle(hDrive, sA, lBytesPerSector, lBytesPerSector, RA())
    End If
    
    'détermine les limites (en offset) des 3 secteurs
    offsetFinSectorBef = IIf(Sb <> -1, Sb * lBytesPerSector, 0) + lBytesPerSector
    offsetFinSectorVis = offsetFinSectorBef + lBytesPerSector
    offsetFinSectorAft = offsetFinSectorVis + lBytesPerSector
    
    'obtient les bytes du secteur visualisé en partie
    Call DirectReadHandle(hDrive, Sector, lBytesPerSector, lBytesPerSector, r())
    
    'créé la liste RD() des bytes visualisés à partir des 3 secteurs dont ont a les bytes
    ReDim RD(lDisplayableBytes)     'redimensionne le tableau au nombre de bytes qui vont être affichés
    
    'met les 3 secteurs bout à bout dans une même liste temporaire
    ReDim RT(lBytesPerSector * (IIf(UBound(RA) > 0, 1, 0) + IIf(UBound(RB) > 0, _
        1, 0) + 2) - 1) 'nombre de bytes lus dans les 3 secteurs lus ou pas
        'le "+2" n'est pas logique pour moi, mais bon, sans çà çà plante (+1 pour moi)

    '//remplit le tableau temporaire contenant la réunion des secteurs lus
        For x = 0 To lBytesPerSector - 1
            If UBound(RB) > 0 Then
                'alors on pioche dans le secteur 1
                RT(x) = RB(x)
            Else
                'alors on pioche dans le secteur 2 (toujours lu)
                RT(x) = r(x)
            End If
        Next x
        If UBound(RT) > lBytesPerSector Then
            For x = lBytesPerSector To 2 * lBytesPerSector - 1
                If UBound(RB) > 0 Then
                    'alors on pioche dans le secteur 2
                    RT(x) = r(x - lBytesPerSector)
                Else
                    'alors on pioche dans le secteur 3
                    RT(x) = RA(x - lBytesPerSector)
                End If
            Next x
        End If
        If UBound(RT) > 2 * lBytesPerSector Then
            For x = 2 * lBytesPerSector To 3 * lBytesPerSector - 1
                'on pioche forcément dans le 3eme secteur
                RT(x) = RA(x - 2 * lBytesPerSector)
            Next x
        End If
        
        
        'affecte au tableau affiché les bytes qui proviennent de la réunion des 3 secteurs
        lDecal = Offset - offsetFinSectorBef + lBytesPerSector
        
        For x = 0 To lDisplayableBytes
            'calcule le décalage
            RD(x) = RT(x + lDecal) 'affecte la valeur qui sera affichée
        Next x
    
    'ajoute les valeurs string/hexa obtenues au HW
    For x = 0 To ByN(UBound(RD()), 16) - 1 Step 16
        s = vbNullString
        For y = 0 To 15
            h = x + y
            s = s & Byte2FormatedString(RD(h))
            HW.AddHexValue 1 + x / 16, y + 1, IIf(Len(Hex$(RD(h))) = 1, "0" & _
                Hex$(RD(h)), Hex$(RD(h)))
        Next y
        HW.AddStringValue 1 + x / 16, s
    Next x
    
    'HW.Refresh  'refresh HW
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "diskPfm.OpenDrive", True
End Sub

'=======================================================
'renvoie si l'offset contient une modification
'=======================================================
Private Function IsOffsetModified(ByVal lOffset As Long, ByRef lPlace As Long) _
    As Boolean
    
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
'renvoie si la case a été modifiée ou non
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
'obtient le nom du fichier à ouvrir, et l'ouvre
'=======================================================
Public Sub GetDrive(ByVal sDrive As String)
Dim l As Currency

    '//on récupère le handle du disque en écriture/lecture
    'TODO MAUVAIS HANDLE (pas écriture)
    hDrive = GetDiskHandleRead(sDrive)

    'récupère les infos sur le disque
    Set cDrive = cFile.GetDrive(Left$(sDrive, 1))
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_OpDisk") & " " & sDrive & " ...")
    
    strDrive = BuildDrive(sDrive)   'formaté
    
    With FV
        .Path = Left$(sDrive, 3)  'affecte le drive sélectionné au path du FV
        FV2.Path = .Path
    End With
    
    '//obtient maintenant les infos sur le Drive
    frmContent.Sb.Panels(1).Text = "Status=[Rerieving disk informations]"
    
    'appelle la classe
    With cDrive
        lBytesPerSector = .BytesPerSector
        lLength = .PartitionLength    'taille totale
    End With
    With HW
        .MaxOffset = lLength 'offset maximal
        .FileSize = lLength
    End With
    Me.Caption = "Disque " & Right$(strDrive, 2) & "\"
    
    'affiche les infos disque dans les textboxes
    With cDrive
        TextBox(8).Text = Lang.GetString("_Disk") & .VolumeLetter & "]"
        TextBox(9).Text = Lang.GetString("_VolName") & .VolumeName & "]"
        TextBox(10).Text = Lang.GetString("_PartType") & .FileSystemName & "]"
        TextBox(11).Text = Lang.GetString("_SN") & Hex$(.VolumeSerialNumber) & "]"
        TextBox(12).Text = Lang.GetString("_DiskType") & .strDriveType & "]"
        TextBox(13).Text = Lang.GetString("_MediaT") & .strMediaType & "]"
        TextBox(14).Text = Lang.GetString("_TotalPhysSpace") & Trim$(Str$(.PartitionLength)) & "]"
        TextBox(15).Text = Lang.GetString("_FreeSpace") & Trim$(Str$(.FreeSpace)) & "]"
        TextBox(0).Text = Lang.GetString("_UsedSpace") & Trim$(Str$(.UsedSpace)) & "]"
        TextBox(1).Text = Lang.GetString("_Percent") & Trim$(Str$(Round(.PercentageFree, 4))) & " %]"
        TextBox(17).Text = Lang.GetString("_Cyl") & Trim$(Str$(.Cylinders)) & "]"
        TextBox(16).Text = Lang.GetString("_TrackPerCyl") & Trim$(Str$(.TracksPerCylinder)) & "]"
        TextBox(7).Text = Lang.GetString("_SecPerT") & Trim$(Str$(.SectorsPerTrack)) & "]"
        TextBox(6).Text = Lang.GetString("_SecLog") & Trim$(Str$(.TotalLogicalSectors)) & "]"
        TextBox(5).Text = Lang.GetString("_SecP") & Trim$(Str$(.TotalPhysicalSectors)) & "]"
        TextBox(4).Text = Lang.GetString("_SecPerC") & Trim$(Str$(.SectorPerCluster)) & "]"
        TextBox(18).Text = Lang.GetString("_HidSec") & Trim$(Str$(.HiddenSectors)) & "]"
        TextBox(19).Text = Lang.GetString("_Clust") & Trim$(Str$(.TotalClusters)) & "]"
        TextBox(2).Text = Lang.GetString("_FClust") & Trim$(Str$(.FreeClusters)) & "]"
        TextBox(3).Text = Lang.GetString("_UClust") & Trim$(Str$(.UsedClusters)) & "]"
        TextBox(27).Text = Lang.GetString("_BPerSec") & Trim$(Str$(.BytesPerSector)) & "]"
        TextBox(26).Text = Lang.GetString("_BPerClust") & Trim$(Str$(.BytesPerCluster)) & "]"
    End With
    TextBox(20).Text = Lang.GetString("_N0Clust") & "0]"
    TextBox(21).Text = Lang.GetString("_N0LogSec") & "0]"
    TextBox(22).Text = Lang.GetString("_N0PhysSec") & "0]"

    frmContent.Sb.Panels(1).Text = "Status=[Ready]"

    'règle la taille de VS
    With VS
        .Min = 0
        .Max = ByN(cDrive.PartitionLength / 16, 16)
        .Value = 0
        .SmallChange = 1
        .LargeChange = NumberPerPage - 1
    End With
    
    'stocke dans les tag les valeurs Max et Min des offsets
    With HW
        .curTag1 = .FirstOffset
        .curTag2 = .MaxOffset
    End With
    
    'MAJ
    Call cmdMAJ_Click
    
    'affichage
    Call OpenDrive
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_DiskO") & " " & sDrive & " " & _
        Lang.GetString("_Opened"))
    
End Sub

Private Sub FV_PathChange(sOldPath As String, sNewPath As String)
    FV2.Path = sNewPath 'met à jour le path du FileView qui affiche les fichiers
    Call DisplayPath 'affiche le path sélectionné
End Sub

Private Sub HW_KeyDown(KeyCode As Integer, Shift As Integer)
'gère les touches qui changent le VS, gère le changement de valeur
    
    On Error GoTo ErrGestion
    
    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
    With HW
         If KeyCode = vbKeyUp Then
            'alors monte
            If .FirstOffset = 0 And .Item.Line = 1 Then Exit Sub  'tout au début déjà
            'on remonte d'une ligne alors
            .Item.Line = .Item.Line - 1
            If .Item.Line = 0 Then
                'alors on remonte le firstoffset
                .Item.Line = 1
                .FirstOffset = .FirstOffset - 16
                VS.Value = VS.Value - 1
                Call VS_Change(VS.Value)
            End If
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, _
                .Item.Col), .SelectionColor, True
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
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, _
                .Item.Col), .SelectionColor, True
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
            .Value = IIf((.Value - NumberPerPage) > .Min, .Value - _
                NumberPerPage, .Min)
            Call VS_Change(.Value)
        End If
        If KeyCode = vbKeyPageDown Then
            'alors descendre de NumberPerPage
            .Value = IIf((.Value + NumberPerPage) < .Max, .Value + _
                NumberPerPage, .Max)
            Call VS_Change(VS.Value)
        End If
    End With
    
    With HW
        If KeyCode = vbKeyLeft Then
            'alors va à gauche
            If .FirstOffset = 0 And .Item.Col = 1 And .Item.Line = 1 Then Exit Sub 'tout au début déjà
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
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, .Item.Col), _
                .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
             
        If KeyCode = vbKeyRight Then
            'alors va à droite
            If .FirstOffset + .Item.Line * 16 - 16 = By16(.MaxOffset) And _
                .Item.Col = 16 Then Exit Sub  'tout à la fin déjà
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
            .ColorItem tHex, .Item.Line, .Item.Col, .Value(.Item.Line, _
                .Item.Col), .SelectionColor, True
            .AddSelection .Item.Line, .Item.Col
        End If
        
        'réenregistre le numéro de l'offset actuel dans hw.item
        .Item.Offset = .FirstOffset + (.Item.Line - 1) * 16
        'affecte les autres valeurs dans Item
        '.Item.tType = tHex
        .Item.Value = .Value(.Item.Line, .Item.Col)
    
    End With
    
    DoEvents
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "diskPfm.KeyDown", True
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
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And _
            KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Then
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
            
            'applique le changement
            Call ModifyData(Chr_(bytHex))
            
            'simule l'appui sur "droite"
            Call HW_KeyDown(vbKeyRight, 0)
        End If
    ElseIf HW.Item.tType = tString Then
        'alors voici la zone STRING
        'on ne tappe QU'UNE SEULE VALEUR
        
            'le nouveau byte est donc désormais KeyAscii
            Call ModifyData(Chr_(KeyAscii))

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
        With frmContent
            .mnuDeleteSelection.Enabled = False
            .mnuCut.Enabled = False
            Me.PopupMenu .rmnuEdit ', X + GD.Left, Y + GD.Top
        End With
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
        
        With Item
            If .Col < 1 Or .Col > 16 Then Exit Sub
                   
            If .Value = vbNullString Then
                'pas de valeur
                txtValue(0).Text = vbNullString
                txtValue(1).Text = vbNullString
                txtValue(2).Text = vbNullString
                txtValue(3).Text = vbNullString
                FrameData.Caption = "No data"
                Exit Sub
            End If
            
            'affiche la donnée sélectionnée dans frmData
            If .tType = tHex Then
                'valeur hexa sélectionnée
                txtValue(0).Text = .Value
                txtValue(1).Text = Hex2Dec(Item.Value)
                txtValue(2).Text = Hex2Str(Item.Value)
                txtValue(3).Text = Hex2Oct(Item.Value)
                FrameData.Caption = "Data=[" & .Value & "]"
            End If
            If .tType = tString Then
                'valeur strind sélectionnée
                s = HW.Value(.Line, .Col)
                txtValue(0).Text = s
                txtValue(1).Text = Hex2Dec(s)
                txtValue(2).Text = Hex2Str(s)
                txtValue(3).Text = Hex2Oct(s)
                FrameData.Caption = "Data=[" & s & "]"
            End If
        End With
        
    ElseIf Button = 4 And Shift = 0 Then
        'click avec la molette, et pas de Shift or Control
        'on ajoute (ou enlève) un signet
        
        With HW
            If .IsSignet(Item.Offset) = False Then
                'on l'ajoute
                .AddSignet Item.Offset
                Me.lstSignets.ListItems.Add Text:=CStr(Item.Offset)
                .TraceSignets
            Else
            
                'alors on l'enlève
                While .IsSignet(.Item.Offset)
                    'on supprime
                    .RemoveSignet Val(.Item.Offset)
                Wend
                
                'enlève du listview
                For r = lstSignets.ListItems.Count To 1 Step -1
                    If lstSignets.ListItems.Item(r).Text = CStr(.Item.Offset) Then
                        lstSignets.ListItems.Remove r
                    End If
                Next r
            
                Call Refresh
            End If
        End With
    ElseIf Button = 4 And Shift = 2 Then
        'click molette + control
        'sélectionne une zone définie
        With frmSelect
            .GetEditFunction 0 'selection mode
            .Show vbModal
        End With
    End If
    
End Sub

Private Sub HW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Sb.Panels(4).Text = Lang.GetString("_Selec") & _
        CStr(HW.NumberOfSelectedItems) & " " & Lang.GetString("_Bytes")
    Label2(9) = Me.Sb.Panels(4).Text
End Sub

Private Sub HW_MouseWheel(ByVal lSens As Long)

    DoEvents    '/!\ IMPORTANT : DO NOT REMOVE
    'it allows to refresh correctly the HW control
    
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

Private Sub lstSignets_ItemClick(ByVal Item As ComctlLib.ListItem)
'va au signet
    If mouseUped Then
        With HW
            .FirstOffset = Val(Item.Text)
            .Refresh
            Me.VS.Value = .FirstOffset / 16
        End With
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
        s = InputBox(Lang.GetString("_AddCommentFor") & " " & tLst.Text, Lang.GetString("_AddCom"))
        If StrPtr(s) <> 0 Then
            'ajoute le commentaire
            tLst.SubItems(1) = s
        End If
    End If
    
    If Button = 4 Then
        'mouse du milieu ==> on supprime le signet
        Set tLst = lstSignets.HitTest(x, y)
        If tLst Is Nothing Then Exit Sub
        
        r = MsgBox(Lang.GetString("_DelSig") & " " & tLst.Text & " ?", vbInformation + vbYesNo, Lang.GetString("_War"))
        If r <> vbYes Then Exit Sub
        
        'on supprime
        Call HW.RemoveSignet(Val(tLst.Text))
        
        'on enlève du listview
        Call lstSignets.ListItems.Remove(tLst.Index)
    End If
        
End Sub

Private Sub pctPath_Change()
    If cFile.FolderExists(cFile.GetFolderName(pctPath.Text & "\")) = False Then
        'couleur rouge
        pctPath.ForeColor = RED_COLOR
    Else
        'c'est un path ok
        pctPath.ForeColor = GREEN_COLOR
    End If
End Sub

Private Sub pctPath_KeyDown(KeyCode As Integer, Shift As Integer)
'valide si entrée
Dim s As String
    If KeyCode = vbKeyReturn Then
        With pctPath
            s = .Text
            If cFile.FolderExists(.Text) Then FV.Path = .Text
            .Text = s
        End With
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

'=======================================================
'change la valeur du VS
'public, car cette sub est aussi appelée pour le refresh
'=======================================================
Public Sub VS_Change(Value As Currency)
Dim lPages As Long

    On Error GoTo ErrGestion
    
    If lBytesPerSector <= 0 Then Exit Sub 'pas prêt

    HW.FirstOffset = VS.Value * 16
    
    'réaffiche la Grid
    Call OpenDrive
    
    'calcule le nbre de pages
    lPages = lLength / (NumberPerPage * 16) + 1
    Me.Sb.Panels(2).Text = "Page=[" & CStr(1 + Int(VS.Value / NumberPerPage)) _
        & "/" & CStr(lPages) & "]"
    Label2(8).Caption = Me.Sb.Panels(2).Text
    
    HW.Refresh
    
    'met à jour les textboxes contenant les numéros de secteur/cluster
    With VS
        TextBox(20).Text = N0Clust & CStr(Int((.Value / cDrive.BytesPerCluster) * 16)) & "]"
        TextBox(21).Text = N0LogSec & CStr(Int((.Value / cDrive.BytesPerSector) * 16)) & "]"
        TextBox(22).Text = N0PhysSec & CStr(Int((.Value / cDrive.BytesPerSector) * 16)) & "]"
    End With
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "diskPfm.VS_Change", True
End Sub

'=======================================================
'ajoute une valeur changée dans la liste des valeurs changées
'=======================================================
Public Sub AddChange(ByVal lOffset As Long, ByVal lCol As Long, _
    ByVal sString As String)
   
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
'procède à la sauvegarde du fichier avec changements à l'emplacement sFile2
'=======================================================
Public Function GetNewFile(ByVal sFile2 As String) As String
'Dim x As Long, s As String
'Dim tmpText As String
'Dim y As Long
'Dim a As Long
'Dim e As Long
'Dim lLen As Long
'Dim lFile2 As Long
'Dim lPlace As Long
'
'    On Error GoTo ErrGestion
'
'    'affiche le message d'attente
'    frmContent.Sb.Panels(1).Text = "Status=[Saving " & Me.Caption & "]"
'
'    lFile2 = FreeFile 'obtient une ouverture dispo
'    Open sFile2 For Binary Access Write As lFile2   'ouvre le fichier sFile2 pour l'enregistrement
'
'    'créé un buffer de longueur divisible par 16
'    'recoupera à la fin pour la longueur exacte du fichier
'    lLen = By16(lLength)
'    tmpText = String$(16, 0)
'
'        For a = 1 To lLen Step 16  'remplit par intervalles de 16
'
'            Get lFile, a, tmpText   'prend un morceau du fichier source (16bytes) à partir de l'octet A
'
'            If IsOffsetModified(a - 1, lPlace) Then
'                'l'offset est modifié
'
'                'la string existe dans la liste des string modifiées
'                'ne formate PAS la string
'
'                'vérifie que l'on est pas dans la dernière ligne, si oui, ne prend que la longueur nécessaire
'                If a + 16 > lLength Then
'                    s = Mid$(ChangeListS(lPlace), 1, lLength - a + 1)
'                Else
'                    s = ChangeListS(lPlace)
'                End If
'            Else
'
'                'pas de modif, prend à partir du fichier source
'                'vérfie que l'on est pas dans la dernière ligne
'
'                If a + 16 > lLength Then
'                    s = Mid$(tmpText, 1, lLength - a + 1)
'                Else
'                    s = tmpText
'                End If
'            End If
'
'            Put lFile2, a, s    'ajoute au fichier résultant
'
'            If (a Mod 160017) = 0 Then
'                'rend un peu la main
'                frmContent.Sb.Panels(1).Text = "Status=[Saving " & Me.Caption & "]" & "[" & Round(100 * a / lLength, 2) & " %]"
'                DoEvents
'            End If
'
'        Next a
'
'    Close lFile2
'
'    'affiche le message de fin de sauvegarde
'    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
'
'    Exit Function
'ErrGestion:
'    clsERREUR.AddError "diskPfm.GetNewFile", True
End Function

'=======================================================
'fonction ayant uniquement pour but d'exister, on l'appelle à partir d'une
'autre fonction pour tester si frmcontent.activeform est
'une form d'édition mémoire, edition de fichier ou de disque
'=======================================================
Public Function Useless() As String
    Useless = "Disk"
End Function

'=======================================================
'renvoie le cDrive du disque ouvert par cette form
'=======================================================
Public Function GetDriveInfos() As FileSystemLibrary.Drive
    Set GetDriveInfos = cDrive
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

        'With frmContent.ActiveForm.HW
            '.AddHexValue I_tem.Line, I_tem.Col, txtValue(0).Text
            '.AddOneStringValue I_tem.Line, I_tem.Col, txtValue(2).Text
            Call ModifyData(txtValue(2).Text)
        'End With
    End If
        
End Sub

'=======================================================
'affiche les infos clusters du fichier sFile
'=======================================================
Private Sub DisplayClustersInfos(ByVal sFile As String)
Dim F As FileClusters
Dim l As Long
Dim clust As Currency
Dim lastOffset  As Currency

    'obtient la localisation (en cluster) d'un fichier sFile
    F = GetFileBitmap(sFile)
    
    lblFrag.Caption = Lang.GetString("_Frag") & F.ExtentsCount & "]"
    Call lstFrag.Clear
    
    'affiche tous les fragments
    lastOffset = 0
    If F.ExtentsCount <> 0 Then
        For l = 0 To F.ExtentsCount - 1
            clust = GetCurrency(F.Extents(l).NextVcn.LowDWORD, _
                F.Extents(l).NextVcn.HighDWORD) - lastOffset
            Call lstFrag.AddItem(CStr(GetCurrency(F.Extents(l).LCN.LowDWORD, _
                F.Extents(l).LCN.HighDWORD)) & "  (" & Trim$(Str$(clust)) & ")")
            lastOffset = GetCurrency(F.Extents(l).NextVcn.LowDWORD, _
                F.Extents(l).NextVcn.HighDWORD)
        Next l
    End If
End Sub

'=======================================================
'des données ont étés modifiées ==> on sauvegarde ces changements
'=======================================================
Private Sub ModifyData(ByVal sNewChar As String)
Dim sActualString As String
Dim iCur As Currency
Dim sBef As String, sAft As String
Dim Sector As Long
Dim sOff As Currency
Dim I_tem As ItemElement

    Set I_tem = HW.Item

    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    'détermine l'offset du caractère à changer
    iCur = I_tem.Offset + I_tem.Col - 1
    
    'détermine le numéro du secteur contenant le caractère à changer
    Sector = Int(iCur / lBytesPerSector)
    
    'récupère dans une string le contenu actuel de ce cluster
    Call DirectReadSHandle(hDrive, Sector, lBytesPerSector, _
        lBytesPerSector, sActualString)
        
    '//modifie cette string avec la nouvelle valeur
    sOff = -Sector
    sOff = sOff * lBytesPerSector
    sOff = sOff + iCur 'décomposition pour éviter le dépassement de capacité
    'sOff = iCur - lBytesPerSector * Sector  'offset DANS LE SECTEUR
    
    'remplace le sOff-ème char par le nouveau char
    sBef = Mid$(sActualString, 1, sOff) 'string d'avant le nouveau char
    sAft = Mid$(sActualString, sOff + 2, lBytesPerSector - sOff)   'string d'après
    sActualString = sBef & sNewChar & sAft   'concatene et créé ainsi la new string
    
    'écrit dans le disque avec WriteFile
    Call DirectWrite(cDrive.VolumeLetter & ":\", Sector, lBytesPerSector, _
        lBytesPerSector, sActualString)
    
    'met à jour le HW
    Call frmContent.ActiveForm.VS_Change(frmContent.ActiveForm.VS.Value)
    
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

Private Sub VS_MouseAction(ByVal lngMouseAction As ExtVS.MOUSE_ACTION)
'alors une action a été effectuée (lance le popup menu)
    If lngMouseAction = RIGHT_CLICK Then
        Me.PopupMenu frmContent.rmnuPos
    End If
End Sub
