VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{9B9A881F-DBDC-4334-BC23-5679E5AB0DC6}#1.2#0"; "FileView_OCX.ocx"
Begin VB.Form frmMain 
   Caption         =   "File Renamer VB"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Navigation dans les dossiers"
      Height          =   6735
      Left            =   120
      TabIndex        =   69
      Top             =   120
      Width           =   10095
      Begin VB.TextBox pctPath 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   74
         Top             =   240
         Width           =   9615
      End
      Begin FileView_OCX.FileView File1 
         Height          =   6015
         Left            =   3240
         TabIndex        =   72
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   10610
         ShowEntirePath  =   0   'False
         Path            =   "C:\"
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
      Begin FileView_OCX.FileView Folder1 
         Height          =   6015
         Left            =   120
         TabIndex        =   71
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   10610
         ShowEntirePath  =   0   'False
         AllowDirectoryDeleting=   0   'False
         AllowFileDeleting=   0   'False
         AllowFileRenaming=   0   'False
         AllowDirectoryRenaming=   0   'False
         AllowReorganisationByColumn=   0   'False
         HideColumnHeaders=   -1  'True
         Path            =   "C:\"
         ShowFiles       =   0   'False
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fichiers à renommer"
      Height          =   6735
      Left            =   10320
      TabIndex        =   68
      Top             =   120
      Width           =   4815
      Begin FileView_OCX.FileView FileR 
         Height          =   6375
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   11245
         ShowEntirePath  =   0   'False
         AllowDirectoryDeleting=   0   'False
         AllowDirectoryEntering=   0   'False
         AllowFileDeleting=   0   'False
         AllowFileRenaming=   0   'False
         AllowDirectoryRenaming=   0   'False
         Path            =   ""
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
   End
   Begin VB.ListBox lstTransfo 
      Height          =   1860
      Left            =   7080
      Style           =   1  'Checkbox
      TabIndex        =   66
      Top             =   7320
      Width           =   4215
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   11520
      TabIndex        =   59
      Top             =   6960
      Width           =   1815
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3135
         Index           =   0
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   1575
         TabIndex        =   60
         Top             =   180
         Width           =   1575
         Begin VB.CommandButton cmdDel 
            Caption         =   "Supprimer"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   615
            Width           =   1575
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "Descendre"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":0B3D
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   1230
            Width           =   1575
         End
         Begin VB.CommandButton cmdUp 
            Caption         =   "Monter"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":0D80
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   0
            Width           =   1575
         End
         Begin VB.CommandButton cmdOpenList 
            Caption         =   "Ouvrir une liste"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":0FC0
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1845
            Width           =   1575
         End
         Begin VB.CommandButton cmdSaveList 
            Caption         =   "Enregistrer la liste"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":121F
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   2460
            Width           =   1575
         End
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   5400
      TabIndex        =   58
      Top             =   9840
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Position"
      Enabled         =   0   'False
      Height          =   1095
      Index           =   2
      Left            =   5160
      TabIndex        =   54
      Top             =   7440
      Width           =   1815
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   735
         Index           =   2
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   1575
         TabIndex        =   55
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton optPos5 
            Caption         =   "Ajouter au début"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   57
            Top             =   120
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optPos5 
            Caption         =   "Ajouter à la fin"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   56
            Top             =   480
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ajouter ..."
      Height          =   1095
      Index           =   0
      Left            =   5160
      TabIndex        =   50
      Top             =   8640
      Width           =   1815
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   765
         Index           =   0
         Left            =   120
         ScaleHeight     =   765
         ScaleWidth      =   1575
         TabIndex        =   51
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton Option1 
            Caption         =   "à l'extension"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   53
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "au nom du fichier"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   52
            Top             =   120
            Value           =   -1  'True
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3375
      Left            =   13320
      TabIndex        =   46
      Top             =   6960
      Width           =   1575
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3015
         Index           =   1
         Left            =   120
         ScaleHeight     =   3015
         ScaleWidth      =   1365
         TabIndex        =   47
         Top             =   240
         Width           =   1370
         Begin VB.CommandButton cmdSimulate 
            Caption         =   "Simulation"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":163B
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton cmdRenamerAll 
            Caption         =   "RENOMMER"
            Height          =   615
            Left            =   0
            MaskColor       =   &H00404040&
            Picture         =   "frmMain.frx":1A5D
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   615
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   5
      Left            =   120
      TabIndex        =   42
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   5
         Left            =   50
         ScaleHeight     =   2655
         ScaleWidth      =   4755
         TabIndex        =   43
         Top             =   120
         Width           =   4755
         Begin VB.ComboBox cb6 
            Height          =   315
            ItemData        =   "frmMain.frx":1E79
            Left            =   120
            List            =   "frmMain.frx":1E92
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label Label2 
            Caption         =   "Type de string à ajouter :"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   45
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   6
      Left            =   120
      TabIndex        =   38
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1815
         Index           =   7
         Left            =   50
         ScaleHeight     =   1815
         ScaleWidth      =   4815
         TabIndex        =   39
         Top             =   120
         Width           =   4815
         Begin VB.ComboBox cb7 
            Height          =   315
            ItemData        =   "frmMain.frx":1EEB
            Left            =   120
            List            =   "frmMain.frx":1EF8
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label Label2 
            Caption         =   "Type de string à ajouter :"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.TextBox txtTransfo 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   7080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   9240
      Width           =   4215
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   3
      Left            =   120
      TabIndex        =   29
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   3
         Left            =   50
         ScaleHeight     =   2655
         ScaleWidth      =   4815
         TabIndex        =   30
         Top             =   120
         Width           =   4815
         Begin VB.ComboBox cb4b 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":1F14
            Left            =   120
            List            =   "frmMain.frx":1F16
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   1320
            Width           =   4575
         End
         Begin VB.ComboBox cb4 
            Height          =   315
            ItemData        =   "frmMain.frx":1F18
            Left            =   120
            List            =   "frmMain.frx":1F3A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   480
            Width           =   4575
         End
         Begin VB.CheckBox chkAddText 
            Caption         =   "N'ajouter que du texte simple"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1920
            Width           =   4455
         End
         Begin VB.TextBox txtString 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   2160
            Width           =   4575
         End
         Begin VB.Label Label2 
            Caption         =   "Option 1 :"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Type de string à ajouter :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.TextBox txtProto 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   9840
      Width           =   4215
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   7440
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2415
         Index           =   0
         Left            =   50
         ScaleHeight     =   2415
         ScaleWidth      =   4815
         TabIndex        =   20
         Top             =   120
         Width           =   4815
         Begin VB.ComboBox cb1b 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMain.frx":1FFB
            Left            =   120
            List            =   "frmMain.frx":1FFD
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1200
            Width           =   4575
         End
         Begin VB.ComboBox cb1 
            Height          =   315
            ItemData        =   "frmMain.frx":1FFF
            Left            =   120
            List            =   "frmMain.frx":2015
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txt1c 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   21
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Appliquer à                    caractères"
            Enabled         =   0   'False
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Option 2 :"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Option 1 :"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Modification à appliquer :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   2
         Left            =   50
         ScaleHeight     =   2655
         ScaleWidth      =   4740
         TabIndex        =   13
         Top             =   120
         Width           =   4740
         Begin VB.TextBox txtStringToChange 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1800
            TabIndex        =   16
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txtNewString 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   480
            Width           =   2895
         End
         Begin VB.CheckBox chkReplaceCase 
            Caption         =   "Respecter la casse"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Label2 
            Caption         =   "Remplacer la string :"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Par la string :"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   4
         Left            =   50
         ScaleHeight     =   2655
         ScaleWidth      =   4815
         TabIndex        =   9
         Top             =   120
         Width           =   4815
         Begin VB.ComboBox cb5 
            Height          =   315
            ItemData        =   "frmMain.frx":20BC
            Left            =   120
            List            =   "frmMain.frx":20D8
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label Label2 
            Caption         =   "Type de string à ajouter :"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   1
         Left            =   50
         ScaleHeight     =   1335
         ScaleWidth      =   4830
         TabIndex        =   1
         Top             =   120
         Width           =   4830
         Begin VB.TextBox txtDepCpt 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   960
            Width           =   735
         End
         Begin VB.ComboBox cb2 
            Height          =   315
            ItemData        =   "frmMain.frx":2134
            Left            =   120
            List            =   "frmMain.frx":213E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   480
            Width           =   4575
         End
         Begin VB.TextBox txtStepCpt 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2760
            TabIndex        =   2
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Départ :"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Type de compteur :"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label2 
            Caption         =   "Pas :"
            Height          =   255
            Index           =   7
            Left            =   1920
            TabIndex        =   5
            Top             =   960
            Width           =   735
         End
      End
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   120
      TabIndex        =   67
      Top             =   6960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   7
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Style"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Compteur"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Remplacer"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Base"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Audio"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Video"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Image"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Liste des transformations à effectuer :"
      Height          =   255
      Left            =   7080
      TabIndex        =   70
      Top             =   6960
      Width           =   4215
   End
   Begin VB.Menu rmnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuChangeFolder1 
         Caption         =   "&Changer le dossier..."
      End
      Begin VB.Menu mnuFileTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu rmnuHelp 
      Caption         =   "&Aide"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Aide..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpTiret 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&A propos"
      End
   End
   Begin VB.Menu mnuPopupFile1 
      Caption         =   "popupFile"
      Visible         =   0   'False
      Begin VB.Menu mnuAddToCurrentList1 
         Caption         =   "&Ajouter à la liste des fichiers courante"
      End
      Begin VB.Menu mnuDisplayProperties1 
         Caption         =   "&Afficher les propriétés"
      End
      Begin VB.Menu mnuPopupFileTiret11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete1 
         Caption         =   "&Supprimer"
      End
      Begin VB.Menu mnuMoveTo1 
         Caption         =   "&Déplacer vers..."
      End
      Begin VB.Menu mnuCopyTo1 
         Caption         =   "&Copier vers..."
      End
      Begin VB.Menu mnuChangeAttr 
         Caption         =   "&Changer l'attribut"
         Begin VB.Menu mnuNormalAttr 
            Caption         =   "&Normal"
         End
         Begin VB.Menu mnuROattr 
            Caption         =   "&Lecture seule"
         End
         Begin VB.Menu mnuHiddenAttr 
            Caption         =   "&Caché"
         End
         Begin VB.Menu mnuSystemAttr 
            Caption         =   "&Système"
         End
         Begin VB.Menu mnuSystemHiddenAttr 
            Caption         =   "&Système (caché)"
         End
      End
   End
   Begin VB.Menu mnuPopupFolder1 
      Caption         =   "popupFile"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenExplorer1 
         Caption         =   "&Ouvrir l'explorateur Windows à cet emplacement"
      End
      Begin VB.Menu mnuPopupFolderTiret11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisplayPropertiesFolder1 
         Caption         =   "&Afficher les propriétés"
      End
      Begin VB.Menu mnuPopupFolderTiret21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFilesFromFolder1 
         Caption         =   "&Ajouter les fichiers du dossier"
         Begin VB.Menu mnuWithoutSubFolder1 
            Caption         =   "&Sans les sous-dossiers"
         End
         Begin VB.Menu mnuWithSubFolders1 
            Caption         =   "&Avec les sous-dossiers"
         End
      End
      Begin VB.Menu mnuPopupFolderTiret31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteFolder1 
         Caption         =   "&Supprimer"
      End
      Begin VB.Menu mnuMoveFolderTo1 
         Caption         =   "&Déplacer vers..."
      End
      Begin VB.Menu mnuCopyFolderTo1 
         Caption         =   "&Copier vers..."
      End
   End
   Begin VB.Menu mnuPopupTransfoList 
      Caption         =   "mnuPopupTransfoList"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Sélectionner tout"
      End
      Begin VB.Menu mnuDeselectAll 
         Caption         =   "&Ne rien sélectionner"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =======================================================
'
' File Renamer VB (part of Hex Editor VB)
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A Windows utility which allows to rename lots of file (part of Hex Editor VB)
'
' Copyright (c) 2006-2007 by Alain Descotes.
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
'FORM PRINCIPALE
'=======================================================

Private Sub cb1_Click()
    If cb1.Text <> vbNullString Then
        'alors une string est sélectionnée, on va pouvoir afficher les options relatives
        Select Case cb1.Text
            Case "Première lettre en majuscule"
                'pas d'option
                Label2(3).Enabled = False
                cb1b.Enabled = False
                Label2(4).Enabled = False
                Label2(17).Enabled = False
                cb1b.Clear
                txt1c.Text = vbNullString
            Case Else
                'alors une option
                Label2(3).Enabled = True
                cb1b.Enabled = True
                If cb1b.ListCount = 0 Then
                    'avant il n'y avait rien, donc on remplit les items
                    cb1b.AddItem "à partir du début"
                    cb1b.AddItem "à partir de la fin"
                    cb1b.AddItem "tous les caractères"
                End If
        End Select
    End If
End Sub

Private Sub cb1b_Click()
    If cb1.Text <> vbNullString Then
        'alors une string est sélectionnée, on va pouvoir afficher les options relatives
        Select Case cb1b.Text
            Case "tous les caractères"
                'pas d'option 2
                Label2(4).Enabled = False
                Label2(17).Enabled = False
                txt1c.Text = vbNullString
                txt1c.Enabled = False
            Case Else
                'alors une option 2
                Label2(4).Enabled = True
                Label2(17).Enabled = True
                txt1c.Enabled = True
        End Select
    End If
End Sub

Private Sub cb4_Click()
    If cb4.Text <> vbNullString Then
        'alors une string est sélectionnée, on va pouvoir afficher les options relatives
        Select Case cb4.Text
            Case "Dossier contenant"
                'pas d'option
                Label2(1).Enabled = False
                cb4b.Enabled = False
                cb4b.Clear
            Case "Extension"
                'pas d'option
                Label2(1).Enabled = False
                cb4b.Enabled = False
                cb4b.Clear
            Case "Nom du fichier (sans le chemin et l'extension)"
                'pas d'option
                Label2(1).Enabled = False
                cb4b.Enabled = False
                cb4b.Clear
            Case "Type"
                'pas d'option
                Label2(1).Enabled = False
                cb4b.Enabled = False
                cb4b.Clear
            Case Else
                'option
                Label2(1).Enabled = True
                cb4b.Enabled = True
                cb4b.Clear
                'avant il n'y avait rien, donc on remplit les items
                If cb4.Text = "Attribut" Then
                    cb4b.AddItem "Lettres"
                    cb4b.AddItem "Numéro"
                ElseIf Left$(cb4.Text, 4) = "Date" Then
                    cb4b.AddItem "JJ/MM/AAAA HH:MM:SS"
                    cb4b.AddItem "MM/JJ/AAAA HH:MM:SS"
                    cb4b.AddItem "JJ/MM/AAAA"
                    cb4b.AddItem "MM/JJ/AAAA"
                    cb4b.AddItem "MM/AAAA"
                    cb4b.AddItem "JJ/MM"
                    cb4b.AddItem "MM/JJ"
                    cb4b.AddItem "HH:MM:SS"
                    cb4b.AddItem "HH:MM"
                    cb4b.AddItem "HH"
                ElseIf cb4.Text = "Taille" Then
                    cb4b.AddItem "En octets"
                    cb4b.AddItem "En Ko"
                    cb4b.AddItem "En Mo"
                    cb4b.AddItem "En Go"
                    cb4b.AddItem "En To"
                    cb4b.AddItem "Unité la mieux adaptée"
                End If
        End Select
    End If
End Sub

Private Sub mnuCopyFolderTo1_Click()
'copie les dossiers sélectionnés
Dim s As String
Dim sTo As String
Dim x As Long

    'browse for folder
    sTo = cFile.BrowseForFolder("Choix du dossier cible", Me.hWnd)
    If cFile.FolderExists(sTo) = False Then Exit Sub
    
    'créé une string qui concatene tous les dossiers
    For x = 1 To Folder1.ListCount
        If Folder1.ListItems.Item(x).Selected Then _
            s = s & Folder1.Path & "\" & Folder1.ListItems.Item(x).Text & vbNullChar
    Next x
    
    Call cFile.CopyFileOrFolder(s, sTo): DoEvents    'copy files
    Call Folder1.Refresh
End Sub

Private Sub mnuCopyTo1_Click()
'copie les fichiers sélectionnés
Dim s As String
Dim sTo As String
Dim x As Long

    'browse for folder
    sTo = cFile.BrowseForFolder("Choix du dossier cible", Me.hWnd)
    If cFile.FolderExists(sTo) = False Then Exit Sub
    
    'créé une string qui concatene tous les fichiers
    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            s = s & Folder1.Path & "\" & File1.ListItems.Item(x).Text & vbNullChar
    Next x
    
    Call cFile.CopyFileOrFolder(s, sTo): DoEvents    'copy files
    Call File1.Refresh
End Sub

Private Sub mnuDelete1_Click()
Dim s As String
Dim x As Long

    'créé une string qui concatene tous les fichiers
    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            s = s & Folder1.Path & "\" & File1.ListItems.Item(x).Text & vbNullChar
    Next x
    
    Call cFile.MoveToTrash(s): DoEvents 'à la poubelle
    Call File1.Refresh
End Sub

Private Sub mnuDeleteFolder1_Click()
Dim s As String
Dim x As Long

    'créé une string qui concatene tous les dossiers
    For x = 1 To Folder1.ListCount
        If Folder1.ListItems.Item(x).Selected Then _
            s = s & Folder1.Path & "\" & Folder1.ListItems.Item(x).Text & vbNullChar
    Next x
    
    Call cFile.MoveToTrash(s): DoEvents 'à la poubelle
    Call Folder1.Refresh
End Sub

Private Sub mnuDisplayProperties1_Click()
'affiche les properties des dossiers sélectionnés
Dim x As Long

    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            Call cFile.ShowFileProperty(Folder1.Path & "\" & File1.ListItems.Item(x).Text, Me.hWnd)
    Next x
End Sub

Private Sub mnuDisplayPropertiesFolder1_Click()
'affiche les properties des dossiers sélectionnés
Dim x As Long

    For x = 1 To Folder1.ListCount
        If Folder1.ListItems.Item(x).Selected Then _
            Call cFile.ShowFileProperty(Folder1.Path & "\" & Folder1.ListItems.Item(x).Text, Me.hWnd)
    Next x
End Sub

Private Sub mnuHiddenAttr_Click()
'attribut caché
Dim x As Long

    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            Call cFile.ChangeAttributes(Folder1.Path & "\" & File1.ListItems.Item(x).Text, FILE_ATTRIBUTE_HIDDEN)
    Next x
End Sub

Private Sub mnuMoveFolderTo1_Click()
'déplace les dossiers sélectionnés
Dim s As String
Dim sTo As String
Dim x As Long

    'browse for folder
    sTo = cFile.BrowseForFolder("Choix du dossier cible", Me.hWnd)
    If cFile.FolderExists(sTo) = False Then Exit Sub
    
    'créé une string qui concatene tous les dossiers
    For x = 1 To Folder1.ListCount
        If Folder1.ListItems.Item(x).Selected Then _
            s = s & Folder1.Path & "\" & Folder1.ListItems.Item(x).Text & vbNullChar
    Next x
    
    Call cFile.MoveFileOrFolder(s, sTo): DoEvents  'move files
    Call Folder1.Refresh
End Sub

Private Sub mnuMoveTo1_Click()
'déplace les fichiers sélectionnés
Dim s As String
Dim sTo As String
Dim x As Long

    'browse for folder
    sTo = cFile.BrowseForFolder("Choix du dossier cible", Me.hWnd)
    If cFile.FolderExists(sTo) = False Then Exit Sub
    
    'créé une string qui concatene tous les fichiers
    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            s = s & Folder1.Path & "\" & File1.ListItems.Item(x).Text & vbNullChar
    Next x
    
    Call cFile.MoveFileOrFolder(s, sTo): DoEvents  'move files
    Call File1.Refresh
End Sub

Private Sub mnuNormalAttr_Click()
'attribut normal
Dim x As Long

    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            Call cFile.ChangeAttributes(Folder1.Path & "\" & File1.ListItems.Item(x).Text, FILE_ATTRIBUTE_NORMAL)
    Next x
End Sub

Private Sub mnuOpenExplorer1_Click()
'ouvre explorer
    Shell "explorer.exe " & Folder1.Path, vbNormalFocus
End Sub

Private Sub mnuROattr_Click()
'attribut lecture seule
Dim x As Long

    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            Call cFile.ChangeAttributes(Folder1.Path & "\" & File1.ListItems.Item(x).Text, FILE_ATTRIBUTE_READONLY)
    Next x
End Sub

Private Sub mnuSystemAttr_Click()
'attribut système
Dim x As Long

    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            Call cFile.ChangeAttributes(Folder1.Path & "\" & File1.ListItems.Item(x).Text, FILE_ATTRIBUTE_SYSTEM)
    Next x
End Sub

Private Sub mnuSystemHiddenAttr_Click()
'attribut système (caché)
Dim x As Long

    For x = 1 To File1.ListCount
        If File1.ListItems.Item(x).Selected Then _
            Call cFile.ChangeAttributes(Folder1.Path & "\" & File1.ListItems.Item(x).Text, FILE_ATTRIBUTE_INVISIBLE_SYSTEM)
    Next x
End Sub

Private Sub mnuWithoutSubFolder1_Click()
'on place tous les fichiers des dossiers sélectionnés dans le LV
Dim x As Long
Dim y As Long
Dim s() As String

    'pour chaque dossier sélectionné
    FileR.Visible = False
    
    For x = 1 To Folder1.ListCount
        If Folder1.ListItems.Item(x).Selected Then
            
            'énumère les fichiers
            Call cFile.EnumFilesFromFolder(Folder1.Path & "\" & Folder1.ListItems.Item(x).Text, s(), False)
            
            'ajoute tout çà à la liste
            For y = 1 To UBound(s())
                'ajoute les fichiers à FileR
                s(y) = Replace$(s(y), "\\", "\", , , vbBinaryCompare)   'vire les deux slash
                FileR.AddItemManually s(y), File, bFillSubItemsAuto:=True
            Next y
            
            DoEvents
        End If
    Next x
            
    FileR.Visible = True
    
End Sub

Private Sub mnuWithSubFolders1_Click()
'on place tous les fichiers des dossiers sélectionnés dans le LV
Dim x As Long
Dim y As Long
Dim s() As String

    'pour chaque dossier sélectionné
    FileR.Visible = False
    
    For x = 1 To Folder1.ListCount
        If Folder1.ListItems.Item(x).Selected Then
            
            'énumère les fichiers
            Call cFile.EnumFilesFromFolder(Folder1.Path & "\" & Folder1.ListItems.Item(x).Text, s(), True)
            
            'ajoute tout çà à la liste
            For y = 1 To UBound(s())
                'ajoute les fichiers à FileR
                s(y) = Replace$(s(y), "\\", "\", , , vbBinaryCompare)   'vire les deux slash
                FileR.AddItemManually s(y), File, bFillSubItemsAuto:=True
            Next y
            
            DoEvents
        End If
    Next x
            
    FileR.Visible = True
End Sub

Private Sub pctPath_Change()
    If cFile.FolderExists(cFile.GetFolderFromPath(pctPath.Text & "\")) = False Then
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
        s = pctPath.Text
        If cFile.FolderExists(s) Then Folder1.Path = s
        File1.Path = s
    End If
End Sub

Private Sub chkAddText_Click()

    If chkAddText.Value Then
        'alors on n'ajoute qu'une string fixée
        txtString.Enabled = True
        Label2(0).Enabled = False
        Label2(1).Enabled = False
        cb4.Enabled = False
        cb4b.Enabled = False
    Else
        'alors on n'ajoute pas du simple texte
        txtString.Enabled = False
        Label2(0).Enabled = True
        Label2(1).Enabled = True
        cb4.Enabled = True
        cb4b.Enabled = True
    End If
        
End Sub

Private Sub cmdAdd_Click()
'ajoute à la listebox des modifications la modification sélectionnée
Dim s As String

    If TB.SelectedItem.Index <> Style And TB.SelectedItem.Index <> Remplacer Then
        'alors on sépare les cas modif début/fin
        If optPos5(3).Value Then
            'alors on ajoute au début
            s = "[debut] "
        Else
            'à la fin
            s = "[fin] "
        End If
    End If
    If Option1(0).Value Then
        'modification appliquées au corps du nom
        s = s & "[nom] "
    Else
        'extension
        s = s & "[ext] "
    End If

    Select Case TB.SelectedItem.Index
        Case Style
            s = s & "Style : " & cb1.Text & IIf(cb1.Text = "Première lettre en majuscule", _
                vbNullString, " " & cb1b.Text) & IIf(cb1b.Text = "tous les caractères", _
                vbNullString, " ," & txt1c.Text & "caractères")
        Case Compteur
            s = s & cb2.Text & " partir de " & txtDepCpt.Text & " avec un pas de " & txtStepCpt.Text
        Case Remplacer
            s = s & "Remplacer [" & txtStringToChange.Text & "] par [" & _
                txtNewString.Text & "] " & IIf(chkReplaceCase.Value, "en respectant", _
                "sans respecter") & " la casse"
        Case Base
            If chkAddText.Value Then
                'simple texte
                s = s & "Ajouter (string fixée) : [" & txtString.Text & "]"
            Else
                s = s & "Ajouter (base) : " & cb4.Text & IIf(cb4b.Text = vbNullString, vbNullString, ", " & cb4b.Text)
            End If
        Case Audio
            s = s & "Ajouter (audio) : " & cb5.Text
        Case Video
            s = s & "Ajouter (video) : " & cb6.Text
        Case TYPE_OF_MODIFICATION.Image     '/!\ DO NOT CHANGE
            s = s & "Ajouter (image) : " & cb7.Text
    End Select
    
    'ajoute la string à la liste
    lstTransfo.AddItem s
    
    DisplaySample   'refresh le sample
End Sub

Private Sub cmdDel_Click()
'supprime tous les éléments sélectionnés dans la listbox
Dim x As Long

    For x = lstTransfo.ListCount To 1 Step -1
        If lstTransfo.Selected(x - 1) Then lstTransfo.RemoveItem x - 1
    Next x
    
    lstTransfo.Refresh
    txtTransfo.Text = lstTransfo.List(lstTransfo.ListIndex)
    
    DisplaySample   'refresh le sample
End Sub

Private Sub cmdDown_Click()
'décale vers le bas les sélections

    DisplaySample   'refresh le sample
End Sub

Private Sub cmdOpenList_Click()
'ouvre une liste

    DisplaySample   'refresh le sample
End Sub

Private Sub cmdRenamerAll_Click()
'lance le renommage
Dim x As Long
Dim i As Long
Dim sOld() As String
Dim sNew() As String

    i = 0
    'compte le nombre de changements à appliquer
    For x = lstTransfo.ListCount To 1 Step -1
        If lstTransfo.Selected(x - 1) Then i = i + 1
    Next x
    
    If i = 0 Then
        'pas de changement
        MsgBox "Aucun changement n'a été sélectionné.", vbInformation + vbOKOnly, "Attention"
        Exit Sub
    End If
    
    'des changements à prévoir ==> demande confirmation
    If MsgBox("Lancer le renommage ? (les changements seront effectués pour de vrai)", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    
    ReDim sOld(FileR.ListCount)
    'remplit sOld (anciens noms)
    For x = 1 To FileR.ListCount
        sOld(x) = FileR.ListItems(x).Text
    Next x
    
    'procède au calcul des nouveaux noms
    RenameMyFiles Me.lstTransfo, sOld(), sNew()
    
    'maintenant renomme tout sOld ==> sNew
    
End Sub

Private Sub cmdSaveList_Click()
'sauvegarde une liste

    DisplaySample   'refresh le sample
End Sub

Private Sub cmdSimulate_Click()
'lance la simulation
Dim x As Long
Dim i As Long

    i = 0
    'compte le nombre de changements à appliquer
    For x = lstTransfo.ListCount To 1 Step -1
        If lstTransfo.Selected(x - 1) Then i = i + 1
    Next x
    
    If i = 0 Then
        'pas de changement
        MsgBox "Aucun changement n'a été sélectionné.", vbInformation + vbOKOnly, "Attention"
        Exit Sub
    End If
    
    frmSimul.Show
End Sub

Private Sub cmdUp_Click()
'décale vers le haut les sélections
    
    DisplaySample   'refresh le sample
End Sub

Private Sub File1_DblClick()
Dim cTag As clsTag
Dim cTg As clsTagInfo
Dim Fic As String
Dim s As String
Dim cVid As clsVideo


    Fic = File1.Path & File1.ListItems(File1.ListIndex)
    
    'Set cTg = New clsTagInfo
    'Set cTag = cTg.GetTagsV1(Fic)
    
    's = s & "Album" & "=" & cTag.Album & vbNewLine
    's = s & "Artist" & "=" & cTag.Artist & vbNewLine
    's = s & "Comment" & "=" & cTag.Comment & vbNewLine
    's = s & "Genre" & "=" & cTag.Genre & vbNewLine
    's = s & "strGenre" & "=" & cTag.strGenre & vbNewLine
    's = s & "Title" & "=" & cTag.Title & vbNewLine
    's = s & "TrackV1" & "=" & cTag.TrackV1 & vbNewLine

    'bmp = GetImageInfos("C:\Documents and Settings\Admin\Mes documents\Mes images\3D\3D.2\Côte d'Azur.jpg")
    
    Set cVid = New clsVideo
    cVid.strFile = Fic
    Call cVid.GetVideoInfo
    
    s = s & "lngMaxBytesPerSec" & "=" & cVid.lngMaxBytesPerSec & vbNewLine
    s = s & "strFileType" & "=" & cVid.strFileType & vbNewLine
    s = s & "strDuration" & "=" & cVid.strDuration & vbNewLine
    s = s & "lngModifiedStreams" & "=" & cVid.lngModifiedStreams & vbNewLine
    s = s & "lngSamplesPerSecond" & "=" & cVid.lngSamplesPerSecond & vbNewLine

    
    'MsgBox s
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then File1.Refresh
End Sub

Private Sub File1_ItemDblSelection(Item As ComctlLib.ListItem)
'ajoute à la liste courante
Dim s As String
    
    ValidateRect File1.hWnd, 0 'gèle l'affichage pour éviter le clignotement
    FileR.AddItemManually File1.Path & "\" & Item.Text, File, bFillSubItemsAuto:=True
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Lst As ListItem

    Set Lst = File1.HitTest(x, y)
    
    If Button = 2 Then
        'popup menu
        If Not (Lst Is Nothing) Then Lst.Selected = True
        Me.PopupMenu Me.mnuPopupFile1
    End If
    
    Set Lst = Nothing
End Sub

Private Sub Folder1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Lst As ListItem

    Set Lst = Folder1.HitTest(x, y)
    
    If Button = 2 Then
        'popup menu
        If Not (Lst Is Nothing) Then Lst.Selected = True
        Me.PopupMenu Me.mnuPopupFolder1
    End If
    
    Set Lst = Nothing
End Sub

Private Sub Folder1_PathChange(sOldPath As String, sNewPath As String)
'mets à jour le combo
    pctPath.Text = sNewPath
    File1.Path = Folder1.Path
End Sub

Private Sub Form_Load()
    
    '//organise les controles
    Folder1.Path = Left$(App.Path, 3)
    FileR.Path = "-inexistant-" 'pas de path
    pctPath.Text = Folder1.Path
    
End Sub

Private Sub lstTransfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'popup menu
    If Button = 2 Then Me.PopupMenu Me.mnuPopupTransfoList
    
    txtTransfo.Text = lstTransfo.List(lstTransfo.ListIndex)
End Sub

Private Sub mnuAbout_Click()
'a propos
    frmAbout.Show vbModal
End Sub

Private Sub mnuAddToCurrentList1_Click()
'ajoute les fichiers à la liste courante
Dim var As Variant
Dim x As Long
    
    'obtient la liste des sélections
    File1.GetSelectedItems var
    
    FileR.Visible = False
    For x = 1 To UBound(var)
        'ajoute les fichiers à FileR
        FileR.AddItemManually File1.Path & var(x).Text, File, bFillSubItemsAuto:=True
    Next x
    FileR.Visible = True

End Sub

Private Sub mnuChangeFolder1_Click()
'changer le dossier 1
Dim s As String

    s = cFile.BrowseForFolder("Choix du répertoire", Me.hWnd)   'browse for folder
    If cFile.FolderExists(s) Then Folder1.Path = s  'change le folder
    File1.Path = s
End Sub

Private Sub mnuDeselectAll_Click()
'déselectionne tous les éléments sélectionnés dans la listbox
Dim x As Long

    For x = lstTransfo.ListCount To 1 Step -1
        lstTransfo.Selected(x - 1) = False
    Next x
    
    lstTransfo.Refresh
End Sub

Private Sub mnuHelp_Click()
'aide

    '/!\ l'aide N'EST PAS disponible pour ce projet hors du contexte de Hex Editor VB
    MsgBox "L'aide est indisponible pour ce projet seul. Svp posez votre question sur www.vbfrance.com.", vbInformation + vbOKOnly, "Aucune aide trouvée"
End Sub

Private Sub mnuQuit_Click()
'quitte le prog
    Unload Me
    EndProg
End Sub

Private Sub mnuSelectAll_Click()
'sélectionne tous les éléments sélectionnés dans la listbox
Dim x As Long

    For x = lstTransfo.ListCount To 1 Step -1
        lstTransfo.Selected(x - 1) = True
    Next x
    
    lstTransfo.Refresh  'évite un bug d'affichage après le cochage des cases
    
End Sub

Private Sub pctPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0  'vire le BEEP
End Sub

Private Sub TB_Click()
'on change le Frame visible
Dim x As Long
Dim i As TYPE_OF_MODIFICATION

    i = TB.SelectedItem.Index

    For x = 0 To Frame3.UBound
        Frame3(x).Visible = False
    Next x
    Frame3(i - 1).Visible = True
    
    'rend enabled ou pas l'option début/fin
    optPos5(3).Enabled = ((i <> Style) And (i <> Remplacer))
    optPos5(2).Enabled = ((i <> Style) And (i <> Remplacer))
    Frame5(2).Enabled = ((i <> Style) And (i <> Remplacer))
End Sub




'=======================================================
'PROCEDURE & FUNCTIONS
'=======================================================

'=======================================================
'permet d'afficher dans txtproto le résultat final des modifications
'/!\ n'affiche que les ajouts de string
'=======================================================
Private Sub DisplaySample()
'
End Sub

Private Sub txt1c_Change()
'formate le texte
    txt1c.Text = CStr(Abs(Int(Val(txt1c.Text))))
End Sub

Private Sub txtDepCpt_Change()
'formate le texte
    txtDepCpt.Text = CStr(Abs(Int(Val(txtDepCpt.Text))))
End Sub

Private Sub txtNewString_Change()
    If IsFileNameOK(txtNewString.Text) = False Then
        'un caractère mauvais
        MsgBox "Les caractères \ /:*?" & Chr$(34) & "<>| sont interdits", vbInformation, "Saisie de texte"
        txtNewString.Text = vbNullString
    End If
End Sub

Private Sub txtStepCpt_Change()
'formate le texte
    txtStepCpt.Text = CStr(Abs(Int(Val(txtStepCpt.Text))))
End Sub

Private Sub txtString_Change()
    If IsFileNameOK(txtString.Text) = False Then
        'un caractère mauvais
        MsgBox "Les caractères \ /:*?" & Chr$(34) & "<>| sont interdits", vbInformation, "Saisie de texte"
        txtString.Text = vbNullString
    End If
End Sub
