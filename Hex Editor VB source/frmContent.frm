VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C9771C4C-85A3-44E9-A790-1B18202DA173}#1.0#0"; "FileView_OCX.ocx"
Object = "{16DCE99A-3937-4772-A07F-3BA5B09FCE6E}#1.1#0"; "vkUserControlsXP.ocx"
Begin VB.MDIForm frmContent 
   BackColor       =   &H8000000C&
   Caption         =   "Hex Editor VB  --- PRE ALPHA v1.6"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9645
   HelpContextID   =   5
   Icon            =   "frmContent.frx":0000
   LinkTopic       =   "Editeur hexad�cimal"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin vkUserContolsXP.vkSysTray vkSysTray 
      Left            =   1440
      Top             =   4920
      _ExtentX        =   794
      _ExtentY        =   794
      Icon            =   "frmContent.frx":5E8A
   End
   Begin VB.PictureBox pctToolbar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      Picture         =   "frmContent.frx":7954
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   643
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   9645
   End
   Begin VB.PictureBox pctConsole 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   9645
      TabIndex        =   5
      Top             =   2535
      Visible         =   0   'False
      Width           =   9645
      Begin RichTextLib.RichTextBox txtE 
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   3000
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmContent.frx":935E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txt 
         Height          =   2895
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5106
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         OLEDragMode     =   0
         OLEDropMode     =   1
         TextRTF         =   $"frmContent.frx":93E1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   89
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":9464
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":97B6
            Key             =   "Fichier|D�sassembler..."
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":9B08
            Key             =   "Outils|D�sassembleur..."
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":9E5A
            Key             =   "Affichage|ConsoleF4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":A3AC
            Key             =   "Fen�tres|Gestion des fen�tres..."
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":A6FE
            Key             =   "Outils|Statistiques du fichier..."
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":AA50
            Key             =   "Outils|R�cup�ration de fichiers..."
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":ADA2
            Key             =   "Outils|Ouvrir avec le bloc-notes"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":B0F4
            Key             =   "Outils|Calculatrice"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":B446
            Key             =   "Nouveau|Nouveau fichier..."
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":B798
            Key             =   "Position|Fin"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":BAEA
            Key             =   "Nouveau|D�marrer un processus..."
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":BE3C
            Key             =   "Outils|Convertisseur..."
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":C18E
            Key             =   "Outils|Renommage massif de fichiers..."
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":C4E0
            Key             =   "Position|Monter d'une page"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":C832
            Key             =   "Fichier|Ex�cuter"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":CB84
            Key             =   "Outils|Ex�cuter le scriptF9"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":CED6
            Key             =   "Position|D�but"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":D228
            Key             =   "Position|Descendre d'une page"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":D57A
            Key             =   "Affichage|Tableau_checked"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":D8CC
            Key             =   "Aide|A propos"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":DE1E
            Key             =   "Ouvrir|Ouvrir un processus en m�moire..."
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":E170
            Key             =   "Ouvrir|Ouvrir un disque..."
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":E4C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":E814
            Key             =   "Signets|Ouvrir une liste de signets..."
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":EB66
            Key             =   "Signets|Enregistrer la liste des signets..."
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":EEB8
            Key             =   "Signets|Ajouter une liste de signets..."
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":F20A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":F55C
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":F8AE
            Key             =   "Outils|D�marrer une t�che..."
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":FC00
            Key             =   "Edition|Coller"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":FF52
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":102A4
            Key             =   "Fichier|Imprimer..."
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":105F6
            Key             =   "Edition|Couper"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":10948
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":10C9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":10FEC
            Key             =   "Fichier|Ouvrir"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1133E
            Key             =   "Outils|Gestion des processus..."
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":11690
            Key             =   "Ouvrir|Ouvrir des fichiers..."
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":119E2
            Key             =   "Ouvrir|Ouvrir un dossier de fichiers..."
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":11D34
            Key             =   "Fichier|Nouveau"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":12086
            Key             =   "Edition|Copier"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":123D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1272A
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":12A7C
            Key             =   "Signets|Basculer un signet"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":12DCE
            Key             =   "Signets|Signet pr�c�dent"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":13120
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":13472
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":137C4
            Key             =   "Signets|Signet suivant"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":13B16
            Key             =   "Edition|Visualiser une partie restreinte..."
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":13E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":141BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1450C
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1485E
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":14BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":14F02
            Key             =   "Outils|D�couper/fusionner des fichiers..."
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":15254
            Key             =   "Signets|Supprimer le signet de l'offset"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":155A6
            Key             =   "Signets|Supprimer tous les signets"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":158F8
            Key             =   "Outils|Suppression de fichiers..."
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":15C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":15F9C
            Key             =   "Position|Aller � l'offset..."
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":162EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":16640
            Key             =   "Rechercher|Chaines de caract�res..."
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":16992
            Key             =   "Outils|Recherche de fichiers..."
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":16CE4
            Key             =   "Aide|Aide...F1"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":17036
            Key             =   "Aide|Rap"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":17388
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":176DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":17A2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":17D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":180D0
            Key             =   "Aide|Faire un don..."
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":18422
            Key             =   "Fichier|Imprimer"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":18774
            Key             =   "Outils|Editeur de script"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":18AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":18E18
            Key             =   "Fichier|Propri�t�s"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1916A
            Key             =   "Edition|Refaire"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":194BC
            Key             =   "Fichier|Enregistrer"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1980E
            Key             =   "Fichier|Enregistrer sous..."
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":19B60
            Key             =   "Edition|Cr�er un fichier depuis la s�lection..."
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":19EB2
            Key             =   "Rechercher|Texte..."
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1A204
            Key             =   "Rechercher|Valeurs hexa..."
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1A556
            Key             =   "Edition|Tout s�lectionner"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1A8A8
            Key             =   "Edition|Remplir la s�lection..."
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1ABFA
            Key             =   "Outils|Editeur de langue..."
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1AF4C
            Key             =   "Affichage|Tableau"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1B29E
            Key             =   "Outils|Options..."
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1B5F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1B942
            Key             =   "Edition|Annuler"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1BC94
            Key             =   "Aide|Hex Editor VB sur Internet"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2760
   End
   Begin VB.PictureBox pctExplorer 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   9645
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   330
      Visible         =   0   'False
      Width           =   9645
      Begin VB.TextBox pctPath 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
      Begin FileView_OCX.FileView LV 
         Height          =   2055
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3625
         Path            =   "C:\"
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
   Begin ComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7410
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14993
            MinWidth        =   14993
            Text            =   "Status=[Ready]"
            TextSave        =   "Status=[Ready]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Ouvertures=[0]"
            TextSave        =   "Ouvertures=[0]"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "23:35"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "29/06/2007"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
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
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1BFE6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1D978
            Key             =   ""
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1F30A
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":20C9C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2262E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":23FC0
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2455A
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":24AF4
            Key             =   "Signet"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":26486
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":27E18
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":297AA
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2B13C
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2CACE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2E460
            Key             =   "Trash"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2FDF2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":31784
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":31D1E
            Key             =   "FileOpen"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":336B0
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":35042
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":355DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":35B76
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":36087
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "Cr�er un nouveau fichier"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenFile"
            Object.ToolTipText     =   "Ouvrir un ou plusieurs fichiers"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HomeOpen"
            Object.ToolTipText     =   "Ouvrir un dossier"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Sauvegarder l'objet ouvert"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Search"
            Object.ToolTipText     =   "Effectuer une recherche dans l'objet"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Cut"
            Object.ToolTipText     =   "Couper la s�lection"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copier la s�lection"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Coller le contenu du presse-papier"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "D�faire"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            Object.ToolTipText     =   "Refaire"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Signet"
            Object.ToolTipText     =   "Basculer le signet � l'offset s�lectionn�"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Up"
            Object.ToolTipText     =   "Aller au signet pr�c�dent"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Down"
            Object.ToolTipText     =   "Aller au signet suivant"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert"
            Object.ToolTipText     =   "Afficher la fen�tre de conversion"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Object.ToolTipText     =   "Configuration du logiciel"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu rmnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuRNew 
         Caption         =   "&Nouveau"
         Begin VB.Menu mnuNew 
            Caption         =   "&Nouveau fichier..."
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuNewProcess 
            Caption         =   "&D�marrer un processus..."
         End
      End
      Begin VB.Menu mnuROpen 
         Caption         =   "&Ouvrir"
         Begin VB.Menu mnuOpen 
            Caption         =   "&Ouvrir des fichiers..."
            Shortcut        =   ^O
         End
         Begin VB.Menu mnuOpenFolder 
            Caption         =   "&Ouvrir un dossier de fichiers..."
         End
         Begin VB.Menu mnuOpenProcess 
            Caption         =   "&Ouvrir un processus en m�moire..."
         End
         Begin VB.Menu mnuOpenDisk 
            Caption         =   "&Ouvrir un disque..."
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Enregistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Enregistrer sous..."
      End
      Begin VB.Menu rmnuExport 
         Caption         =   "&Exporter"
         Begin VB.Menu mnuExport 
            Caption         =   "&Le fichier entier..."
         End
         Begin VB.Menu mnuExportSel 
            Caption         =   "&La s�lection..."
         End
      End
      Begin VB.Menu mnuFileTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "&Ex�cuter"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDisAsmThisFile 
         Caption         =   "&D�sassembler..."
      End
      Begin VB.Menu munFileTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileTiret10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperty 
         Caption         =   "&Propri�t�s"
      End
      Begin VB.Menu mnuFileTiret3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "&Tout fermer"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu rmnuEdit 
      Caption         =   "&Edition"
      Enabled         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "&Annuler"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Refaire"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "&Couper"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copier"
         Begin VB.Menu mnuCopyASCII 
            Caption         =   "&Valeurs ASCII format�es"
         End
         Begin VB.Menu mnuCopyASCII2 
            Caption         =   "&Valeurs ASCII format�es bas niveau"
         End
         Begin VB.Menu mnuCopyASCIIReal 
            Caption         =   "&Valeurs ASCII r�elles"
            Shortcut        =   ^C
         End
         Begin VB.Menu mnuCopyhexa 
            Caption         =   "&Valeurs hexa"
         End
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Coller"
         Begin VB.Menu mnuPasteHexa 
            Caption         =   "&Valeurs hexa"
         End
         Begin VB.Menu mnuPasteASCII 
            Caption         =   "&Valeurs ASCII"
            Shortcut        =   ^V
         End
      End
      Begin VB.Menu mnuEditTiret212 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThisIsTheBeginnig 
         Caption         =   "&D�signer comme d�but de s�lection"
      End
      Begin VB.Menu mnuThisIsTheEnd 
         Caption         =   "D�signer comme fin de s�lection"
      End
      Begin VB.Menu mnuEditTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Ins�rer..."
      End
      Begin VB.Menu mnuEditTiret3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Tout s�lectionner"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelectZone 
         Caption         =   "&S�lectionner une zone..."
      End
      Begin VB.Menu mnuSelectFromByte 
         Caption         =   "&S�lectionner � partir du byte..."
      End
      Begin VB.Menu mnuEditTiret41 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFillSelection 
         Caption         =   "&Remplir la s�lection..."
      End
      Begin VB.Menu mnuDeleteSelection 
         Caption         =   "&Supprimer la s�lection"
      End
      Begin VB.Menu mnuCreateFileFromSelelection 
         Caption         =   "&Cr�er un fichier depuis la s�lection..."
      End
      Begin VB.Menu mnuEditTiret4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowNotAllFile 
         Caption         =   "&Visualiser une partie restreinte..."
      End
   End
   Begin VB.Menu rmnuFind 
      Caption         =   "&Rechercher"
      Enabled         =   0   'False
      Begin VB.Menu mnuSearchT 
         Caption         =   "&Texte..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchH 
         Caption         =   "&Valeurs hexa..."
      End
      Begin VB.Menu mnuFinTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplaceT 
         Caption         =   "&Remplacer texte..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuReplaceH 
         Caption         =   "&Remplacer valeurs hexa..."
      End
      Begin VB.Menu mnuFindTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchForString 
         Caption         =   "&Chaines de caract�res..."
      End
      Begin VB.Menu mnuFindTiret3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoOn 
         Caption         =   "Continuer la recherche"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu rmnuDisplay 
      Caption         =   "&Affichage"
      Begin VB.Menu mnuTab 
         Caption         =   "&Tableau"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInformations 
         Caption         =   "&Informations"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditTools 
         Caption         =   "&Donn�e"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExploreDisplay 
         Caption         =   "&Explorateur de fichiers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExploreDisk 
         Caption         =   "&Explorateur de disque"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuShowIcons 
         Caption         =   "&Icones du fichier"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileViewMode 
         Caption         =   "&Mode ""Lecture de fichier"""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisplayTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTdec2Ascii 
         Caption         =   "&Table hex<-->ASCII"
      End
      Begin VB.Menu mnuTableMulti 
         Caption         =   "&Table multibase"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuDisplayTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowConsole 
         Caption         =   "&Console"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuDisplayTiret215 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusOK 
         Caption         =   "&R�initialiser le status"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuRefreh 
         Caption         =   "&Rafra�chir"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu rmnuPos 
      Caption         =   "&Position"
      Enabled         =   0   'False
      Begin VB.Menu mnuDown 
         Caption         =   "&Descendre d'une page"
      End
      Begin VB.Menu muUp 
         Caption         =   "&Monter d'une page"
      End
      Begin VB.Menu mnuPosTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeginning 
         Caption         =   "&Fin"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&D�but"
      End
      Begin VB.Menu mnuPosTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToOffset 
         Caption         =   "&Aller � l'offset..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuMoveOffset 
         Caption         =   "&D�placer l'offset..."
      End
   End
   Begin VB.Menu rmnuSignets 
      Caption         =   "&Signets"
      Enabled         =   0   'False
      Begin VB.Menu mnuAddSignet 
         Caption         =   "&Basculer un signet"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRemoveSignet 
         Caption         =   "&Supprimer le signet de l'offset"
      End
      Begin VB.Menu mnuSignetTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveAll 
         Caption         =   "&Supprimer tous les signets"
      End
      Begin VB.Menu mnuSignetTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSignetPrev 
         Caption         =   "&Signet pr�c�dent"
      End
      Begin VB.Menu mnuSignetNext 
         Caption         =   "&Signet suivant"
      End
      Begin VB.Menu mnuSignetsTiret3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenSignetsList 
         Caption         =   "&Ouvrir une liste de signets..."
      End
      Begin VB.Menu mnuAddSignetIn 
         Caption         =   "&Ajouter une liste de signets..."
      End
      Begin VB.Menu mnuSaveSignets 
         Caption         =   "&Enregistrer la liste des signets..."
      End
      Begin VB.Menu mnuTiretAgainAndAgain 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGestSignets 
         Caption         =   "&Gestionnaire de signets..."
      End
   End
   Begin VB.Menu rmnuTools 
      Caption         =   "&Outils"
      Begin VB.Menu mnuHome 
         Caption         =   "&D�marrer une t�che..."
      End
      Begin VB.Menu mnuToolsTiret_moins1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditScript 
         Caption         =   "&Editeur de script"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExecuteScript 
         Caption         =   "&Ex�cuter le script"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuToolsTiret0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "&Calculatrice"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "&Convertisseur..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuOpenInBN 
         Caption         =   "&Ouvrir avec le bloc-notes"
      End
      Begin VB.Menu mnuStats 
         Caption         =   "&Statistiques du fichier..."
      End
      Begin VB.Menu mnuInterpretAdvanced 
         Caption         =   "&Conversion avanc�e..."
      End
      Begin VB.Menu mnuToolsTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompareFiles 
         Caption         =   "&Comparaison de fichiers..."
      End
      Begin VB.Menu mnuChangeDates 
         Caption         =   "&Changer les dates d'un fichier..."
      End
      Begin VB.Menu mnuShredder 
         Caption         =   "&Suppression de fichiers..."
      End
      Begin VB.Menu mnuSanitDisk 
         Caption         =   "&Sanitization..."
      End
      Begin VB.Menu mnuRecoverFiles 
         Caption         =   "&R�cup�ration de fichiers..."
      End
      Begin VB.Menu mnuToolsTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProcesses 
         Caption         =   "&Gestion des processus..."
      End
      Begin VB.Menu mnuDiskInfos 
         Caption         =   "&Informations sur les disques..."
      End
      Begin VB.Menu mnuToolsTiret3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRenamer 
         Caption         =   "&Renommage massif de fichiers..."
      End
      Begin VB.Menu mnuCutCopyFiles 
         Caption         =   "&D�couper/fusionner des fichiers..."
      End
      Begin VB.Menu mnuFileSearch 
         Caption         =   "&Recherche de fichiers..."
      End
      Begin VB.Menu mnuCreateISOFile 
         Caption         =   "&Cr�er un fichier ISO depuis un disque..."
      End
      Begin VB.Menu mnuToolsTiret4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisAsm 
         Caption         =   "&D�sassembleur..."
      End
      Begin VB.Menu mnuToolsTiret41 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLangEditor 
         Caption         =   "&Editeur de langue..."
      End
      Begin VB.Menu mnuToolsTiret4112 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu rmnuWindow 
      Caption         =   "&Fen�tres"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&En cascade"
      End
      Begin VB.Menu mnuMH 
         Caption         =   "Mosa�que &horizontale"
      End
      Begin VB.Menu mnuMV 
         Caption         =   "Mosa�que &verticale"
      End
      Begin VB.Menu mnuReorganize 
         Caption         =   "&R�organiser les icones"
      End
      Begin VB.Menu mnuWindowsTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGestWindows 
         Caption         =   "&Gestion des fen�tres..."
      End
   End
   Begin VB.Menu rmnuHelp 
      Caption         =   "&Aide"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Aide..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpTiret7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLangMenu 
         Caption         =   "&Langue"
         Begin VB.Menu mnuLang 
            Caption         =   "&Fran�ais"
            Index           =   1
         End
      End
      Begin VB.Menu mnuHelpTiret 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErr 
         Caption         =   "&Rapport d'erreurs"
      End
      Begin VB.Menu mnuHelpTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternetSection 
         Caption         =   "&Hex Editor VB sur Internet"
         Begin VB.Menu mnuHelpForum 
            Caption         =   "&Forum de demande d'aide..."
         End
         Begin VB.Menu mnuFreeForum 
            Caption         =   "&Forum de discussion..."
         End
         Begin VB.Menu mnuSourceForge 
            Caption         =   "&Page SourceForge.net du projet..."
         End
         Begin VB.Menu mnuVbfrance 
            Caption         =   "&Hex Editor VB sur vbfrance.com..."
         End
      End
      Begin VB.Menu mnuInternetTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVersions 
         Caption         =   "&Versions..."
      End
      Begin VB.Menu mnuInternetTiret10101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Rechercher une mise � jour..."
      End
      Begin VB.Menu mnuInternetTiret17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&A propos"
      End
   End
   Begin VB.Menu mnuPopupExplore 
      Caption         =   "Popup_Explore"
      Visible         =   0   'False
      Begin VB.Menu mnuEditSelection 
         Caption         =   "&Editer les fichiers s�lectionn�s"
      End
      Begin VB.Menu mnuStatsPopup 
         Caption         =   "&Statistiques du fichier..."
      End
      Begin VB.Menu mnuPopupTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenSelectedFiles 
         Caption         =   "&Ouvrir les fichiers s�lectionn�s"
      End
      Begin VB.Menu mnuPopupTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenExplorer 
         Caption         =   "&Ouvrir Explorer � cet endroit..."
      End
   End
   Begin VB.Menu mnuPopupDisk 
      Caption         =   "mnuPopupDisk"
      Visible         =   0   'False
      Begin VB.Menu mnuPrevClust 
         Caption         =   "&Cluster pr�c�dent"
      End
      Begin VB.Menu mnuNextClust 
         Caption         =   "&Cluster suivant"
      End
      Begin VB.Menu mnuPopupTiret178 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrevSect 
         Caption         =   "&Secteur pr�c�dent"
      End
      Begin VB.Menu mnuNextSect 
         Caption         =   "&Secteur suivant"
      End
      Begin VB.Menu mnuPopupTiret278 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeginingPart 
         Caption         =   "&D�but de partition"
      End
      Begin VB.Menu mnuEndPart 
         Caption         =   "&Fin de partition"
      End
   End
   Begin VB.Menu mnuPopupIcon 
      Caption         =   "mnuPopupIcon"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveIconAsBitmap 
         Caption         =   "&Enregistrer l'icone en bitmap..."
      End
      Begin VB.Menu mnuCopyBitmapToClipBoard 
         Caption         =   "&Copier l'image dans le presse papier"
      End
   End
End
Attribute VB_Name = "frmContent"
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
'FORM PARENT QUI CONTIENT LES FORM D'EDITION
'FICHIER/MEMOIRE
'CONTIENT LES MENUS
'=======================================================

Implements IOverMenuEvent
Private bDonneeForm As Boolean
Private clsPref As clsIniForm
Public Lang As New clsLang


Private Sub cSubEvent_MenuOver(ByVal strCaption As String)
    'cet event est lib�r� lors du survol des menus
    Sb.Panels(1).Text = strCaption
End Sub

'=======================================================
'sub qui sera activ�e lors du survol du menu
'=======================================================
Private Sub IOverMenuEvent_OnMenuOver(ByVal strCaption As String)
    Me.Caption = strCaption
End Sub

Private Sub LV_ItemDblSelection(Item As ComctlLib.ListItem)
Dim Frm As Form

    'ouvre un fichier

    On Error GoTo ErrGestion

    'demande le fichier

    If cFile.FileExists(Item.Tag) = False Then Exit Sub
    
    'affiche une nouvelle fen�tre
    Set Frm = New Pfm
    Call Frm.GetFile(Item.Tag)
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
    
    Exit Sub
ErrGestion:
End Sub

'=======================================================
'affiche le path du fichier s�lectionn� dans la picturbox
'=======================================================
Private Sub DisplayPath()
Dim S As String
Dim l As Long

    'r�cup�re le texte � afficher
    S = LV.Path & "\"
    S = Replace$(S, "\\", "\")  'vire le double slash
    
    'enl�ve la partie apr�s le vbNullChar de la string
    l = InStr(1, S, vbNullChar)
    If l > 0 Then
        S = Left$(S, l)
    End If
    
    'affiche la string dans la picturebox
    pctPath.Text = cFile.GetFolderName(S)
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'suppression des fichiers
    If KeyCode = vbKeyDelete Then
        Call LV.DeleteSelectedItemsFromDisk(False, , , True, True)
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'popup menu
        Me.PopupMenu Me.mnuPopupExplore, , X + LV.Left, Y + LV.Top + 300
    End If
End Sub

Private Sub LV_PathChange(sOldPath As String, sNewPath As String)
    Call DisplayPath 'affiche le path dans la "barre d'adresse"
End Sub

Private Sub MDIForm_Activate()

    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entr�es dans les menus
    
    'ferme le splash screen si il �tait encore ouvert
    bEndSplash = True
        
End Sub

Private Sub MDIForm_DblClick()
'montre la form de d�marrage rapide
    frmHome.Show
    Call SetFormForeBackGround(frmHome, SetFormForeGround)
End Sub

Private Sub MDIForm_Load()
Dim X As Long

    On Error Resume Next
    
    Set clsPref = New clsIniForm
    
    frmSplash.lblState.Caption = "R�cup�ration des fichiers de langue..."
    
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
    
    'chargement des menus de langue (sLang())
    For X = 1 To UBound(sLang())
        'ajoute une entr�e au menu
        Load Me.mnuLang(X)
        Me.mnuLang(X).Caption = Left$(cFile.GetFileName(sLang(X)), _
            Len(cFile.GetFileName(sLang(X))) - 4)
    Next X
    
    'coche le bon menu
    For X = 1 To mnuLang.Count
        If Replace$(Me.mnuLang(X).Caption, "&", vbNullString) = cPref.env_Lang _
            Then Me.mnuLang(X).Checked = True
    Next X
    
    'colorise les menus
    'Call ColorFormMenu(Me, cPref.general_MenuBackColor)
    
    'applique une image dans le Toolbar
    'If cPref.general_ToolbarPCT Then Call ColorToolbar(Me.Toolbar1, _
        Me.pctToolbar.Picture.Handle)
        
    'on met l'icone dans le TraySystem
    With vkSysTray
        .BalloonTipString = Me.Caption
        Call .AddToTray(0)
    End With

    'loading des preferences
    frmSplash.lblState.Caption = Lang.GetString("_LoadingPref")
    Call clsPref.GetFormSettings(App.Path & "\Preferences\FrmContent.ini", Me)
    
    'valeurs par d�faut
    ReDim strConsoleText(0)
    lngConsolePos = 0
    txt.SelColor = vbWhite
    txt.Refresh
    
    'lance le subclassing pour le resize des pictureboxes
    frmSplash.lblState.Caption = "Starting subclassing..."
    Call HookPictureResizement(Me.pctConsole, 1)
    Call HookPictureResizement(Me.pctExplorer)
    
    #If USE_FRMC_SUBCLASSING Then
        'instancie les classes
        Set cSub = New clsFrmSubClass
        
        'd�marre le hook de la form
        Call cSub.HookFormMenu(Me, True)
    #End If
    
    frmSplash.lblState.Caption = Lang.GetString("_CheckForEXEs")
    'v�rifie la pr�sence de FileRenamer.exe
    If cFile.FileExists(App.Path & "\FileRenamer.exe") = False Then
        Me.mnuFileRenamer.Enabled = False
    Else
        Me.mnuFileRenamer.Enabled = True
    End If
    'v�rifie la pr�sence de Disassembler.exe
    If cFile.FileExists(App.Path & "\Disassembler.exe") = False Then
        Me.mnuDisAsm.Enabled = False
    Else
        Me.mnuDisAsm.Enabled = True
    End If
    'v�rifie la pr�sence de LangEditor
    If cFile.FileExists(App.Path & "\LangEditor.exe") = False Then
        Me.mnuLangEditor.Enabled = False
    Else
        Me.mnuLangEditor.Enabled = True
    End If
    
    'ajoute les icones aux menus
    frmSplash.lblState.Caption = Lang.GetString("_AddIconsMenus")
    Call AddIconsToMenus(Me.hWnd, Me.ImageList2)
        
    lNbChildFrm = 0
    
    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entr�es dans les menus
    
    frmSplash.lblState.Caption = Lang.GetString("_ReadPref")
    With cPref
        Me.mnuEditTools.Checked = .general_DisplayData
        Me.mnuInformations.Checked = .general_DisplayInfos
        Me.mnuShowConsole.Checked = .console_Load
        Me.pctConsole.Visible = CBool(.general_DisplayExplore)
        Me.mnuExploreDisplay.Checked = .general_DisplayExplore
        Me.mnuShowIcons.Checked = .general_DisplayIcon
    End With
    
    'loading des pref de la console
    With cPref
        Me.pctConsole.BackColor = .console_BackColor
        frmContent.txt.BackColor = .console_BackColor
        frmContent.txtE.BackColor = .console_BackColor
        Me.pctConsole.Height = .console_Heigth
        Me.mnuShowConsole.Checked = CBool(.console_Load)
        Me.pctConsole.Visible = CBool(.console_Load)
    End With
    
    
    If cPref.general_DisplayExplore Then
        frmSplash.lblState.Caption = Lang.GetString("_ExploLau")
        
        'loading de la taille de l'explorer
        Me.pctExplorer.Height = cPref.explo_Height
        
        'charge les prefs de l'explorer
        '/!\ C'est ce code qui fait charger le logiciel lentement
        '==> on cache le LV
        With LV
            .BlockDisplay = True    'empeche de refresh plusieurs fois
            .Visible = False
            
            .Height = cPref.explo_Height - 145
            
            If cPref.explo_DefaultPath = Lang.GetString("_ProgramDir!") Then
                'alors c'est dans app.path
                .Path = App.Path
            Else
                'alors un dossier perso
                .Path = cPref.explo_DefaultPath
            End If
            .ShowEntirePath = CBool(cPref.explo_ShowPath)
            .ShowHiddenDirectories = CBool(cPref.explo_ShowHiddenFolders)
            .ShowHiddenFiles = CBool(cPref.explo_ShowHiddenFiles)
            .ShowSystemDirectories = CBool(cPref.explo_ShowSystemFodlers)
            .ShowSystemFiles = CBool(cPref.explo_ShowSystemFiles)
            .ShowReadOnlyDirectories = CBool(cPref.explo_ShowROFolders)
            .ShowReadOnlyFiles = CBool(cPref.explo_ShowROFiles)
            .AllowMultiSelect = CBool(cPref.explo_AllowMultipleSelection)
            .AllowFileDeleting = CBool(cPref.explo_AllowFileSuppression)
            .Pattern = cPref.explo_Pattern
            .HideColumnHeaders = CBool(cPref.explo_HideColumnTitle)
            Select Case cPref.explo_IconType
                Case 0
                    .DisplayIcons = BasicIcons
                Case 1
                    .DisplayIcons = FileIcons
                Case 2
                    .DisplayIcons = NoIcons
            End Select
            
            .BlockDisplay = False 'lib�re le refresh
            'Call .Refresh
            .Visible = True
            .RefreshListViewOnly    '/!\ DO NOT REMOVE
        End With
    End If

    'ajoute du texte � la console
    Call AddTextToConsole("Hex Editor VB is ready")
    
    #If Not (FINAL_VERSION) Then
        lngTimeLoad = GetTickCount - lngTimeLoad
        Call AddTextToConsole("Application loaded in " & Trim$(Str$(lngTimeLoad)) & " ms")
    #End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Frm As Form

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_Closing"))
    
    Call SaveQuickBackupINIFile     'permet de sauver (si n�cessaire) l'�tat du programme
    
    'on cache notre form pour �viter de montrer que le programme met du temps � se fermer...
    Me.Hide

    On Error Resume Next
    
    'ferme toute les fen�tres
    For Each Frm In Forms
        If (TypeOf Frm Is Pfm) Or (TypeOf Frm Is diskPfm) Or (TypeOf Frm Is MemPfm) _
            Or (TypeOf Frm Is physPfm) Then
            Call SendMessage(Frm.hWnd, WM_CLOSE, 0, 0)
        End If
    Next Frm
    
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\FrmContent.ini", _
        frmContent)

End Sub

Private Sub mnuCopyBitmapToClipBoard_Click()
'enregistre l'icone de l'active form en bitmap
Dim S As String

    If Me.ActiveForm Is Nothing Then Exit Sub
    If TypeOfForm(Me.ActiveForm) <> "Fichier" And TypeOfForm(Me.ActiveForm) <> _
        "Processus" Then Exit Sub
    
    'sauvegarder l'icone s�lectionn�e en bitmap
    If Me.ActiveForm.lvIcon.SelectedItem Is Nothing Then Exit Sub
    
    'pose l'image sur le picturebox
    ImageList_Draw Me.ActiveForm.IMG.hImageList, Me.ActiveForm.lvIcon.SelectedItem.Index - 1, _
        Me.ActiveForm.pct.hdc, 0, 0, ILD_TRANSPARENT
   
    If Me.ActiveForm.pct.Picture Is Nothing Then Exit Sub
    
    'copie dans le presse papier
    Call Clipboard.Clear
    Call Clipboard.SetData(Me.ActiveForm.pct.Image)
    
    Set Me.ActiveForm.pct.Picture = Nothing
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_IconCopy"))
End Sub

Private Sub mnuCreateISOFile_Click()
'cr�ation d'un fichier ISO
    frmISO.Show vbModal
End Sub

Public Sub mnuCut_Click()
'coupe la s�lection

End Sub

Private Sub mnuDeleteSelection_Click()
'efface les �l�ments de la s�lection
    
    'affiche la boite de demande de confirmation
    frmCreateBackup.Show vbModal
    
    'proc�de � la suppression de la zone et sauvegarde dans un backup
    If bAcceptBackup Then frmContent.ActiveForm.DeleteZone

    'ajoute du texte � la console
    If bAcceptBackup Then Call AddTextToConsole(Lang.GetString("_SelDeled"))
End Sub

Private Sub mnuDisAsm_Click()
'lance Disassembler.exe
    Call cFile.ShellOpenFile(App.Path & "\Disassembler.exe", Me.hWnd, , App.Path)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_DisASMLau"))

End Sub

Private Sub mnuDisAsmThisFile_Click()
'd�sassemble le fichier ouvert
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    On Error Resume Next
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_DisASMFileOk"))
    
    'lance l'exe de d�sassemblage avce le fichier en param�tre d'ouverture
    Shell App.Path & "\Disassembler.exe " & Chr_(34) & Me.ActiveForm.Caption & Chr_(34), vbNormalFocus
    
End Sub

Private Sub mnuExport_Click()
'exporte les valeurs hexa du fichier entier
    Call frmExport.IsEntireFile
    frmExport.Show vbModal
End Sub

Private Sub mnuExportSel_Click()
'exporte la s�lection

End Sub

Private Sub mnuGestSignets_Click()
'affiche la liste des signets
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    frmSignets.Show vbModal
End Sub

Private Sub mnuLang_Click(Index As Integer)
'on change de langue
Dim S As String
Dim X As Long
Dim cPRE As clsIniFile

    'd�termine le path du dossier
    If App.LogMode = 0 Then
        S = LANG_PATH
    Else
        S = App.Path & "\Lang"
    End If
    
    S = S & "\" & mnuLang(Index).Caption & ".ini"
    S = Replace$(S, "&", vbNullString)
    
    'v�rifie la pr�sence du fichier
    If cFile.FileExists(S) = False Then MsgBox Lang.GetString("_LangFileNot"), _
        vbCritical, Lang.GetString("_Error"): Exit Sub
    
    'on d�coche tout les menus
    For X = 1 To UBound(sLang())
        mnuLang(X).Checked = False
    Next X
    
    'on coche celui s�lectionn�
    mnuLang(Index).Checked = True
    
    'on affiche un message comme quoi il faut red�marrer
    MsgBox Lang.GetString("_HaveTo1") & vbNewLine & Lang.GetString("_HaveTo2"), _
        vbInformation, Lang.GetString("_War")
    
    'on change les pref
    cPref.env_Lang = mnuLang(Index).Caption
    Set cPRE = New clsIniFile
    Call cPRE.SaveIniFile(cPref)
    Set cPRE = Nothing
    
    'on ferme si pas dans l'IDE
    If App.LogMode <> 0 Then Call EndProgram
End Sub

Private Sub mnuLangEditor_Click()
'lance Disassembler.exe
    Call cFile.ShellOpenFile(App.Path & "\LangEditor.exe", Me.hWnd, , App.Path)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_LangEditorLau"))
End Sub

Private Sub mnuRecoverFiles_Click()
'r�cup�ration de fichiers
    frmRecoverFiles.Show
End Sub

Private Sub mnuSanitDisk_Click()
'sanitization de disque
    frmSanitization.Show vbModal
End Sub

Private Sub mnuSave_Click()
'lance la sauvegarde (�crasement)
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    If TypeOfForm(Me.ActiveForm) = "Fichier" Then
        'alors c'est un fichier
        
        If cPref.general_ShowAlert Then
            'alors il nous faut demander confirmation
            frmSave.Show vbModal
        Else
            'alors lance la sauvegarde
            
        End If
    End If
End Sub

Private Sub mnuShowConsole_Click()
'affiche la console
    pctConsole.Visible = Not (pctConsole.Visible)
    Me.mnuShowConsole.Checked = pctConsole.Visible
End Sub

Private Sub mnuUpdate_Click()
    frmUpdate.Show vbModal
End Sub

Private Sub mnuVersions_Click()
    frmComponents.Show vbModal
End Sub

Private Sub pctConsole_Resize()
    On Error Resume Next
    'resize des 2 RTF
    With txtE
        .Left = 0
        .Width = pctConsole.Width
        .Height = 220
        .Top = pctConsole.Height - 250
    End With
    With txt
        .Left = 0
        .Top = 0
        .Width = pctConsole.Width
        .Height = pctConsole.Height - 250
    End With
    cPref.console_Heigth = Me.pctConsole.Height  'sauvegarde la position actuelle
End Sub

Private Sub pctExplorer_Resize()
    Call MDIForm_Resize
    cPref.explo_Height = Me.pctExplorer.Height  'sauvegarde la position actuelle
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
'valide si entr�e
Dim S As String
    If KeyCode = vbKeyReturn Then
        S = pctPath.Text
        If cFile.FolderExists(pctPath.Text) Then LV.Path = pctPath.Text
        pctPath.Text = S
    End If
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'affiche un popup
        Me.PopupMenu Me.rmnuTools
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'alors on r�cup�re un drag&drop de fichiers
Dim i As Long
Dim i2 As Long
Dim m() As String
Dim Frm As Form

    'ajoute chaque fichier si existant, et le contenu de chaque dossier si dossier
    For i = 1 To Data.Files.Count
    
        If cFile.FileExists(Data.Files.Item(i)) Then
            'alors on ajoute le fichier
            'affiche une nouvelle fen�tre
            Set Frm = New Pfm
            Call Frm.GetFile(Data.Files.Item(i))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            
            Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & _
                CStr(lNbChildFrm) & "]"

        ElseIf cFile.FolderExists(Data.Files.Item(i)) Then
            'alors on ajoute le contenu du dossier
            
            'liste les fichiers
            m() = cFile.EnumFilesStr(Data.Files.Item(i), CBool(cPref.general_OpenSubFiles))
            If UBound(m()) < 1 Then Exit Sub
            
            'les ouvre un par un
            For i2 = 1 To UBound(m)
                If cFile.FileExists(m(i2)) Then
                    Set Frm = New Pfm
                    Call Frm.GetFile(m(i2))
                    Frm.Show
                    lNbChildFrm = lNbChildFrm + 1
                    Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & _
                        CStr(lNbChildFrm) & "]"
                    DoEvents
                End If
            Next i2

        End If
        
    Next i

End Sub

Public Sub MDIForm_Resize()
    On Error Resume Next
    
    'If Me.WindowState = vbMinimized Then frmData.Hide Else If Me.mnuEditTools.Checked Then frmData.Show
    
    With LV
        .Width = Me.Width - 120
        .Height = Me.pctExplorer.Height - 370
    End With
    
    'positionne le pctPath
    With pctPath
        .Height = 200
        .Top = LV.Height + 100
        .Left = 50
        .Width = LV.Width - 300
    End With
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Set clsPref = Nothing
    Set Lang = Nothing
    
    Call UnHookPictureResizement(Me.pctConsole.hWnd, 1)
    Call UnHookPictureResizement(Me.pctExplorer.hWnd)
    
    'vire l'icone du Tray
    Call Me.vkSysTray.RemoveFromTray(0)
    
    #If USE_FRMC_SUBCLASSING Then
        'enl�ve le hook de la form
        Call cSub.UnHookFormMenu(Me.hWnd)
        Set cSub = Nothing
    #End If
        
    Unload Me
    
    'lance la proc�dure d'arr�t
    Call EndProgram
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuAddSignet_Click()
Dim r As Long

    'on ajoute (ou enl�ve) un signet
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    With Me.ActiveForm
        If .HW.IsSignet(.HW.Item.Offset) = False Then
            'on l'ajoute
            Call .HW.AddSignet(.HW.Item.Offset)
            .lstSignets.ListItems.Add Text:=CStr(.HW.Item.Offset)
            .HW.TraceSignets
        Else
        
            'alors on l'enl�ve
            While .HW.IsSignet(.HW.Item.Offset)
                'on supprime
                .HW.RemoveSignet Val(.HW.Item.Offset)
            Wend
            
            'enl�ve du listview
            For r = .lstSignets.ListItems.Count To 1 Step -1
                If .lstSignets.ListItems.Item(r).Text = CStr(.HW.Item.Offset) _
                    Then .lstSignets.ListItems.Remove r
            Next r
            
            .HW.TraceSignets
        End If
    End With
                
End Sub

Private Sub mnuAddSignetIn_Click()
'ajoute une liste de signets
    Call AddSignetIn(False)
End Sub

Private Sub mnuBeginning_Click()
'aller tout � la fin
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Max
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Public Sub mnuCalc_Click()
    'lance la calcultarice
    On Error Resume Next
    Shell cFile.GetSpecialFolder(CSIDL_WINDOWS) & "\System32\calc.exe", vbNormalFocus
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuChangeDates_Click()
'changer les dates

    If TypeOfForm(Me.ActiveForm) = "Fichier" Then
        frmDates.txtFile.Text = Me.ActiveForm.Caption
        Call frmDates.GetFile(Me.ActiveForm.TheFile)
    End If
    
    frmDates.Show vbModal
    
End Sub

Private Sub mnuCloseAll_Click()
'ferme toutes les fen�tres
Dim lRep As Long
Dim X As Long
    
    If Me.ActiveForm Is Nothing Then Exit Sub
        
    lRep = MsgBox(Lang.GetString("_CloseAllW"), vbYesNo + vbInformation, Lang.GetString("_War"))
    If Not (lRep = vbYes) Then Exit Sub
    
    Do While Not (Me.ActiveForm Is Nothing)
        Unload Me.ActiveForm
    Loop
    
End Sub

Private Sub mnuCompareFiles_Click()
    frmCPF.Show
End Sub

Private Sub mnuConvert_Click()
    frmConvert.Show
End Sub

Private Sub mnuCopyASCII_Click()
'copier la s�lection (strings) format�e
Dim X As Long
Dim Y As Long
Dim S As String
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curSize As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    S = vbNullString    'contiendra la string � copier
    
    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"
        
    'vide le clipboard
    Call Clipboard.Clear
    
    With Me.ActiveForm
    
        'd�termine la taille
        curSize = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
            .HW.FirstSelectionItem.Offset - .HW.FirstSelectionItem.Col + 1
        
        'd�termine la position du premier offset
        curPos = .HW.FirstSelectionItem.Offset + .HW.FirstSelectionItem.Col - 1
        
        Select Case TypeOfForm(frmContent.ActiveForm)
            Case "Fichier"
                '�dition d'un fichier ==> va piocher avec ReadFile
    
                'r�cup�re la string
                S = GetBytesFromFile(.Caption, curSize, curPos)
                
            Case "Processus"
            
                'r�cup�re la string
                S = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), _
                    CLng(curPos), CLng(curSize))
    
            Case "Disque"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadS(.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
            
            Case "Disque physique"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadSPhys(Val(.Tag), _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
                
        End Select
    
    End With

    'formate la string
    S = FormatednString(S)
            
    Clipboard.SetText S
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub
Private Sub mnuCopyASCII2_Click()
'copier la s�lection (strings) format�e en bas niveau
Dim X As Long
Dim Y As Long
Dim S As String
Dim curSize As Currency
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    S = vbNullString    'contiendra la string � copier
    
    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"
        
    'vide le clipboard
    Call Clipboard.Clear

    With Me.ActiveForm
    
        'd�termine la taille
        curSize = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
            .HW.FirstSelectionItem.Offset - .HW.FirstSelectionItem.Col + 1
        
        'd�termine la position du premier offset
        curPos = .HW.FirstSelectionItem.Offset + .HW.FirstSelectionItem.Col - 1
        
        Select Case TypeOfForm(frmContent.ActiveForm)
            Case "Fichier"
                '�dition d'un fichier ==> va piocher avec ReadFile
                            
                'r�cup�re la string
                S = GetBytesFromFile(.Caption, curSize, curPos)
                
            Case "Processus"
            
                'r�cup�re la string
                S = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))
              
            Case "Disque"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadS(.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
    
            Case "Disque physique"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadSPhys(Val(.Tag), _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
                
        End Select
    End With

    'formate la string
    S = Replace$(S, vbNullChar, Chr_(32), , , vbBinaryCompare)
    
    Clipboard.SetText S
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub
Private Sub mnuCopyASCIIReal_Click()
'copie les valeurs ASCII r�elles vers le clipboard
'/!\ NULL TERMINATED STRING
Dim X As Long
Dim Y As Long
Dim S As String
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curSize As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    S = vbNullString    'contiendra la string � copier

    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"
    
    'vide le clipboard
    Call Clipboard.Clear
    
    With Me.ActiveForm
    
        'd�termine la taille
        curSize = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
            .HW.FirstSelectionItem.Offset - .HW.FirstSelectionItem.Col + 1
        
        'd�termine la position du premier offset
        curPos = .HW.FirstSelectionItem.Offset + .HW.FirstSelectionItem.Col - 1
                
        Select Case TypeOfForm(frmContent.ActiveForm)
            Case "Fichier"
                '�dition d'un fichier ==> va piocher avec ReadFile
                
                'r�cup�re la string
                S = GetBytesFromFile(.Caption, curSize, curPos)
            
            Case "Processus"
                
                'r�cup�re la string
                S = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))
            
            Case "Disque"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadS(.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
    
            Case "Disque physique"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadSPhys(Val(.Tag), _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
                
        End Select
    End With

    Clipboard.SetText S, vbCFText   'format fichier texte
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub
Private Sub mnuCopyhexa_Click()
'copier la s�lection (hexa)
Dim X As Long
Dim Y As Long
Dim S As String
Dim s2 As String
Dim curSize As Currency
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    S = vbNullString    'contiendra la string � copier
    
    'vide le clipboard
    Call Clipboard.Clear
    
    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"

    With Me.ActiveForm
    
        'd�termine la taille
        curSize = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
            .HW.FirstSelectionItem.Offset - .HW.FirstSelectionItem.Col + 1
        
        'd�termine la position du premier offset
        curPos = .HW.FirstSelectionItem.Offset + .HW.FirstSelectionItem.Col - 1
        
        Select Case TypeOfForm(frmContent.ActiveForm)
            Case "Fichier"
                '�dition d'un fichier ==> va piocher avec ReadFile
                            
                'r�cup�re la string
                S = GetBytesFromFile(.Caption, curSize, curPos)
                
            Case "Processus"
            
                'r�cup�re la string
                S = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))
    
            Case "Disque"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadS(.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
                
            Case "Disque physique"
            
                'red�finit correctement la position et la taille (doivent �tre multiple du nombre
                'de bytes par secteur)
                curPos2 = ByND(curPos, .GetDriveInfos.BytesPerSector)
                curSize2 = .HW.SecondSelectionItem.Offset + .HW.SecondSelectionItem.Col - _
                    curPos2  'recalcule la taille en partant du d�but du secteur
                curSize2 = ByN(curSize2, .GetDriveInfos.BytesPerSector)
    
                'r�cup�re la string
                Call DirectReadSPhys(Val(.Tag), _
                    curPos2 / .GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    .GetDriveInfos.BytesPerSector, S)
                    
                'recoupe la string pour r�cup�rer ce qui int�resse vraiment
                S = Mid$(S, curPos - curPos2 + 1, curSize)
                
        End Select
    End With

    'formate la string
    s2 = vbNullString
    For X = 1 To Len(S)
        If (X Mod 1000) = 0 Then DoEvents 'rend la main
        s2 = s2 & Str2Hex_(Mid$(S, X, 1))
    Next X
            
    Clipboard.SetText s2
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuCreateFileFormSel_Click()
'cr�� un fichier � partir de la s�lection
Dim curSize As Currency
Dim bOver As Boolean

    On Error GoTo CancelPushed

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'calcule la taille du fichier r�sultat
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset - _
        Me.ActiveForm.HW.FirstSelectionItem.Offset
    
    If curSize > 200000000 Then
        'fichier >200Mo, demande de confirmation
        If MsgBox(Lang.GetString("_FileWillBeLARGE"), vbInformation + vbYesNo, _
            Lang.GetString("_War")) <> vbYes Then Exit Sub
    End If
    
    'demande le fichier r�sultat
    With CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SavingSel")
        .Filter = Lang.GetString("_All") & " |*.*|" & Lang.GetString("_DatFile") & "|*.dat"
        .FileName = vbNullString
        .ShowSave
        If cFile.FileExists(.FileName) Then
            'le fichier existe d�j�
            If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + _
                vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
        End If
        'cr�� un fichier vide
        Call cFile.CreateEmptyFile(.FileName, True)
    End With
    

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_CreatingFileCour"))
    
    If TypeOfForm(Me.ActiveForm) = "Fichier" Then
        'alors on sauvegarde en utilisant ReadFile dans un fichier
        
    ElseIf TypeOfForm(Me.ActiveForm) = "Disque" Then
        'alors disque ==> readfile
        
    Else
        'alors lecture en m�moire
        
    End If
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_CreaDone"))
    
CancelPushed:
End Sub

Private Sub mnuCreateFileFromSelelection_Click()
Dim X As Long

    'cr�� un fichier depuis la s�lection
    X = MsgBox(Lang.GetString("_WannaNewFile"), vbQuestion + vbYesNo, _
        Lang.GetString("_SaveType"))
    
    Call CreateFileFromCurrentSelection(X)
End Sub

Private Sub mnuCutCopyFiles_Click()
    'd�coupage/collage
    frmCut.Show vbModal
End Sub

Private Sub mnuDiskInfos_Click()
'infos disques
    frmDiskInfos.Show
End Sub

Private Sub mnuDown_Click()
'descend d'une page
Dim l As Long

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    l = Me.ActiveForm.HW.NumberPerPage
    
    Me.ActiveForm.VS.Value = IIf((Me.ActiveForm.VS.Value + l) < _
        Me.ActiveForm.VS.Max, Me.ActiveForm.VS.Value + l, Me.ActiveForm.VS.Max)
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuEditScript_Click()
'�diteur de script
    frmScript.Show
End Sub

Private Sub mnuEditSelection_Click()
'�dite les fichiers s�lectionn�s dans le LV
Dim Frm As Form
Dim sFile() As ComctlLib.ListItem
Dim X As Long

    'On Error GoTo ErrGestion
    Call LV.GetSelectedItems(sFile)
    
    For X = 1 To UBound(sFile)
        If cFile.FileExists(sFile(X).Tag) Then
            'affiche une nouvelle fen�tre
            Set Frm = New Pfm
            Call Frm.GetFile(sFile(X).Tag)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
        End If
        DoEvents
    Next X
    
    Exit Sub
    
ErrGestion:
End Sub

Private Sub mnuEditTools_Click()
'affiche ou non le frame Data
    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.mnuEditTools.Checked = Not (Me.mnuEditTools.Checked)
    Me.ActiveForm.FrameData.Visible = Me.mnuEditTools.Checked
    Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuEnd_Click()
'tout au d�but

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Min
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuBeginingPart_Click()
'va tout au d�but de la partition
    
    If Me.ActiveForm Is Nothing Then Exit Sub

    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Min
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuEndPart_Click()
'va tout � la fin de la partition

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Max
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuExecuteScript_Click()
'execute le script actif


    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_ScriptEXE"))
End Sub

Private Sub mnuFileRenamer_Click()

    'lance FileRenamer.exe
    Call cFile.ShellOpenFile(App.Path & "\FileRenamer.exe", Me.hWnd, , App.Path)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_FileRenamerLau"))
    
End Sub

Private Sub mnuFileSearch_Click()
'lance la recherche de fichiers
    frmFileSearch.Show
End Sub

Private Sub mnuFreeForum_Click()
'forum de discussion
    Call cFile.ShellOpenFile("http://sourceforge.net/forum/forum.php?forum_id=654034", Me.hWnd, , App.Path)
End Sub

Private Sub mnuHelpForum_Click()
'forum de demande d'aide
    Call cFile.ShellOpenFile("http://sourceforge.net/forum/forum.php?forum_id=654035", Me.hWnd, , App.Path)
End Sub

Private Sub mnuHome_Click()
'lance frmHome
    frmHome.Show
    Call SetFormForeBackGround(frmHome, SetFormForeGround)
End Sub

Private Sub mnuMoveOffset_Click()
'va � une valeur particuli�re de l'offset (d�placement relatif)
Dim S As String
Dim l As Currency

    On Error Resume Next
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    S = InputBox(Lang.GetString("_HowMove"), Lang.GetString("_OffChange"))
    If StrPtr(S) = 0 Then Exit Sub  'cancel
    
    'alors on va � l'offset
    l = By16(Me.ActiveForm.HW.Item.Offset + Me.ActiveForm.HW.Item.Col + _
        Int(Val(S)))
    
    If l <= By16(Me.ActiveForm.HW.MaxOffset) And l >= 0 Then
        'alors c'est ok
        Me.ActiveForm.VS.Value = (l / 16)
        Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)    'refresh
    End If
    
End Sub

Private Sub mnuNextClust_Click()
'va au cluster suivant
Dim ActualClust As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    'd�termine le cluster actuel
    With Me.ActiveForm
        ActualClust = Int((.VS.Value / .GetDriveInfos.BytesPerCluster) * 16)
        ActualClust = ActualClust + 1
        
        .VS.Value = Int((ActualClust / 16) * .GetDriveInfos.BytesPerCluster)
        If .VS.Value > .VS.Max Then .VS.Value = .VS.Max
        
        Call .VS_Change(.VS.Value)
    End With
    
End Sub

Private Sub mnuNextSect_Click()
'va au secteur suivant
Dim ActualSect As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    'd�termine le cluster actuel
    With Me.ActiveForm
        ActualSect = Int((.VS.Value / .GetDriveInfos.BytesPerSector) * 16)
        ActualSect = ActualSect + 1
        
        .VS.Value = Int((ActualSect / 16) * .GetDriveInfos.BytesPerSector)
        If .VS.Value > .VS.Max Then .VS.Value = .VS.Max
        
        Call .VS_Change(.VS.Value)
    End With
End Sub

Private Sub mnuPrevClust_Click()
'va au cluster pr�d�dent
Dim ActualClust As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    'd�termine le cluster actuel
    With Me.ActiveForm
        ActualClust = Int((.VS.Value / .GetDriveInfos.BytesPerCluster) * 16)
        ActualClust = ActualClust - 1
        
        .VS.Value = Int((ActualClust / 16) * .GetDriveInfos.BytesPerCluster)
        If .VS.Value < .VS.Min Then .VS.Value = .VS.Min
        
        Call .VS_Change(.VS.Value)
    End With
    
End Sub

Private Sub mnuPrevSect_Click()
'va au secteur pr�d�dent
Dim ActualSect As Long

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    With Me.ActiveForm
        'd�termine le cluster actuel
        ActualSect = Int((.VS.Value / .GetDriveInfos.BytesPerSector) * 16)
        ActualSect = ActualSect - 1
        
        .VS.Value = Int((ActualSect / 16) * .GetDriveInfos.BytesPerSector)
        If .VS.Value < .VS.Min Then .VS.Value = .VS.Min
        
        Call .VS_Change(.VS.Value)
    End With
End Sub

Private Sub mnuErr_Click()
'affiche la form de rapport d'erreur
    frmLogErr.Show vbModal
End Sub

Private Sub mnuExecute_Click()
'ex�cute le fichier (temporaire)
Dim sExt As String

    If ActiveForm Is Nothing Then Exit Sub
    
    'obtient la termaison
    sExt = cFile.GetFileExtension(Me.ActiveForm.Caption)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_CreaTempCour"))
    
    Call ExecuteTempFile(Me.hWnd, Me.ActiveForm, sExt)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_CreaTempOk"))
End Sub

Public Sub mnuExit_Click()
    'quitte
    Call MDIForm_Unload(0)
End Sub

Private Sub mnuExploreDisk_Click()
'affiche ou pas l'explorer de disque

    If TypeOfForm(Me.ActiveForm) = "Disque" Then
        With Me.ActiveForm
            mnuExploreDisk.Checked = Not (mnuExploreDisk.Checked)
            .FV.Visible = mnuExploreDisk.Checked
            .FV2.Visible = mnuExploreDisk.Checked
            .pctPath.Visible = mnuExploreDisk.Checked
            .FrameFrag.Visible = mnuExploreDisk.Checked
        End With
    End If
    Call frmContent.ActiveForm.ResizeMe
       
End Sub

Private Sub mnuExploreDisplay_Click()

    mnuExploreDisplay.Checked = Not (mnuExploreDisplay.Checked)
    pctExplorer.Visible = mnuExploreDisplay.Checked
    
    If pctExplorer.Visible = True Then
        'alors il faut rafraichir le Expore
        
        'loading de la taille de l'explorer
        Me.pctExplorer.Height = cPref.explo_Height
    
        With LV
            .BlockDisplay = True    'bloque l'affichage pour �viter le refresh � chaque changement de property
            .Visible = False
            
            .Height = cPref.explo_Height - 370
            
            If cPref.explo_DefaultPath = Lang.GetString("_ProgramDir!") Then
                'alors c'est dans app.path
                .Path = App.Path
            Else
                'alors un dossier perso
                .Path = cPref.explo_DefaultPath
            End If
            .ShowEntirePath = CBool(cPref.explo_ShowPath)
            .ShowHiddenDirectories = CBool(cPref.explo_ShowHiddenFolders)
            .ShowHiddenFiles = CBool(cPref.explo_ShowHiddenFiles)
            .ShowSystemDirectories = CBool(cPref.explo_ShowSystemFodlers)
            .ShowSystemFiles = CBool(cPref.explo_ShowSystemFiles)
            .ShowReadOnlyDirectories = CBool(cPref.explo_ShowROFolders)
            .ShowReadOnlyFiles = CBool(cPref.explo_ShowROFiles)
            .AllowMultiSelect = CBool(cPref.explo_AllowMultipleSelection)
            .AllowFileDeleting = CBool(cPref.explo_AllowFileSuppression)
            .Pattern = cPref.explo_Pattern
            .HideColumnHeaders = CBool(cPref.explo_HideColumnTitle)
            Select Case cPref.explo_IconType
                Case 0
                    .DisplayIcons = BasicIcons
                Case 1
                    .DisplayIcons = FileIcons
                Case 2
                    .DisplayIcons = NoIcons
            End Select
            
            .BlockDisplay = False
            Call .Refresh
            .Visible = True
            .RefreshListViewOnly    '/!\ DO NOT REMOVE
        End With
    End If
    
    If (frmContent.ActiveForm Is Nothing) = False Then Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuFileViewMode_Click()
'affiche le mode "lecture de fichier", cad que les valeurs hexa

End Sub

Private Sub mnuFillSelection_Click()
'remplit la s�lection du HW

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    frmFillSelection.Show

End Sub

Private Sub mnuGestWindows_Click()
'affiche la form de gestion des fen�tres
    frmGestWindows.Show vbModal
End Sub

Private Sub mnuGoToOffset_Click()
'va � une valeur particuli�re de l'offset
Dim S As String
Dim l As Currency

    On Error Resume Next
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    S = InputBox(Lang.GetString("_MoveToWich"), Lang.GetString("_OffChange"))
    If StrPtr(S) = 0 Then Exit Sub  'cancel
    
    'alors on va � l'offset (si possible)
    l = By16(Int(Abs(Val(S))) - 15)  'formatage de l'offset
    
    If l <= By16(Me.ActiveForm.HW.MaxOffset) Then
        'alors c'est ok
        Me.ActiveForm.VS.Value = (l / 16)
        Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)    'refresh
    End If
    
End Sub

Private Sub mnuHelp_Click()
'affiche l'aide
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_HelpDis"))
    
    'affiche
    Call cFile.ShellOpenFile(App.Path & "\Help.chm", Me.hWnd)
    
End Sub

Private Sub mnuInformations_Click()
'affiche ou non les infos
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.mnuInformations.Checked = Not (Me.mnuInformations.Checked)
    Me.ActiveForm.FrameInfos.Visible = Me.mnuInformations.Checked
    If (TypeOfForm(Me.ActiveForm) = "Disque") Or (TypeOfForm(Me.ActiveForm) = "Disque physique") Then Me.ActiveForm.FrameInfo2.Visible = Me.mnuInformations.Checked
    Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuInterpretAdvanced_Click()
'conversion avanc�e
    frmAdvancedConversion.Show
End Sub

Private Sub mnuMH_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuMV_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuNew_Click()
'nouveau fichier
    frmNew.Show vbModal
End Sub

Private Sub mnuNewProcess_Click()
'invite � demarrer un nouveau processus
    Call cFile.ShowRunBox(Me.hWnd, Lang.GetString("_StartProcessTitle"), _
        Lang.GetString("_StartProcessMsg"))
End Sub

Private Sub mnuOpen_Click()
'ajoute un fichier � la liste � supprimer
Dim S() As String
Dim s2 As String
Dim X As Long
Dim Frm As Form
    
    ReDim S(0)
    
    s2 = cFile.ShowOpen(Lang.GetString("_SelFileToOpen"), Me.hWnd, _
        Lang.GetString("_All") & "|*.*", , , , , OFN_EXPLORER + _
        OFN_ALLOWMULTISELECT, S())
    
    For X = 1 To UBound(S())
        If cFile.FileExists(S(X)) Then
            Set Frm = New Pfm
            Call Frm.GetFile(S(X))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
        End If
        DoEvents    '/!\ IMPORTANT DO NOT REMOVE
    Next X
    
    'dans le cas d'un fichier simple
    If cFile.FileExists(s2) Then
        Set Frm = New Pfm
        Call Frm.GetFile(s2)
        Frm.Show
        lNbChildFrm = lNbChildFrm + 1
    End If

    Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
    
End Sub

Private Sub mnuOpenDisk_Click()
'ouvre un disque physique
    frmDrive.Show vbModal
End Sub

Private Sub mnuOpenExplorer_Click()
'ouvre explorer.exe � l'emplacement point� par LV

    Shell "explorer.exe " & LV.Path, vbNormalFocus
End Sub

Private Sub mnuOpenFolder_Click()
'ouvre un dossier
Dim m() As String
Dim sDir As String
Dim Frm As Form
Dim X As Long

    's�lectionne un r�pertoire
    sDir = cFile.BrowseForFolder(Lang.GetString("_SelADir"), Me.hWnd)
    
    'teste la validit� du r�pertoire
    If cFile.FolderExists(sDir) = False Then Exit Sub
    
    'liste les fichiers
    m() = cFile.EnumFilesStr(sDir, CBool(cPref.general_OpenSubFiles))
    If UBound(m()) < 1 Then Exit Sub

    'les ouvre un par un
    For X = 1 To UBound(m)
        If cFile.FileExists(m(X)) Then
            Set Frm = New Pfm
            Call Frm.GetFile(m(X))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
            DoEvents
        End If
    Next X
  
    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entr�es dans les menus

End Sub

Private Sub mnuOpenInBN_Click()
'ouvre le fichier dans le bloc notes
Dim X As Long

    If ActiveForm Is Nothing Then Exit Sub
    
    If cFile.FileExists(Me.ActiveForm.Caption) = False Then
        'pas de fichier
        MsgBox Lang.GetString("_FileAbsent"), vbInformation, Lang.GetString("_CanNotOpen")
    End If
    
    If cFile.GetFileSize(Me.ActiveForm.Caption) > 1000000 Then
        'fichier de plus de 700Ko
        X = MsgBox(Lang.GetString("_FileMakeMoreThan") & vbNewLine & Lang.GetString("_ShoudNotBN"), vbInformation + vbYesNo, Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_BNLau"))
    
    Shell "notepad " & Me.ActiveForm.Caption, vbNormalFocus
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_BNOk"))
End Sub

Private Sub mnuOpenProcess_Click()
'ouvre un processus en m�moire
  
    'affiche la liste des process
    frmProcesses.Show vbModal

End Sub

Private Sub mnuOpenSelectedFiles_Click()
'ouvre les fichiers s�lectionn�s dans le LV
Dim sFile() As ListItem
Dim X As Long

    'obtient la liste des s�lections
    Call LV.GetSelectedItems(sFile)
    
    For X = 1 To UBound(sFile)
        Call cFile.ShellOpenFile(sFile(X).Tag, Me.hWnd)
    Next X
    
End Sub

Private Sub mnuOpenSignetsList_Click()
'ouvre une liste de signet
    Call AddSignetIn(True)
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPrint_Click()
'impression
    frmPrint.Show vbModal
End Sub

Private Sub mnuProcesses_Click()
'gestionnaire tr�s simple de processus
    frmProcess.Show
End Sub

Private Sub mnuProperty_Click()
'affiche les propri�t�s du fichier
    frmPropertyShow.Show
End Sub

Public Sub mnuRedo_Click()
    If Me.ActiveForm Is Nothing Then Exit Sub
    Call Me.ActiveForm.RedoM
End Sub

Public Sub mnuRefreh_Click()
'refresh

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'on refresh le HW
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuRemoveAll_Click()
'supprime tous les signets ==> demande confirmation
   
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'confirmation
    If MsgBox(Lang.GetString("_SureDelAllSig"), vbInformation + vbYesNo, _
        Lang.GetString("_War")) <> vbYes Then Exit Sub
    
    With Me.ActiveForm
        .HW.RemoveAllSignets
        .lstSignets.ListItems.Clear
    End With
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SigDeled"))
End Sub

Private Sub mnuRemoveSignet_Click()
'supprime un signet, si existant
Dim X As Long

    
    If Me.ActiveForm Is Nothing Then Exit Sub

    With Me.ActiveForm
        If .HW.IsSignet(.HW.Item.Offset) Then
        
            While .HW.IsSignet(.HW.Item.Offset)
                'on supprime
                .HW.RemoveSignet Val(.HW.Item.Offset)
            Wend
            
            'enl�ve du listview
            For X = .lstSignets.ListItems.Count To 1 Step -1
                If .lstSignets.ListItems.Item(X).Text = CStr(.HW.Item.Offset) Then
                    .lstSignets.ListItems.Remove X
                End If
            Next X
        End If
    End With
    
End Sub

Private Sub mnuReorganize_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuSaveIconAsBitmap_Click()
'enregistre l'icone de l'active form en bitmap
Dim S As String

    If Me.ActiveForm Is Nothing Then Exit Sub
    If TypeOfForm(Me.ActiveForm) <> "Fichier" And TypeOfForm(Me.ActiveForm) <> "Processus" Then Exit Sub
    
    With Me.ActiveForm
        'sauvegarder l'icone s�lectionn�e en bitmap
        If .lvIcon.SelectedItem Is Nothing Then Exit Sub
        
        'pose l'image sur le picturebox
        Call ImageList_Draw(.IMG.hImageList, .lvIcon.SelectedItem.Index - 1, _
            .pct.hdc, 2, 2, ILD_TRANSPARENT) 'tente de recentrer l'image avec 2,2
        
        If .pct.Picture Is Nothing Then Exit Sub
    End With
   
    'demande la sauvegarde du fichier
    On Error GoTo Err
    
    'affiche la boite de dialogue "sauvegarder"
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SaveBMP")
        .Filter = "Bitmap Image|*.bmp|"
        .FileName = vbNullString
        .ShowSave
        S = .FileName
    End With
    
    'rajoute l'extension si n�cessaire
    If LCase$(Right$(S, 4)) <> ".bmp" Then S = S & ".bmp"
    
    'lance la sauvegarde
    Call SavePicture(Me.ActiveForm.pct.Image, S)
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_IconSaved"))
    
Err:
    Set Me.ActiveForm.pct.Picture = Nothing
End Sub

Private Sub mnuSaveSignets_Click()
'enregistre la liste des signets de la form active
Dim S As String
Dim lFile As Long
Dim X As Long

    On Error GoTo ErrGestion
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    If Me.ActiveForm.lstSignets.ListItems.Count = 0 Then Exit Sub 'pas de signets
    
    'enregistrement ==> choix du fichier
    With CMD
        .CancelError = True
        .FileName = Me.ActiveForm.Caption & ".sig"
        .DialogTitle = Lang.GetString("_SaveSigList")
        .Filter = Lang.GetString("_SigList") & " |*.sig|"
        .InitDir = App.Path
        .FileName = vbNullString
        .ShowSave
        S = .FileName
    End With

    If cFile.FileExists(S) Then
        'message de confirmation
        X = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If
    
    'ouvre le fchier
    lFile = FreeFile
    Open S For Output As lFile
    
    'enregistre les entr�es
    For X = 1 To Me.ActiveForm.lstSignets.ListItems.Count
        Write #lFile, Me.ActiveForm.lstSignets.ListItems.Item(X) & "|" & Me.ActiveForm.lstSignets.ListItems.Item(X).SubItems(1)
    Next X
    
    Close lFile

    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SigSavedOk"))
    
ErrGestion:
    
End Sub

Private Sub mnuSearchForString_Click()
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    frmStringSearch.Show
End Sub

Private Sub mnuSearchH_Click()
'recherche valeur hexa
    If Me.ActiveForm Is Nothing Then Exit Sub
    frmSearch.Show
    frmSearch.Option4(0).Value = True   'valeur hexa
End Sub

Private Sub mnuSearchT_Click()
'recherche du texte
    If Me.ActiveForm Is Nothing Then Exit Sub
    frmSearch.Show
    frmSearch.Option4(1).Value = True   'valeur hexa
End Sub

Private Sub mnuSelectAll_Click()
'tout s�lectionner
    
    If Me.ActiveForm Is Nothing Then Exit Sub

    With Me.ActiveForm
        .HW.SelectZone 0, 0, 16 - (By16(.HW.MaxOffset) - .HW.MaxOffset) - 1, _
            By16(.HW.MaxOffset) - 16
        .HW.Refresh
        
        'refresh le label qui contient la taille de la s�lection
        .Sb.Panels(4).Text = "S�lection=[" & CStr(.HW.NumberOfSelectedItems) & _
            " bytes]"
        .Label2(9) = .Sb.Panels(4).Text
    End With
End Sub

Private Sub mnuSelectFromByte_Click()
's�lection � partir d'un byte
    frmSelect2.Show vbModal
End Sub

Private Sub mnuSelectZone_Click()
's�lectionne une zone d�finie
    frmSelect.GetEditFunction 0 'selection mode
    frmSelect.Show vbModal
End Sub

Private Sub mnuShowIcons_Click()
'affiche ou pas les icones du fichier
    If TypeOfForm(frmContent.ActiveForm) = "Processus" Or TypeOfForm(frmContent.ActiveForm) = "Fichier" Then
        'alors c'est bon, il existe le FrameIcon
        Me.mnuShowIcons.Checked = Not (Me.mnuShowIcons.Checked)
        frmContent.ActiveForm.FrameIcon.Visible = Me.mnuShowIcons.Checked
    End If
    Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuShowNotAllFile_Click()
'visualisation restreinte
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    frmSelect.GetEditFunction 1 'recoupage mode
    frmSelect.Show vbModal
End Sub

Private Sub mnuShredder_Click()
'supprime d�finitivement des fichiers
    frmShredd.Show vbModal
End Sub

Private Sub mnuSignetNext_Click()
'signet suivant

    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.HW.FirstOffset = Me.ActiveForm.HW.GetNextSignet(Me.ActiveForm.HW.Item.Offset)
    Me.ActiveForm.HW.Refresh
    Me.ActiveForm.VS.Value = Me.ActiveForm.HW.FirstOffset / 16
End Sub

Private Sub mnuSignetPrev_Click()
'signet pr�c�dent

    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.HW.FirstOffset = Me.ActiveForm.HW.GetPrevSignet(Me.ActiveForm.HW.Item.Offset)
    Me.ActiveForm.HW.Refresh
    Me.ActiveForm.VS.Value = Me.ActiveForm.HW.FirstOffset / 16
End Sub

Public Sub mnuSourceForge_Click()
'page source forge
    Call cFile.ShellOpenFile("http://sourceforge.net/projects/hexeditorvb/", Me.hWnd, , App.Path)
End Sub

Public Sub mnuStats_Click()
'affiche les statistiques du fichier
Dim Frm As Form

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'affiche la form d'analyse
    Set Frm = New frmAnalys
    Frm.GetFile Me.ActiveForm.Caption
    Frm.Show
    
End Sub

Private Sub mnuStatsPopup_Click()
'affiche les stats des fichiers s�lectionn�s dans LV
Dim Frm As Form
Dim sFile() As ListItem
Dim X As Long

    'On Error GoTo ErrGestion

    LV.GetSelectedItems sFile
    
    For X = 1 To UBound(sFile)
        If cFile.FileExists(sFile(X).Tag) Then
            'affiche une nouvelle fen�tre
            Set Frm = New frmAnalys
            Call Frm.GetFile(sFile(X).Tag)
            Call Frm.cmdAnalyse_Click   'lance l'analyse
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
        End If
        DoEvents
    Next X
    
    Exit Sub
    
ErrGestion:
End Sub

Private Sub mnuStatusOK_Click()
'r�initialise le status
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuTab_Click()
'affiche ou non le HW
    mnuTab.Checked = Not (mnuTab.Checked)
    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.HW.Visible = mnuTab.Checked
    Me.ActiveForm.VS.Visible = mnuTab.Checked
    If TypeOfForm(Me.ActiveForm) = "Processus" Then Me.ActiveForm.MemTB.Visible = mnuTab.Checked
    Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuTableMulti_Click()
'affiche la table
    frmTable.Show
    Call frmTable.CreateTable(AllTables)
End Sub

Private Sub mnuTdec2Ascii_Click()
'affiche la table
    frmTable.Show
    Call frmTable.CreateTable(HEX_ASCII)
End Sub

Private Sub mnuSaveAs_Click()
'enregistrer sous
Dim sFile As String
Dim sPath As String
Dim lFile As Long
Dim X As Long

    On Error GoTo GestionErr

    If Me.ActiveForm Is Nothing Then Exit Sub

    'il faut sauvegarder en prenant compte des 2 changelist de pfm
    
    With CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_SaveAs")
        .Filter = Lang.GetString("_All") & "|*.*"
        .FileName = vbNullString
        .ShowSave
        sPath = .FileName
    End With
    
    If cFile.FileExists(sPath) Then
        'message de confirmation
        X = MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War"))
        If Not (X = vbYes) Then Exit Sub
    End If
    
    'efface le pr�c�dent fichier
    Call cFile.DeleteFile(sPath)
    
    'cr�� le fichier
    Call Me.ActiveForm.GetNewFile(sPath)

GestionErr:
End Sub

Private Sub mnuThisIsTheBeginnig_Click()
'marque le d�but de la s�lection � cet offset
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    frmContent.ActiveForm.HW.FirstSelectionItem.Offset = frmContent.ActiveForm.HW.Item.Offset

End Sub

Private Sub mnuThisIsTheEnd_Click()
'marque la fin de la s�lection � cet offset
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    frmContent.ActiveForm.HW.SecondSelectionItem.Offset = frmContent.ActiveForm.HW.Item.Offset
End Sub

Public Sub mnuUndo_Click()
    If Me.ActiveForm Is Nothing Then Exit Sub
    Call Me.ActiveForm.UndoM
End Sub

Public Sub mnuVbfrance_Click()
'vbfrance.com
    Call cFile.ShellOpenFile("http://www.vbfrance.com/auteurdetail.aspx?ID=523601&print=1", Me.hWnd, , App.Path)
End Sub

Private Sub muUp_Click()
'monte d'une page
Dim l As Long

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    l = Me.ActiveForm.HW.NumberPerPage
    
    Me.ActiveForm.VS.Value = IIf((Me.ActiveForm.VS.Value - l) > Me.ActiveForm.VS.Min, Me.ActiveForm.VS.Value - l, Me.ActiveForm.VS.Min)
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub rmnuHelp_Click()
'affiche le nombre d'erreurs enregistr�es dans le menu "Rapport..."
    Me.mnuErr.Caption = Lang.GetString("_ErrorReport") & " (" & Trim$(Str$(clsERREUR.NumberOfErrorInLogFile)) & ")"
End Sub

Private Sub Timer1_Timer()
    Call frmContent.ChangeEnabledMenus  'active ou pas certaines entr�es dans les menus
    Call RefreshToolbarEnableState  'active ou pas certain boutons dans la toolbar
    
    'r�affiche le choix des dossiers (si choix = true)
    frmContent.pctExplorer.Visible = frmContent.mnuExploreDisplay.Checked And Not (TypeOfActiveForm = "Disk")
    
    'actualise les fonctions Undo/Redo et les fonctions Signet pr�c�dent/suivant
    If Not (Me.ActiveForm Is Nothing) Then
        Call ModifyHistoEnabled
        Call RefreshBookMarkEnabled
    Else
        'pas de fichier ouvert ==> enabled=false
        Me.mnuRedo.Enabled = False
        Me.mnuUndo.Enabled = False
        With Me.Toolbar1.Buttons
            .Item(12).Enabled = False
            .Item(13).Enabled = False
            .Item(16).Enabled = False
            .Item(17).Enabled = False
        End With
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'appui sur les icones

    Select Case Button.Key
    
        Case "OpenFile"
            Call mnuOpen_Click
        Case "HomeOpen"
            'affiche la boite de dialogue Home (choix des diff�rentes actions � faire)
            frmHome.Show
            Call SetFormForeBackGround(frmHome, SetFormForeGround)
        Case "New"
            mnuNew_Click
        Case "Signet"
            If Me.ActiveForm Is Nothing Then Exit Sub
            Call mnuAddSignet_Click
        Case "Up"
            If Me.ActiveForm Is Nothing Then Exit Sub
            Me.ActiveForm.HW.FirstOffset = Me.ActiveForm.HW.GetPrevSignet(Me.ActiveForm.HW.Item.Offset)
            Call Me.ActiveForm.HW.Refresh
            Me.ActiveForm.VS.Value = Me.ActiveForm.HW.FirstOffset / 16
        Case "Down"
            If Me.ActiveForm Is Nothing Then Exit Sub
            Me.ActiveForm.HW.FirstOffset = Me.ActiveForm.HW.GetNextSignet(Me.ActiveForm.HW.Item.Offset)
            Call Me.ActiveForm.HW.Refresh
            Me.ActiveForm.VS.Value = Me.ActiveForm.HW.FirstOffset / 16
        Case "Copy"
            Call mnuCopyASCIIReal_Click
        Case "Convert"
            frmConvert.Show
        Case "Settings"
            frmOptions.Show vbModal
        Case "Print"
            Call mnuPrint_Click
        Case "Search"
            Call mnuSearchT_Click
        Case "Undo"
            Call Me.ActiveForm.UndoM
        Case "Redo"
            Call Me.ActiveForm.RedoM
        Case "Save"
            Call mnuSaveAs_Click
        Case "Print"
            Call mnuPrint_Click
        Case "Cut"
            Call mnuCut_Click
        Case "Paste"
            
    End Select

End Sub

'=======================================================
'ajoute (ou ouvre si overwrite) une liste de signets
'=======================================================
Private Sub AddSignetIn(ByVal bOverWrite As Boolean)
Dim S As String
Dim lFile As Long
Dim X As Long
Dim sTemp As String
Dim l As Long

    On Error GoTo ErrGestion
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'ouverture ==> choix du fichier
    With CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_OpenSigList")
        .Filter = Lang.GetString("_SigList") & " |*.sig|"
        .InitDir = App.Path
        .ShowOpen
        S = .FileName
    End With
    
    If bOverWrite Then
        Me.ActiveForm.lstSignets.ListItems.Clear
        Me.ActiveForm.HW.RemoveAllSignets
    End If
    
    'ouvre le fchier
    lFile = FreeFile
    Open S For Input As lFile
    While Not EOF(lFile)
        Input #lFile, sTemp
        l = InStr(1, sTemp, "|", vbBinaryCompare)
        If l <> 0 Then
            'ajoute aussi un commentaire
            Me.ActiveForm.lstSignets.ListItems.Add Text:=Left$(sTemp, l - 1)
            Me.ActiveForm.HW.AddSignet Val(Left$(sTemp, l - 1))
            Me.ActiveForm.lstSignets.ListItems.Item(Me.ActiveForm.lstSignets.ListItems.Count).SubItems(1) = Right$(sTemp, Len(sTemp) - l)
        End If
    Wend
    
    Me.ActiveForm.HW.Refresh
    
    Close lFile
    
    'ajoute du texte � la console
    Call AddTextToConsole(Lang.GetString("_SigAdded"))
ErrGestion:
End Sub

'=======================================================
'permet de masquer ou d'afficher les menus en fonction du type de form qui est active
'=======================================================
Public Function ChangeEnabledMenus()

    If TypeOfActiveForm = "Mem" Then
        'ActiveForm=MemPfm
        'alors on masque certaines options des menus
        Me.mnuDisAsmThisFile.Enabled = False
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = False
        Me.mnuExecute.Enabled = False
        Me.mnuCut.Enabled = False
        Me.mnuOpenInBN = False
        Me.mnuStats.Enabled = False
        Me.mnuDeleteSelection.Enabled = False
        Me.mnuShowIcons.Enabled = True
        Me.mnuSaveAs.Enabled = True
        Me.mnuPrint.Enabled = True
        Me.mnuProperty.Enabled = True
        Me.mnuCloseAll.Enabled = True
        Me.rmnuEdit.Enabled = True
        Me.mnuTab.Enabled = True
        Me.mnuInformations.Enabled = True
        Me.mnuEditTools.Enabled = True
        Me.mnuTdec2Ascii.Enabled = True
        Me.mnuTableMulti.Enabled = True
        Me.mnuRefreh.Enabled = True
        Me.rmnuPos.Enabled = True
        Me.rmnuSignets.Enabled = True
        Me.rmnuFind.Enabled = True
        Me.mnuGoOn.Enabled = True
        Me.rmnuWindow.Enabled = True
        Me.mnuExportSel.Enabled = True
        Me.mnuExport.Enabled = False
        Me.rmnuExport.Enabled = True
    ElseIf (Me.ActiveForm Is Nothing) = False And TypeOfActiveForm = "Pfm" Then
        'ActiveForm=Pfm
        'alors on affiche les options qui auraient pu �tre cach�es
        Me.mnuDisAsmThisFile.Enabled = True
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = True
        Me.mnuExecute.Enabled = True
        Me.mnuCut.Enabled = True
        Me.mnuOpenInBN = True
        Me.mnuStats.Enabled = True
        Me.mnuDeleteSelection.Enabled = True
        Me.mnuShowIcons.Enabled = True
        Me.mnuSaveAs.Enabled = True
        Me.mnuPrint.Enabled = True
        Me.mnuProperty.Enabled = True
        Me.mnuCloseAll.Enabled = True
        Me.rmnuEdit.Enabled = True
        Me.mnuTab.Enabled = True
        Me.mnuInformations.Enabled = True
        Me.mnuEditTools.Enabled = True
        Me.mnuTdec2Ascii.Enabled = True
        Me.mnuTableMulti.Enabled = True
        Me.mnuRefreh.Enabled = True
        Me.rmnuPos.Enabled = True
        Me.rmnuSignets.Enabled = True
        Me.rmnuFind.Enabled = True
        Me.mnuGoOn.Enabled = True
        Me.rmnuWindow.Enabled = True
        Me.mnuExport.Enabled = True
        Me.mnuExportSel.Enabled = True
        Me.rmnuExport.Enabled = True
    ElseIf Me.ActiveForm Is Nothing Then
        'ActiveForm = nothing
        Me.mnuDisAsmThisFile.Enabled = False
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = False
        Me.mnuExecute.Enabled = False
        Me.mnuCut.Enabled = False
        Me.mnuOpenInBN = False
        Me.mnuStats.Enabled = False
        Me.mnuDeleteSelection.Enabled = False
        Me.mnuShowIcons.Enabled = False
        Me.mnuSaveAs.Enabled = False
        Me.mnuPrint.Enabled = False
        Me.mnuProperty.Enabled = False
        Me.mnuCloseAll.Enabled = False
        Me.rmnuEdit.Enabled = False
        Me.mnuTab.Enabled = False
        Me.mnuInformations.Enabled = False
        Me.mnuEditTools.Enabled = False
        Me.mnuTdec2Ascii.Enabled = False
        Me.mnuTableMulti.Enabled = False
        Me.mnuRefreh.Enabled = False
        Me.rmnuPos.Enabled = False
        Me.rmnuSignets.Enabled = False
        Me.rmnuFind.Enabled = False
        Me.mnuGoOn.Enabled = False
        Me.rmnuWindow.Enabled = False
        Me.rmnuExport.Enabled = False
    ElseIf (Me.ActiveForm Is Nothing) = False And TypeOfActiveForm = "Disk" Then
        'diskfrm
        Me.mnuDisAsmThisFile.Enabled = False
        Me.mnuExploreDisk.Enabled = True
        Me.mnuSave.Enabled = False
        Me.mnuExecute.Enabled = True
        Me.mnuCut.Enabled = True
        Me.mnuOpenInBN = True
        Me.mnuStats.Enabled = True
        Me.mnuDeleteSelection.Enabled = True
        Me.mnuShowIcons.Enabled = False
        Me.mnuSaveAs.Enabled = True
        Me.mnuPrint.Enabled = True
        Me.mnuProperty.Enabled = True
        Me.mnuCloseAll.Enabled = True
        Me.rmnuEdit.Enabled = True
        Me.mnuTab.Enabled = True
        Me.mnuInformations.Enabled = True
        Me.mnuEditTools.Enabled = True
        Me.mnuTdec2Ascii.Enabled = True
        Me.mnuTableMulti.Enabled = True
        Me.mnuRefreh.Enabled = True
        Me.rmnuPos.Enabled = True
        Me.rmnuSignets.Enabled = True
        Me.rmnuFind.Enabled = True
        Me.mnuGoOn.Enabled = True
        Me.rmnuWindow.Enabled = True
        Me.rmnuExport.Enabled = True
        Me.mnuExportSel.Enabled = True
        Me.mnuExport.Enabled = False
    Else
        'phyPfm
        Me.mnuDisAsmThisFile.Enabled = False
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = False
        Me.mnuExecute.Enabled = True
        Me.mnuCut.Enabled = False
        Me.mnuOpenInBN = False
        Me.mnuStats.Enabled = True
        Me.mnuDeleteSelection.Enabled = False
        Me.mnuShowIcons.Enabled = False
        Me.mnuSaveAs.Enabled = True
        Me.mnuPrint.Enabled = True
        Me.mnuProperty.Enabled = True
        Me.mnuCloseAll.Enabled = True
        Me.rmnuEdit.Enabled = True
        Me.mnuTab.Enabled = True
        Me.mnuInformations.Enabled = True
        Me.mnuEditTools.Enabled = True
        Me.mnuTdec2Ascii.Enabled = True
        Me.mnuTableMulti.Enabled = True
        Me.mnuRefreh.Enabled = True
        Me.rmnuPos.Enabled = True
        Me.rmnuSignets.Enabled = True
        Me.rmnuFind.Enabled = True
        Me.mnuGoOn.Enabled = True
        Me.rmnuWindow.Enabled = True
        Me.rmnuExport.Enabled = True
        Me.mnuExportSel.Enabled = True
        Me.mnuExport.Enabled = False
    End If
    
End Function

'=======================================================
'permet d'activer ou non les boutons de la ToolBar
'=======================================================
Private Sub RefreshToolbarEnableState()

    With Me.Toolbar1.Buttons
        If Me.ActiveForm Is Nothing Then
            'alors pas de Copier/coller/rechercher/couper/signets
            .Item(4).Enabled = False
            .Item(5).Enabled = False
            .Item(7).Enabled = False
            .Item(8).Enabled = False
            .Item(9).Enabled = False
            .Item(10).Enabled = False
            .Item(15).Enabled = False
        Else
            'on active
            If TypeOfForm(Me.ActiveForm) = "Fichier" Then .Item(4).Enabled = True
            .Item(5).Enabled = True
            .Item(7).Enabled = True
            .Item(8).Enabled = True
            .Item(9).Enabled = True
            .Item(10).Enabled = True
            .Item(15).Enabled = True
        End If
    End With
        
End Sub

Private Sub txtE_Change()
'on applique la couleur RGB(192,192,192)
    With txtE
        .SelStart = 0
        .SelLength = Len(txtE.Text)
        .SelColor = cPref.console_ForeColor
        .SelStart = Len(txtE.Text)
    End With
End Sub

Private Sub txtE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        'alors on r�cup�re la pr�c�dente commande
        If lngConsolePos > 1 Then lngConsolePos = lngConsolePos - 1
        txtE.Text = GetCommand
    ElseIf KeyCode = vbKeyDown Then
        'on r�cup�re la commande suivante
        If lngConsolePos < UBound(strConsoleText()) Then lngConsolePos = lngConsolePos + 1
        txtE.Text = GetCommand
    ElseIf KeyCode = vbKeyReturn Then
        'alors on a valid� une commande
        If txtE.Text <> vbNullString Then
            lngConsolePos = lngConsolePos + 1
            Call LaunchCommand
            txtE.Text = vbNullString
        End If
    End If
End Sub

Private Sub txtE_KeyPress(KeyAscii As Integer)
'�vite le 'BIP' lors de l'appui sur la touche entr�e
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub vkSysTray_MouseDblClick(Button As MouseButtonConstants, id As Long)
    Call frmHome.Show
End Sub

Private Sub vkSysTray_MouseUp(Button As MouseButtonConstants, id As Long)
    If Button = vbRightButton Then _
        Call Me.PopupMenu(Me.rmnuFichier)   'popup menu du Tray
End Sub
