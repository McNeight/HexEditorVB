VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{9B9A881F-DBDC-4334-BC23-5679E5AB0DC6}#1.1#0"; "FileView_OCX.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Begin VB.MDIForm frmContent 
   BackColor       =   &H8000000C&
   Caption         =   "Editeur hexadécimal"
   ClientHeight    =   7785
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   9870
   Icon            =   "frmContent.frx":0000
   LinkTopic       =   "Editeur hexadécimal"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   1440
      Top             =   3720
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2760
   End
   Begin VB.PictureBox pctExplorer 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   2200
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   9870
      TabIndex        =   1
      Top             =   330
      Width           =   9870
      Begin FileView_OCX.FileView LV 
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   120
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
      TabIndex        =   0
      Top             =   7530
      Width           =   9870
      _ExtentX        =   17410
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
            TextSave        =   "23:01"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "14/02/2007"
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
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
            Object.ToolTipText     =   "Créer un nouveau fichier"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OpenFile"
            Object.ToolTipText     =   "Ouvrir un fichier"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HomeOpen"
            Object.ToolTipText     =   "Ouvrir un dossier"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Sauvegarder l'objet ouvert"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Effectuer une recherche dans l'objet"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Couper la sélection"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copier la sélection"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Coller le contenu du presse-papier"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Défaire"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Refaire"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Signet"
            Object.ToolTipText     =   "Ajouter un signet à l'offset sélectionné"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up"
            Object.ToolTipText     =   "Aller au signet précédent"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Down"
            Object.ToolTipText     =   "Aller au signet suivant"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert"
            Object.ToolTipText     =   "Afficher la fenêtre de conversion"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3360
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
            Picture         =   "frmContent.frx":628A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":7C1C
            Key             =   ""
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":95AE
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":AF40
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":C8D2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":E264
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":FBF6
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":11588
            Key             =   "Signet"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":12F1A
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":148AC
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1623E
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":17BD0
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":19562
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1AEF4
            Key             =   "Trash"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1C886
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1E218
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":1FBAA
            Key             =   "FileOpen"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":2153C
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":22ECE
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":24860
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":26CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContent.frx":271C3
            Key             =   ""
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
            Caption         =   "&Nouveau processus..."
         End
      End
      Begin VB.Menu mnuROpen 
         Caption         =   "&Ouvrir"
         Begin VB.Menu mnuOpen 
            Caption         =   "&Ouvrir un fichier..."
            Shortcut        =   ^O
         End
         Begin VB.Menu mnuOpenFolder 
            Caption         =   "&Ouvrir un dossier de fichiers..."
         End
         Begin VB.Menu mnuOpenProcess 
            Caption         =   "&Ouvrir un processus en mémoire..."
         End
         Begin VB.Menu mnuOpenDisk 
            Caption         =   "&Ouvrir un disque physique..."
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Enregistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Enregistrer sous..."
      End
      Begin VB.Menu mnuFileTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "&Exécuter"
         Shortcut        =   ^E
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
         Caption         =   "&Propriétés"
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
            Caption         =   "&Valeurs ASCII formatées"
         End
         Begin VB.Menu mnuCopyASCII2 
            Caption         =   "&Valeurs ASCII formatées bas niveau"
         End
         Begin VB.Menu mnuCopyASCIIReal 
            Caption         =   "&Valeurs ASCII réelles"
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
      Begin VB.Menu mnuEditTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insérer..."
      End
      Begin VB.Menu mnuEditTiret3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Tout sélectionner"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelectZone 
         Caption         =   "&Sélectionner une zone..."
      End
      Begin VB.Menu mnuSelectFromByte 
         Caption         =   "&Sélectionner à partir du byte..."
      End
      Begin VB.Menu mnuEditTiret41 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFillSelection 
         Caption         =   "&Remplir la sélection..."
      End
      Begin VB.Menu mnuCreateFileFromSelelection 
         Caption         =   "&Créer un fichier depuis la sélection..."
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
         Caption         =   "&Valeur hexa..."
      End
      Begin VB.Menu mnuFinTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplaceT 
         Caption         =   "&Remplacer texte..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuReplaceH 
         Caption         =   "&Remplacer valeur hexa..."
      End
      Begin VB.Menu mnuFindTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchForString 
         Caption         =   "&Chaine de caractères..."
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
         Caption         =   "&Donnée"
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
      Begin VB.Menu mnuStatusOK 
         Caption         =   "&Réinitialiser le status"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuRefreh 
         Caption         =   "&Rafraîchir"
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
         Caption         =   "&Début"
      End
      Begin VB.Menu mnuPosTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToOffset 
         Caption         =   "&Aller à l'offset..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu rmnuSignets 
      Caption         =   "&Signets"
      Enabled         =   0   'False
      Begin VB.Menu mnuAddSignet 
         Caption         =   "&Ajouter un signet"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRemoveSignet 
         Caption         =   "&Supprimer les signets de l'offset"
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
         Caption         =   "&Signet précédent"
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
   End
   Begin VB.Menu rmnuTools 
      Caption         =   "&Outils"
      Begin VB.Menu mnuHome 
         Caption         =   "&Démarrer une tâche..."
      End
      Begin VB.Menu mnuToolsTiret_moins1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditScript 
         Caption         =   "&Editeur de script"
      End
      Begin VB.Menu mnuExecuteScript 
         Caption         =   "&Exécuter le script"
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
         Caption         =   "&Conversion avancée..."
      End
      Begin VB.Menu mnuToolsTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompareFiles 
         Caption         =   "&Comparaison de fichiers..."
      End
      Begin VB.Menu mnuChangeDates 
         Caption         =   "&Changer les dates..."
      End
      Begin VB.Menu mnuShredder 
         Caption         =   "&Suppression de fichiers..."
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
         Caption         =   "&Découper/fusionner des fichiers..."
      End
      Begin VB.Menu mnuToolsTiret4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu rmnuWindow 
      Caption         =   "&Fenêtres"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "&En cascade"
      End
      Begin VB.Menu mnuMH 
         Caption         =   "Mosaïque &horizontale"
      End
      Begin VB.Menu mnuMV 
         Caption         =   "Mosaïque &verticale"
      End
      Begin VB.Menu mnuReorganize 
         Caption         =   "&Réorganiser les icones"
      End
      Begin VB.Menu mnuWindowsTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGestWindows 
         Caption         =   "&Gestion des fenêtres..."
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
      Begin VB.Menu mnuAbout 
         Caption         =   "&A propos"
      End
   End
   Begin VB.Menu mnuPopupGrid 
      Caption         =   "Popup_Grid"
      Visible         =   0   'False
      Begin VB.Menu mnuThisIsTheBeginnig 
         Caption         =   "&Désigner comme début de sélection"
      End
      Begin VB.Menu mnuThisIsTheEnd 
         Caption         =   "&Désigner comme fin de sélection"
      End
      Begin VB.Menu mnuPopupGridTiret0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCut 
         Caption         =   "&Couper"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copier"
         Begin VB.Menu mnuPopupCopy2 
            Caption         =   "&Valeurs ASCII formatées"
         End
         Begin VB.Menu mnuPopupCopy22 
            Caption         =   "&Valeurs ASCII formatées bas niveau"
         End
         Begin VB.Menu mnuPopupCopy3 
            Caption         =   "&Valeurs ASCII réelles"
         End
         Begin VB.Menu mnuPopupCopy 
            Caption         =   "&Valeurs hexa"
         End
      End
      Begin VB.Menu mnuPopupPaste 
         Caption         =   "&Coller"
         Begin VB.Menu mnuPasteHexaPopup 
            Caption         =   "&Valeurs hexa"
         End
         Begin VB.Menu mnuPasteASCIIPopup 
            Caption         =   "&Valeurs ASCII"
         End
      End
      Begin VB.Menu mnuPopupGridTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Supprimer la sélection..."
      End
      Begin VB.Menu mnuFillSelectionPopUp 
         Caption         =   "&Remplir la sélection..."
      End
      Begin VB.Menu mnuCreateFileFromSel2 
         Caption         =   "&Créer un fichier à partir de la sélection..."
      End
   End
   Begin VB.Menu mnuPopupExplore 
      Caption         =   "Popup_Explore"
      Visible         =   0   'False
      Begin VB.Menu mnuEditSelection 
         Caption         =   "&Editer les fichiers sélectionnés"
      End
      Begin VB.Menu mnuStatsPopup 
         Caption         =   "&Statistiques du fichier..."
      End
      Begin VB.Menu mnuPopupTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenSelectedFiles 
         Caption         =   "&Ouvrir les fichiers sélectionnés"
      End
      Begin VB.Menu mnuPopupTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenExplorer 
         Caption         =   "&Ouvrir Explorer à cet endroit..."
      End
   End
   Begin VB.Menu mnuPopupDisk 
      Caption         =   "mnuPopupDisk"
      Visible         =   0   'False
      Begin VB.Menu mnuPrevClust 
         Caption         =   "&Cluster précédent"
      End
      Begin VB.Menu mnuNextClust 
         Caption         =   "&Cluster suivant"
      End
      Begin VB.Menu mnuPopupTiret178 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrevSect 
         Caption         =   "&Secteur précédent"
      End
      Begin VB.Menu mnuNextSect 
         Caption         =   "&Secteur suivant"
      End
      Begin VB.Menu mnuPopupTiret278 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBeginingPart 
         Caption         =   "&Début de partition"
      End
      Begin VB.Menu mnuEndPart 
         Caption         =   "&Fin de partition"
      End
   End
End
Attribute VB_Name = "frmContent"
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
'FORM PARENT QUI CONTIENT LES FORM D'EDITION
'FICHIER/MEMOIRE
'CONTIENT LES MENUS
'-------------------------------------------------------

Private bDonneeForm As Boolean

Private Sub LV_ItemDblSelection(Item As ComctlLib.ListItem)
Dim Frm As Form
    'ouvre un fichier

    On Error GoTo ErrGestion

    'demande le fichier

    If cFile.FileExists(Item.Text) = False Then Exit Sub
    
    'affiche une nouvelle fenêtre
    Set Frm = New Pfm
    Call Frm.GetFile(Item.Text)
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
    
    Exit Sub
ErrGestion:
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
'suppression des fichiers
    If KeyCode = vbKeyDelete Then
        LV.DeleteSelectedItemsFromDisk False, , , True, True
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'popup menu
        Me.PopupMenu Me.mnuPopupExplore, , x + LV.Left, y + LV.Top + 300
    End If
End Sub

Private Sub MDIForm_Activate()

    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus
    
    'ferme le splash screen si il était encore ouvert
    bEndSplash = True
    
End Sub

Private Sub MDIForm_DblClick()
    Me.PopupMenu Me.rmnuTools
End Sub

Private Sub MDIForm_Load()

    'vérifie la présence de FileRenamer.exe
    If cFile.FileExists(App.Path & "\FileRenamer.exe") = False Then
        Me.mnuFileRenamer.Enabled = False
    Else
        Me.mnuFileRenamer.Enabled = True
    End If
        
    lNbChildFrm = 0
    Lang.LangFolder = App.Path & "\Lang"
    
    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus
    
    'loading de la taille de l'explorer
    Me.pctExplorer.Height = cPref.explo_Height
    
    'charge les prefs de l'explorer
    '/!\ C'est ce code qui fait charger le logiciel lentement
    '==> on cache le LV
    With LV
        .Visible = False
        
        
        .Height = cPref.explo_Height - 145
        
        If cPref.explo_DefaultPath = "Dossier du programme" Then
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
        
        .Visible = True
        .RefreshListViewOnly    '/!\ DO NOT REMOVE
    End With

End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        'affiche un popup
        Me.PopupMenu Me.rmnuTools
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'alors on récupère un drag&drop de fichiers
Dim i As Long
Dim i2 As Long
Dim m() As String
Dim Frm As Form

    'ajoute chaque fichier si existant, et le contenu de chaque dossier si dossier
    For i = 1 To Data.Files.Count
    
        If cFile.FileExists(Data.Files.Item(i)) Then
            'alors on ajoute le fichier
            'affiche une nouvelle fenêtre
            Set Frm = New Pfm
            Call Frm.GetFile(CMD.Filename)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"

        ElseIf cFile.FolderExists(Data.Files.Item(i)) Then
            'alors on ajoute le contenu du dossier
            
            'liste les fichiers
            If cFile.GetFolderFiles(Data.Files.Item(i), m, CBool(cPref.general_OpenSubFiles)) < 1 Then Exit Sub
            
            'les ouvre un par un
            For i2 = 1 To UBound(m)
                If cFile.FileExists(m(i2)) Then
                    Set Frm = New Pfm
                    Call Frm.GetFile(m(i2))
                    Frm.Show
                    lNbChildFrm = lNbChildFrm + 1
                    Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
                    DoEvents
                End If
            Next i2

        End If
        
    Next i
    
    'vire le contenu
    Data.Clear
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    
    'If Me.WindowState = vbMinimized Then frmData.Hide Else If Me.mnuEditTools.Checked Then frmData.Show
    
    LV.Width = Me.Width - 300
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload Me
    'Unload frmData
    EndProgram
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuAddSignet_Click()
'ajoute un signet sur l'offset actif
    
    'ajoute le signet au contrôle HW
    Me.ActiveForm.HW.AddSignet Me.ActiveForm.HW.Item.Offset
    Me.ActiveForm.HW.TraceSignets
    
    'ajoute le signet à la listview
    Me.ActiveForm.lstSignets.ListItems.Add Text:=CStr(Me.ActiveForm.HW.Item.Offset)
    
End Sub

Private Sub mnuAddSignetIn_Click()
'ajoute une liste de signets
    AddSignetIn False
End Sub

Private Sub mnuBeginning_Click()
'aller tout à la fin
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Max
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuCalc_Click()
    'lance la calcultarice
    Shell cFile.GetSpecialFolder(CSIDL_WINDOWS) & "\System32\calc.exe", vbNormalFocus
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuChangeDates_Click()
'changer les dates

    If Not (Me.ActiveForm Is Nothing) Then
        frmDates.txtFile.Text = Me.ActiveForm.Caption
    End If
    frmDates.Show vbModal
    
End Sub

Private Sub mnuCloseAll_Click()
'ferme toutes les fenêtres
Dim lRep As Long
Dim x As Long
    
    If Me.ActiveForm Is Nothing Then Exit Sub
        
    lRep = MsgBox("Fermer toutes les fenêtres ?", vbYesNo + vbInformation, "Attention")
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
    mnuPopupCopy2_Click
End Sub
Private Sub mnuCopyASCII2_Click()
    mnuPopupCopy22_Click
End Sub
Private Sub mnuCopyASCIIReal_Click()
    mnuPopupCopy3_Click
End Sub
Private Sub mnuCopyhexa_Click()
    mnuPopupCopy_Click
End Sub

Private Sub mnuCreateFileFormSel_Click()
'créé un fichier à partir de la sélection
Dim curSize As Currency
Dim bOver As Boolean

    On Error GoTo CancelPushed

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'calcule la taille du fichier résultat
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset - Me.ActiveForm.HW.FirstSelectionItem.Offset
    
    If curSize > 200000000 Then
        'fichier >200Mo, demande de confirmation
        If MsgBox("Le fichier créé aura une taille supérieure à 200Mo. Continuer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If
    
    'demande le fichier résultat
    With CMD
        .CancelError = True
        .DialogTitle = "Sauvegarde de la sélection"
        .Filter = "Tous |*.*|Fichier de donnée|*.dat"
        .ShowSave
        If cFile.FileExists(.Filename) Then
            'le fichier existe déjà
            If MsgBox("Le fichier existe déjà, l'écraser ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
        End If
        'créé un fichier vide
        cFile.CreateEmptyFile .Filename, True
    End With
    
    
    If TypeOfForm(Me.ActiveForm) = "Fichier" Then
        'alors on sauvegarde en utilisant ReadFile dans un fichier
        
    ElseIf TypeOfForm(Me.ActiveForm) = "Disque" Then
        'alors disque ==> readfile
        
    Else
        'alors lecture en mémoire
        
    End If
    
CancelPushed:
End Sub

Private Sub mnuCreateFileFromSel2_Click()
'créé un fichier depuis la sélection
Dim x As Long
Dim y As Long
Dim s2 As String
Dim s() As String
Dim sFile As String
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curSize As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    On Error GoTo CancelPushed

    Me.Sb.Panels(1).Text = "Status=[Creating file from selection]"
    
    'détermine la taille
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
        Me.ActiveForm.HW.FirstSelectionItem.Offset - Me.ActiveForm.HW.FirstSelectionItem.Col + 1
    
    'détermine la position du premier offset
    curPos = Me.ActiveForm.HW.FirstSelectionItem.Offset + Me.ActiveForm.HW.FirstSelectionItem.Col - 1
        
    With CMD
        .CancelError = True
        .DialogTitle = "Sélection du fichier à sauvegarder"
        .Filter = "Tous|*.*"
        .ShowSave
        sFile = .Filename
    End With
    
    If cFile.FileExists(sFile) Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà. Le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile et sauvegarde à la volée (buffers de 500Ko)
            
            If curSize <= 512000 Then
                'alors tout rentre dans un buffer
                'récupère la string
                s2 = GetBytesFromFile(Me.ActiveForm.Caption, curSize, curPos)
                GoTo CreateMyFileFromOneBuffer
            Else
                'plusieurs buffers nécessaire
                
                
                GoTo CreateMyFileFromBuffers
            End If
        
        Case "Processus"
            'sauvegarde avec un buffer de 50Ko
            If curSize <= 512000 Then
                'alors tout rentre dans un buffer
                s2 = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))
                GoTo CreateMyFileFromOneBuffer
            Else
                'alors plusieurs buffers nécessaires
                
                
                
                GoTo CreateMyFileFromBuffers
            End If
            
        Case "Disque"
            'sauvegarde avec un buffer de 50Ko
            
            'redéfinit correctement la position et la taille (doivent être multiple du nombre
            'de bytes par secteur)
            curPos2 = ByND(curPos, Me.ActiveForm.GetDriveInfos.BytesPerSector)
            curSize2 = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, Me.ActiveForm.GetDriveInfos.BytesPerSector)
            
            If curSize2 <= 512000 Then
                'alors tout rentre dans un buffer
                'récupère la string
                DirectReadS Me.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / Me.ActiveForm.GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    Me.ActiveForm.GetDriveInfos.BytesPerSector, s2
                    
                'recoupe la string pour récupérer ce qui intéresse vraiment
                s2 = Mid$(s2, curPos - curPos2 + 1, curSize)
                GoTo CreateMyFileFromOneBuffer
            Else
                'plusieurs buffers nécessaires
                
                
                GoTo CreateMyFileFromBuffers
            End If
    End Select

CreateMyFileFromOneBuffer:
    'sauvegarde le fichier (un seul buffer)
    cFile.SaveDATAinFile sFile, s2, True   'lance la sauvegarde
    
    GoTo CancelPushed
    
CreateMyFileFromBuffers:
    'sauvegarde le fichier (plusieurs buffers)

CancelPushed:
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuCreateFileFromSelelection_Click()
    Call mnuCreateFileFromSel2_Click    'créé fichier à partir de la sélection
End Sub

Private Sub mnuCut_Click()
'coupe la sélection
    Call mnuPopupCut_Click
End Sub

Private Sub mnuCutCopyFiles_Click()
    'découpage/collage
    frmCut.Show vbModal
End Sub

Private Sub mnuDelete_Click()
'efface les éléments de la sélection
    
    'affiche la boite de demande de confirmation
    frmCreateBackup.Show vbModal
    
    'procède à la suppression de la zone et sauvegarde dans un backup
    If bAcceptBackup Then frmContent.ActiveForm.DeleteZone
    
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
    
    Me.ActiveForm.VS.Value = IIf((Me.ActiveForm.VS.Value + l) < Me.ActiveForm.VS.Max, Me.ActiveForm.VS.Value + l, Me.ActiveForm.VS.Max)
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuEditScript_Click()
'éditeur de script
    frmScript.Show
End Sub

Private Sub mnuEditSelection_Click()
'édite les fichiers sélectionnés dans le LV
Dim Frm As Form
Dim sFile() As ComctlLib.ListItem
Dim x As Long

    'On Error GoTo ErrGestion
    LV.GetSelectedItems sFile
    
    For x = 1 To UBound(sFile)
        If cFile.FileExists(sFile(x).Text) Then
            'affiche une nouvelle fenêtre
            Set Frm = New Pfm
            Call Frm.GetFile(sFile(x).Text)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
        End If
        DoEvents
    Next x
    
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
'tout au début

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Min
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuBeginingPart_Click()
'va tout au début de la partition
    
    If Me.ActiveForm Is Nothing Then Exit Sub

    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Min
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuEndPart_Click()
'va tout à la fin de la partition

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Max
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuExecuteScript_Click()
'execute le script actif
End Sub

Private Sub mnuFileRenamer_Click()
'lance FileRenamer.exe
    ShellExecute Me.hwnd, "Open", App.Path & "\FileRenamer.exe", "", App.Path, 1
End Sub

Private Sub mnuFreeForum_Click()
'forum de discussion
    ShellExecute Me.hwnd, "Open", "https://sourceforge.net/forum/forum.php?forum_id=654034", "", App.Path, 1
End Sub

Private Sub mnuHelpForum_Click()
'forum de demande d'aide
    ShellExecute Me.hwnd, "Open", "https://sourceforge.net/forum/forum.php?forum_id=654035", "", App.Path, 1
End Sub

Private Sub mnuHome_Click()
'lance frmHome
    frmHome.Show
    PremierPlan frmHome, MettreAuPremierPlan
End Sub

Private Sub mnuNextClust_Click()
'va au cluster suivant
Dim ActualClust As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    'détermine le cluster actuel
    ActualClust = Int((Me.ActiveForm.VS.Value / Me.ActiveForm.GetDriveInfos.BytesPerCluster) * 16)
    ActualClust = ActualClust + 1
    
    Me.ActiveForm.VS.Value = Int((ActualClust / 16) * Me.ActiveForm.GetDriveInfos.BytesPerCluster)
    If Me.ActiveForm.VS.Value > Me.ActiveForm.VS.Max Then Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Max
    
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
    
End Sub

Private Sub mnuNextSect_Click()
'va au secteur suivant
Dim ActualSect As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    'détermine le cluster actuel
    ActualSect = Int((Me.ActiveForm.VS.Value / Me.ActiveForm.GetDriveInfos.BytesPerSector) * 16)
    ActualSect = ActualSect + 1
    
    Me.ActiveForm.VS.Value = Int((ActualSect / 16) * Me.ActiveForm.GetDriveInfos.BytesPerSector)
    If Me.ActiveForm.VS.Value > Me.ActiveForm.VS.Max Then Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Max
    
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuPrevClust_Click()
'va au cluster prédédent
Dim ActualClust As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    'détermine le cluster actuel
    ActualClust = Int((Me.ActiveForm.VS.Value / Me.ActiveForm.GetDriveInfos.BytesPerCluster) * 16)
    ActualClust = ActualClust - 1
    
    Me.ActiveForm.VS.Value = Int((ActualClust / 16) * Me.ActiveForm.GetDriveInfos.BytesPerCluster)
    If Me.ActiveForm.VS.Value < Me.ActiveForm.VS.Min Then Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Min
    
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
    
End Sub

Private Sub mnuPrevSect_Click()
'va au secteur prédédent
Dim ActualSect As Long

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'détermine le cluster actuel
    ActualSect = Int((Me.ActiveForm.VS.Value / Me.ActiveForm.GetDriveInfos.BytesPerSector) * 16)
    ActualSect = ActualSect - 1
    
    Me.ActiveForm.VS.Value = Int((ActualSect / 16) * Me.ActiveForm.GetDriveInfos.BytesPerSector)
    If Me.ActiveForm.VS.Value < Me.ActiveForm.VS.Min Then Me.ActiveForm.VS.Value = Me.ActiveForm.VS.Min
    
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuErr_Click()
'affiche la form de rapport d'erreur
    frmLogErr.Show vbModal
End Sub

Private Sub mnuExecute_Click()
'exécute le fichier (temporaire)
Dim sExt As String

    If ActiveForm Is Nothing Then Exit Sub
    
    'obtient la termaison
    sExt = cFile.GetFileExtension(Me.ActiveForm.Caption)
    
    ExecuteTempFile Me.hwnd, Me.ActiveForm, sExt
End Sub

Private Sub mnuExit_Click()
    'quitte
    EndProgram
End Sub

Private Sub mnuExploreDisk_Click()
'affiche ou pas l'explorer de disque

    If TypeOfForm(Me.ActiveForm) = "Disque" Then
        mnuExploreDisk.Checked = Not (mnuExploreDisk.Checked)
        Me.ActiveForm.FV.Visible = mnuExploreDisk.Checked
        Me.ActiveForm.FV2.Visible = mnuExploreDisk.Checked
        Me.ActiveForm.pctPath.Visible = mnuExploreDisk.Checked
        Me.ActiveForm.FrameFrag.Visible = mnuExploreDisk.Checked
    End If
    Call frmContent.ActiveForm.ResizeMe
       
End Sub

Private Sub mnuExploreDisplay_Click()
    mnuExploreDisplay.Checked = Not (mnuExploreDisplay.Checked)
    pctExplorer.Visible = mnuExploreDisplay.Checked
    Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuFileViewMode_Click()
'affiche le mode "lecture de fichier", cad que les valeurs hexa

End Sub

Private Sub mnuFillSelection_Click()
'remplit la sélection du HW

    If Me.ActiveForm Is Nothing Then Exit Sub

    frmFillSelection.Show vbModal

End Sub

Private Sub mnuFillSelectionPopUp_Click()
    mnuFillSelection_Click
End Sub

Private Sub mnuGestWindows_Click()
'affiche la form de gestion des fenêtres
    frmGestWindows.Show vbModal
End Sub

Private Sub mnuGoToOffset_Click()
'va à une valeur particulière de l'offset
Dim s As String
Dim l As Long

    On Error Resume Next
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    s = InputBox("Aller à quel offset ?", "Changement d'offset")
    If StrPtr(s) = 0 Then Exit Sub  'cancel
    
    'alors on va à l'offset (si possible)
    l = By16(Int(Abs(Val(s))))  'formatage de l'offset
    
    If l <= By16(Me.ActiveForm.HW.MaxOffset) Then
        'alors c'est ok
        Me.ActiveForm.VS.Value = (l / 16)
        Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)    'refresh
    End If
    
End Sub

Private Sub mnuHelp_Click()
'affiche l'aide
    On Error GoTo 5
    'ShellExecute Me.hWnd, "open", App.Path & "\aide.chm", vbNullString, vbNullString, 1
    
    ' Dim s() As Byte
    
    ' DirectRead "l:\", 0, 512, s()
    ' Dim x As Byte
    'For x = 0 To 15
    '     MsgBox s(10), , s(5)
    'Next x
    ' frmSaveProcess.Show vbModal
    'Randomize
    'Lang.WriteIniFileFormIDEform
    
    'switch basique de langue
    'If Lang.Language = "French" Then
    '    Lang.Language = "English"
    'Else
    '    Lang.Language = "French"
    'End If
    
    
    
    'Err.Raise Int(Rnd * 50)
5
    'clsERREUR.AddError "mnuHelp_Click"
End Sub

Private Sub mnuInformations_Click()
'affiche ou non les infos
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.mnuInformations.Checked = Not (Me.mnuInformations.Checked)
    Me.ActiveForm.FrameInfos.Visible = Me.mnuInformations.Checked
    If TypeOfForm(Me.ActiveForm) = "Disque" Then Me.ActiveForm.FrameInfo2.Visible = Me.mnuInformations.Checked
    Call frmContent.ActiveForm.ResizeMe
End Sub

Private Sub mnuInterpretAdvanced_Click()
'conversion avancée
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
'invite à demarrer un nouveau processus
    ShowRunBox Me.hwnd
End Sub

Private Sub mnuOpen_Click()
Dim Frm As Form
    'ouvre un fichier

    On Error GoTo ErrGestion

    'demande le fichier
    CMD.Filter = "Tous |*.*"
    CMD.DialogTitle = "Sélection du fichier à ouvrir"
    CMD.CancelError = True
    CMD.ShowOpen

    If cFile.FileExists(CMD.Filename) = False Then Exit Sub
    
    'affiche une nouvelle fenêtre
    Set Frm = New Pfm
    Call Frm.GetFile(CMD.Filename)
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
    
    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus
    
    Exit Sub
ErrGestion:
End Sub

Private Sub mnuOpenDisk_Click()
'ouvre un disque physique

    frmDrive.Show vbModal
End Sub

Private Sub mnuOpenExplorer_Click()
'ouvre explorer.exe à l'emplacement pointé par LV

    Shell "explorer.exe " & LV.Path, vbNormalFocus
End Sub

Private Sub mnuOpenFolder_Click()
'ouvre un dossier
Dim m() As String
Dim sDir As String
Dim Frm As Form
Dim x As Long

    'sélectionne un répertoire
    sDir = cFile.BrowseForFolder("Sélectionner un répertoire", Me.hwnd)
    
    'teste la validité du répertoire
    If cFile.FolderExists(sDir) = False Then Exit Sub
    
    'liste les fichiers
    If cFile.GetFolderFiles(sDir, m, CBool(cPref.general_OpenSubFiles)) < 1 Then Exit Sub
    
    'les ouvre un par un
    For x = 1 To UBound(m)
        If cFile.FileExists(m(x)) Then
            Set Frm = New Pfm
            Call Frm.GetFile(m(x))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
            DoEvents
        End If
    Next x
  
    'Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus

End Sub

Private Sub mnuOpenInBN_Click()
'ouvre le fichier dans le bloc notes
Dim x As Long

On Error Resume Next

    If ActiveForm Is Nothing Then Exit Sub
    
    If cFile.FileExists(Me.ActiveForm.Caption) = False Then
        'pas de fichier
        MsgBox "Fichier inexistant", vbInformation, "Impossible d'ouvrir"
    End If
    
    If cFile.GetFileSize(Me.ActiveForm.Caption) > 1000000 Then
        'fichier de plus de 700Ko
        x = MsgBox("Votre fichier fait plus de 1Mo." & vbNewLine & "Il n'est pas conseillé d'ouvrir un fichier de cette taille" & vbNewLine & "avec le bloc-notes. Continuer ?", vbInformation + vbYesNo, "Attention")
        If Not (x = vbYes) Then Exit Sub
    End If
        
    Shell "notepad " & Me.ActiveForm.Caption, vbNormalFocus
End Sub

Private Sub mnuOpenProcess_Click()
'ouvre un processus en mémoire
  
    'affiche la liste des process
    frmProcesses.Show vbModal

End Sub

Private Sub mnuOpenSelectedFiles_Click()
'ouvre les fichiers sélectionnés dans le LV
Dim sFile() As ListItem
Dim x As Long

    'obtient la liste des sélections
    LV.GetSelectedItems sFile
    
    For x = 1 To UBound(sFile)
        ShellExecute Me.hwnd, "open", sFile(x).Text, vbNullString, vbNullString, 1
    Next x
    
End Sub

Private Sub mnuOpenSignetsList_Click()
'ouvre une liste de signet
    AddSignetIn True
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPopupCopy_Click()
'copier la sélection (hexa)
Dim x As Long
Dim y As Long
Dim s As String
Dim s2 As String
Dim curSize As Currency
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    s = vbNullString    'contiendra la string à copier
    
    'vide le clipboard
    Clipboard.Clear
    
    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"

    'détermine la taille
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
        Me.ActiveForm.HW.FirstSelectionItem.Offset - Me.ActiveForm.HW.FirstSelectionItem.Col + 1
    
    'détermine la position du premier offset
    curPos = Me.ActiveForm.HW.FirstSelectionItem.Offset + Me.ActiveForm.HW.FirstSelectionItem.Col - 1
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile
                        
            'récupère la string
            s = GetBytesFromFile(Me.ActiveForm.Caption, curSize, curPos)
            
        Case "Processus"
        
            'récupère la string
            s = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))

        Case "Disque"
        
            'redéfinit correctement la position et la taille (doivent être multiple du nombre
            'de bytes par secteur)
            curPos2 = ByND(curPos, Me.ActiveForm.GetDriveInfos.BytesPerSector)
            curSize2 = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, Me.ActiveForm.GetDriveInfos.BytesPerSector)

            'récupère la string
            DirectReadS Me.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                curPos2 / Me.ActiveForm.GetDriveInfos.BytesPerSector, CLng(curSize2), _
                Me.ActiveForm.GetDriveInfos.BytesPerSector, s
                
            'recoupe la string pour récupérer ce qui intéresse vraiment
            s = Mid$(s, curPos - curPos2 + 1, curSize)
            
    End Select

    'formate la string
    s2 = vbNullString
    For x = 1 To Len(s)
        If (x Mod 1000) = 0 Then DoEvents 'rend la main
        s2 = s2 & Str2Hex_(Mid$(s, x, 1))
    Next x
            
    Clipboard.SetText s2
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuPopupCopy22_Click()
'copier la sélection (strings) formatée en bas niveau
Dim x As Long
Dim y As Long
Dim s As String
Dim curSize As Currency
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    s = vbNullString    'contiendra la string à copier
    
    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"
        
    'vide le clipboard
    Clipboard.Clear

    'détermine la taille
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
        Me.ActiveForm.HW.FirstSelectionItem.Offset - Me.ActiveForm.HW.FirstSelectionItem.Col + 1
    
    'détermine la position du premier offset
    curPos = Me.ActiveForm.HW.FirstSelectionItem.Offset + Me.ActiveForm.HW.FirstSelectionItem.Col - 1
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile
                        
            'récupère la string
            s = GetBytesFromFile(Me.ActiveForm.Caption, curSize, curPos)
            
        Case "Processus"
        
            'récupère la string
            s = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))
          
        Case "Disque"
        
            'redéfinit correctement la position et la taille (doivent être multiple du nombre
            'de bytes par secteur)
            curPos2 = ByND(curPos, Me.ActiveForm.GetDriveInfos.BytesPerSector)
            curSize2 = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, Me.ActiveForm.GetDriveInfos.BytesPerSector)

            'récupère la string
            DirectReadS Me.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                curPos2 / Me.ActiveForm.GetDriveInfos.BytesPerSector, CLng(curSize2), _
                Me.ActiveForm.GetDriveInfos.BytesPerSector, s
                
            'recoupe la string pour récupérer ce qui intéresse vraiment
            s = Mid$(s, curPos - curPos2 + 1, curSize)
            
    End Select

    'formate la string
    s = Replace$(s, vbNullChar, Chr$(32), , , vbBinaryCompare)
    
    Clipboard.SetText s
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuPopupCopy2_Click()
'copier la sélection (strings) formatée
Dim x As Long
Dim y As Long
Dim s As String
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curSize As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    s = vbNullString    'contiendra la string à copier
    
    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"
        
    'vide le clipboard
    Clipboard.Clear
    
    'détermine la taille
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
        Me.ActiveForm.HW.FirstSelectionItem.Offset - Me.ActiveForm.HW.FirstSelectionItem.Col + 1
    
    'détermine la position du premier offset
    curPos = Me.ActiveForm.HW.FirstSelectionItem.Offset + Me.ActiveForm.HW.FirstSelectionItem.Col - 1
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile

            'récupère la string
            s = GetBytesFromFile(Me.ActiveForm.Caption, curSize, curPos)
            
        Case "Processus"
        
            'récupère la string
            s = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))

        Case "Disque"
        
            'redéfinit correctement la position et la taille (doivent être multiple du nombre
            'de bytes par secteur)
            curPos2 = ByND(curPos, Me.ActiveForm.GetDriveInfos.BytesPerSector)
            curSize2 = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, Me.ActiveForm.GetDriveInfos.BytesPerSector)

            'récupère la string
            DirectReadS Me.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                curPos2 / Me.ActiveForm.GetDriveInfos.BytesPerSector, CLng(curSize2), _
                Me.ActiveForm.GetDriveInfos.BytesPerSector, s
                
            'recoupe la string pour récupérer ce qui intéresse vraiment
            s = Mid$(s, curPos - curPos2 + 1, curSize)
            
    End Select

    'formate la string
    s = FormatednString(s)
            
    Clipboard.SetText s
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuPopupCopy3_Click()
'copie les valeurs ASCII réelles vers le clipboard
'/!\ NULL TERMINATED STRING
Dim x As Long
Dim y As Long
Dim s As String
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curSize As Currency
Dim curPos As Currency

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    s = vbNullString    'contiendra la string à copier

    Me.Sb.Panels(1).Text = "Status=[Copying to ClipBoard]"
    
    'vide le clipboard
    Clipboard.Clear
    
    'détermine la taille
    curSize = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
        Me.ActiveForm.HW.FirstSelectionItem.Offset - Me.ActiveForm.HW.FirstSelectionItem.Col + 1
    
    'détermine la position du premier offset
    curPos = Me.ActiveForm.HW.FirstSelectionItem.Offset + Me.ActiveForm.HW.FirstSelectionItem.Col - 1
            
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile
            
            'récupère la string
            s = GetBytesFromFile(Me.ActiveForm.Caption, curSize, curPos)
        
        Case "Processus"
            
            'récupère la string
            s = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos), CLng(curSize))
        
        Case "Disque"
        
            'redéfinit correctement la position et la taille (doivent être multiple du nombre
            'de bytes par secteur)
            curPos2 = ByND(curPos, Me.ActiveForm.GetDriveInfos.BytesPerSector)
            curSize2 = Me.ActiveForm.HW.SecondSelectionItem.Offset + Me.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, Me.ActiveForm.GetDriveInfos.BytesPerSector)

            'récupère la string
            DirectReadS Me.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                curPos2 / Me.ActiveForm.GetDriveInfos.BytesPerSector, CLng(curSize2), _
                Me.ActiveForm.GetDriveInfos.BytesPerSector, s
                
            'recoupe la string pour récupérer ce qui intéresse vraiment
            s = Mid$(s, curPos - curPos2 + 1, curSize)
            
    End Select

    Clipboard.SetText s, vbCFText   'format fichier texte
    Me.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

Private Sub mnuPopupCut_Click()
'couper la sélection

End Sub

Private Sub mnuPopupPaste_Click()
'coller

End Sub

Private Sub mnuPrint_Click()
'impression
    frmPrint.Show vbModal
End Sub

Private Sub mnuProcesses_Click()
'gestionnaire très simple de processus
    frmProcess.Show
End Sub

Private Sub mnuProperty_Click()
'affiche les propriétés du fichier
    frmPropertyShow.Show
End Sub

Private Sub mnuRedo_Click()
    Call Me.ActiveForm.RedoM
End Sub

Private Sub mnuRefreh_Click()
'refresh

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'on refresh le HW
    Call Me.ActiveForm.VS_Change(Me.ActiveForm.VS.Value)
End Sub

Private Sub mnuRemoveAll_Click()
'supprime tous les signets ==> demande confirmation
   
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'confirmation
    If MsgBox("Êtes vous sur de vouloir supprimer tous les signets ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    
    Me.ActiveForm.HW.RemoveAllSignets
    Me.ActiveForm.lstSignets.ListItems.Clear
End Sub

Private Sub mnuRemoveSignet_Click()
'supprime un signet, si existant
Dim x As Long

    If Me.ActiveForm Is Nothing Then Exit Sub

    If Me.ActiveForm.HW.IsSignet(Me.ActiveForm.HW.Item.Offset) Then
    
        While Me.ActiveForm.HW.IsSignet(Me.ActiveForm.HW.Item.Offset)
            'on supprime
            Me.ActiveForm.HW.RemoveSignet Val(Me.ActiveForm.HW.Item.Offset)
        Wend
        
        'enlève du listview
        For x = Me.ActiveForm.lstSignets.ListItems.Count To 1 Step -1
            If Me.ActiveForm.lstSignets.ListItems.Item(x).Text = CStr(Me.ActiveForm.HW.Item.Offset) Then
                Me.ActiveForm.lstSignets.ListItems.Remove x
            End If
        Next x
    End If
    
End Sub

Private Sub mnuReorganize_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuSaveSignets_Click()
'enregistre la liste des signets de la form active
Dim s As String
Dim lFile As Long
Dim x As Long

    On Error GoTo ErrGestion
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    If Me.ActiveForm.lstSignets.ListItems.Count = 0 Then Exit Sub 'pas de signets
    
    'enregistrement ==> choix du fichier
    With CMD
        .CancelError = True
        .Filename = Me.ActiveForm.Caption & ".sig"
        .DialogTitle = "Enregistrement de la liste des signets"
        .Filter = "Liste de signets |*.sig|"
        .InitDir = App.Path
        .ShowSave
        s = .Filename
    End With

    If cFile.FileExists(s) Then
        'message de confirmation
        x = MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention")
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'ouvre le fchier
    lFile = FreeFile
    Open s For Output As lFile
    
    'enregistre les entrées
    For x = 1 To Me.ActiveForm.lstSignets.ListItems.Count
        Write #lFile, Me.ActiveForm.lstSignets.ListItems.Item(x) & "|" & Me.ActiveForm.lstSignets.ListItems.Item(x).SubItems(1)
    Next x
    
    Close lFile
    
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
'tout sélectionner
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    Me.ActiveForm.HW.SelectZone 0, 0, 16 - (By16(Me.ActiveForm.HW.MaxOffset) _
    - Me.ActiveForm.HW.MaxOffset), By16(Me.ActiveForm.HW.MaxOffset) - 16
    Me.ActiveForm.HW.Refresh
    
    'refresh le label qui contient la taille de la sélection
    Me.ActiveForm.Sb.Panels(4).Text = "Sélection=[" & CStr(Me.ActiveForm.HW.NumberOfSelectedItems) & " bytes]"
    Me.ActiveForm.Label2(9) = Me.ActiveForm.Sb.Panels(4).Text
End Sub

Private Sub mnuSelectFromByte_Click()
'sélection à partir d'un byte
    frmSelect2.Show vbModal
End Sub

Private Sub mnuSelectZone_Click()
'sélectionne une zone définie
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
'supprime définitivement des fichiers
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
'signet précédent

    If Me.ActiveForm Is Nothing Then Exit Sub
    Me.ActiveForm.HW.FirstOffset = Me.ActiveForm.HW.GetPrevSignet(Me.ActiveForm.HW.Item.Offset)
    Me.ActiveForm.HW.Refresh
    Me.ActiveForm.VS.Value = Me.ActiveForm.HW.FirstOffset / 16
End Sub

Private Sub mnuSourceForge_Click()
'page source forge
    ShellExecute Me.hwnd, "Open", "https://sourceforge.net/projects/hexeditorvb/", "", App.Path, 1
End Sub

Private Sub mnuStats_Click()
'affiche les statistiques du fichier
Dim Frm As Form

    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'affiche la form d'analyse
    Set Frm = New frmAnalys
    Frm.GetFile Me.ActiveForm.Caption
    Frm.Show
    
End Sub

Private Sub mnuStatsPopup_Click()
'affiche les stats des fichiers sélectionnés dans LV
Dim Frm As Form
Dim sFile() As ListItem
Dim x As Long

    'On Error GoTo ErrGestion

    LV.GetSelectedItems sFile
    
    For x = 1 To UBound(sFile)
        If cFile.FileExists(sFile(x).Text) Then
            'affiche une nouvelle fenêtre
            Set Frm = New frmAnalys
            Call Frm.GetFile(sFile(x).Text)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            Me.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
        End If
        DoEvents
    Next x
    
    Exit Sub
    
ErrGestion:
End Sub

Private Sub mnuStatusOK_Click()
'réinitialise le status
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
    frmTable.CreateTable AllTables
End Sub

Private Sub mnuTdec2Ascii_Click()
'affiche la table
    frmTable.Show
    frmTable.CreateTable HEX_ASCII
End Sub

Private Sub mnuSaveAs_Click()
'enregistrer sous
Dim sFile As String
Dim sPath As String
Dim lFile As Long
Dim x As Long

    On Error GoTo GestionErr

    If Me.ActiveForm Is Nothing Then Exit Sub

    'il faut sauvegarder en prenant compte des 2 changelist de pfm
    
    With CMD
        .CancelError = True
        .DialogTitle = "Sauvegarder sous..."
        .Filter = "Tous|*.*"
        .ShowSave
        sPath = .Filename
    End With
    
    If cFile.FileExists(sPath) Then
        'message de confirmation
        x = MsgBox("Le fichier existe déjà, le remplacer ?", vbInformation + vbYesNo, "Attention")
        If Not (x = vbYes) Then Exit Sub
    End If
    
    'efface le précédent fichier
    cFile.KillFile sPath
    
    'créé le fichier
    Call Me.ActiveForm.GetNewFile(sPath)

GestionErr:
End Sub

Private Sub mnuThisIsTheBeginnig_Click()
'marque le début de la sélection à cet offset
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    frmContent.ActiveForm.HW.FirstSelectionItem.Offset = frmContent.ActiveForm.HW.Item.Offset

End Sub

Private Sub mnuThisIsTheEnd_Click()
'marque la fin de la sélection à cet offset
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    frmContent.ActiveForm.HW.SecondSelectionItem.Offset = frmContent.ActiveForm.HW.Item.Offset
End Sub

Private Sub mnuUndo_Click()
    Call Me.ActiveForm.UndoM
End Sub

Private Sub mnuVbfrance_Click()
'vbfrance.com
    ShellExecute Me.hwnd, "Open", "http://www.vbfrance.com/auteurdetail.aspx?ID=523601&print=1", "", App.Path, 1
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
'affiche le nombre d'erreurs enregistrées dans le menu "Rapport..."
    Me.mnuErr.Caption = "Rapport d'erreurs (" & Trim$(Str$(clsERREUR.NumberOfErrorInLogFile)) & ")"
End Sub

Private Sub Timer1_Timer()
    Call frmContent.ChangeEnabledMenus  'active ou pas certaines entrées dans les menus
    
    'réaffiche le choix des dossiers (si choix = true)
    frmContent.pctExplorer.Visible = frmContent.mnuExploreDisplay.Checked And Not (TypeOfActiveForm = "Disk")
    
    'actualise les fonctions Undo/Redo
    If Not (Me.ActiveForm Is Nothing) Then
        ModifyHistoEnabled
    Else
        'pas de fichier ouvert ==> enabled=false
        Me.mnuRedo.Enabled = False
        Me.mnuUndo.Enabled = False
        Me.Toolbar1.Buttons.Item(12).Enabled = False
        Me.Toolbar1.Buttons.Item(13).Enabled = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'appui sur les icones

    Select Case Button.Key
    
        Case "OpenFile"
            Call mnuOpen_Click
        Case "HomeOpen"
            'affiche la boite de dialogue Home (choix des différentes actions à faire)
            frmHome.Show
            PremierPlan frmHome, MettreAuPremierPlan
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
            Call mnuPopupCopy2_Click
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
            Call mnuPopupCut_Click
        Case "Paste"
            
    End Select

End Sub

'-------------------------------------------------------
'ajoute (ou ouvre si overwrite) une liste de signets
'-------------------------------------------------------
Private Sub AddSignetIn(ByVal bOverWrite As Boolean)
Dim s As String
Dim lFile As Long
Dim x As Long
Dim sTemp As String
Dim l As Long

    On Error GoTo ErrGestion
    
    If Me.ActiveForm Is Nothing Then Exit Sub
    
    'ouverture ==> choix du fichier
    With CMD
        .CancelError = True
        .DialogTitle = "Ouverture d'une liste de signets"
        .Filter = "Liste de signets |*.sig|"
        .InitDir = App.Path
        .ShowOpen
        s = .Filename
    End With
    
    If bOverWrite Then
        Me.ActiveForm.lstSignets.ListItems.Clear
        Me.ActiveForm.HW.RemoveAllSignets
    End If
    
    'ouvre le fchier
    lFile = FreeFile
    Open s For Input As lFile
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
    
ErrGestion:
End Sub

'-------------------------------------------------------
'permet de masquer ou d'afficher les menus en fonction du type de form qui est active
'-------------------------------------------------------
Public Function ChangeEnabledMenus()

    If TypeOfActiveForm = "Mem" Then
        'ActiveForm=MemPfm
        'alors on masque certaines options des menus
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = False
        Me.mnuExecute.Enabled = False
        Me.mnuCut.Enabled = False
        Me.mnuOpenInBN = False
        Me.mnuStats.Enabled = False
        Me.mnuPopupCut.Enabled = False
        Me.mnuDelete.Enabled = False
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
    ElseIf (Me.ActiveForm Is Nothing) = False And TypeOfActiveForm = "Pfm" Then
        'ActiveForm=Pfm
        'alors on affiche les options qui auraient pu être cachées
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = True
        Me.mnuExecute.Enabled = True
        Me.mnuCut.Enabled = True
        Me.mnuOpenInBN = True
        Me.mnuStats.Enabled = True
        Me.mnuPopupCut.Enabled = True
        Me.mnuDelete.Enabled = True
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
    ElseIf Me.ActiveForm Is Nothing Then
        'ActiveForm = nothing
        Me.mnuExploreDisk.Enabled = False
        Me.mnuSave.Enabled = False
        Me.mnuExecute.Enabled = False
        Me.mnuCut.Enabled = False
        Me.mnuOpenInBN = False
        Me.mnuStats.Enabled = False
        Me.mnuPopupCut.Enabled = False
        Me.mnuDelete.Enabled = False
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
    Else
        'diskfrm
        Me.mnuExploreDisk.Enabled = True
        Me.mnuSave.Enabled = True
        Me.mnuExecute.Enabled = True
        Me.mnuCut.Enabled = True
        Me.mnuOpenInBN = True
        Me.mnuStats.Enabled = True
        Me.mnuPopupCut.Enabled = True
        Me.mnuDelete.Enabled = True
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
    End If
    
End Function
