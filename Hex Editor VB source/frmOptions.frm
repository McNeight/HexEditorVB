VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C60799F1-7AA3-45BA-AFBF-5BEAB08BC66C}#1.0#0"; "HexViewer_OCX.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7455
      Index           =   0
      Left            =   240
      TabIndex        =   65
      Top             =   720
      Width           =   9495
      Begin VB.PictureBox pctCauzeOfManifest 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   9495
         TabIndex        =   67
         Top             =   4080
         Width           =   9500
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   11
            Left            =   8400
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   82
            Top             =   720
            Width           =   375
         End
         Begin VB.ComboBox cbGrid 
            Height          =   315
            ItemData        =   "frmOptions.frx":058A
            Left            =   5640
            List            =   "frmOptions.frx":05A0
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Tag             =   "lang_ok"
            ToolTipText     =   "Type de grille à afficher"
            Top             =   960
            Width           =   3855
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   7
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   80
            Top             =   1680
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   6
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   79
            Top             =   1440
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   10
            Left            =   8400
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   78
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   9
            Left            =   8400
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   77
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   8
            Left            =   8400
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   76
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   5
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   75
            Top             =   1200
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   4
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   74
            Top             =   960
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   3
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   73
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   2
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   72
            Top             =   480
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   1
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   71
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   0
            Left            =   3120
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   70
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton optHex 
            Caption         =   "Offsets en hexadécimal"
            Height          =   195
            Left            =   4800
            TabIndex        =   69
            ToolTipText     =   "Affiche les offsets en base hexadécimale"
            Top             =   1440
            Width           =   3135
         End
         Begin VB.OptionButton optDec 
            Caption         =   "Offsets en décimal"
            Height          =   255
            Left            =   4800
            TabIndex        =   68
            ToolTipText     =   "Affiche les offsets en base décimale"
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur des signets"
            Height          =   255
            Index           =   12
            Left            =   4800
            TabIndex        =   95
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Grille"
            Height          =   255
            Index           =   11
            Left            =   4800
            TabIndex        =   94
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de fond de titre"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   93
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur des lignes"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   92
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur des éléments modifiés sélectionnés"
            Height          =   255
            Index           =   10
            Left            =   4800
            TabIndex        =   91
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur des éléments modifiés"
            Height          =   255
            Index           =   9
            Left            =   4800
            TabIndex        =   90
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la sélection"
            Height          =   255
            Index           =   8
            Left            =   4800
            TabIndex        =   89
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la police de la base"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   88
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la police du titre Offset"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   87
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la police des strings"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   86
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la police des valeurs hexa"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   85
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la police de l'offset"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   84
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de fond"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   2055
         End
      End
      Begin HexViewer_OCX.HexViewer HW 
         Height          =   3735
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   6588
         strTag1         =   "0"
         strTag2         =   "0"
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Index           =   1
      Left            =   120
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox pctManifest 
         BorderStyle     =   0  'None
         Height          =   5895
         Index           =   0
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   9435
         TabIndex        =   61
         Top             =   360
         Width           =   9435
         Begin VB.CheckBox chkContextMenu 
            Caption         =   "Mettre une entrée au menu contextuel de Windows pour les fichiers"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   64
            ToolTipText     =   "Ajoute une entrée au menu contextuel de Windows pour les fichiers"
            Top             =   120
            Width           =   6855
         End
         Begin VB.CheckBox chkSendTo 
            Caption         =   "Mettre une entrée dans le menu ""Envoyer vers"" de Windows"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            ToolTipText     =   "Ajoute une entrée 'Envoyer vers --> Hex Editor VB'"
            Top             =   1080
            Width           =   5175
         End
         Begin VB.CheckBox chkContextMenu 
            Caption         =   "Mettre une entrée au menu contextuel de Windows pour les dossiers"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   62
            ToolTipText     =   "Ajoute une entrée au menu contextuel de Windows pour les dossiers"
            Top             =   600
            Width           =   6855
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Index           =   3
      Left            =   240
      TabIndex        =   54
      Top             =   480
      Visible         =   0   'False
      Width           =   9615
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   120
         ScaleHeight     =   5175
         ScaleWidth      =   7455
         TabIndex        =   55
         Top             =   240
         Width           =   7455
         Begin VB.ComboBox cbOS 
            Height          =   315
            ItemData        =   "frmOptions.frx":062F
            Left            =   2040
            List            =   "frmOptions.frx":0639
            Style           =   2  'Dropdown List
            TabIndex        =   57
            ToolTipText     =   "Système d'exploitation utilisant le logiciel"
            Top             =   240
            Width           =   4215
         End
         Begin VB.ComboBox cbLang 
            Height          =   315
            ItemData        =   "frmOptions.frx":066C
            Left            =   2040
            List            =   "frmOptions.frx":0673
            Style           =   2  'Dropdown List
            TabIndex        =   56
            ToolTipText     =   "Langue par défaut"
            Top             =   1080
            Width           =   4215
         End
         Begin VB.Label Label7 
            Caption         =   "Système d'exploitation :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "Langue par défaut :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7575
      Index           =   2
      Left            =   0
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   9855
      Begin VB.PictureBox pctManifest 
         BorderStyle     =   0  'None
         Height          =   5895
         Index           =   1
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   9435
         TabIndex        =   38
         Top             =   240
         Width           =   9435
         Begin VB.CheckBox Check2 
            Caption         =   "Afficher la liste des icones par défaut"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "Affiche la liste des icones par défaut (fichier et processus)"
            Top             =   480
            Width           =   6615
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Afficher les données par défaut"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Afficher la zone de changement rapide de donnée lors de l'ouverture des fenêtres d'édition"
            Top             =   840
            Width           =   6615
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Afficher les informations fichier par défaut"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Affiche les informations sur les fichiers dans les fenêtres d'édition"
            Top             =   1200
            Width           =   6615
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Permettre plusieurs instances du programme"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            ToolTipText     =   "Permet au logiciel de se lancer plusieurs fois en même temps"
            Top             =   1560
            Width           =   6615
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Ne pas changer les dates des fichiers modifiés"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            ToolTipText     =   "Conserve les dates originelles du fichiers après sa modification et sa sauvegarde"
            Top             =   1920
            Width           =   6615
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   4680
            TabIndex        =   46
            Text            =   "640"
            ToolTipText     =   "Largeur"
            Top             =   4200
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   5520
            TabIndex        =   45
            Text            =   "480"
            ToolTipText     =   "Hauteur"
            Top             =   4200
            Width           =   495
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Ouvrir également les fichiers des sous-dossiers lors de l'ouverture d'un dossier"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            ToolTipText     =   "Liste et ouvre tous les fichiers des sous dossiers lors de l'ouverture d'un dossier (lent - déconseillé)"
            Top             =   2280
            Width           =   6615
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Fermer la fenêtre de démarrage après le choix d'un objet à ouvrir"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Referme la fenêtre de démarrage rapide après le choix d'une action"
            Top             =   2640
            Width           =   6615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Maximiser les fenêtres à leur ouverture"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Lance les fenêtres d'édition en grand lors de leur ouverture"
            Top             =   120
            Width           =   6615
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Afficher le splash screen"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Autorise l'affichage du splash screen au démarrage du logiciel"
            Top             =   3000
            Width           =   6615
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Ouvrir Hex Editor VB dans le même état qu'en partant"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   $"frmOptions.frx":067F
            Top             =   3360
            Width           =   4455
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Afficher les messages de confirmation"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Si cette case est cochée, les messages de confirmation seront affichés (recommandé)"
            Top             =   3720
            Width           =   4455
         End
         Begin VB.Label Label4 
            Caption         =   "Résolution de sauvegarde des images d'analyse des fichiers :"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   4200
            Width           =   4575
         End
         Begin VB.Label Label5 
            Caption         =   "X"
            Height          =   255
            Left            =   5280
            TabIndex        =   52
            Top             =   4200
            Width           =   135
         End
      End
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "Annuler"
      Height          =   495
      Left            =   3000
      TabIndex        =   36
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Par défaut"
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdSauvegarder 
      Caption         =   "OK"
      Height          =   495
      Left            =   1320
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Index           =   4
      Left            =   2280
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   7935
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   120
         ScaleHeight     =   5175
         ScaleWidth      =   7455
         TabIndex        =   24
         Top             =   240
         Width           =   7455
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   12
            Left            =   3360
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   28
            Top             =   3120
            Width           =   375
         End
         Begin VB.PictureBox pctColor 
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            Height          =   200
            Index           =   13
            Left            =   3360
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   27
            Top             =   3360
            Width           =   375
         End
         Begin VB.TextBox txtC 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2400
            TabIndex        =   26
            Text            =   "2000"
            ToolTipText     =   "Hauteur de la console"
            Top             =   3720
            Width           =   2655
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Afficher la console par défaut"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            ToolTipText     =   "Affiche la console par défaut"
            Top             =   4080
            Width           =   4575
         End
         Begin RichTextLib.RichTextBox txt 
            Height          =   2655
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   4683
            _Version        =   393217
            BackColor       =   0
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            OLEDragMode     =   0
            OLEDropMode     =   1
            TextRTF         =   $"frmOptions.frx":0711
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
         Begin VB.Label Label1 
            Caption         =   "Couleur de fond"
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   32
            Top             =   3120
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Couleur de la police"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   31
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Hauteur du composant :"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   30
            Top             =   3720
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Index           =   5
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   9855
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   120
         ScaleHeight     =   6135
         ScaleWidth      =   9615
         TabIndex        =   1
         Top             =   120
         Width           =   9615
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les paths entiers"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "Affiche le chemin du fichier avec le nom du fichier"
            Top             =   360
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les dossiers système"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "Autorise l'affichage des dossiers avec l'attribut 'système' dans l'explorateur de fichiers"
            Top             =   1440
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Autoriser la suppression de dossiers"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "Permet de supprimer des dossiers dans l'explorateur de fichiers"
            Top             =   1080
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les dossiers en lecture seule"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   15
            ToolTipText     =   "Autorise l'affichage des dossiers avec l'attribut 'lecture seule' dans l'explorateur de fichiers"
            Top             =   2160
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les dossiers cachés"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "Autorise l'affichage des dossiers avec l'attribut 'caché' dans l'explorateur de fichiers"
            Top             =   1800
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les fichiers cachés"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   13
            ToolTipText     =   "Autorise l'affichage des fichiers avec l'attribut 'caché' dans l'explorateur de fichiers"
            Top             =   2880
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les fichiers système"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   12
            ToolTipText     =   "Autorise l'affichage des fichiers avec l'attribut 'système' dans l'explorateur de fichiers"
            Top             =   2520
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Autoriser la sélection multiple"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "Autorise la sélection multiple dans l'exporateur de fichiers"
            Top             =   3600
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher les fichiers en lecture seule"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "Autorise l'affichage des fichiers avec l'attribut 'lecture seule' dans l'explorateur de fichiers"
            Top             =   3240
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Masquer les en-têtes des colonnes"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   9
            ToolTipText     =   "Masque les en têtes des colonnes (Taille, Nom, Date...) dans l'explorateur de fichiers"
            Top             =   3960
            Width           =   8655
         End
         Begin VB.ComboBox cbExpIcon 
            Height          =   315
            ItemData        =   "frmOptions.frx":0794
            Left            =   2520
            List            =   "frmOptions.frx":07A1
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "lang_ok"
            ToolTipText     =   "Type d'icones à afficher dans l'explorateur de fichiers"
            Top             =   4320
            Width           =   2535
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Autoriser la suppression de fichiers"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   7
            ToolTipText     =   "Permet de supprimer des fichiers dans l'explorateur de fichiers"
            Top             =   720
            Value           =   1  'Checked
            Width           =   6495
         End
         Begin VB.ComboBox cbExpInitDir 
            Height          =   315
            ItemData        =   "frmOptions.frx":07D8
            Left            =   2520
            List            =   "frmOptions.frx":07E2
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Tag             =   "lang_ok"
            ToolTipText     =   "Type de chemin par défaut de l'explorateur de fichiers"
            Top             =   4680
            Width           =   2535
         End
         Begin VB.TextBox txtExpPattern 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Text            =   "*.*"
            ToolTipText     =   "Filtre de l'explorateur de fichiers"
            Top             =   5520
            Width           =   2535
         End
         Begin VB.TextBox txtHeight 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2520
            TabIndex        =   4
            Text            =   "2200"
            ToolTipText     =   "Hauteur de l'explorateur de fichiers"
            Top             =   5160
            Width           =   2655
         End
         Begin VB.TextBox txtExpPath 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   5280
            TabIndex        =   3
            Text            =   "C:\"
            ToolTipText     =   $"frmOptions.frx":0806
            Top             =   4680
            Width           =   2535
         End
         Begin VB.CheckBox chkEx 
            Caption         =   "Afficher l'explorateur par défaut"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   2
            ToolTipText     =   "Affiche l'explorateur au chargement du logiciel"
            Top             =   5880
            Value           =   1  'Checked
            Width           =   8655
         End
         Begin VB.Label Label1 
            Caption         =   "Affichage des icones :"
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   22
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Path par défaut :"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   21
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Hauteur du composant :"
            Height          =   255
            Index           =   15
            Left            =   360
            TabIndex        =   20
            Top             =   5160
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Filtre :"
            Height          =   255
            Index           =   16
            Left            =   360
            TabIndex        =   19
            Top             =   5520
            Width           =   2055
         End
      End
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Tag             =   "lang_ok"
      Top             =   30
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   6
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Apparence du tableau"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Intégration dans Explorer"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Options générales"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Environnement"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Console"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Explorateur de fichiers"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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
'FORM QUI AFFICHE LES OPTIONS
'=======================================================

Private Lang As New clsLang

Private Sub cbExpInitDir_Click()
    txtExpPath.Enabled = (cbExpInitDir.ListIndex = 1)
End Sub

Private Sub cbGrid_Click()
'applique la grille souhaitée ou HW
                    
    HW.Grid = cbGrid.ListIndex
   
End Sub

Private Sub chkContextMenu_Click(Index As Integer)
'ajoute (ou pas) le context menu Index=0 ==> fichier, Index=1 ==> dossier
    
    If chkContextMenu(Index).Value Then
        'ajoute
        Call AddContextMenu(Index + 1)
    Else
        'retire
        Call RemoveContextMenu(Index + 1)
    End If
    
End Sub

Private Sub chkSendTo_Click()
'créé ou supprime l'entrée Send To...
    
    If chkSendTo.Value Then
        'créé
        Call Shortcut(True)
    Else
        'supprime
        Call Shortcut(False)
    End If
    
End Sub

Private Sub cmdDefault_Click()
'remet tout par défaut
Dim X As Long
Dim Y As Long
Dim s As String

    With HW
        .BackColor = vbWhite
        .OffsetForeColor = 16737380
        .HexForeColor = &H6F6F6F
        .StringForeColor = &H6F6F6F
        .TitleBackGround = &H8000000F
        .OffsetTitleForeColor = 16737380
        .BaseTitleForeColor = 16737380
        .LineColor = &H8000000C
        .FirstOffset = 0
        .NumberPerPage = 20
        .SelectionColor = &HE0E0E0
        .Grid = None
        .SignetColor = &H8080FF
        .ModifiedItemColor = &HFF&
        .ModifiedSelectedItemColor = &HFF&
        .Refresh
        .UseHexOffset = True
        pctColor(0).BackColor = .BackColor
        pctColor(1).BackColor = .OffsetForeColor
        pctColor(2).BackColor = .HexForeColor
        pctColor(3).BackColor = .StringForeColor
        pctColor(4).BackColor = .OffsetTitleForeColor
        pctColor(5).BackColor = .BaseTitleForeColor
        pctColor(6).BackColor = .BackColor
        pctColor(7).BackColor = .LineColor
        pctColor(8).BackColor = .SelectionColor
        pctColor(9).BackColor = .ModifiedItemColor
        pctColor(10).BackColor = .ModifiedSelectedItemColor
        pctColor(11).BackColor = .SignetColor
        pctColor(12).BackColor = vbBlack
        pctColor(13).BackColor = 12632256
    End With
    
    txtC.Text = "2000"
    optHex.Value = True
        
    'affiche un exemple de valeurs Offset, String et Hexa dans le HW
    HW.NumberPerPage = 13
    Randomize
    For X = 1 To 13
        s = vbNullString
        For Y = 1 To 16
            HW.AddHexValue X, Y, Hex$(Y - 1) & "0"
            s = s & Byte2FormatedString(Int(Rnd * 256))
        Next Y
        HW.AddStringValue X, s
    Next X
    
    With HW
        .FillText
        .Refresh
    End With
    
    With txt
        .BackColor = pctColor(12).BackColor
        .SelStart = 0
        .Text = Lang.GetString("_ConsoleExample") & vbNewLine & vbNewLine & Lang.GetString("_SecondLine")
        .SelLength = Len(.Text)
        .SelColor = pctColor(13).BackColor
        .SelStart = Len(.Text)
    End With
    
    cbGrid.ListIndex = HW.Grid
        
    Check1.Value = 1
    Check2.Value = 1
    Check3.Value = 1
    Check4.Value = 1
    Check5.Value = 0
    Check6.Value = 1
    Check7.Value = 0
    Check8.Value = 0
    Check9.Value = 1
    Check10.Value = 1
    Check11.Value = 1
    Check12.Value = 0
    Text3.Text = 640
    Text4.Text = 480

    chkContextMenu(0).Value = 1
    chkContextMenu(1).Value = 1
    chkSendTo.Value = 1
        
    cbOS.ListIndex = 1  '==> A CHANGER
        
    chkEx(0).Value = 0
    chkEx(1).Value = 1
    chkEx(2).Value = 1
    chkEx(3).Value = 0
    chkEx(4).Value = 1
    chkEx(5).Value = 1
    chkEx(6).Value = 1
    chkEx(7).Value = 1
    chkEx(8).Value = 1
    chkEx(9).Value = 1
    chkEx(10).Value = 1
    chkEx(11).Value = 0
    txtHeight.Text = 2200
    txtExpPattern.Text = "*.*"
    
    cbExpIcon.ListIndex = 1
    
    txtExpPath.Enabled = False
    cbExpInitDir.ListIndex = 0
    
    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_ByDefaut"))
End Sub

Private Sub cmdQuitter_Click()
    Unload Me
End Sub

Private Sub cmdSauvegarder_Click()
Dim X As Form
Dim s As String

    'sauvegarde les options

    'affecte à cPref toutes ses valeurs
        With cPref
            .app_BackGroundColor = pctColor(0).BackColor
            .app_OffsetForeColor = pctColor(1).BackColor
            .app_HexaForeColor = pctColor(2).BackColor
            .app_StringsForeColor = pctColor(3).BackColor
            .app_OffsetTitleForeColor = pctColor(4).BackColor
            .app_BaseForeColor = pctColor(5).BackColor
            .app_TitleBackGroundColor = pctColor(6).BackColor
            .app_LinesColor = pctColor(7).BackColor
            .app_SelectionColor = pctColor(8).BackColor
            .app_ModifiedItems = pctColor(9).BackColor
            .app_ModifiedSelectedItems = pctColor(10).BackColor
            .app_BookMarkColor = pctColor(11).BackColor
            .app_OffsetsHex = Abs(CLng(optHex.Value))
            
            .app_Grid = cbGrid.ListIndex
            
            .general_MaximizeWhenOpen = Check1.Value
            .general_DisplayExplore = chkEx(10).Value
            .general_DisplayIcon = Check2.Value
            .general_DisplayData = Check3.Value
            .general_DisplayInfos = Check4.Value
            .general_ShowAlert = Check11.Value
            .general_AllowMultipleInstances = Check5.Value
            .general_DoNotChangeDates = Check6.Value
            .general_OpenSubFiles = Check7.Value
            .general_CloseHomeWhenChosen = Check8.Value
            .general_ResoX = Text3.Text
            .general_ResoY = Text4.Text
            .general_Splash = Check9.Value
            .general_QuickBackup = Check10.Value
            
            .integ_FileContextual = chkContextMenu(0).Value
            .integ_FolderContextual = chkContextMenu(1).Value
            .integ_SendTo = chkSendTo.Value
            
            
            .env_OS = cbOS.ListIndex
            
            .explo_ShowPath = chkEx(0).Value
            .explo_AllowFileSuppression = chkEx(1).Value
            .explo_ShowSystemFodlers = chkEx(2).Value
            .explo_AllowFolderSuppression = chkEx(3).Value
            .explo_ShowROFolders = chkEx(4).Value
            .explo_ShowHiddenFolders = chkEx(5).Value
            .explo_ShowHiddenFiles = chkEx(6).Value
            .explo_ShowSystemFiles = chkEx(7).Value
            .explo_AllowMultipleSelection = chkEx(8).Value
            .explo_ShowROFiles = chkEx(9).Value
            .explo_HideColumnTitle = chkEx(11).Value
            .explo_Height = txtHeight.Text
            .explo_Pattern = txtExpPattern.Text
            
            .explo_IconType = cbExpIcon.ListIndex
            .explo_DefaultPath = IIf(txtExpPath.Enabled, txtExpPath.Text, Lang.GetString("_ProgPath"))
            
            .console_BackColor = pctColor(12).BackColor
            .console_ForeColor = pctColor(13).BackColor
            .console_Heigth = Val(txtC.Text)
            .console_Load = Check12.Value
            
        End With
    
    'lance la sauvegarde
    Call clsPref.SaveIniFile(cPref)
    
    
    'On Error Resume Next
    'on change l'apparence de tous les HW de toutes les forms
    For Each X In Forms
        If (TypeOf X Is Pfm) Or (TypeOf X Is diskPfm) Or (TypeOf X Is MemPfm) _
            Or (TypeOf X Is physPfm) Then

                With X.HW
                    'on applique ces couleurs au HW de CETTE form
                    .BackColor = cPref.app_BackGroundColor
                    .OffsetForeColor = cPref.app_OffsetForeColor
                    .HexForeColor = cPref.app_HexaForeColor
                    .StringForeColor = cPref.app_StringsForeColor
                    .OffsetTitleForeColor = cPref.app_OffsetTitleForeColor
                    .BaseTitleForeColor = cPref.app_BaseForeColor
                    .TitleBackGround = cPref.app_TitleBackGroundColor
                    .LineColor = cPref.app_LinesColor
                    .SelectionColor = cPref.app_SelectionColor
                    .ModifiedItemColor = cPref.app_ModifiedItems
                    .ModifiedSelectedItemColor = cPref.app_ModifiedSelectedItems
                    .SignetColor = cPref.app_BookMarkColor
                    .Grid = cPref.app_Grid
                    .UseHexOffset = CBool(cPref.app_OffsetsHex)
                    .Refresh
                End With
                
                'change les Visible des frames de toutes les forms active
                X.FrameData.Visible = CBool(cPref.general_DisplayData)
                X.FrameInfos.Visible = CBool(cPref.general_DisplayInfos)
                If (TypeOf X Is diskPfm) Or (TypeOf X Is physPfm) Then X.FrameInfo2.Visible = CBool(cPref.general_DisplayInfos)
            'End If
        End If
    Next X
              
    On Error Resume Next
    
    'on change la taille du Explorer
    With frmContent
        .pctExplorer.Height = cPref.explo_Height
        .LV.Height = cPref.explo_Height - 145
        
        'apparence de la console
        .pctConsole.BackColor = cPref.console_BackColor
        .txt.BackColor = cPref.console_BackColor
        .txtE.BackColor = cPref.console_BackColor
        .pctConsole.Height = cPref.console_Heigth
        .txt.BackColor = pctColor(12).BackColor
        .txt.SelStart = 0
        .txt.SelLength = Len(.txt.Text)
        .txt.SelColor = pctColor(13).BackColor
        .txt.SelStart = Len(.txt.Text)
        .txtE.BackColor = pctColor(12).BackColor
        .txtE.SelStart = 0
        .txtE.SelLength = Len(.txtE.Text)
        .txtE.SelColor = pctColor(13).BackColor
        .txtE.SelStart = Len(.txtE.Text)
    End With

    'créé ou supprime les menus contextuels de Windows en fonction des nouvelles prefs.
    If CBool(cPref.integ_FileContextual) = False Then
        'enlève
        Call RemoveContextMenu(1)
    Else
        'ajoute
        Call AddContextMenu(1)
    End If
    If CBool(cPref.integ_FolderContextual) = False Then
        'enlève
        Call RemoveContextMenu(0)
    Else
        'ajoute
        Call AddContextMenu(0)
    End If
    
    'créé ou pas le raccourci
    Call Shortcut(CBool(cPref.integ_SendTo))


    'change les settings du Explorer
    With frmContent.LV
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
    End With
    
    Call frmContent.MDIForm_Resize

    'ajoute du texte à la console
    Call AddTextToConsole(Lang.GetString("_OptSaved"))
    
    Unload Me
    
End Sub

Private Sub Form_Load()
Dim X As Long
Dim Y As Long
Dim s As String

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
    
    TB.ZOrder vbSendToBack  'dernier plan
    
    'remet/redimensionne les frames à leur place et redimensionne la form
    For X = 0 To Frame1.Count - 1
        Frame1(X).Top = 430
        Frame1(X).Width = 9855
        Frame1(X).Height = 6375
        Frame1(X).Left = 50
    Next X
    
    With Me
        Me.Width = 10065
        .Height = 7900
        .cmdDefault.Left = 1000
        .cmdQuitter.Left = 7800
        .cmdSauvegarder.Left = 6300
        .cmdDefault.Top = 6900
        .cmdQuitter.Top = 6900
        .cmdSauvegarder.Top = 6900
    End With
    
    
    '//LECTURE DES PREFERENCES
        
        
        
        '//APPARENCE DU HW
        With cPref
            pctColor(0).BackColor = .app_BackGroundColor
            pctColor(1).BackColor = .app_OffsetForeColor
            pctColor(2).BackColor = .app_HexaForeColor
            pctColor(3).BackColor = .app_StringsForeColor
            pctColor(4).BackColor = .app_OffsetTitleForeColor
            pctColor(5).BackColor = .app_BaseForeColor
            pctColor(6).BackColor = .app_TitleBackGroundColor
            pctColor(7).BackColor = .app_LinesColor
            pctColor(8).BackColor = .app_SelectionColor
            pctColor(9).BackColor = .app_ModifiedItems
            pctColor(10).BackColor = .app_ModifiedSelectedItems
            pctColor(11).BackColor = .app_BookMarkColor
            optHex.Value = CBool(.app_OffsetsHex)
            optDec.Value = Not (optHex.Value)
        End With
        
        With HW
            'on applique ces couleurs au HW de CETTE form
            .BackColor = pctColor(0).BackColor
            .OffsetForeColor = pctColor(1).BackColor
            .HexForeColor = pctColor(2).BackColor
            .StringForeColor = pctColor(3).BackColor
            .OffsetTitleForeColor = pctColor(4).BackColor
            .BaseTitleForeColor = pctColor(5).BackColor
            .TitleBackGround = pctColor(6).BackColor
            .LineColor = pctColor(7).BackColor
            .SelectionColor = pctColor(8).BackColor
            .ModifiedItemColor = pctColor(9).BackColor
            .ModifiedSelectedItemColor = pctColor(10).BackColor
            .SignetColor = pctColor(11).BackColor
            .UseHexOffset = optHex.Value
        End With
        
        'affiche un exemple de valeurs Offset, String et Hexa dans le HW
        HW.NumberPerPage = 13
        Randomize
        For X = 1 To 13
            s = vbNullString
            For Y = 1 To 16
                HW.AddHexValue X, Y, Hex$(Y - 1) & "0"
                s = s & Byte2FormatedString(Int(Rnd * 256))
            Next Y
            HW.AddStringValue X, s
        Next X
        
        'affiche la bonne valeur de grille dans le combobox
        Select Case cPref.app_Grid
            Case 0
                HW.Grid = 0
            Case Horizontal
                HW.Grid = Horizontal
            Case HorizontalHexOnly
                HW.Grid = HorizontalHexOnly
            Case VerticalHex
                HW.Grid = VerticalHex
            Case HorizontalHexOnly_VerticalHex
                HW.Grid = HorizontalHexOnly_VerticalHex
            Case Horizontal_VerticalHex
                HW.Grid = Horizontal_VerticalHex
        End Select
        cbGrid.ListIndex = cPref.app_Grid
        
        HW.FillText
        HW.Refresh
        
        
        '//APPARENCE DE LA CONSOLE
        With cPref
            pctColor(12).BackColor = .console_BackColor
            pctColor(13).BackColor = .console_ForeColor
            txtC.Text = Trim$(Str$(frmContent.pctConsole.Height))
        End With
        With txt
            .BackColor = cPref.console_BackColor
            .SelStart = 0
            .Text = Lang.GetString("_ConsoleExample") & vbNewLine & vbNewLine & Lang.GetString("_SecondLine")
            .SelLength = Len(.Text)
            .SelColor = cPref.console_ForeColor
            .SelStart = Len(.Text)
        End With
        
        
        '//GENERAL
        With cPref
            Check1.Value = .general_MaximizeWhenOpen
            Check2.Value = .general_DisplayIcon
            Check3.Value = .general_DisplayData
            Check4.Value = .general_DisplayInfos
            Check5.Value = .general_AllowMultipleInstances
            Check6.Value = .general_DoNotChangeDates
            Check7.Value = .general_OpenSubFiles
            Check8.Value = .general_CloseHomeWhenChosen
            Check9.Value = .general_Splash
            Check10.Value = .general_QuickBackup
            Check11.Value = .general_ShowAlert
            Check12.Value = .console_Load
            Text3.Text = .general_ResoX
            Text4.Text = .general_ResoY
        End With
        
        
        '//INTEGRATION
        With cPref
            chkContextMenu(0).Value = .integ_FileContextual
            chkContextMenu(1).Value = .integ_FolderContextual
            chkSendTo.Value = .integ_SendTo
        End With
        
        
        '//ENVIRONNEMENT
        With cPref
            If InStr(1, .env_OS, "Vista") Then
                'alors c'est vista
                cbOS.ListIndex = 1
            Else
                cbOS.ListIndex = 0
            End If
            
            'LANGUE ==> ??
        End With
        
        
        '//EXPLORATEUR
        With cPref
            chkEx(0).Value = .explo_ShowPath
            chkEx(1).Value = .explo_AllowFileSuppression
            chkEx(2).Value = .explo_ShowSystemFodlers
            chkEx(3).Value = .explo_AllowFolderSuppression
            chkEx(4).Value = .explo_ShowROFolders
            chkEx(5).Value = .explo_ShowHiddenFolders
            chkEx(6).Value = .explo_ShowHiddenFiles
            chkEx(7).Value = .explo_ShowSystemFiles
            chkEx(8).Value = .explo_AllowMultipleSelection
            chkEx(9).Value = .explo_ShowROFiles
            chkEx(10).Value = .general_DisplayExplore
            chkEx(11).Value = .explo_HideColumnTitle
            txtHeight.Text = .explo_Height
            txtExpPattern.Text = .explo_Pattern
            
            cbExpIcon.ListIndex = .explo_IconType
            
            If .explo_DefaultPath = Lang.GetString("_ProgPath") Then
                'alors c'est le dossier du programme
                txtExpPath.Enabled = False
                cbExpInitDir.ListIndex = 0
            Else
                txtExpPath.Enabled = True
                txtExpPath.Text = .explo_DefaultPath
                cbExpInitDir.ListIndex = 1
            End If
        End With
        
End Sub

Private Sub HW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, Item As HexViewer_OCX.ItemElement)
 
    If Button = 4 And Shift = 0 Then
        'click avec la molette, et pas de Shift or Control
        'on ajoute (ou enlève) un signet
    
        If HW.IsSignet(Item.Offset) = False Then
            'on l'ajoute
            Call HW.AddSignet(Item.Offset)
            Call HW.TraceSignets
        ElseIf HW.IsSignet(Item.Offset) Then
        
            'alors on l'enlève
            While HW.IsSignet(HW.Item.Offset)
                'on supprime
                Call HW.RemoveSignet(Val(HW.Item.Offset))
            Wend
        End If
    End If
End Sub

Private Sub pctColor_Click(Index As Integer)
'alors ouvre un CMD pour pouvoir choisir une couleur

    On Error GoTo ErrGestion
    
    'affiche la couleur dans la picturebox
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = Lang.GetString("_ColorChoice")
        .ShowColor
        pctColor(Index).BackColor = .Color
    End With
    
    If Index < 12 Then
        With HW
            'maintenant, change la couleur dans le HW
            .BackColor = pctColor(0).BackColor
            .OffsetForeColor = pctColor(1).BackColor
            .HexForeColor = pctColor(2).BackColor
            .StringForeColor = pctColor(3).BackColor
            .OffsetTitleForeColor = pctColor(4).BackColor
            .BaseTitleForeColor = pctColor(5).BackColor
            .TitleBackGround = pctColor(6).BackColor
            .LineColor = pctColor(7).BackColor
            .SelectionColor = pctColor(8).BackColor
            '.ModifiedItemColor = pctColor(9).BackColor
            '.ModifiedSelectedItemColor = pctColor(10).BackColor
        End With
    Else
        With txt
            .BackColor = pctColor(12).BackColor
            .SelStart = 0
            .Text = Lang.GetString("_ConsoleExample") & vbNewLine & vbNewLine & Lang.GetString("_SecondLine")
            .SelLength = Len(.Text)
            .SelColor = pctColor(13).BackColor
            .SelStart = Len(.Text)
        End With
    End If
    
ErrGestion:
End Sub

Private Sub TB_Click()
'change le frame Visible
Dim X As Long

    'rend invisible tout les frames
    For X = 0 To Frame1.Count - 1
        Frame1(X).Visible = False
    Next X
    
    'affiche le bon en fonction du tab
    Frame1(TB.SelectedItem.Index - 1).Visible = True
    
End Sub
