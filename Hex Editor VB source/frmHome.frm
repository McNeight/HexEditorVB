VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{9B9A881F-DBDC-4334-BC23-5679E5AB0DC6}#1.1#0"; "FileView_OCX.ocx"
Begin VB.Form frmHome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu principal"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   4
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   36
         Top             =   240
         Width           =   6495
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2160
            TabIndex        =   41
            Tag             =   "pref"
            Text            =   "100"
            ToolTipText     =   "Taille"
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cdUnit 
            Height          =   315
            ItemData        =   "frmHome.frx":08CA
            Left            =   3360
            List            =   "frmHome.frx":08DA
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Tag             =   "pref"
            ToolTipText     =   "Unité"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtNewFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Nouveau fichier à créer"
            Top             =   480
            Width           =   5415
         End
         Begin VB.CommandButton cmdBrowseNew 
            Caption         =   "..."
            Height          =   255
            Left            =   5640
            TabIndex        =   38
            ToolTipText     =   "Choix du fichier à créer"
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Taille du fichier"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Création d'un nouveau fichier"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Annuler"
      Height          =   435
      Left            =   4125
      TabIndex        =   28
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ouvrir"
      Height          =   435
      Left            =   1125
      TabIndex        =   27
      Top             =   5160
      Width           =   1815
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir fichier"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un fichier"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir dossier"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un dossier de fichiers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir disque"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un disque"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ouvrir processus"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ouvrir un processus"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Nouveau fichier"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   3
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   8
         Top             =   240
         Width           =   6495
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   3615
            Index           =   1
            Left            =   2880
            TabIndex        =   23
            Top             =   360
            Width           =   3495
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   3255
               Index           =   1
               Left            =   120
               ScaleHeight     =   3255
               ScaleWidth      =   3255
               TabIndex        =   24
               Top             =   240
               Width           =   3255
               Begin VB.TextBox txtProcessInfos 
                  BorderStyle     =   0  'None
                  Height          =   3135
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   34
                  Top             =   0
                  Width           =   3255
               End
            End
         End
         Begin ComctlLib.ListView LV 
            Height          =   3495
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   6165
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
               Text            =   "PID"
               Object.Width           =   706
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Processus"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du processus à ouvrir :"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   1
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   4
         Top             =   240
         Width           =   6495
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   2295
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   6255
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1935
               Index           =   3
               Left            =   120
               ScaleHeight     =   1935
               ScaleWidth      =   6015
               TabIndex        =   30
               Top             =   240
               Width           =   6015
               Begin VB.TextBox txtFolderInfos 
                  BorderStyle     =   0  'None
                  Height          =   1815
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   31
                  Top             =   120
                  Width           =   6015
               End
            End
         End
         Begin VB.OptionButton optFolderSub 
            Caption         =   "N'ouvrir que les fichiers contenus directement dans le dossier"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Tag             =   "pref"
            ToolTipText     =   "Ne sélectionne que les fichiers qui sont dans la racine du dossier"
            Top             =   1320
            Value           =   -1  'True
            Width           =   5775
         End
         Begin VB.OptionButton optFolderSub 
            Caption         =   "Ouvrir également les fichiers des sous-dossiers"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Tag             =   "pref"
            ToolTipText     =   "Ouvre également tous les fichiers des sous dossiers (lent)"
            Top             =   960
            Width           =   5775
         End
         Begin VB.CommandButton cmdBrowseFolder 
            Caption         =   "..."
            Height          =   255
            Left            =   5640
            TabIndex        =   14
            ToolTipText     =   "Choix du dossier à ouvrir"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFolder 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Dossier choisi pour l'ouverture"
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du dossier à ouvrir :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   2
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   6
         Top             =   240
         Width           =   6495
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   3735
            Index           =   0
            Left            =   2640
            TabIndex        =   19
            Top             =   360
            Width           =   3735
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   3375
               Index           =   0
               Left            =   120
               ScaleHeight     =   3375
               ScaleWidth      =   3495
               TabIndex        =   20
               Top             =   240
               Width           =   3495
               Begin VB.TextBox txtDiskInfos 
                  BorderStyle     =   0  'None
                  Height          =   3375
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   33
                  Top             =   0
                  Width           =   3495
               End
            End
         End
         Begin FileView_OCX.FileView FV 
            Height          =   3615
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   6376
            AllowDirectoryEntering=   0   'False
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
            Caption         =   "Choix du disque à ouvrir :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   0
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   2
         Top             =   240
         Width           =   6495
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   3135
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   6255
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   2775
               Index           =   2
               Left            =   120
               ScaleHeight     =   2775
               ScaleWidth      =   6015
               TabIndex        =   26
               Top             =   240
               Width           =   6015
               Begin VB.TextBox txtFileInfos 
                  BorderStyle     =   0  'None
                  Height          =   2775
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   32
                  Top             =   0
                  Width           =   6015
               End
            End
         End
         Begin VB.CommandButton cmdBrowseFile 
            Caption         =   "..."
            Height          =   255
            Left            =   5640
            TabIndex        =   11
            ToolTipText     =   "Choix du fichier à ouvrir"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Fichier choisi pour l'ouverture"
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du fichier à ouvrir :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   4095
         End
      End
   End
End
Attribute VB_Name = "frmHome"
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
'FORM DE DEMARRAGE (CHOIX DES ACTIONS A EFFECTUER)
'-------------------------------------------------------

Private clsPref As clsIniForm

Private Sub cmdBrowseFile_Click()
'browse un fichier
    txtFile.Text = cFile.ShowOpen("Sélectionner le fichier à ouvrir", Me.hwnd, "Tous|*.*")
End Sub

Private Sub cmdBrowseFolder_Click()
'browse un dossier
    txtFolder.Text = cFile.BrowseForFolder("Sélectionner le dossier à ouvrir", Me.hwnd)
End Sub

Private Sub cmdBrowseNew_Click()
'browse un dossier
    txtNewFile.Text = cFile.ShowSave("Sélectionner le fichier à créer", Me.hwnd, "Tous|*.*", App.Path)
End Sub

Private Sub cmdOk_Click()
'ouvre l'élément sélectionné
Dim m() As String
Dim Frm As Form
Dim x As Long
Dim sDrive As String
Dim cDr As clsDiskInfos
Dim lH As Long
Dim lFile As Long
Dim lLen As Double

    On Error GoTo ErrGestion

    Select Case TB.SelectedItem.Index
        Case 1
            'fichier
            If cFile.FileExists(txtFile.Text) = False Then Exit Sub
    
            'affiche une nouvelle fenêtre
            Set Frm = New Pfm
            Call Frm.GetFile(txtFile.Text)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
            
        Case 2
            'dossier
            
            'teste la validité du répertoire
            If cFile.FolderExists(txtFolder.Text) = False Then Exit Sub
            
            'liste les fichiers
            If cFile.GetFolderFiles(txtFolder.Text, m, optFolderSub(0).Value) < 1 Then Exit Sub
            
            'les ouvre un par un
            For x = 1 To UBound(m)
                If cFile.FileExists(m(x)) Then
                    Set Frm = New Pfm
                    Call Frm.GetFile(m(x))
                    Frm.Show
                    lNbChildFrm = lNbChildFrm + 1
                    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
                    DoEvents
                End If
            Next x
    
        Case 3
            'disque
            'vérifie que le drive est accessible
            Set cDr = New clsDiskInfos
            If cDr.IsLogicalDriveAccessible(FV.ListItems.Item(FV.ListIndex).Text) = False Then
                Set cDr = Nothing   'inaccessible, alors on sort de cette procédure
                Exit Sub
            End If
            Set cDr = Nothing
            
            'affiche une nouvelle fenêtre
            Set Frm = New diskPfm
            
            Call Frm.GetDrive(FV.ListItems.Item(FV.ListIndex).Text) 'renseigne sur le path sélectionné
            
            Frm.Show    'affiche la nouvelle
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
            
        Case 4
            'processus

            'vérfie que le processus est ouvrable
            lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(LV.SelectedItem.Text))
            CloseHandle lH
            
            If lH = 0 Then
                'pas possible
                MsgBox "Accès impossible à ce processus", vbInformation, "Erreur"
                Exit Sub
            End If
            
            'possible affiche une nouvelle fenêtre
            Set Frm = New MemPfm
            Call Frm.GetFile(Val(LV.SelectedItem.Text), LV.SelectedItem.SubItems(1))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
            
        Case 5
            
            'vérifie que le fichier n'existe pas
            If cFile.FileExists(txtNewFile.Text) Then
                'fichier déjà existant
                If MsgBox("Le fichier sélectionné existe déjà. Le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
            End If
            cFile.CreateEmptyFile txtNewFile.Text, True
            
            'création du fichier
            lLen = Abs(Val(txtSize.Text))
            If cdUnit.Text = "Ko" Then lLen = lLen * 1024
            If cdUnit.Text = "Mo" Then lLen = (lLen * 1024) * 1024
            If cdUnit.Text = "Go" Then lLen = ((lLen * 1024) * 1024) * 1024
            
            'obtient un numéro de fichier disponible
            lFile = FreeFile
            
            Open txtNewFile.Text For Binary Access Write As lFile
                Put lFile, , String$(lLen, vbNullChar)
            Close lFile
            
            'affiche une nouvelle fenêtre
            Set Frm = New Pfm
            Call Frm.GetFile(txtNewFile.Text)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"
        
    End Select
    
    If cPref.general_CloseHomeWhenChosen Then
        'alors on ferme cette form car on a ouvert un fichier
        Unload Me
    End If
    
    Exit Sub
    
ErrGestion:
    clsERREUR.AddError "frmHome.cmdOk_Click", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    clsPref.SaveFormSettings App.Path & "\Preferences\HomeWindow.ini", Me
    Set clsPref = Nothing
End Sub

Private Sub cmdQuit_Click()
    'annule
    Unload Me
End Sub

'-------------------------------------------------------
'FORM HOME ==> CHOIX DE L'OBJET A OUVRIR
'-------------------------------------------------------
Private Sub Form_Load()
Dim x As Long
Dim p() As ProcessItem
Dim clsProc As clsProcess   'appel à une classe de gestion de processus

    'loading des preferences
    Set clsPref = New clsIniForm
    clsPref.GetFormSettings App.Path & "\Preferences\HomeWindow.ini", Me
    optFolderSub(1).Value = Not (optFolderSub(0).Value)
    
    'réorganise les Frames
    For x = 0 To Frame1.Count - 1
        Frame1(x).Left = 120
        Frame1(x).Top = 480
        Frame1(x).Width = 6600
        Frame1(x).Height = 4500
    Next x
    
    'prépare le FileView
    With FV
        .Path = Left$(App.Path, 3)  'affiche un drive existant (nécessaire à AddDrives)
        .ShowFiles = False
        .ShowDirectories = False
        .ShowDrives = True 'affiche QUE les drives
        .AllowDirectoryEntering = False 'ne RENTRE PAS dans les dossiers (drives)
        .View = lvwList 'pas de columns
        .AddDrives    'affiche les drives
    End With
    
    'fait la liste des processus en mémoire
    Set clsProc = New clsProcess

    LV.ListItems.Clear

    'énumération
    clsProc.EnumerateProcesses p()
    
    'affiche la liste
    For x = 0 To UBound(p) - 1
        LV.ListItems.Add Text:=p(x).th32ProcessID
        LV.ListItems.Item(x + 1).SubItems(1) = p(x).szExeFile
    Next x
    
    'affiche un seul frame
    MaskFrames 0
End Sub

'-------------------------------------------------------
'masque tous les frames sauf un
'-------------------------------------------------------
Private Sub MaskFrames(ByVal lFrame As Long)
Dim x As Long

    For x = 0 To Frame1.Count - 1
        Frame1(x).Visible = False
    Next x
    Frame1(lFrame).Visible = True
    
    If lFrame = 4 Then
        'création de fichier
        cmdOk.Caption = "Créer et ouvrir"
    Else
        cmdOk.Caption = "Ouvrir"
    End If

End Sub

Private Sub FV_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim cDrive As clsDrive
Dim s As String
    
    'vérifie la disponibilité du disque
    If cDisk.IsLogicalDriveAccessible(Item.Text) = False Then
        'inaccessible
        txtDiskInfos.Text = "Disque inaccessible"
        Exit Sub
    End If
    
    'affichage des infos disque
    Set cDrive = cDisk.GetLogicalDrive(Item.Text)
    
    With cDrive
        s = "Lecteur=[" & .VolumeLetter & "]"
        s = s & vbNewLine & "Nom du volume=[" & CStr(.VolumeName) & "]"
        s = s & vbNewLine & "Numéro de série=[" & Hex$(.VolumeSerialNumber) & "]"
        s = s & vbNewLine & "Système de fichier=[" & CStr(.FileSystemName) & "]"
        s = s & vbNewLine & "Type de lecteur=[" & CStr(.strDriveType) & "]"
        s = s & vbNewLine & "Type de média=[" & CStr(.strMediaType) & "]"
        s = s & vbNewLine & "Taille de la partition=[" & CStr(.PartitionLength) & "]"
        s = s & vbNewLine & "Taille totale=[" & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
        s = s & vbNewLine & "Taille libre=[" & CStr(.FreeSpace) & "  <--> " & FormatedSize(.FreeSpace, 10) & " ]"
        s = s & vbNewLine & "Taille utilisée=[" & CStr(.UsedSpace) & "  <--> " & FormatedSize(.UsedSpace, 10) & " ]"
        s = s & vbNewLine & "Pourcentage de taille libre=[" & CStr(.PercentageFree) & " %]"
        s = s & vbNewLine & "Nombre de secteurs logiques=[" & CStr(.TotalLogicalSectors) & "]"
        s = s & vbNewLine & "Nombre de secteurs physiques=[" & CStr(.TotalPhysicalSectors) & "]"
        s = s & vbNewLine & "Nombre de secteurs cachés=[" & CStr(.HiddenSectors) & "]"
        s = s & vbNewLine & "Octets par secteur=[" & CStr(.BytesPerSector) & "]"
        s = s & vbNewLine & "Secteurs par cluster=[" & CStr(.SectorPerCluster) & "]"
        s = s & vbNewLine & "Nombre de clusters=[" & CStr(.TotalClusters) & "]"
        s = s & vbNewLine & "Clusters libres=[" & CStr(.FreeClusters) & "]"
        s = s & vbNewLine & "Clusters utilisés=[" & CStr(.UsedClusters) & "]"
        s = s & vbNewLine & "Octets par cluster=[" & CStr(.BytesPerCluster) & "]"
        s = s & vbNewLine & "Nombre de cylindres=[" & CStr(.Cylinders) & "]"
        s = s & vbNewLine & "Pistes par cylindre=[" & CStr(.TracksPerCylinder) & "]"
        s = s & vbNewLine & "Secteurs par piste=[" & CStr(.SectorsPerTrack) & "]"
        s = s & vbNewLine & "Offset de départ=[" & CStr(.StartingOffset) & "]"
    End With
    
    txtDiskInfos.Text = s
    
    'libère la mémoire
    Set cDrive = Nothing
End Sub

Private Sub FV_ItemDblSelection(Item As ComctlLib.ListItem)
    cmdOk_Click
End Sub

Private Sub LV_DblClick()
    If LV.SelectedItem Is Nothing Then Exit Sub
    cmdOk_Click
End Sub

Private Sub LV_ItemClick(ByVal Item As ComctlLib.ListItem)
'met à jour les infos
Dim pProcess As ProcessItem
Dim cFic As clsFile
Dim s As String

    On Error Resume Next
    
    'vérifie l'existence du processus
    If cProc.DoesPIDExist(Val(Item.Text)) = False Then
        'existe pas
        txtProcessInfos.Text = "Processus inaccessible"
        Exit Sub
    End If
    
    'récupère les infos sur le processus
    Set pProcess = cProc.GetProcess(Val(Item.Text), False, False, True)
    Set cFic = cFile.GetFile(pProcess.szImagePath)
    
    'infos fichier cible
    With cFic
        s = "-------------------------------------------"
        s = s & vbNewLine & "-------------- Fichier cible  -------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "Fichier=[" & .File & "]"
        s = s & vbNewLine & "Taille=[" & CStr(.FileSize) & " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        s = s & vbNewLine & "Attribut=[" & CStr(.FileAttributes) & "]"
        s = s & vbNewLine & "Date de création=[" & .CreationDate & "]"
        s = s & vbNewLine & "Date de dernier accès=[" & .LastAccessDate & "]"
        s = s & vbNewLine & "Date de dernière modification=[" & .LastModificationDate & "]"
        s = s & vbNewLine & "Version=[" & .EXEFileVersion & "]"
        s = s & vbNewLine & "Description=[" & .EXEFileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .EXELegalCopyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .EXECompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .EXEInternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .EXEOriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .EXEProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .EXEProductVersion & "]"
        s = s & vbNewLine & "Taille compressée=[" & .FileCompressedSize & "]"
        s = s & vbNewLine & "Programme associé=[" & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & "Répertoire contenant=[" & .FileDirectory & "]"
        s = s & vbNewLine & "Lecteur contenant=[" & .FileDrive & "]"
        s = s & vbNewLine & "Type de fichier=[" & .FileType & "]"
        s = s & vbNewLine & "Extension du fichier=[" & .FileExtension & "]"
        s = s & vbNewLine & "Nom court=[" & .ShortName & "]"
        s = s & vbNewLine & "Chemin court=[" & .ShortPath & "]"
    End With
    
    'info process
    With pProcess
        s = s & vbNewLine & vbNewLine & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "--------------- Processus --------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & "Processus parent=[" & .th32ParentProcessID & "   " & .procParentProcess.szImagePath & "]"
        s = s & vbNewLine & "Threads=[" & .cntThreads & "]"
        s = s & vbNewLine & "Priorité=[" & .pcPriClassBase & "]"
        s = s & vbNewLine & "Mémoire utilisée=[" & .procMemory.WorkingSetSize & "]"
        s = s & vbNewLine & "Pic de mémoire utilisée=[" & .procMemory.PeakWorkingSetSize & "]"
        s = s & vbNewLine & "Utilisation du SWAP=[" & .procMemory.PagefileUsage & "]"
        s = s & vbNewLine & "Pic d'utilisation du SWAP=[" & .procMemory.PeakPagefileUsage & "]"
        s = s & vbNewLine & "QuotaPagedPoolUsage=[" & .procMemory.QuotaPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaNonPagedPoolUsage=[" & .procMemory.QuotaNonPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakPagedPoolUsage=[" & .procMemory.QuotaPeakPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakNonPagedPoolUsage=[" & .procMemory.QuotaPeakNonPagedPoolUsage & "]"
        s = s & vbNewLine & "Erreurs de page=[" & .procMemory.PageFaultCount & "]"
    End With
    
    txtProcessInfos.Text = s
    
    'libère mémoire
    Set cFic = Nothing
    Set pProcess = Nothing
End Sub

Private Sub TB_Click()
'change le frame visible
    MaskFrames TB.SelectedItem.Index - 1
End Sub

Private Sub txtFile_Change()
Dim cFil As clsFile
Dim s As String

    'met à jour les infos du fichier si ce dernier existe
    If cFile.FileExists(txtFile.Text) = False Then
        'existe pas
        txtFileInfos.Text = "Fichier inexistant"
        Exit Sub
    End If
    
    'alors le fichier existe
    'récupère les infos
    Set cFil = cFile.GetFile(txtFile.Text)
    
    With cFil
        s = "Fichier=[" & .File & "]"
        s = s & vbNewLine & "Taille=[" & CStr(.FileSize) & " Octets  -  " & CStr(Round(.FileSize / 1024, 3)) & " Ko" & "]"
        s = s & vbNewLine & "Attribut=[" & CStr(.FileAttributes) & "]"
        s = s & vbNewLine & "Date de création=[" & .CreationDate & "]"
        s = s & vbNewLine & "Date de dernier accès=[" & .LastAccessDate & "]"
        s = s & vbNewLine & "Date de dernière modification=[" & .LastModificationDate & "]"
        s = s & vbNewLine & "Version=[" & .EXEFileVersion & "]"
        s = s & vbNewLine & "Description=[" & .EXEFileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .EXELegalCopyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .EXECompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .EXEInternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .EXEOriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .EXEProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .EXEProductVersion & "]"
        s = s & vbNewLine & "Taille compressée=[" & .FileCompressedSize & "]"
        s = s & vbNewLine & "Programme associé=[" & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & "Répertoire contenant=[" & .FileDirectory & "]"
        s = s & vbNewLine & "Lecteur contenant=[" & .FileDrive & "]"
        s = s & vbNewLine & "Type de fichier=[" & .FileType & "]"
        s = s & vbNewLine & "Extension du fichier=[" & .FileExtension & "]"
        s = s & vbNewLine & "Nom court=[" & .ShortName & "]"
        s = s & vbNewLine & "Chemin court=[" & .ShortPath & "]"
    End With
    
    
    txtFileInfos.Text = s
    
    Set cFil = Nothing  'libère
End Sub

Private Sub txtFolder_Change()
Dim cFol As clsFolder
Dim s As String

    'met à jour les infos du dossier si ce dernier existe
    If cFile.FolderExists(txtFolder.Text) = False Then
        'existe pas
        txtFolderInfos.Text = "Dossier inexistant"
        Exit Sub
    End If
    
    'alors le dossier existe
    'récupère les infos
    Set cFol = cFile.GetFolder(txtFolder.Text)
    
    With cFol
        s = "Dossier=[" & .Folder & "]"
        s = s & vbNewLine & "Date de création=[" & .CreationDate & "]"
        s = s & vbNewLine & "Date de dernier accès=[" & .LastAccessDate & "]"
        s = s & vbNewLine & "Date de dernière modification=[" & .LastModificationDate & "]"
        s = s & vbNewLine & "Nom court=[" & .ShortPath & "]"
        s = s & vbNewLine & "Attribut normal=[" & .IsNormal & "]"
        s = s & vbNewLine & "Attribut caché=[" & .IsHidden & "]"
        s = s & vbNewLine & "Attribut lecture seule=[" & .IsReadOnly & "]"
        s = s & vbNewLine & "Attribut système=[" & .IsSystem & "]"
    End With
    
    
    txtFolderInfos.Text = s
    
    Set cFol = Nothing  'libère
    
End Sub
