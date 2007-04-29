VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Object = "{3AF19019-2368-4F9C-BBFC-FD02C59BD0EC}#1.0#0"; "DriveView_OCX.ocx"
Object = "{2245E336-2835-4C1E-B373-2395637023C8}#1.0#0"; "ProcessView_OCX.ocx"
Begin VB.Form frmHome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu principal"
   ClientHeight    =   5700
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Height          =   435
      Left            =   4125
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Height          =   435
      Left            =   1125
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
   Begin ComctlLib.TabStrip TB 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Tag             =   "lang_ok"
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
      TabIndex        =   23
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   3
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   6495
         Begin ProcessView_OCX.ProcessView PV 
            Height          =   3495
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   6165
            Sorted          =   0   'False
         End
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   3615
            Index           =   1
            Left            =   2880
            TabIndex        =   31
            Top             =   360
            Width           =   3495
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   3255
               Index           =   1
               Left            =   120
               ScaleHeight     =   3255
               ScaleWidth      =   3255
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   240
               Width           =   3255
               Begin VB.TextBox txtProcessInfos 
                  BorderStyle     =   0  'None
                  Height          =   3135
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   11
                  Top             =   0
                  Width           =   3255
               End
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du processus à ouvrir :"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   1
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   6495
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   2295
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   6255
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   1935
               Index           =   3
               Left            =   120
               ScaleHeight     =   1935
               ScaleWidth      =   6015
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   240
               Width           =   6015
               Begin VB.TextBox txtFolderInfos 
                  BorderStyle     =   0  'None
                  Height          =   1815
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
            Tag             =   "pref"
            ToolTipText     =   "Ouvre également tous les fichiers des sous dossiers (lent)"
            Top             =   960
            Width           =   5775
         End
         Begin VB.CommandButton cmdBrowseFolder 
            Height          =   255
            Left            =   5640
            TabIndex        =   6
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFolder 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du dossier à ouvrir :"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   2
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   6495
         Begin DriveView_OCX.DriveView DV 
            Height          =   3615
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   6376
         End
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   3735
            Index           =   0
            Left            =   2640
            TabIndex        =   28
            Top             =   360
            Width           =   3735
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   3375
               Index           =   0
               Left            =   120
               ScaleHeight     =   3375
               ScaleWidth      =   3495
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   240
               Width           =   3495
               Begin VB.TextBox txtDiskInfos 
                  BorderStyle     =   0  'None
                  Height          =   3375
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   10
                  Top             =   0
                  Width           =   3495
               End
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du disque à ouvrir :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   2775
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   0
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   6495
         Begin VB.Frame Frame2 
            Caption         =   "Informations"
            Height          =   3135
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   6255
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Height          =   2775
               Index           =   2
               Left            =   120
               ScaleHeight     =   2775
               ScaleWidth      =   6015
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   240
               Width           =   6015
               Begin VB.TextBox txtFileInfos 
                  BorderStyle     =   0  'None
                  Height          =   2775
                  Left            =   0
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   4
                  Top             =   0
                  Width           =   6015
               End
            End
         End
         Begin VB.CommandButton cmdBrowseFile 
            Height          =   255
            Left            =   5640
            TabIndex        =   3
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   "Choix du fichier à ouvrir :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   4095
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   4
      Left            =   120
      TabIndex        =   37
      Top             =   480
      Width           =   6615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4095
         Index           =   4
         Left            =   50
         ScaleHeight     =   4095
         ScaleWidth      =   6495
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   6495
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2160
            TabIndex        =   14
            Tag             =   "pref"
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cdUnit 
            Height          =   315
            ItemData        =   "frmHome.frx":058A
            Left            =   3360
            List            =   "frmHome.frx":058C
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Tag             =   "pref lang_ok"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtNewFile 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   5415
         End
         Begin VB.CommandButton cmdBrowseNew 
            Height          =   255
            Left            =   5640
            TabIndex        =   13
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Taille du fichier"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Création d'un nouveau fichier"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   3255
         End
      End
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
End
Attribute VB_Name = "frmHome"
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
'FORM DE DEMARRAGE (CHOIX DES ACTIONS A EFFECTUER)
'=======================================================

Private clsPref As clsIniForm
Private bFirst As Boolean

Private Sub cmdBrowseFile_Click()
'browse un fichier
    txtFile.Text = cFile.ShowOpen(Lang.GetString("_FileToOpen"), Me.hWnd, _
        Lang.GetString("_All") & "|*.*")
End Sub

Private Sub cmdBrowseFolder_Click()
'browse un dossier
    txtFolder.Text = cFile.BrowseForFolder(Lang.GetString("_DirToOpen"), Me.hWnd)
End Sub

Private Sub cmdBrowseNew_Click()
'browse un dossier
    txtNewFile.Text = cFile.ShowSave(Lang.GetString("_FileToCreate"), _
        Me.hWnd, Lang.GetString("_All") & "|*.*", App.Path)
End Sub

Private Sub cmdOk_Click()
'ouvre l'élément sélectionné
Dim m() As String
Dim Frm As Form
Dim x As Long
Dim sDrive As String
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
            frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
            
        Case 2
            'dossier
            
            'teste la validité du répertoire
            If cFile.FolderExists(txtFolder.Text) = False Then Exit Sub
            
            'liste les fichiers
            m() = cFile.EnumFilesStr(txtFolder.Text, optFolderSub(0).Value)
            If UBound(m()) < 1 Then Exit Sub
            
            'les ouvre un par un
            For x = 1 To UBound(m)
                If cFile.FileExists(m(x)) Then
                    Set Frm = New Pfm
                    Call Frm.GetFile(m(x))
                    Frm.Show
                    lNbChildFrm = lNbChildFrm + 1
                    frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                    DoEvents
                End If
            Next x
    
        Case 3
        
            If DV.SelectedItem Is Nothing Then Exit Sub
                
            'on check si c'est un disque logique ou un disque physique
            If Left$(DV.SelectedItem.Key, 3) = "log" Then
            
                'disque logique
            
                'vérifie que le drive est accessible
                If DV.IsSelectedDriveAccessible = False Then Exit Sub
                
                'affiche une nouvelle fenêtre
                Set Frm = New diskPfm
                
                Call Frm.GetDrive(DV.SelectedItem.Text)  'renseigne sur le path sélectionné
                
                Unload Me   'quitte cette form
                
                Frm.Show    'affiche la nouvelle
                lNbChildFrm = lNbChildFrm + 1
                frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
            Else
            
                'disque physique
                
                'vérifie que le drive est accessible
                If DV.IsSelectedDriveAccessible = False Then Exit Sub
                
                'affiche une nouvelle fenêtre
                Set Frm = New physPfm
                
                Call Frm.GetDrive(Val(Mid$(DV.SelectedItem.Text, 3, 1)))   'renseigne sur le path sélectionné
                
                Unload Me   'quitte cette form
                
                Frm.Show    'affiche la nouvelle
                lNbChildFrm = lNbChildFrm + 1
                frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
            End If

        Case 4
            'processus

            'vérfie que le processus est ouvrable
            lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(PV.SelectedItem.Tag))
            Call CloseHandle(lH)
            
            If lH = 0 Then
                'pas possible
                Me.Caption = Lang.GetString("_AccessDen")
                Exit Sub
            End If
            
            'possible affiche une nouvelle fenêtre
            Set Frm = New MemPfm
            Call Frm.GetFile(Val(PV.SelectedItem.Tag))
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
            
        Case 5
            
            'vérifie que le fichier n'existe pas
            If cFile.FileExists(txtNewFile.Text) Then
                'fichier déjà existant
                If MsgBox(Lang.GetString("_FileAlreadyExists"), vbInformation + vbYesNo, Lang.GetString("_War")) <> vbYes Then Exit Sub
            End If
            
            Call cFile.CreateEmptyFile(txtNewFile.Text, True)
            
            'création du fichier
            lLen = Abs(Val(txtSize.Text))
            With Lang
                If cdUnit.Text = .GetString("_Ko") Then lLen = lLen * 1024
                If cdUnit.Text = .GetString("_Mo") Then lLen = (lLen * 1024) * 1024
                If cdUnit.Text = .GetString("_Go") Then lLen = ((lLen * 1024) * 1024) * 1024
            End With
            
            'ajoute du texte à la console
            Call AddTextToConsole(Lang.GetString("_CreatingFile"))
    
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
            frmContent.Sb.Panels(2).Text = Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
        
    End Select
    
    If cPref.general_CloseHomeWhenChosen Then
        'alors on ferme cette form car on a ouvert un fichier
        Unload Me
    End If
    
    Exit Sub
    
ErrGestion:
    clsERREUR.AddError "frmHome.cmdOk_Click", True
End Sub

Private Sub DV_DblClick()
    If DV.SelectedItem Is Nothing Then Exit Sub
    If DV.SelectedItem.Children <> 0 Then Exit Sub
    cmdOk_Click
End Sub

Private Sub DV_NodeClick(ByVal Node As ComctlLib.Node)
Dim cDrive As FileSystemLibrary.Drive
Dim cDisk As FileSystemLibrary.PhysicalDisk
Dim s As String
    
    If DV.SelectedItem Is Nothing Then Exit Sub
    
    'on check si c'est un disque logique ou un disque physique
    If Left$(DV.SelectedItem.Key, 3) = "log" Then
        
        'disque logique
        
        'vérifie la disponibilité du disque
        If cFile.IsDriveAvailable(Left$(Node.Text, 1)) = False Then
            'inaccessible
            txtDiskInfos.Text = Lang.GetString("_DiskNon")
            Exit Sub
        End If
        
        'affichage des infos disque
        Set cDrive = cFile.GetDrive(Left$(Node.Text, 1))
        
        With cDrive
            s = Lang.GetString("_Drive") & .VolumeLetter & "]"
            s = s & vbNewLine & Lang.GetString("_VolName") & CStr(.VolumeName) & "]"
            s = s & vbNewLine & Lang.GetString("_Serial") & Hex$(.VolumeSerialNumber) & "]"
            s = s & vbNewLine & Lang.GetString("_FileS") & CStr(.FileSystemName) & "]"
            s = s & vbNewLine & Lang.GetString("_DriveT") & CStr(.strDriveType) & "]"
            s = s & vbNewLine & Lang.GetString("_MedT") & CStr(.strMediaType) & "]"
            s = s & vbNewLine & Lang.GetString("_PartSize") & CStr(.PartitionLength) & "]"
            s = s & vbNewLine & Lang.GetString("_TotalSize") & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
            s = s & vbNewLine & Lang.GetString("_FreeSize") & CStr(.FreeSpace) & "  <--> " & FormatedSize(.FreeSpace, 10) & " ]"
            s = s & vbNewLine & Lang.GetString("_UsedSize") & CStr(.UsedSpace) & "  <--> " & FormatedSize(.UsedSpace, 10) & " ]"
            s = s & vbNewLine & Lang.GetString("_Percent") & CStr(.PercentageFree) & " %]"
            s = s & vbNewLine & Lang.GetString("_LogCount") & CStr(.TotalLogicalSectors) & "]"
            s = s & vbNewLine & Lang.GetString("_PhysCount") & CStr(.TotalPhysicalSectors) & "]"
            s = s & vbNewLine & Lang.GetString("_Hid") & CStr(.HiddenSectors) & "]"
            s = s & vbNewLine & Lang.GetString("_BPerSec") & CStr(.BytesPerSector) & "]"
            s = s & vbNewLine & Lang.GetString("_SPerClust") & CStr(.SectorPerCluster) & "]"
            s = s & vbNewLine & Lang.GetString("_Clust") & CStr(.TotalClusters) & "]"
            s = s & vbNewLine & Lang.GetString("_FreeClust") & CStr(.FreeClusters) & "]"
            s = s & vbNewLine & Lang.GetString("_UsedClust") & CStr(.UsedClusters) & "]"
            s = s & vbNewLine & Lang.GetString("_BPerClust") & CStr(.BytesPerCluster) & "]"
            s = s & vbNewLine & Lang.GetString("_Cyl") & CStr(.Cylinders) & "]"
            s = s & vbNewLine & Lang.GetString("_TPerCyl") & CStr(.TracksPerCylinder) & "]"
            s = s & vbNewLine & Lang.GetString("_SPerT") & CStr(.SectorsPerTrack) & "]"
            s = s & vbNewLine & Lang.GetString("_OffDep") & CStr(.StartingOffset) & "]"
        End With
        
        Set cDrive = Nothing
    Else
        
        'disque physique
        
        'vérifie la disponibilité du disque
        If cFile.IsPhysicalDiskAvailable(Val(Mid$(Node.Text, 3, 1))) = False Then
            'inaccessible
            txtDiskInfos.Text = Lang.GetString("_DiskPNon")
            Exit Sub
        End If
        
        'affichage des infos disque
        Set cDisk = cFile.GetPhysicalDisk(Val(Mid$(Node.Text, 3, 1)))
    
        With cDisk
            s = Lang.GetString("_Drive") & .DiskNumber & "]"
            s = s & vbNewLine & Lang.GetString("_VolName") & CStr(.DiskName) & "]"
            s = s & vbNewLine & Lang.GetString("_MedT") & CStr(.strMediaType) & "]"
            s = s & vbNewLine & Lang.GetString("_TotalSize") & CStr(.TotalSpace) & "  <--> " & FormatedSize(.TotalSpace, 10) & " ]"
            s = s & vbNewLine & Lang.GetString("_PhysCount") & CStr(.TotalPhysicalSectors) & "]"
            s = s & vbNewLine & Lang.GetString("_BPerSec") & CStr(.BytesPerSector) & "]"
            s = s & vbNewLine & Lang.GetString("_Cyl") & CStr(.Cylinders) & "]"
            s = s & vbNewLine & Lang.GetString("_TPerCyl") & CStr(.TracksPerCylinder) & "]"
            s = s & vbNewLine & Lang.GetString("_SPerT") & CStr(.SectorsPerTrack) & "]"
        End With
        
        Set cDisk = Nothing
    End If
    
    txtDiskInfos.Text = s
End Sub

Private Sub Form_Activate()
Dim ND As Node
    
    On Error Resume Next
    
    Call MarkUnaccessibleDrives(Me.DV)  'marque les drives inaccessibles
    
    'on expand si c'est la première activation de la form
    If bFirst = False Then
        bFirst = True
        
        'on extend tous les noeuds
        With PV
            For Each ND In .Nodes
                ND.Expanded = True
            Next ND
            
            'met en surbrillance le premier
            .Nodes.Item(1).Selected = True
        End With
        
        With DV
            For Each ND In .Nodes
                ND.Expanded = True
            Next ND
            
            'met en surbrillance le premier
            .Nodes.Item(1).Selected = True
        End With
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'sauvegarde des preferences
    Call clsPref.SaveFormSettings(App.Path & "\Preferences\HomeWindow.ini", Me)
    Set clsPref = Nothing
End Sub

Private Sub cmdQuit_Click()
    'annule
    Unload Me
End Sub

'=======================================================
'FORM HOME ==> CHOIX DE L'OBJET A OUVRIR
'=======================================================
Private Sub Form_Load()
Dim x As Long

    Set clsPref = New clsIniForm
    
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
        .Language = cPref.env_Lang
        .LoadControlsCaption
        
        DV.PhysicalDrivesString = .GetString("_PhysStr")
        DV.LogicalDrivesString = .GetString("_LogStr")
    End With
    
    bFirst = False
    
    'loading des preferences
    Call clsPref.GetFormSettings(App.Path & "\Preferences\HomeWindow.ini", Me)
    optFolderSub(1).Value = Not (optFolderSub(0).Value)
    
    'réorganise les Frames
    For x = 0 To Frame1.Count - 1
        With Frame1(x)
            .Left = 120
            .Top = 480
            .Width = 6600
            .Height = 4500
        End With
    Next x
    
    'affiche un seul frame
    Call MaskFrames(0)
End Sub

'=======================================================
'masque tous les frames sauf un
'=======================================================
Private Sub MaskFrames(ByVal lFrame As Long)
Dim x As Long

    For x = 0 To Frame1.Count - 1
        Frame1(x).Visible = False
    Next x
    Frame1(lFrame).Visible = True
    
    If lFrame = 4 Then
        'création de fichier
        cmdOK.Caption = Lang.GetString("_CreateAndOpen!")
    Else
        cmdOK.Caption = Lang.GetString("_Open!")
    End If

End Sub

Private Sub PV_NodeClick(ByVal Node As ComctlLib.Node)
'met à jour les infos
Dim pProcess As ProcessItem
Dim cFic As FileSystemLibrary.File
Dim s As String

    On Error Resume Next
    
    'vérifie l'existence du processus
    If cProc.DoesPIDExist(Val(Node.Tag)) = False Then
        'existe pas
        txtProcessInfos.Text = Lang.GetString("_AccDen")
        Exit Sub
    End If
    
    'récupère les infos sur le processus
    Set pProcess = cProc.GetProcess(Val(Node.Tag), False, False, True)
    Set cFic = cFile.GetFile(pProcess.szImagePath)
    
    'infos fichier cible
    With cFic
        s = "-------------------------------------------"
        s = s & vbNewLine & "-------------- " & Lang.GetString("_Cible") & "  -------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & Lang.GetString("_File") & .Path & "]"
        s = s & vbNewLine & Lang.GetString("_Size") & CStr(.FileSize) & " " & Lang.GetString("_Bytes") & "  -  " & CStr(Round(.FileSize / 1024, 3)) & " " & Lang.GetString("_Ko") & "]"
        s = s & vbNewLine & Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        s = s & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        s = s & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        s = s & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        s = s & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        s = s & vbNewLine & Lang.GetString("_Descr") & .FileVersionInfos.FileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .FileVersionInfos.CompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .FileVersionInfos.InternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .FileVersionInfos.OriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .FileVersionInfos.ProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .FileVersionInfos.ProductVersion & "]"
        s = s & vbNewLine & Lang.GetString("_CompS") & .FileCompressedSize & "]"
        s = s & vbNewLine & Lang.GetString("_AssocP") & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & Lang.GetString("_Fold") & .FolderName & "]"
        s = s & vbNewLine & Lang.GetString("_DriveC") & .DriveName & "]"
        s = s & vbNewLine & Lang.GetString("_FileType") & .FileType & "]"
        s = s & vbNewLine & Lang.GetString("_FileExt") & .FileExtension & "]"
        s = s & vbNewLine & Lang.GetString("_ShortN") & .ShortName & "]"
        s = s & vbNewLine & Lang.GetString("_ShortP") & .ShortPath & "]"
    End With
    
    'info process
    With pProcess
        s = s & vbNewLine & vbNewLine & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "--------------- " & Lang.GetString("_Process") & " --------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & Lang.GetString("_ParentP") & .th32ParentProcessID & "   " & .procParentProcess.szImagePath & "]"
        s = s & vbNewLine & "Threads=[" & .cntThreads & "]"
        s = s & vbNewLine & Lang.GetString("_Prior") & .pcPriClassBase & "]"
        s = s & vbNewLine & Lang.GetString("_MemUsed") & .procMemory.WorkingSetSize & "]"
        s = s & vbNewLine & Lang.GetString("_PicMemUsed") & .procMemory.PeakWorkingSetSize & "]"
        s = s & vbNewLine & Lang.GetString("_SwapU") & .procMemory.PagefileUsage & "]"
        s = s & vbNewLine & Lang.GetString("_PicSwapU") & .procMemory.PeakPagefileUsage & "]"
        s = s & vbNewLine & "QuotaPagedPoolUsage=[" & .procMemory.QuotaPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaNonPagedPoolUsage=[" & .procMemory.QuotaNonPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakPagedPoolUsage=[" & .procMemory.QuotaPeakPagedPoolUsage & "]"
        s = s & vbNewLine & "QuotaPeakNonPagedPoolUsage=[" & .procMemory.QuotaPeakNonPagedPoolUsage & "]"
        s = s & vbNewLine & Lang.GetString("_PageF") & .procMemory.PageFaultCount & "]"
    End With
    
    txtProcessInfos.Text = s
    
    'libère mémoire
    Set cFic = Nothing
    Set pProcess = Nothing
End Sub

Private Sub TB_Click()
'change le frame visible
    Call MaskFrames(TB.SelectedItem.Index - 1)
End Sub

Private Sub txtFile_Change()
Dim cFil As FileSystemLibrary.File
Dim s As String

    'met à jour les infos du fichier si ce dernier existe
    If cFile.FileExists(txtFile.Text) = False Then
        'existe pas
        txtFileInfos.Text = Lang.GetString("_FileDoesnt")
        Exit Sub
    End If
    
    'alors le fichier existe
    'récupère les infos
    Set cFil = cFile.GetFile(txtFile.Text)
    
    With cFil
        s = Lang.GetString("_File") & .Path & "]"
        s = s & vbNewLine & Lang.GetString("_Size") & CStr(.FileSize) & " " & Lang.GetString("_Bytes") & "  -  " & CStr(Round(.FileSize / 1024, 3)) & " " & Lang.GetString("_Ko") & "]"
        s = s & vbNewLine & Lang.GetString("_Attr") & CStr(.Attributes) & "]"
        s = s & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        s = s & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        s = s & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        s = s & vbNewLine & Lang.GetString("_Version") & .FileVersionInfos.FileVersion & "]"
        s = s & vbNewLine & Lang.GetString("_Descr") & .FileVersionInfos.FileDescription & "]"
        s = s & vbNewLine & "Copyright=[" & .FileVersionInfos.Copyright & "]"
        s = s & vbNewLine & "CompanyName=[" & .FileVersionInfos.CompanyName & "]"
        s = s & vbNewLine & "InternalName=[" & .FileVersionInfos.InternalName & "]"
        s = s & vbNewLine & "OriginalFileName=[" & .FileVersionInfos.OriginalFileName & "]"
        s = s & vbNewLine & "ProductName=[" & .FileVersionInfos.ProductName & "]"
        s = s & vbNewLine & "ProductVersion=[" & .FileVersionInfos.ProductVersion & "]"
        s = s & vbNewLine & Lang.GetString("_CompS") & .FileCompressedSize & "]"
        s = s & vbNewLine & Lang.GetString("_AssocP") & .AssociatedExecutableProgram & "]"
        s = s & vbNewLine & Lang.GetString("_Fold") & .FolderName & "]"
        s = s & vbNewLine & Lang.GetString("_DriveC") & .DriveName & "]"
        s = s & vbNewLine & Lang.GetString("_FileType") & .FileType & "]"
        s = s & vbNewLine & Lang.GetString("_FileExt") & .FileExtension & "]"
        s = s & vbNewLine & Lang.GetString("_ShortN") & .ShortName & "]"
        s = s & vbNewLine & Lang.GetString("_ShortP") & .ShortPath & "]"
    End With
    
    txtFileInfos.Text = s
    
    Set cFil = Nothing  'libère
End Sub

Private Sub txtFolder_Change()
Dim cFol As FileSystemLibrary.Folder
Dim s As String

    'met à jour les infos du dossier si ce dernier existe
    If cFile.FolderExists(txtFolder.Text) = False Then
        'existe pas
        txtFolderInfos.Text = Lang.GetString("_DirDoesnt")
        Exit Sub
    End If
    
    'alors le dossier existe
    'récupère les infos
    Set cFol = cFile.GetFolder(txtFolder.Text)
    
    With cFol
        s = Lang.GetString("_DirIs") & .Path & "]"
        s = s & vbNewLine & Lang.GetString("_Crea") & .DateCreated & "]"
        s = s & vbNewLine & Lang.GetString("_Access") & .DateLastAccessed & "]"
        s = s & vbNewLine & Lang.GetString("_Modif") & .DateLastModified & "]"
        s = s & vbNewLine & Lang.GetString("_ShortN") & .ShortPath & "]"
        s = s & vbNewLine & Lang.GetString("_NormalAttr") & .IsNormal & "]"
        s = s & vbNewLine & Lang.GetString("_HidAttr") & .IsHidden & "]"
        s = s & vbNewLine & Lang.GetString("_ROAttr") & .IsReadOnly & "]"
        s = s & vbNewLine & Lang.GetString("_SysAttr") & .IsSystem & "]"
    End With
    
    
    txtFolderInfos.Text = s
    
    Set cFol = Nothing  'libère
    
End Sub
