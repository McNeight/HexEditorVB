VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C77F04DF-B546-4EBA-AFE7-F46C1BA9BCF4}#1.0#0"; "LanguageTranslator.ocx"
Object = "{2245E336-2835-4C1E-B373-2395637023C8}#1.0#0"; "ProcessView_OCX.ocx"
Begin VB.Form frmProcess 
   Caption         =   "Gestionnaire de processus"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin ProcessView_OCX.ProcessView PV 
      Height          =   2535
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4471
   End
   Begin VB.PictureBox pctIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4080
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   3240
   End
   Begin ComctlLib.ListView LV 
      Height          =   2655
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "IMG"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":0E42
            Key             =   "Processus|Autoriser le processus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":1194
            Key             =   "Menu|RafraichirF5"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":14E6
            Key             =   "Processus|Bloquer le processus"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":1838
            Key             =   "Processus|Propriétés"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":1B8A
            Key             =   "Processus|Ouvrir explorer à l'emplacement du fichier..."
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":1EDC
            Key             =   "Processus|Terminer le processus"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":222E
            Key             =   "Processus|Rechercher sur Internet..."
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcess.frx":2580
            Key             =   "Menu|Exécuter..."
         EndProperty
      EndProperty
   End
   Begin LanguageTranslator.ctrlLanguage Lang 
      Left            =   0
      Top             =   0
      _ExtentX        =   1402
      _ExtentY        =   1402
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   360
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmProcess.frx":28D2
            Key             =   "noIcon"
            Object.Tag             =   "Pas d'icone dans le fichier qui utilisera cette image"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuRootMnu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuExecute 
         Caption         =   "&Exécuter..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuMenuTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPremierPlan 
         Caption         =   "&Toujours au premier plan"
      End
      Begin VB.Menu mnuIconesDisplay 
         Caption         =   "&Afficher les icones"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuChangeDisplayType 
         Caption         =   "&Afficher une arborescence"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMenuTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefrehNOW 
         Caption         =   "&Rafraichir"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRefreshAuto 
         Caption         =   "&Rafraichissement automatique"
         Begin VB.Menu mnuDeActivate 
            Caption         =   "&Désactiver"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuMenuTiret3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRefreshRapide 
            Caption         =   "&Rapide"
         End
         Begin VB.Menu mnuMoyen 
            Caption         =   "&Moyen"
         End
         Begin VB.Menu mnuLent 
            Caption         =   "&Lent"
         End
      End
      Begin VB.Menu mnuMenuTiret4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Processus"
      Begin VB.Menu mnuTerminate 
         Caption         =   "&Terminer le processus"
      End
      Begin VB.Menu mnuBlockProcess 
         Caption         =   "&Bloquer le processus"
      End
      Begin VB.Menu mnuAutorizeProc 
         Caption         =   "&Autoriser le processus"
      End
      Begin VB.Menu mnuProcTiret1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPriority 
         Caption         =   "&Priorité"
         Begin VB.Menu mnuRealTimeP 
            Caption         =   "&Temps réel"
         End
         Begin VB.Menu mnuHighP 
            Caption         =   "&Haute"
         End
         Begin VB.Menu mnuAboveP 
            Caption         =   "&Supérieure à la normale"
         End
         Begin VB.Menu mnuNormalP 
            Caption         =   "&Normale"
         End
         Begin VB.Menu mnuBelowP 
            Caption         =   "&Inférieure à la normale"
         End
         Begin VB.Menu mnuIdleP 
            Caption         =   "&Basse"
         End
      End
      Begin VB.Menu mnuProcTiret2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Propriétés"
      End
      Begin VB.Menu mnuSearchInternet 
         Caption         =   "&Rechercher sur Internet..."
      End
      Begin VB.Menu mnuOpenExplorer 
         Caption         =   "&Ouvrir explorer à l'emplacement du fichier"
      End
      Begin VB.Menu mnuOpenHexa 
         Caption         =   "&Editer le processus"
         Begin VB.Menu mnuMemoryEdit 
            Caption         =   "&En mémoire"
         End
         Begin VB.Menu mnuDiskEdit 
            Caption         =   "&Sur le disque"
         End
      End
   End
End
Attribute VB_Name = "frmProcess"
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
'FORM DE GESTION SIMPLIFIEE DES PROCESSUS
'=======================================================


Private Sub Form_Load()
'ajoute les en-têtes de colonne

    #If MODE_DEBUG Then
        If App.LogMode = 0 Then
            'on créé le fichier de langue français
            Lang.Language = "French"
            Lang.LangFolder = LANG_PATH
            Lang.WriteIniFileFormIDEform
        End If
    #End If
    
    If App.LogMode = 0 Then
        'alors on est dans l'IDE
        Lang.LangFolder = LANG_PATH
    Else
        Lang.LangFolder = App.Path & "\Lang"
    End If
    
    'applique la langue désirée aux controles
    Lang.Language = MyLang
    Lang.LoadControlsCaption
    
    With LV
        .ColumnHeaders.Add , , "Processus", 1550
        .ColumnHeaders.Add , , "PID", 600
        .ColumnHeaders.Add , , "Path", 3340
        .ColumnHeaders.Add , , "Mémoire utilisée", 1500
        .ColumnHeaders.Add , , "Pic de mémoire utilisée", 1500
        .ColumnHeaders.Add , , "Utilisation du Swap", 1500
        .ColumnHeaders.Add , , "Pic util. swap", 1500
        .ColumnHeaders.Add , , "Erreurs de page", 1500
        .ColumnHeaders.Add , , "Réserve non paginée", 1500
        .ColumnHeaders.Add , , "Pic de réserve non paginée", 1500
        .ColumnHeaders.Add , , "Réserve paginée", 1500
        .ColumnHeaders.Add , , "Pic de réserve paginée", 1500
        .ColumnHeaders.Add , , "Processus parent", 1600
        .ColumnHeaders.Add , , "Threads", 1000
        .ColumnHeaders.Add , , "Priorité", 1000
    End With
    
    'ajoute les icones
    Call AddIconsToMenus(Me.hWnd, Me.ImageList2)
    
    'refresh
    RefreshProcList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With LV
        .Top = 0
        .Left = 0
        .Width = Me.Width - 100
        .Height = Me.Height - 820
    End With
    With PV
        .Top = 0
        .Left = 0
        .Width = LV.Width
        .Height = LV.Height
    End With
End Sub

Private Sub LV_Click()
Dim s As String

    'obtient la string contenant la priorité
    s = Left$(LV.SelectedItem.SubItems(14), InStr(1, LV.SelectedItem.SubItems(14), "[") - 2)
    
    Me.mnuHighP.Checked = False
    Me.mnuRealTimeP.Checked = False
    Me.mnuAboveP.Checked = False
    Me.mnuBelowP.Checked = False
    Me.mnuIdleP.Checked = False
    Me.mnuNormalP.Checked = False
    
    Select Case s
        Case "Temps réel"
            Me.mnuRealTimeP.Checked = True
        Case "Haute"
            Me.mnuHighP.Checked = True
        Case "Supérieure à la normale"
            Me.mnuAboveP.Checked = True
        Case "Normale"
            Me.mnuNormalP.Checked = True
        Case "Inférieure à la normale"
            Me.mnuBelowP.Checked = True
        Case "Basse"
            Me.mnuIdleP.Checked = True
    End Select
    
    DoEvents
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim It As ListItem
Dim s As String


    If Button = 2 Then
        'alors tout d'abord, on sélectionne l'élément sous le curseur
        LV.SelectedItem.Selected = False
        
        Set It = LV.HitTest(X, Y)
        If Not (It Is Nothing) Then It.Selected = True
        
        LV_Click
        
        'affiche le popup menu
        Me.PopupMenu Me.mnuPopUp
    End If
        
End Sub

Private Sub mnuAboveP_Click()
'change la priorité
    cProc.ChangePriority Val(LV.SelectedItem.SubItems(1)), ABOVE_NORMAL_PRIORITY
    RefreshPriority
    LV_Click
End Sub

Private Sub mnuAutorizeProc_Click()
Dim pr() As ProcessItem
Dim tmp As ProcessItem
Dim X As Long

    'autorise le process
    cProc.ResumeProcess Val(LV.SelectedItem.SubItems(1))
    
    'supprime le process de la liste des process bloqués si il est présent
    'ajoute dans pr tous les processus bloqués différents de celui qu'on débloque
    ReDim pr(0)
    
    'récupère le process qui est libéré
    Set tmp = cProc.GetProcess(Val(LV.SelectedItem.SubItems(1)))
    For X = 1 To UBound(JailedProcess())
        With JailedProcess(X)
            If tmp.szImagePath = .szImagePath And tmp.th32ProcessID = .th32ProcessID And _
                tmp.th32ParentProcessID = .th32ParentProcessID Then
                'alors on considère que les processus sont les mêmes (même PID, même process parent
                'et même exécutable)
                'donc dans ce cas on ne garde pas ce process dans la liste des process Jailes
            Else
                'alors là on récupère le process
                ReDim Preserve pr(UBound(pr()) + 1)
                Set pr(UBound(pr())) = JailedProcess(X)
            End If
        End With
    Next X
    
    'on sauvegarde pr dans JailedProcess
    ReDim JailedProcess(UBound(pr()))
    For X = 1 To UBound(pr())
        Set JailedProcess(X) = pr(X)
    Next X
        
    'libère
    Set tmp = Nothing
    
End Sub

Private Sub mnuBelowP_Click()
'change la priorité
    cProc.ChangePriority Val(LV.SelectedItem.SubItems(1)), BELOW_NORMAL_PRIORITY
    RefreshPriority
    LV_Click
End Sub

Private Sub mnuBlockProcess_Click()

    'bloque le processus
    cProc.SuspendProcess Val(LV.SelectedItem.SubItems(1))
    
    'sauvegarde ce processus bloqué dans la liste des process bloqués
    ReDim Preserve JailedProcess(UBound(JailedProcess()) + 1)
    Set JailedProcess(UBound(JailedProcess())) = cProc.GetProcess(Val(LV.SelectedItem.SubItems(1)))
    
End Sub

Private Sub mnuChangeDisplayType_Click()
    If PV.Visible Then
        'alors on change
        PV.Visible = False
        LV.Visible = True
        mnuChangeDisplayType.Caption = "Afficher une arborescence"
    Else
        PV.Visible = True
        LV.Visible = False
        mnuChangeDisplayType.Caption = "Afficher une liste"
    End If
    Call mnuRefrehNOW_Click
End Sub

Private Sub mnuDeActivate_Click()
    mnuRefreshRapide.Checked = False
    Me.mnuLent.Checked = False
    Me.mnuMoyen.Checked = False
    mnuDeActivate.Checked = True
    Timer1.Enabled = False
End Sub

Private Sub mnuDiskEdit_Click()
Dim Frm As Form
    
    'affiche une nouvelle fenêtre
    Set Frm = New Pfm
    Call Frm.GetFile(LV.SelectedItem.SubItems(2))
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"

End Sub

Private Sub mnuExecute_Click()
    ShowRunBox Me.hWnd  'affiche la boite de dialogue Executer...
End Sub

Private Sub mnuHighP_Click()
'change la priorité
    cProc.ChangePriority Val(LV.SelectedItem.SubItems(1)), HIGH_PRIORITY
    RefreshPriority
    LV_Click
End Sub

Private Sub mnuIconesDisplay_Click()
    mnuIconesDisplay.Checked = Not (mnuIconesDisplay.Checked)
    RefreshProcList 'refresh list
End Sub

Private Sub mnuIdleP_Click()
'change la priorité
    cProc.ChangePriority Val(LV.SelectedItem.SubItems(1)), IDLE_PRIORITY
    RefreshPriority
    LV_Click
End Sub

Private Sub mnuLent_Click()
    mnuRefreshRapide.Checked = False
    Me.mnuMoyen.Checked = False
    mnuLent.Checked = True
    mnuDeActivate.Checked = False
    Timer1.Enabled = True
    Timer1.Interval = 3500
End Sub

Private Sub mnuMemoryEdit_Click()
'édition en mémoire du processus
Dim lH As Long
Dim Frm As Form

    'vérfie que le processus est ouvrable
    lH = OpenProcess(PROCESS_ALL_ACCESS, False, Val(LV.SelectedItem.SubItems(1)))
    CloseHandle lH
    
    If lH = 0 Then
        'pas possible
        MsgBox "Accès impossible à ce processus", vbInformation, "Erreur"
        Exit Sub
    End If
        
    'possible affiche une nouvelle fenêtre
    Set Frm = New MemPfm
    Call Frm.GetFile(Val(LV.SelectedItem.SubItems(1)))
    Frm.Show
    lNbChildFrm = lNbChildFrm + 1
    frmContent.Sb.Panels(2).Text = "Ouvertures=[" & CStr(lNbChildFrm) & "]"

End Sub

Private Sub mnuMoyen_Click()
    mnuMoyen.Checked = True
    mnuRefreshRapide.Checked = False
    Me.mnuLent.Checked = False
    mnuDeActivate.Checked = False
    Timer1.Enabled = True
    Timer1.Interval = 1200
End Sub

Private Sub mnuNormalP_Click()
'change la priorité
    cProc.ChangePriority Val(LV.SelectedItem.SubItems(1)), NORMAL_PRIORITY
    RefreshPriority
    LV_Click
End Sub

Private Sub mnuOpenExplorer_Click()
'ouvre explorer à l'endroit du *.exe
    Shell "explorer.exe " & cFile.GetFolderFromPath(LV.SelectedItem.SubItems(2)), vbNormalFocus
End Sub

Private Sub mnuPremierPlan_Click()
'mettre (ou non) au premier plan
    Me.mnuPremierPlan.Checked = Not (Me.mnuPremierPlan.Checked)
    If Me.mnuPremierPlan.Checked Then PremierPlan Me, MettreAuPremierPlan Else _
        PremierPlan Me, MettreNormal
End Sub

Private Sub mnuProperties_Click()
'affiche les propriétés du fichier
    cFile.ShowFileProperty LV.SelectedItem.SubItems(2), Me.hWnd
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRealTimeP_Click()
'change la priorité
    cProc.ChangePriority Val(LV.SelectedItem.SubItems(1)), REALTIME_PRIORITY
    RefreshPriority
    LV_Click
End Sub

Private Sub mnuRefrehNOW_Click()
'rafraichissement
    Call Timer1_Timer
End Sub

Private Sub mnuRefreshRapide_Click()
    mnuRefreshRapide.Checked = True
    Me.mnuLent.Checked = False
    Me.mnuMoyen.Checked = False
    mnuDeActivate.Checked = False
    Timer1.Enabled = True
    Timer1.Interval = 500
End Sub

Private Sub mnuSearchInternet_Click()
Dim sURL As String  'url de recherche
Dim sSearch As String

    'lance la recherche sur le net
    sSearch = LV.SelectedItem.Text  'texte à rechercher
    
    'uniquement Google de dispo (suffisant vu l'utilité de la fonction)
    'formate la string pour la recherche
    sURL = "http://www.google.com/search?hl=en&q=%22" & sSearch & "%22"

    cFile.ShellOpenFile sURL, Me.hWnd
    
End Sub

Private Sub mnuTerminate_Click()
'termine le processus sélectionné
    If cProc.TerminateProcessByPID(Val(LV.SelectedItem.SubItems(1)), True) Then
        DoEvents: RefreshProcList
    End If
End Sub

Private Sub Timer1_Timer()
'rafraichissement
    If LV.Visible Then
        'on refresh la liste
        RefreshProcList
    Else
        'on refresh le PV
        PV.Refresh
    End If
End Sub

'=======================================================
'rafraichissement de la priorité du SelectedItem
'=======================================================
Private Sub RefreshPriority()
Dim p As ProcessItem

    'obtient le process désiré
    Set p = cProc.GetProcess(Val(LV.SelectedItem.SubItems(1)), False, False, False)
    
    'affichage dans le LV
    '/!\ on GELE l'affichage pour éviter le clignotement
    ValidateRect LV.hWnd, 0&

    LV.SelectedItem.SubItems(14) = PriorityFromLong(p.pcPriClassBase) & " [" & p.pcPriClassBase & "]"
    
    InvalidateRect LV.hWnd, 0&, 0&   'dégèle le display
    
    Set p = Nothing

End Sub

'=======================================================
'rafraichissement du LV
'=======================================================
Private Sub RefreshProcList()
Dim p() As ProcessItem
Dim lCount As Long
Dim X As Long
Dim sKey As String

    On Error GoTo ErrGestion
    
    'énumération des processus
    lCount = cProc.EnumerateProcesses(p(), False, False, True)
    
    'affichage dans le LV
    '/!\ on GELE l'affichage pour éviter le clignotement
    ValidateRect LV.hWnd, 0&
    
    LV.ListItems.Clear
    
    If mnuIconesDisplay.Checked Then
        'on affiche les icones

        For X = 0 To lCount - 1
            With LV.ListItems
                
                'ajoute la clé, et l'icone, au IMG
                sKey = "_" & p(X).szImagePath
                
                If DoesKeyExist(sKey) Then
                    'clé existe deja, on rajoute pas
                    .Add Text:=p(X).szExeFile, SmallIcon:="_" & p(X).szImagePath
                ElseIf AddIconToIMG(p(X).szImagePath, "_" & p(X).szImagePath) Then
                    'clé inexistante, on l'a ajoutée
                    
                    'la clé a été correctement ajoutée, on ajoute l'icone correspondant à sKey
                    .Add Text:=p(X).szExeFile, SmallIcon:="_" & p(X).szImagePath
                Else
                    'la clé ne peut être ajoutée (exemple : [system process])
                    .Add Text:=p(X).szExeFile, SmallIcon:="noIcon"
                End If
                
                .Item(X + 1).SubItems(1) = p(X).th32ProcessID
                .Item(X + 1).SubItems(2) = p(X).szImagePath
                .Item(X + 1).SubItems(3) = p(X).procMemory.WorkingSetSize
                .Item(X + 1).SubItems(4) = p(X).procMemory.PeakWorkingSetSize
                .Item(X + 1).SubItems(5) = p(X).procMemory.PagefileUsage
                .Item(X + 1).SubItems(6) = p(X).procMemory.PeakPagefileUsage
                .Item(X + 1).SubItems(7) = p(X).procMemory.PageFaultCount
                .Item(X + 1).SubItems(8) = p(X).procMemory.QuotaNonPagedPoolUsage
                .Item(X + 1).SubItems(9) = p(X).procMemory.QuotaPeakNonPagedPoolUsage
                .Item(X + 1).SubItems(10) = p(X).procMemory.QuotaPagedPoolUsage
                .Item(X + 1).SubItems(11) = p(X).procMemory.QuotaPeakPagedPoolUsage
                .Item(X + 1).SubItems(12) = cProc.GetProcessNameFromPID(p(X).th32ParentProcessID) & "[" & p(X).th32ParentProcessID & "]"
                .Item(X + 1).SubItems(13) = p(X).cntThreads
                .Item(X + 1).SubItems(14) = PriorityFromLong(p(X).pcPriClassBase) & " [" & p(X).pcPriClassBase & "]"
            End With
        Next X
        
    Else
        'pas d'icones
    
        For X = 0 To lCount - 1
            With LV.ListItems
                .Add Text:=p(X).szExeFile
                .Item(X + 1).SubItems(1) = p(X).th32ProcessID
                .Item(X + 1).SubItems(2) = p(X).szImagePath
                .Item(X + 1).SubItems(3) = p(X).procMemory.WorkingSetSize
                .Item(X + 1).SubItems(4) = p(X).procMemory.PeakWorkingSetSize
                .Item(X + 1).SubItems(5) = p(X).procMemory.PagefileUsage
                .Item(X + 1).SubItems(6) = p(X).procMemory.PeakPagefileUsage
                .Item(X + 1).SubItems(7) = p(X).procMemory.PageFaultCount
                .Item(X + 1).SubItems(8) = p(X).procMemory.QuotaNonPagedPoolUsage
                .Item(X + 1).SubItems(9) = p(X).procMemory.QuotaPeakNonPagedPoolUsage
                .Item(X + 1).SubItems(10) = p(X).procMemory.QuotaPagedPoolUsage
                .Item(X + 1).SubItems(11) = p(X).procMemory.QuotaPeakPagedPoolUsage
                .Item(X + 1).SubItems(12) = cProc.GetProcessNameFromPID(p(X).th32ParentProcessID) & "[" & p(X).th32ParentProcessID & "]"
                .Item(X + 1).SubItems(13) = p(X).cntThreads
                .Item(X + 1).SubItems(14) = PriorityFromLong(p(X).pcPriClassBase) & " [" & p(X).pcPriClassBase & "]"
            End With
        Next X
    End If
    
    InvalidateRect LV.hWnd, 0&, 0&   'dégèle le display
    Me.Caption = "Gestionnaire de processus --- " & CStr(lCount) & " processus"
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "frmProcess.RefreshProcList", True
    
End Sub

'=======================================================
'détermine si une clé existe deja ou pas dans le IMG
'=======================================================
Private Function DoesKeyExist(ByVal sKey As String) As Boolean
'renvoie si la clé existe ou non deja dans IMG
Dim l As Long

    DoesKeyExist = False
    
    On Error GoTo ErrGest
    
    l = IMG.ListImages(sKey).Index

    DoesKeyExist = True
    
    Exit Function
    
ErrGest:
'la clé n'existait pas
End Function

'=======================================================
'ajoute une icone au IMG, en fonction du fichier (obtient l'icone de l'executable)
'=======================================================
Private Function AddIconToIMG(ByVal sFile As String, ByVal sKey As String) As Boolean
Dim lstImg As ListImage
Dim hIcon As Long
Dim ShInfo As SHFILEINFO
Dim pct As IPictureDisp

    On Error GoTo ErrGestion
    
    AddIconToIMG = False
    
    If cFile.FileExists(sFile) = False Then Exit Function 'fichier introuvable
    If DoesKeyExist(sKey) Then Exit Function 'clé existe déjà
    
    'obtient le handle de l'icone
    hIcon = SHGetFileInfo(sFile, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
    'prépare la picturebox
    pctIcon.Picture = Nothing
    
    'trace l'image
    ImageList_Draw hIcon, ShInfo.iIcon, pctIcon.hdc, 0, 0, ILD_TRANSPARENT
    
    'ajout de l'icone à l'imagelist
    IMG.ListImages.Add Key:=sKey, Picture:=pctIcon.Image
    
    AddIconToIMG = True

    Exit Function
ErrGestion:
    clsERREUR.AddError "frmProcess.AddIconToImg", True
End Function
