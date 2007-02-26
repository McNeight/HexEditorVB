VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
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
   ScaleHeight     =   5415
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2520
      Top             =   3720
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
   Begin ComctlLib.ImageList IMG 
      Left            =   480
      Top             =   3720
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
            Picture         =   "frmProcess.frx":08CA
            Key             =   "noIcon"
            Object.Tag             =   "Pas d'icone dans le fichier qui utilisera cette image"
         EndProperty
      EndProperty
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
   Begin VB.Menu mnuRootMnu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuExecute 
         Caption         =   "&Executer..."
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
    
    'refresh
    RefreshProcList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    LV.Top = 0
    LV.Left = 0
    LV.Width = Me.Width - 100
    LV.Height = Me.Height - 820
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

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim It As ListItem
Dim s As String


    If Button = 2 Then
        'alors tout d'abord, on sélectionne l'élément sous le curseur
        LV.SelectedItem.Selected = False
        
        Set It = LV.HitTest(x, y)
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
Dim x As Long

    'autorise le process
    cProc.ResumeProcess Val(LV.SelectedItem.SubItems(1))
    
    'supprime le process de la liste des process bloqués si il est présent
    'ajoute dans pr tous les processus bloqués différents de celui qu'on débloque
    ReDim pr(0)
    
    'récupère le process qui est libéré
    Set tmp = cProc.GetProcess(Val(LV.SelectedItem.SubItems(1)))
    For x = 1 To UBound(JailedProcess())
        With JailedProcess(x)
            If tmp.szImagePath = .szImagePath And tmp.th32ProcessID = .th32ProcessID And _
                tmp.th32ParentProcessID = .th32ParentProcessID Then
                'alors on considère que les processus sont les mêmes (même PID, même process parent
                'et même exécutable)
                'donc dans ce cas on ne garde pas ce process dans la liste des process Jailes
            Else
                'alors là on récupère le process
                ReDim Preserve pr(UBound(pr()) + 1)
                Set pr(UBound(pr())) = JailedProcess(x)
            End If
        End With
    Next x
    
    'on sauvegarde pr dans JailedProcess
    ReDim JailedProcess(UBound(pr()))
    For x = 1 To UBound(pr())
        Set JailedProcess(x) = pr(x)
    Next x
        
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
    Call Frm.GetFile(Val(LV.SelectedItem.SubItems(1)), LV.SelectedItem.Text)
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
    RefreshProcList
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
    If cProc.TerminateProc(Val(LV.SelectedItem.SubItems(1)), True) Then
        DoEvents: RefreshProcList
    End If
End Sub

Private Sub Timer1_Timer()
'rafraichissement
    RefreshProcList
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
Dim x As Long
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

        For x = 0 To lCount - 1
            With LV.ListItems
                
                'ajoute la clé, et l'icone, au IMG
                sKey = "_" & p(x).szImagePath
                
                If DoesKeyExist(sKey) Then
                    'clé existe deja, on rajoute pas
                    .Add Text:=p(x).szExeFile, SmallIcon:="_" & p(x).szImagePath
                ElseIf AddIconToIMG(p(x).szImagePath, "_" & p(x).szImagePath) Then
                    'clé inexistante, on l'a ajoutée
                    
                    'la clé a été correctement ajoutée, on ajoute l'icone correspondant à sKey
                    .Add Text:=p(x).szExeFile, SmallIcon:="_" & p(x).szImagePath
                Else
                    'la clé ne peut être ajoutée (exemple : [system process])
                    .Add Text:=p(x).szExeFile, SmallIcon:="noIcon"
                End If
                
                .Item(x + 1).SubItems(1) = p(x).th32ProcessID
                .Item(x + 1).SubItems(2) = p(x).szImagePath
                .Item(x + 1).SubItems(3) = p(x).procMemory.WorkingSetSize
                .Item(x + 1).SubItems(4) = p(x).procMemory.PeakWorkingSetSize
                .Item(x + 1).SubItems(5) = p(x).procMemory.PagefileUsage
                .Item(x + 1).SubItems(6) = p(x).procMemory.PeakPagefileUsage
                .Item(x + 1).SubItems(7) = p(x).procMemory.PageFaultCount
                .Item(x + 1).SubItems(8) = p(x).procMemory.QuotaNonPagedPoolUsage
                .Item(x + 1).SubItems(9) = p(x).procMemory.QuotaPeakNonPagedPoolUsage
                .Item(x + 1).SubItems(10) = p(x).procMemory.QuotaPagedPoolUsage
                .Item(x + 1).SubItems(11) = p(x).procMemory.QuotaPeakPagedPoolUsage
                .Item(x + 1).SubItems(12) = cProc.GetProcessFromPID(p(x).th32ParentProcessID) & "[" & p(x).th32ParentProcessID & "]"
                .Item(x + 1).SubItems(13) = p(x).cntThreads
                .Item(x + 1).SubItems(14) = PriorityFromLong(p(x).pcPriClassBase) & " [" & p(x).pcPriClassBase & "]"
            End With
        Next x
        
    Else
        'pas d'icones
    
        For x = 0 To lCount - 1
            With LV.ListItems
                .Add Text:=p(x).szExeFile
                .Item(x + 1).SubItems(1) = p(x).th32ProcessID
                .Item(x + 1).SubItems(2) = p(x).szImagePath
                .Item(x + 1).SubItems(3) = p(x).procMemory.WorkingSetSize
                .Item(x + 1).SubItems(4) = p(x).procMemory.PeakWorkingSetSize
                .Item(x + 1).SubItems(5) = p(x).procMemory.PagefileUsage
                .Item(x + 1).SubItems(6) = p(x).procMemory.PeakPagefileUsage
                .Item(x + 1).SubItems(7) = p(x).procMemory.PageFaultCount
                .Item(x + 1).SubItems(8) = p(x).procMemory.QuotaNonPagedPoolUsage
                .Item(x + 1).SubItems(9) = p(x).procMemory.QuotaPeakNonPagedPoolUsage
                .Item(x + 1).SubItems(10) = p(x).procMemory.QuotaPagedPoolUsage
                .Item(x + 1).SubItems(11) = p(x).procMemory.QuotaPeakPagedPoolUsage
                .Item(x + 1).SubItems(12) = cProc.GetProcessFromPID(p(x).th32ParentProcessID) & "[" & p(x).th32ParentProcessID & "]"
                .Item(x + 1).SubItems(13) = p(x).cntThreads
                .Item(x + 1).SubItems(14) = PriorityFromLong(p(x).pcPriClassBase) & " [" & p(x).pcPriClassBase & "]"
            End With
        Next x
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
