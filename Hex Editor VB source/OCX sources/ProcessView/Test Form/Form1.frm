VERSION 5.00
Object = "{88A64AB7-8026-47F4-8E67-1A0451E8679C}#1.0#0"; "ProcessView_OCX.ocx"
Begin VB.Form Form1 
   Caption         =   "Test"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPath 
      Caption         =   "Afficher le path entier"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CommandButton cmdProcParentInfos 
      Caption         =   "Infos sur processus parent"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3000
      Width           =   4335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Rafraichir"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Débloquer le processus sélectionné"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmdBlock 
      Caption         =   "Bloquer le processus sélectionné"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "Tuer le processus sélectionné"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin ProcessView_OCX.ProcessView PV 
      Height          =   5895
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10398
      DisplayPath     =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private Sub chkPath_Click()
    PV.DisplayPath = CBool(chkPath.Value)
End Sub

Private Sub cmdBlock_Click()
Dim lPID As Long

    'PID de l'élément sélectionné
    lPID = PV.SelectedProcess.th32ProcessID
    
    'on bloque
    Call PV.Processes.SuspendProcess(lPID)
End Sub

Private Sub cmdKill_Click()
'on tue, MAIS DEMANDE DE CONFIRMATION
Dim lPID As Long

    'PID de l'élément sélectionné
    lPID = PV.SelectedProcess.th32ProcessID
    
    'on kill
    Call PV.Processes.TerminateProcessByPID(lPID, True)

End Sub

Private Sub cmdProcParentInfos_Click()
'on affiche les infos sur le process parent
Dim s As String
Dim tProcParent As ProcessView_OCX.ProcessItem

    On Error Resume Next    'évite les bugs quand on récupère des infos sur un parent qui n'existe plus
    
    If PV.SelectedItem Is Nothing Then
        MsgBox "SVP sélectionnez un processus", vbInformation, "Attention"
        Exit Sub
    End If
    
    'récupère les infos sur le parent EN DEMANDANT LES INFOS MEMOIRES
    Set tProcParent = PV.Processes.GetProcess(PV.SelectedProcess.th32ParentProcessID, _
        False, False, True)
        
    If tProcParent Is Nothing Then Exit Sub
    
    With tProcParent
        s = "-------------------------------------------"
        s = s & vbNewLine & "--------------- Processus --------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & "Fichier=[" & .szImagePath & "]"
        s = s & vbNewLine & "ProcessName=[" & .szExeFile & "]"
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
    
    MsgBox s, vbInformation, "Propriétés du parent"
End Sub

Private Sub cmdRefresh_Click()
Dim lTim As Long
    
    'petit bench pour le refresh
    lTim = GetTickCount
    
    'refresh
    Call PV.Refresh
    
    lTim = GetTickCount - lTim
    Me.Caption = "Affichage en " & Str$(lTim) & " ms"

End Sub

Private Sub cmdResume_Click()
Dim lPID As Long

    'PID de l'élément sélectionné
    lPID = PV.SelectedProcess.th32ProcessID
    
    'on débloque
    Call PV.Processes.ResumeProcess(lPID)
End Sub

Private Sub PV_NodeClick(ByVal Node As ComctlLib.INode)
'on affiche les infos sur le process
Dim s As String
    
    'DEMANDE D'INFOS SUR LA MEMOIRE
    With PV.SelectedProcess(False, False, True)
        s = "-------------------------------------------"
        s = s & vbNewLine & "--------------- Processus --------------"
        s = s & vbNewLine & "-------------------------------------------"
        s = s & vbNewLine & "PID=[" & .th32ProcessID & "]"
        s = s & vbNewLine & "Fichier=[" & .szImagePath & "]"
        s = s & vbNewLine & "ProcessName=[" & .szExeFile & "]"
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
    
    txt.Text = s
End Sub
