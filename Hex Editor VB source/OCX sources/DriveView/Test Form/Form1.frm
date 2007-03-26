VERSION 5.00
Object = "{8536549B-3A21-4343-AAB6-D03F570E5256}#23.1#0"; "DriveView_OCX.ocx"
Begin VB.Form Form1 
   Caption         =   "Test"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5355
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
   ScaleHeight     =   5715
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Disques physiques"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Disques logiques"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Langue"
      Height          =   1215
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "Anglais"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Français"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin DriveView_OCX.DriveView DriveView1 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9128
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    DriveView1.DisplayLogicalDrives = CBool(Check1.Value)
End Sub

Private Sub Check2_Click()
    DriveView1.DisplayPhysicalDrives = CBool(Check2.Value)
End Sub

Private Sub DriveView1_NodeClick(ByVal Node As ComctlLib.INode)
    If Node.Text <> DriveView1.LogicalDrivesString And Node.Text <> DriveView1.PhysicalDrivesString Then _
    MsgBox "Disque " & Node.Text & " " & IIf(DriveView1.IsSelectedDriveAccessible, vbNullString, "in") & "accessible. Taille = " & CStr(DriveView1.GetSelectedDrive.TotalSpace), vbInformation, "Test"
End Sub

Private Sub Option1_Click()
    With DriveView1
        .LogicalDrivesString = "Disques logiques"
        .PhysicalDrivesString = "Disques physiques"
    End With
End Sub

Private Sub Option2_Click()
    With DriveView1
        .LogicalDrivesString = "Logical disks"
        .PhysicalDrivesString = "Physical disks"
    End With
End Sub
