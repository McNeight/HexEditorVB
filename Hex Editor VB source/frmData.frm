VERSION 5.00
Begin VB.Form frmData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Donnée"
   ClientHeight    =   1470
   ClientLeft      =   -72960
   ClientTop       =   315
   ClientWidth     =   1770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frame 
      Caption         =   "Valeur"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   50
         ScaleHeight     =   1095
         ScaleWidth      =   1605
         TabIndex        =   1
         Top             =   240
         Width           =   1600
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   2
            Left            =   960
            MaxLength       =   1
            TabIndex        =   5
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   1
            Left            =   960
            MaxLength       =   3
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   0
            Left            =   960
            MaxLength       =   2
            TabIndex        =   3
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txtValue 
            BorderStyle     =   0  'None
            Height          =   195
            Index           =   3
            Left            =   960
            TabIndex        =   2
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblValue 
            Caption         =   "String :"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Decimal :"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Hexa :"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblValue 
            Caption         =   "Octal :"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   6
            Top             =   720
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------
'FORM PERMETTANT DE VISUALISER/CHANGER LES VALEURS
'CONTENUES DANS LE HW DE LA FORM ACTIVE
'-------------------------------------------------------

Private I_tem As ItemElement

'-------------------------------------------------------
'obtient l'item sélectionné dans une autre form
'cette sub est appelée à partir de l'ActiveForm et
'renseigne sur la position actuelle du curseur dans le HW
'-------------------------------------------------------
Public Sub GetItem(iItem As ItemElement)
Set I_tem = New ItemElement

    'obtient l'item sélectionné
    I_tem.Col = iItem.Col
    I_tem.Line = iItem.Line
    I_tem.tType = iItem.tType
    I_tem.Value = iItem.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmContent.mnuEditTools.Checked = False
End Sub

Private Sub txtValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'change la valeur de l'item sélectionné du HW de frmcontent.activeform (i_tem)

    If KeyCode = 13 Then
        'alors appui sur "enter"
    
            If Index = 0 Then
                'alors on change les autres champs que le champ "Hexa"
                frmData.txtValue(1).Text = Hex2Dec(txtValue(0).Text)
                frmData.txtValue(2).Text = Hex2Str(txtValue(0).Text)
                frmData.txtValue(3).Text = Hex2Oct(txtValue(0).Text)
            End If
            If Index = 1 Then
                'alors on change les autres champs que le champ "decimal"
                frmData.txtValue(0).Text = Hex$(Val(txtValue(1).Text))
                frmData.txtValue(2).Text = Byte2FormatedString(Val(txtValue(1).Text))
                frmData.txtValue(3).Text = Oct$(Val(txtValue(1).Text))
            End If
            If Index = 2 Then
                'alors on change les autres champs que le champ "string"
                frmData.txtValue(0).Text = Str2Hex(Val(txtValue(2).Text))
                frmData.txtValue(1).Text = Str2Dec(txtValue(2).Text)
                frmData.txtValue(3).Text = Str2Oct(Val(txtValue(2).Text))
            End If

        If IsActiveFormAMemoryEdition Then
            'alors on modifie directement en mémoire
            Call frmContent.ActiveForm.AddAChange(Hex2Dec(txtValue(0).Text))
        Else
            'alors on applique le changement différement
            With frmContent.ActiveForm.HW
                .AddHexValue I_tem.Line, I_tem.Col, txtValue(0).Text
                .AddOneStringValue I_tem.Line, I_tem.Col, txtValue(2).Text
                ModifyData
            End With
        End If
    End If
        
End Sub

'-------------------------------------------------------
'des données ont étés modifiées ==> on sauvegarde ces changements
'-------------------------------------------------------
Private Sub ModifyData()
Dim s As String
Dim x As Long

    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    'définit s (nouvelle string)
    s = vbNullString
    For x = 1 To 16
        s = s & Hex2Str_(frmContent.ActiveForm.HW.Value(I_tem.Line, x))
    Next x

    frmContent.ActiveForm.AddChange frmContent.ActiveForm.HW.FirstOffset + 16 * (I_tem.Line - 1), I_tem.Col, s
    
End Sub
