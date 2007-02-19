VERSION 5.00
Begin VB.UserControl pgrBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.PictureBox tpn 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.PictureBox backImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox imgNull 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox frontImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3360
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "pgrBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'//CONTROLE PERSONNALISE
'//Par Violent_ken
'//PROGRESS BAR AVANCEE --> pgrBar
'=======================================================

'=======================================================
'Historique d'avancement du contôle
'=======================================================
'****v1.0
'initial release
'****v1.1
'-choix de couleur des dégradés
'-contour
'-texte affichable
'-gestion de l'interaction de changement de valeur avec mousemove sur le contrôle
'-différents modes de dégradé
'****v1.2
'-ajout des évênements MouseMove, MouseDown, MouseUp, KeyDown, KeyUp
'-Click, DoubleClick, Changement de valeur, KeyPress, InteractionComplete
'-ValueIsMax, ValueIsMin
'****v1.3
'-ajout de l'icone du contrôle
'-ajout de la vue 3D
'-ajout du choix de couleur avec la palette
'****v1.4
'-ajout du choix d'alignement du caption
'-ajout des positionnements en offset X et Y du caption
'****v1.5
'-choix de la police par boite de dialogue
'****v1.6
'-optimisation de la création des dégradés
'-ajout du choix "Steps" dans le mode d'affichage du texte
'****v1.7
'-ajout du "a propos"
'-possibilité d'ajout d'une image en fond et/ou qui apparait progressivement



'=======================================================
'TO DO
'=======================================================
'-choix de l'orientation G/D/H/B de la barre de valeur
'-ajout de la form de propriétés
'-ajout d'une liste de valeurs (par l'utilisateur) créant un évênement



'=======================================================
'DESCRIPTIF du composant
'=======================================================

'Cette progress bar possède un look stylé XP
'Quelques propriétés graphiques :
'-affichage Smooth (obligatoire)
'-affichage d'un contour (facultatif, couleur à choisir)
'-affichage d'un dégradé dans le fond (les 2 couleurs sont au choix)
'-dégradé de la barre de valeur
'-2 modes de dégradé pour la valeur
'-affichage de texte (valeur ou pourcentage) avec fonte personnalisée
'-choix du nombre de décimales du pourcentage

'Particularité de ce composant :
'-il gère le changement de valeur lors du passage de la souris.
'Possibilité d'activer (ou non) cette fonction, et de définir le bouton
'à utiliser lors du MouseMove sur le composant pour changer sa valeur.
'-gestion de valeurs DECIMALES

'Evênement gérés :
'-MouseMove
'-MouseDown
'-MouseUp
'-KeyDown
'-KeyUp
'-Click
'-DoubleClick
'-Change
'-KeyPress
'-InteractionComplete
'-ValueIsMax
'-ValueIsMin

'Propriétés du composant :
'-Min --> valeur minimale (double)
'-Max --> valeur maximale (double)
'-Value --> valeur active (double)
'-InteractiveControl --> activer ou non la gestion de la valeur lors d'un mousemove
'de la souris sur le contrôle (boolean)
'-ShowLabel --> type d'affichage du label. No=rien, PercentageMode=pourcentage
'et ValueMode=valeur active et Steps=valeur/max
'-RightColor --> couleur de droite du dégradé de la barre de valeur (long)
'-LeftColor --> couleur de gauche du dégradé de la barre de valeur (long)
'-BackColorBottom --> couleur du bas du dégradé du fond de contrôle (long)
'-BackColorTop --> couleur du haut du dégradé du fond de contrôle (long)
'-Fonte --> choix de la fonte (stdFont)
'-Degrade --> type de dégradé de la barre de valeur. AllLengh=dégradé avec
'bord droit de couleur RightColor pour value=max.  OnlyValue=dégradé avec
'bord droit de couleur RightColor pour toutes les values.
'-InteractiveButton --> bouton gérant l'interaction
'NoButton=MouseMove uniquement, les autres boutons correspondent aux boutons
'de la souris.
'-RoundColor --> afficher le contour du controle (boolean)
'-RoundColorValue --> couleur du contour du contrôle (long)
'-LabelColor --> couleur du texte à afficher (long)
'-LabelDecimals --> nombre de décimales à afficher pour le pourcentage
'-BorderStyle --> style d'affichage (3D ou non)
'-Alignement --> position du label dans le control
'-OffSetX --> valeur de décalage horizontal du caption en PIXEL (long)
'-OffSetY --> valeur de décalage vertical du caption en PIXEL (long)
'les offsets positifs décalent vers le bas et le haut
'-BackPicture --> définit la picture affichée en fond de contrôle
'-FrontPicture --> définit la picture affichée en tant que barre de progression






'=======================================================
'SOURCE du composant
'=======================================================

'=======================================================
'APIs
'=======================================================
'pour appliquer des bitmaps
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'=======================================================
'EVENTs publics
'=======================================================
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Change(NewValue As Double, OldValue As Double)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Click()
Public Event DblClick()
Public Event InteractionComplete(NewValue As Double, OldValue As Double)
Public Event ValueIsMax(Value As Double)
Public Event ValueIsMin(Value As Double)

'=======================================================
'VARIABLEs privée
'=======================================================
Private lBackColorTop As Long 'couleur dégradé 1
Private lBackColorBottom As Long 'couleur dégradé 2
Private lLeftColor As Long 'couleur valeur dégradé 1
Private lRightColor As Long 'couleur valeur dégradé 2
Private dMin As Double 'minimum
Private dMax As Double 'maximum
Private bIsInteractive As Boolean 'contrôle intéractif ou non
Private bShowLabel As LABEL_MODE 'type d'affichae
Private dValue As Double 'value
Private mdDeg As MODE_DEGRADE   'type de dégradé
Private btButton As BUTTON_TYPE 'bouton d'interaction
Private brndColor As Boolean    'couleur du contour
Private lLabelColor As Long 'forecolor
Private lPercentDecimal As Long 'nombre de décimales au pourcentage (label)
Private b3D As BORDER  'affichage en 3D
Private taAlign As TEXT_ALIGMENT   'alignement du texte
Private lOSx As Long    'offsetX
Private lOSy As Long    'offsetY

'=======================================================
'ENUMs de choix de propriétés
'=======================================================
Enum LABEL_MODE
    No = 0
    PercentageMode = 1
    ValueMode = 2
    Steps = 3
End Enum
Enum BORDER
    None = 0
    FixedSingle = 1
End Enum
Enum MODE_DEGRADE
    OnlyValue = 1
    AllLengh = 2
End Enum
Enum BUTTON_TYPE
    NoButton = 0
    LeftButton = 1
    RightButton = 2
    MiddleButton = 4
End Enum
Enum TEXT_ALIGMENT
    TopLeft = 1
    TopCenter = 2
    TopRight = 3
    MiddleLeft = 4
    MiddleCenter = 5
    MiddleRight = 6
    BottomLeft = 7
    BottomCenter = 8
    BottomRight = 9
End Enum


'=======================================================
'PROPERTIES
'=======================================================
Public Property Get Alignment() As TEXT_ALIGMENT: Alignment = taAlign: End Property
Public Property Let Alignment(Alignment As TEXT_ALIGMENT): taAlign = Alignment: Refresh: End Property
Public Property Get LabelColor() As OLE_COLOR: LabelColor = lLabelColor: End Property
Public Property Let LabelColor(LabelColor As OLE_COLOR): lLabelColor = LabelColor: Refresh: End Property
Public Property Get Value() As Double: Value = dValue: End Property
Public Property Let Value(Value As Double): Dim lOld As Double
    lOld = dValue
    If Value < Min Then
        dValue = Min
    ElseIf Value > dMax Then
        dValue = dMax
    Else
        dValue = Value
    End If
    RaiseEvent Change(dValue, lOld)
    If dValue = dMin Then RaiseEvent ValueIsMin(dMin)
    If dValue = dMax Then RaiseEvent ValueIsMax(dMax)
    Refresh
End Property
Public Property Get InteractiveButton() As BUTTON_TYPE: InteractiveButton = btButton: End Property
Public Property Let InteractiveButton(InteractiveButton As BUTTON_TYPE): btButton = InteractiveButton: End Property
Public Property Get InteractiveControl() As Boolean: InteractiveControl = bIsInteractive: End Property
Public Property Let InteractiveControl(InteractiveControl As Boolean): bIsInteractive = InteractiveControl: End Property
Public Property Get ShowLabel() As LABEL_MODE: ShowLabel = bShowLabel: End Property
Public Property Let ShowLabel(ShowLabel As LABEL_MODE): bShowLabel = ShowLabel: Refresh: End Property
Public Property Get Font() As StdFont: Set Font = UserControl.Font: End Property
Public Property Set Font(Font As StdFont): Set UserControl.Font = Font: End Property
Public Property Get RoundColor() As Boolean: RoundColor = brndColor: UserControl_Resize: End Property
Public Property Let RoundColor(RoundColor As Boolean): brndColor = RoundColor: End Property
Public Property Get BackPicture() As StdPicture
'BackPicture = lLeftColor
    Set BackPicture = tpn.Picture
    Set BackPicture = backImg.Picture
End Property
Public Property Set BackPicture(ByVal BackPicture As StdPicture)
    Set tpn.Picture = BackPicture
    Set backImg.Picture = BackPicture
End Property
Public Property Get ValuePicture() As StdPicture
    Set ValuePicture = frontImg.Picture
End Property
Public Property Set ValuePicture(ByVal ValuePicture As StdPicture): Set frontImg.Picture = ValuePicture: End Property
Public Property Get LeftColor() As OLE_COLOR: LeftColor = lLeftColor: End Property
Public Property Let LeftColor(LeftColor As OLE_COLOR): lLeftColor = LeftColor: Refresh: End Property
Public Property Get BorderStyle() As BORDER: BorderStyle = b3D: End Property
Public Property Let BorderStyle(BorderStyle As BORDER): b3D = BorderStyle: pct.BorderStyle = b3D: End Property
Public Property Get LabelDecimals() As Long: LabelDecimals = lPercentDecimal: End Property
Public Property Let LabelDecimals(LabelDecimals As Long)
    If LabelDecimals >= 23 Then
        lPercentDecimal = 22
    Else
        lPercentDecimal = LabelDecimals
    End If
    Refresh
End Property
Public Property Get RightColor() As OLE_COLOR: RightColor = lRightColor: End Property
Public Property Let RightColor(RightColor As OLE_COLOR): lRightColor = RightColor: Refresh: End Property
Public Property Get BackColorTop() As OLE_COLOR: BackColorTop = lBackColorTop: End Property
Public Property Let BackColorTop(BackColorTop As OLE_COLOR): lBackColorTop = BackColorTop: Degrader: End Property
Public Property Get BackColorBottom() As OLE_COLOR: BackColorBottom = lBackColorBottom: End Property
Public Property Let BackColorBottom(BackColorBottom As OLE_COLOR): lBackColorBottom = BackColorBottom: Degrader: End Property
Public Property Get OffSetX() As Long: OffSetX = lOSx: End Property
Public Property Let OffSetX(OffSetX As Long): lOSx = OffSetX: Refresh: End Property
Public Property Get OffSetY() As Long: OffSetY = lOSy: End Property
Public Property Let OffSetY(OffSetY As Long): lOSy = OffSetY: Refresh: End Property
Public Property Get Min() As Double: Min = dMin: End Property
Public Property Let Min(Min As Double): dMin = Min: End Property
Public Property Get Max() As Double: Max = dMax: End Property
Public Property Let Max(Max As Double)
If Max > dMin Then dMax = Max
End Property
Public Property Get RoundColorValue() As OLE_COLOR: RoundColorValue = UserControl.BackColor: End Property
Public Property Let RoundColorValue(RoundColorValue As OLE_COLOR): UserControl.BackColor = RoundColorValue: Refresh: End Property
Public Property Get Degrade() As MODE_DEGRADE: Degrade = mdDeg: End Property
Public Property Let Degrade(Degrade As MODE_DEGRADE): mdDeg = Degrade: Refresh: End Property
'=======================================================
'ABOUT
'=======================================================
Public Sub About()
Dim s As String
    s = "prgBar v1.7 par violent_ken (septembre 2006)" & vbNewLine & "Remplace la progressbar de Windows Common Controls"
    MsgBox Prompt:=s, Title:="A propos"
End Sub



'=======================================================
'EVENEMENTS
'=======================================================
Private Sub pct_Click()
    RaiseEvent Click
End Sub
Private Sub pct_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub pct_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub pct_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub pct_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub pct_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub pct_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub


'=======================================================
'USERCONTROL
'=======================================================
Private Sub UserControl_InitProperties()

'valeurs par défaut
    Me.Min = 1
    Me.Max = 100
    Me.Value = 50
    Me.InteractiveControl = False
    Me.ShowLabel = PercentageMode
    Me.RightColor = 16770790
    Me.LeftColor = 12941855
    Me.BackColorBottom = &HEFEFEF
    Me.BackColorTop = &HC6C6C6
    Me.Degrade = AllLengh
    Me.InteractiveButton = LeftButton
    Me.RoundColor = True
    Me.RoundColorValue = &HFFC0C0
    Me.LabelColor = vbWhite
    Me.LabelDecimals = 2
    Me.BorderStyle = None
    Me.Font = UserControl.Font
    Me.Alignment = MiddleCenter
    Me.OffSetX = 0
    Me.OffSetY = 0
    
    'refresh value
    Refresh
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.RoundColor = PropBag.ReadProperty("RoundColor", True)
    Me.LabelDecimals = PropBag.ReadProperty("LabelDecimals", 2)
    Me.RoundColorValue = PropBag.ReadProperty("RoundColorValue", &HFFC0C0)
    Me.BackColorTop = PropBag.ReadProperty("BackColorTop", &HEFEFEF)
    Me.BackColorBottom = PropBag.ReadProperty("BackColorBottom", &HC6C6C6)
    Me.LeftColor = PropBag.ReadProperty("LeftColor", 12941855)
    Me.RightColor = PropBag.ReadProperty("RightColor", 16770790)
    Me.Min = PropBag.ReadProperty("Min", 1)
    Me.Max = PropBag.ReadProperty("Max", 100)
    Me.OffSetX = PropBag.ReadProperty("OffSetX", 0)
    Me.OffSetY = PropBag.ReadProperty("OffSetY", 0)
    Me.Value = PropBag.ReadProperty("Value", 50)
    Me.InteractiveControl = PropBag.ReadProperty("InteractiveControl", False)
    Me.ShowLabel = PropBag.ReadProperty("ShowLabel", PercentageMode)
    Me.Degrade = PropBag.ReadProperty("Degrade", AllLengh)
    Me.LabelColor = PropBag.ReadProperty("LabelColor", vbWhite)
    Me.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Me.Alignment = PropBag.ReadProperty("Alignment", 5)
    Me.InteractiveButton = PropBag.ReadProperty("InteractiveButton", LeftButton)
    Set backImg.Picture = PropBag.ReadProperty("BackPicture", imgNull.Picture)
    Set frontImg.Picture = PropBag.ReadProperty("FrontPicture", imgNull.Picture)
    If backImg.Picture <> 0 Then Set tpn.Picture = backImg.Picture: Refresh
    If frontImg.Picture <> 0 Then Refresh
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("RoundColorValue", Me.RoundColorValue, &HFFC0C0)
    Call PropBag.WriteProperty("RoundColor", Me.RoundColor, True)
    Call PropBag.WriteProperty("LabelDecimals", Me.LabelDecimals, 2)
    Call PropBag.WriteProperty("BackColorTop", Me.BackColorTop, &HEFEFEF)
    Call PropBag.WriteProperty("BackColorBottom", Me.BackColorBottom, &HC6C6C6)
    Call PropBag.WriteProperty("LeftColor", Me.LeftColor, 12941855)
    Call PropBag.WriteProperty("RightColor", Me.RightColor, 16770790)
    Call PropBag.WriteProperty("Min", Me.Min, 1)
    Call PropBag.WriteProperty("Max", Me.Max, 100)
    Call PropBag.WriteProperty("OffSetX", Me.OffSetX, 0)
    Call PropBag.WriteProperty("OffSetY", Me.OffSetY, 0)
    Call PropBag.WriteProperty("Alignment", Me.Alignment, 5)
    Call PropBag.WriteProperty("Value", Me.Value, 50)
    Call PropBag.WriteProperty("InteractiveControl", Me.InteractiveControl, False)
    Call PropBag.WriteProperty("ShowLabel", Me.ShowLabel, PercentageMode)
    Call PropBag.WriteProperty("Degrade", Me.Degrade, AllLengh)
    Call PropBag.WriteProperty("LabelColor", Me.LabelColor, vbWhite)
    Call PropBag.WriteProperty("BorderStyle", Me.BorderStyle, 0)
    Call PropBag.WriteProperty("InteractiveButton", Me.InteractiveButton, LeftButton)
    Call PropBag.WriteProperty("BackPicture", UserControl.backImg, imgNull.Picture)
    Call PropBag.WriteProperty("FrontPicture", UserControl.frontImg, imgNull.Picture)
End Sub



Private Sub pct_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'actionne l'interaction
Dim lOld As Double

    'pour ne pas sortir du composant
    If x > pct.ScaleWidth Then x = pct.ScaleWidth
    If x < 0 Then x = 0
    If y > pct.ScaleHeight Then y = pct.ScaleHeight
    If y < 0 Then y = 0

    RaiseEvent MouseMove(Button, Shift, x, y)
    
    If bIsInteractive And btButton = Button Then
        'alors on change la valeur
        lOld = dValue
        dValue = (dMax / pct.ScaleWidth) * x
        
        'met à 100% quand on sélectionne tout à droite
        If x = (pct.ScaleWidth - 1) Then dValue = dMax
        
        Refresh 'réaffiche la barre

        'évênements
        RaiseEvent InteractionComplete(dValue, lOld)
        RaiseEvent Change(dValue, lOld)
        If dValue = dMin Then RaiseEvent ValueIsMin(dMin)
        If dValue = dMax Then RaiseEvent ValueIsMax(dMax)
        
    End If
End Sub

Private Sub UserControl_Initialize()
'initialisation du controle
    Degrader
    Refresh
End Sub

Private Sub UserControl_Resize()
'resize le composant
Dim lDif As Long    'marge

    'calcule la marge (=0 si pas de bordure)
    If brndColor Then lDif = 30 Else lDif = 0

    'redimensionne les composants
    pct.Left = lDif / 2
    pct.Top = lDif / 2
    pct.Height = UserControl.Height - lDif
    pct.Width = UserControl.Width - lDif
    tpn.Left = -75000
    tpn.Height = UserControl.Height - lDif
    tpn.Width = UserControl.Width - lDif
    backImg.Left = pct.Left
    backImg.Top = pct.Top
    backImg.Width = pct.Width
    backImg.Height = pct.Height

    
    'rafraichit les pictures
    Degrader
    Refresh
End Sub

'=======================================================
'transforme une couleur long en RGB
'=======================================================
Private Function LongToRGB(ByVal lLong As Long, ByRef lRed As Long, ByRef lGreen As Long, ByRef lBlue As Long)
    lBlue = CLng(Int(lLong / 65536))
    lGreen = CLng(Int((lLong - CLng(lBlue) * 65536) / 256))
    lRed = CLng(lLong - CLng(lBlue) * 65536 - CLng(lGreen) * 256)
End Function

'=======================================================
'créé le dégradé
'=======================================================
Private Sub Degrader()
Dim pxlHeight As Long   'hauteur (pixel) du picturebox
Dim pxlWidth As Long    'largeur (pixel) du picturebox
Dim x As Long
Dim y As Long
Dim rUp As Long 'composante rouge couleur du haut
Dim gUp As Long 'composante verte couleur du haut
Dim bUp As Long 'composante bleue couleur du haut
Dim rDown As Long 'composante rouge couleur du bas
Dim gDown As Long 'composante verte couleur du bas
Dim bDown As Long 'composante bleue couleur du bas
Dim dIncrRed As Double  'incrémentation de la composante rouge
Dim dIncrGreen As Double  'incrémentation de la composante verte
Dim dIncrBlue As Double  'incrémentation de la composante bleue


    'on trace dans un picturebox tampon (évite de redissiner ce dégradé à chaque refresh)

    If Me.BackPicture <> 0 Then
        'alors une picture est affichée en fond, donc on redessine pas
        Exit Sub
    End If
    
    'clear Picture
    tpn.Cls
    
    'récupère les valeurs RBG des deux couleurs du dégradé
    LongToRGB BackColorTop, rUp, gUp, bUp
    LongToRGB BackColorBottom, rDown, gDown, bDown
    
    'récupère les dimensions (pixel) du picturebox tampon
    pxlHeight = pct.ScaleHeight
    pxlWidth = pct.ScaleWidth
    
    'calcule les incrémentations pour chaque composante
    dIncrRed = (rUp - rDown) / pxlHeight
    dIncrGreen = (gUp - gDown) / pxlHeight
    dIncrBlue = (bUp - bDown) / pxlHeight
    
    'trace le dégradé
    For x = 0 To pxlHeight
        tpn.ForeColor = RGB(CByte(rDown + x * dIncrRed), CByte(gDown + x * dIncrGreen), CByte(bDown + x * dIncrBlue))
        tpn.Line (0, x)-(pxlWidth, x)
    Next x
    
End Sub

'=======================================================
'rafraichit le contrôle
'=======================================================
Private Sub Refresh()
Dim pxlHeight As Long   'hauteur du picturebox
Dim pxlWidth As Long    'largeur du picturebox
Dim x As Long
Dim y As Long
Dim lValueWidth As Double   'largeur (pixel) de la valeur
Dim lPlage As Double  'nombre de valeurs différentes possibles pour value (en long)
Dim rLeft As Long 'composante rouge couleur de gauche
Dim gLeft As Long 'composante verte couleur de gauche
Dim bLeft As Long 'composante bleue couleur de gauche
Dim rRight As Long 'composante rouge couleur de droite
Dim gRight As Long 'composante verte couleur de droite
Dim bRight As Long 'composante bleue couleur de droite
Dim dIncrRed As Double  'incrémentation de la composante rouge
Dim dIncrGreen As Double  'incrémentation de la composante verte
Dim dIncrBlue As Double  'incrémentation de la composante bleue
Dim lDif As Long    'marge de resizement
Dim txtWidth As Long    'largeur du texte
Dim txtHeight As Long    'hauteur du texte
Dim sText As String 'texte à afficher
Dim lRet As Long    'retour de l'API

    
    'On Error Resume Next
    
    '//affichage du dégradé
    
    
    If frontImg.Picture <> 0 Then
        'alors on a une image à afficher dans la barre de progression
        
        'calcule la largeur de la barre à afficher
        pxlWidth = pct.ScaleWidth
        
        'calcule la largeur (pixel) de la valeur à afficher
        lPlage = dMax - dMin
        
        If lPlage = 0 Then Exit Sub 'pas encore initialisé
    
        'largeur de la picture à poser
        lValueWidth = ((dValue - dMin) / lPlage) * pxlWidth + 0.0001
                
        'efface la picturebox
        pct.Cls
        'ajoute le dégradé de fond
        pct.Picture = tpn.Image
        
        'plaque la picturebox de devant sur la picturebox contenant la barre
        StretchBlt pct.hdc, 0, 0, Int(15 * lValueWidth), pct.Height, frontImg.hdc, 0, _
        0, frontImg.Width, frontImg.Height, &HCC0020
        
        'pct.Picture = frontImg.Picture
        
        
        GoTo BarreDone
    End If
        
        
    'efface la picturebox
    pct.Cls
    'ajoute le dégradé de fond
    pct.Picture = tpn.Image
    
    'obtient les dimensions du picturebox
    pxlHeight = pct.ScaleHeight
    pxlWidth = pct.ScaleWidth
    
    'calcule la largeur (pixel) de la valeur à afficher
    lPlage = dMax - dMin
    
    If lPlage = 0 Then Exit Sub 'pas encore initialisé
    
    lValueWidth = ((dValue - dMin) / lPlage) * pxlWidth + 0.0001
    
    'récupère les valeurs RBG des deux couleurs du dégradé
    LongToRGB LeftColor, rLeft, gLeft, bLeft
    LongToRGB RightColor, rRight, gRight, bRight
    
    'calcule le pas en fonction du mode de dégradé
    If mdDeg = AllLengh Then
        'dégradé sur toute la longueur
        'calcul des incrémentations
        dIncrRed = (rLeft - rRight) / pxlWidth
        dIncrGreen = (gLeft - gRight) / pxlWidth
        dIncrBlue = (bLeft - bRight) / pxlWidth
    Else
        'dégradé uniquement sur la plage affichée
        dIncrRed = (rLeft - rRight) / lValueWidth
        dIncrGreen = (gLeft - gRight) / lValueWidth
        dIncrBlue = (bLeft - bRight) / lValueWidth
    End If
    
    'affichage des dégradés
    For x = 0 To Int(lValueWidth)
        pct.ForeColor = RGB(CByte(rLeft - x * dIncrRed), CByte(gLeft - x * dIncrGreen), CByte(bLeft - x * dIncrBlue))
        pct.Line (x, 0)-(x, pxlHeight)
    Next x
    
    
BarreDone:

    
    '//affichage (ou pas) du texte
    
    'affiche le texte dans le label (en autosize) pour pouvoir _
    'calculer la largeur à prévoir
    'If bShowLabel = No Then GoTo NoText
    If bShowLabel = PercentageMode Then
        'pourcentage
        sText = CStr(Round(100 * (dValue - dMin) / lPlage, lPercentDecimal)) & " %"
    ElseIf bShowLabel = ValueMode Then
        'valeur
        sText = CStr(Round(dValue, lPercentDecimal))
    ElseIf bShowLabel = Steps Then
        'avancement en pas
        sText = CStr(Round(dValue, lPercentDecimal)) & "/" & CStr(dMax)
    End If
    
    lbl.Caption = sText
    
    'récupère la dimension (en pixels)
    txtWidth = lbl.Width / 15
    txtHeight = lbl.Height / 15
    
    'affiche le texte dans pct en le centrant
    'positionnement
    Select Case taAlign
        Case MiddleCenter
            pct.CurrentX = Int((pct.ScaleWidth - txtWidth) / 2) + lOSx
            pct.CurrentY = Int((pct.ScaleHeight - txtHeight) / 2) + lOSy
        Case MiddleLeft
            pct.CurrentX = lOSx
            pct.CurrentY = Int((pct.ScaleHeight - txtHeight) / 2) + lOSy
        Case MiddleRight
            pct.CurrentX = pct.ScaleWidth - txtWidth + lOSx
            pct.CurrentY = Int((pct.ScaleHeight - txtHeight) / 2) + lOSy
        Case TopLeft
            pct.CurrentX = lOSx
            pct.CurrentY = lOSy
        Case TopCenter
            pct.CurrentX = Int((pct.ScaleWidth - txtWidth) / 2) + lOSx
            pct.CurrentY = lOSy
        Case TopRight
            pct.CurrentX = pct.ScaleWidth - txtWidth + lOSx
            pct.CurrentY = lOSy
        Case BottomLeft
            pct.CurrentX = lOSx
            pct.CurrentY = pct.ScaleHeight - txtHeight + lOSy
        Case BottomRight
            pct.CurrentX = pct.ScaleWidth - txtWidth + lOSx
            pct.CurrentY = pct.ScaleHeight - txtHeight + lOSy
        Case BottomCenter
            pct.CurrentX = Int((pct.ScaleWidth - txtWidth) / 2) + lOSx
            pct.CurrentY = pct.ScaleHeight - txtHeight + lOSy
    End Select
    
    'affichage du texte
    pct.ForeColor = LabelColor
    pct.Font = UserControl.Font
    
    pct.Print sText
    
    
NoText:
    
    '//resize le composant
    
    'calcule la marge (=0 si pas de bordure)
    If brndColor Then lDif = 30 Else lDif = 0

    'redimensionne les composants
    pct.Left = lDif / 2
    pct.Top = lDif / 2
    pct.Height = UserControl.Height - lDif
    pct.Width = UserControl.Width - lDif
    tpn.Left = -75000
    tpn.Height = UserControl.Height - lDif
    tpn.Width = UserControl.Width - lDif

End Sub


