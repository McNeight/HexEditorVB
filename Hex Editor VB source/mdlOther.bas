Attribute VB_Name = "mdlOther"
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



Private MinX As Single
Private MinY As Single
Private AdressWinProc As Long



'=======================================================
'//MODULE CONTENANT DIFFERENTES SUBS ET FUNCTIONS NECESSAIRES
'=======================================================


'=======================================================
'met au premier plan ou non une form
'=======================================================
Public Function PremierPlan(Frm As Form, IsPremierPlan As ModePlan) As Long
    Select Case IsPremierPlan
        Case True
            PremierPlan = SetWindowPos(Frm.hWnd, -1, 0, 0, 0, 0, VISIBLEFLAGS)
        Case False
            PremierPlan = SetWindowPos(Frm.hWnd, -2, 0, 0, 0, 0, VISIBLEFLAGS)
        End Select
End Function

'=======================================================
'modifie les enabled des buttons/menus pour les undo/redo
'=======================================================
Public Sub ModifyHistoEnabled()
Dim l As Long
Dim c As Long

    On Error Resume Next
    
    'numéro de l'item sélectionné
    
    With frmContent
        l = .ActiveForm.lstHisto.SelectedItem.Index
        c = .ActiveForm.lstHisto.ListItems.Count
        
        If l > 1 And l < c Then
            'alors des élements avant/après
            .mnuUndo.Enabled = True
            .mnuRedo.Enabled = True
            .Toolbar1.Buttons.Item(12).Enabled = True
            .Toolbar1.Buttons.Item(13).Enabled = True
            Exit Sub
        End If
      
        If l = 1 Then
            'alors pas d'élément sélectionné (Undo=max)
            .mnuUndo.Enabled = (.ActiveForm.lstHisto.ListItems.Item(1).Selected)
            .mnuRedo.Enabled = True
            .Toolbar1.Buttons.Item(12).Enabled = (.ActiveForm.lstHisto.ListItems.Item(1).Selected)
            .Toolbar1.Buttons.Item(13).Enabled = True
            Exit Sub
        End If
        
        If c = l And .ActiveForm.lstHisto.ListItems.Item(l).Selected = False Then
            'alors pas d'élément sélectionné (Redo=max)
            .mnuUndo.Enabled = True
            .mnuRedo.Enabled = False
            .Toolbar1.Buttons.Item(12).Enabled = True
            .Toolbar1.Buttons.Item(13).Enabled = False
            Exit Sub
        End If
        
        If c = 0 Or l = -1 Then
            'alors rien de sélectionné
            .mnuUndo.Enabled = False
            .mnuRedo.Enabled = False
            .Toolbar1.Buttons.Item(12).Enabled = False
            .Toolbar1.Buttons.Item(13).Enabled = False
            Exit Sub
        End If
        
        If l = c And l = 1 Then
            'alors qu'un seul élément
            .mnuUndo.Enabled = True
            .mnuRedo.Enabled = True
            .Toolbar1.Buttons.Item(12).Enabled = True
            .Toolbar1.Buttons.Item(13).Enabled = True
            Exit Sub
        End If
        
        If l < 1 And c > 1 Then
            'alors seulement redo possible
            .mnuUndo.Enabled = False
            .mnuRedo.Enabled = True
            .Toolbar1.Buttons.Item(12).Enabled = False
            .Toolbar1.Buttons.Item(13).Enabled = True
            Exit Sub
        End If
        
        If c = l And c > 1 Then
            'alors seulement undo possible
            .mnuUndo.Enabled = True
            .mnuRedo.Enabled = False
            .Toolbar1.Buttons.Item(12).Enabled = True
            .Toolbar1.Buttons.Item(13).Enabled = False
            Exit Sub
        End If

    End With
        
End Sub

'=======================================================
'ajoute l'entrée du menu contextuel
'=======================================================
Public Sub AddContextMenu(ByVal tType As Byte)
Dim cReg As clsRegistry

    Set cReg = New clsRegistry
    'tType=1 ==> Fichier
    'tType=2 ==> Dossier

    If tType = 1 Then
        'créé les clés registre nécessaires pour ajouter une entrée au menu contextuel (fichier)
        cReg.CreateKey HKEY_CLASSES_ROOT, "*\Shell\hexeditShellMenu\Command"
        cReg.WriteValue HKEY_CLASSES_ROOT, "*\Shell\hexeditShellMenu", "", "Ouvrir avec HexEditor", REG_SZ
        cReg.WriteValue HKEY_CLASSES_ROOT, "*\Shell\hexeditShellMenu\Command\", "", Chr$(34) & App.Path & "\HexEditor.exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), REG_SZ
    Else
        'dossier
        cReg.CreateKey HKEY_CLASSES_ROOT, "Folder\Shell\hexeditShellMenu\Command"
        cReg.WriteValue HKEY_CLASSES_ROOT, "Folder\Shell\hexeditShellMenu", "", "Ouvrir avec HexEditor", REG_SZ
        cReg.WriteValue HKEY_CLASSES_ROOT, "Folder\Shell\hexeditShellMenu\Command\", "", Chr$(34) & App.Path & "\HexEditor.exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34), REG_SZ
    End If
    
    Set cReg = Nothing
End Sub

'=======================================================
'enlève l'entrée du menu contextuel
'=======================================================
Public Sub RemoveContextMenu(ByVal tType As Byte)
Dim cReg As clsRegistry

    Set cReg = New clsRegistry
    
    If tType = 1 Then
        'fichier
        cReg.DelKey HKEY_CLASSES_ROOT, "*\Shell\hexeditShellMenu\Command"
        cReg.DelKey HKEY_CLASSES_ROOT, "*\Shell\hexeditShellMenu"
    Else
        'dossier
        cReg.DelKey HKEY_CLASSES_ROOT, "Folder\Shell\hexeditShellMenu\Command"
        cReg.DelKey HKEY_CLASSES_ROOT, "Folder\Shell\hexeditShellMenu"
    End If
    
    Set cReg = Nothing
    
End Sub

'=======================================================
'renvoie a^b (plus rapide que a^b)
'=======================================================
Public Function AexpB(ByVal a As Long, ByVal b As Long) As Currency
Dim x As Long
Dim l As Long

    On Error Resume Next

    If b = 0 Then
        AexpB = 1
        Exit Function
    End If
    
    l = 1
    For x = 1 To b
        l = l * a
    Next x
    AexpB = l

End Function

'=======================================================
'renvoie une valeur divisible par 16 (supérieure à l)
'=======================================================
Public Function By16(ByVal l As Currency) As Currency
Dim r As Currency

    r = l
    While Int(r / 16) < (r / 16)
        r = r + 1
    Wend
    By16 = r

End Function

'=======================================================
'renvoie une valeur divisible par 16 (inférieure à l)
'=======================================================
Public Function By16D(ByVal l As Currency) As Currency
Dim r As Currency

    If (l Mod 16) = 0 Then
        By16D = l
        Exit Function
    End If

    r = l
    While Int(r / 16) < (r / 16)
        r = r + 1
    Wend
    By16D = r - 16

End Function

'=======================================================
'effectue un modulo sans dépassement de capacité
'très peu optimisé, mais utile pour les grandes valeurs de cur
'=======================================================
Public Function Mod2(ByVal cur As Currency, lng As Long) As Currency
    Mod2 = cur - Int(cur / lng) * lng
End Function

'=======================================================
'renvoie une valeur divisible par n (inférieure à l)
'=======================================================
Public Function ByND(ByVal l As Currency, ByVal n As Long) As Currency
Dim r As Currency

    If Mod2(l, n) = 0 Then
        ByND = l
        Exit Function
    End If

    r = l
    While Int(r / n) < (r / n)
        r = r + 1
    Wend
    ByND = r - n

End Function

'=======================================================
'renvoie une valeur divisible par n (supérieure à l)
'=======================================================
Public Function ByN(ByVal l As Currency, ByVal n As Long) As Currency
Dim r As Currency

    r = l
    While Int(r / n) < (r / n)
        r = r + 1
    Wend
    ByN = r

End Function

'=======================================================
'renvoie le type d'activeform
'=======================================================
Public Function TypeOfActiveForm() As String
    
    On Error Resume Next
    
    TypeOfActiveForm = vbNullString
    TypeOfActiveForm = frmContent.ActiveForm.Useless

End Function

'=======================================================
'renvoie le type d'une form
'=======================================================
Public Function TypeOfForm(Frm As Form) As String
    
    On Error Resume Next
    
    TypeOfForm = IIf(Frm.Useless = "Pfm", "Fichier", vbNullString) & IIf(Frm.Useless = "Mem", "Processus", vbNullString) & IIf(Frm.Useless = "Disk", "Disque", vbNullString)

End Function

'=======================================================
'divise une currency en 2 long ==> créé une LARGE_INTEGER
'=======================================================
Public Sub GetLargeInteger(ByVal curVar As Currency, ByRef lngLowPart As Long, ByRef lngHighPart As Long)
    lngLowPart = 0: lngHighPart = 0
    Do
        If curVar < 2 ^ 32 Then Exit Do
        curVar = curVar - 2 ^ 32: lngHighPart = lngHighPart + 1
        'If lngHighPart >= (2 ^ 31) Then lngHighPart = lngHighPart - (2 ^ 32)
    Loop
    If curVar >= 2 ^ 31 Then curVar = curVar - 2 ^ 32
    lngLowPart = CLng(curVar)
End Sub

'=======================================================
'rassemble de long pour former une currency
'=======================================================
Public Function GetCurrency(ByVal lngLowPart As Long, ByVal lngHighPart As Long) As Currency
    GetCurrency = 2 ^ 32 * lngHighPart + lngLowPart
End Function

'=======================================================
'transforme un largeinterger en currency
'=======================================================
Public Function LI2Currency(liInput As LARGE_INTEGER) As Currency
    CopyMemory LI2Currency, liInput, LenB(liInput)
End Function

'=======================================================
'transforme un filetime en currency
'=======================================================
Public Function FT2Currency(FT As FILETIME) As Currency
    CopyMemory FT2Currency, FT, LenB(FT)
End Function

'=======================================================
'affichage de la boite de dialogue Executer...
'=======================================================
Public Function ShowRunBox(ByVal hWnd As Long) As Long
    ShowRunBox = SHRunDialog(hWnd, 0, 0, StrConv("Exécuter", vbUnicode), _
        StrConv("Sélectionnez un élément à lancer (dossier, document, exécutable...) et Process Guardian l'ouvrira pour vous.", _
        vbUnicode), 0)
End Function

'=======================================================
'récupère l'icone associée à un fichier
'sortie en type IPictureDisp
'=======================================================
Public Function CreateIcon(ByVal sFile As String) As IPicture
Dim vSHFI As SHFILEINFO
Dim lAttr As Long
Dim vStruct As PICTDESC
Dim vGuid   As GUID

    On Error GoTo ErrGestion

    'prend la LargeIcon ==> définit un Flag correspondant
    lAttr = SHGFI_LARGEICON Or SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES Or SHGFI_TYPENAME
    
    'obtient infos sur le fichier (ici l'icone est utilisée)
    SHGetFileInfo sFile, FILE_ATTRIBUTE_NORMAL, vSHFI, Len(vSHFI), lAttr

    If vSHFI.hIcon = 0 Then Exit Function
    
    'prépare la structure contenant l'icone
    With vStruct
       .dwType = vbPicTypeIcon
       .dwSize = Len(vStruct)
       .hImage = vSHFI.hIcon
    End With
    
    'affectation de l'icone sous form de IPicture à la fonction CreateIcon
    'en fonction de la structure définie
    If CLSIDFromString(StrPtr(IID_IICON), vGuid) = 0 Then _
    OleCreatePictureIndirect vStruct, vGuid, True, CreateIcon

    Exit Function
ErrGestion:
    clsERREUR.AddError "mdlOther.CreateIcon", True
End Function

'=======================================================
'ajoute les icones du fichier sFile au Listview spécifié
'utilise une picture et une IMG pour tracer les images
'=======================================================
Public Sub LoadIconesToLV(ByVal sFile As String, LV As ListView, pct As PictureBox, IMG As ImageList)
Dim lIcon As Long
Dim x As Long
Dim b As Long
    
    On Error GoTo ErrGestion
    
    LV.ListItems.Clear
    lIcon = 1: x = 0
    
    'tant que l'on trouve des icones
    Do
        lIcon = ExtractIcon(App.hInstance, sFile, x)
        
        If lIcon = 0 Then Exit Do   'plus d'icone, quitte
        
        pct.Cls 'clear picture
        
        DrawIconEx pct.hdc, 0, 0, lIcon, 0, 0, 0, 0, &H1 Or &H2 'trace la picture
        SimpleAddToLV "_" & CStr(lIcon), pct.Image, IMG 'ajoute au LV
        
        ValidateRect LV.hWnd, 0&    'gèle l'affichage
        
        'Incrementation de l'emplacement de l'icone pour l'extraction
        x = x + 1
        
        LV.ListItems.Add , , vbNullString, , "_" & CStr(lIcon)    'ajout de l'icone
        DoEvents
        
        DestroyIcon lIcon   'décharge l'icone
    Loop
    
    InvalidateRect LV.hWnd, 0&, 0&  'dégèle l'affichage
    
    Exit Sub
ErrGestion:
    clsERREUR.AddError "mdlOther.LoadIconesToLV", True
End Sub

'=======================================================
'ajout d'une image au ImageList
'pas d'erreur en cas de clé déjà existante
'=======================================================
Public Sub SimpleAddToLV(ByVal sKey As String, IMG As Picture, ImageL As ImageList)
Dim lst As ListImage

    'On Error Resume Next

    Set lst = ImageL.ListImages.Add(Key:=sKey, Picture:=IMG)

End Sub

'=======================================================
'enregsitre le type de fichier *.hescr
'=======================================================
Public Sub Reg_HESCR_file()
Dim cReg As clsRegistry
    
    'instancie la classe
    Set cReg = New clsRegistry
    
    'associe l'icone
    With cReg
        .CreateKey HKEY_CLASSES_ROOT, "HexEditor VB.hescr"
        .WriteValue HKEY_CLASSES_ROOT, "HexEditor VB.hescr", "", "Script Hex Editor VB", REG_SZ
        .CreateKey HKEY_CLASSES_ROOT, "HexEditor VB.hescr\DefaultIcon"
        .WriteValue HKEY_CLASSES_ROOT, "HexEditor VB.hescr\DefaultIcon", "", App.Path & "\Other\hescr.ico", REG_SZ
        .CreateKey HKEY_CLASSES_ROOT, "HexEditor VB.hescr\Shell\Modifier avec HexEditor VB\Command"
        .WriteValue HKEY_CLASSES_ROOT, "HexEditor VB.hescr\Shell\Modifier avec HexEditor VB\Command", "", """" & App.Path & "\HexEditor.exe" & """" & " " & """" & " %script" & """", REG_SZ
        .WriteValue HKEY_CLASSES_ROOT, "HexEditor VB.hescr\Shell", "", "Modifier avec Hex Editor VB", REG_SZ
        .CreateKey HKEY_CLASSES_ROOT, ".hescr"
        .WriteValue HKEY_CLASSES_ROOT, ".hescr", "", "HexEditor VB.hescr", REG_SZ
        .CreateKey HKEY_CLASSES_ROOT, "HexEditor VB.hescr\Shell\Exécuter\Command"
        .WriteValue HKEY_CLASSES_ROOT, "HexEditor VB.hescr\Shell\Exécuter\Command", "", """" & App.Path & "\HexEditor.exe" & """" & " " & """" & " %script" & """", REG_SZ
        .WriteValue HKEY_CLASSES_ROOT, "HexEditor VB.hescr\Shell", "", "Exécuter", REG_SZ
    End With
    
    'libère la classe
    Set cReg = Nothing

End Sub

'=======================================================
'fonction pour le subclassing (utilisé pour limiter le resize)
'=======================================================
Public Function MaWinProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO
    
    'Intercepte le Message Windows de redimensionnement de fenêtre
    If uMsg = WM_GETMINMAXINFO Then
        CopyMemory MinMax, ByVal lParam, Len(MinMax)
        MinMax.ptMinTrackSize.x = MinX \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.y = MinY \ Screen.TwipsPerPixelY

        CopyMemory ByVal lParam, MinMax, Len(MinMax)
        'Code de retour pour signaler à Windows que le traitement s'est correctement effectué
        MaWinProc = 1
        Exit Function
    End If
    
    'Laisse les autres Messages à traiter à Windows
    MaWinProc = CallWindowProc(AdressWinProc, hWnd, uMsg, wParam, lParam)
End Function

'=======================================================
'limitation du resize d'une form
'=======================================================
Public Function LoadResizing(ByRef hWnd As Long, ByRef MinWidth As Single, ByRef MinHeight As Single)
    MinX = MinWidth
    MinY = MinHeight
    AdressWinProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MaWinProc)
End Function

'=======================================================
'désubclasse
'=======================================================
Public Function RestoreResizing(ByRef hWnd As Long)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AdressWinProc)
End Function

'=======================================================
'récupère la version de Windows
'=======================================================
Public Function GetWindowsVersion(Optional ByRef sWindowsVersion As String, Optional ByRef lBuildNumber As Long) As WINDOWS_VERSION
Dim OS As OSVERSIONINFO
Dim s As String, l As Long

    'taille de la structure
    OS.dwOSVersionInfoSize = Len(OS)
    
    'récupère l'info sur la version
    If GetVersionEx(OS) = 0 Then
        'échec
        sWindowsVersion = "Cannot retrieve information"
        GetWindowsVersion = UnKnown_OS
        Exit Function
    End If
        
    'numéro de la build
    lBuildNumber = OS.dwBuildNumber
    
    'récupère la version en fonction de Major et Minor
    Select Case OS.dwMajorVersion
        Case 6
            GetWindowsVersion = [Windows Vista]
            sWindowsVersion = "Windows Vista"
            Exit Function
        Case 5
            If OS.dwMinorVersion = 2 Then
                GetWindowsVersion = [Windows Server 2003]
                sWindowsVersion = "Windows Server 2003"
            ElseIf OS.dwMinorVersion = 1 Then
                GetWindowsVersion = [Windows XP]
                sWindowsVersion = "Windows XP"
            ElseIf OS.dwMinorVersion = 0 Then
                GetWindowsVersion = [Windows 2000]
                sWindowsVersion = "Windows 2000"
            End If
            Exit Function
        Case 4
            If OS.dwMinorVersion = 90 Then
                GetWindowsVersion = [Windows Me]
                sWindowsVersion = "Windows ME"
            ElseIf OS.dwMinorVersion = 10 Then
                GetWindowsVersion = [Windows 98]
                sWindowsVersion = "Windows 98"
            ElseIf OS.dwMinorVersion = 0 Then
                GetWindowsVersion = [Windows 95]
                sWindowsVersion = "Windows 95"
            End If
            Exit Function
    End Select
    
    GetWindowsVersion = [UnKnown_OS]
    
End Function

'=======================================================
'récupère le nom de l'utilisateur
'=======================================================
Public Function GetUserName() As String
Dim strS As String
Dim Ret As Long

    'créé un buffer
    strS = String$(200, 0)
    
    'récupère le Name
    Ret = GetUserNameA(strS, 199)
    If Ret <> 0 Then GetUserName = Left$(strS, 199) Else GetUserName = vbNullString
End Function

'=======================================================
'transforme une string (date) en currency
'=======================================================
Public Function DateString2Currency(ByVal sDate As String) As Currency
Dim FT As FILETIME
Dim d As Date
Dim t As Date
Dim ST As SYSTEMTIME

    d = DateValue(sDate)
    t = TimeValue(sDate)

    'transfome déjà en systemtime
    With ST
        .wDay = Day(d)
        .wMonth = Month(d)
        .wYear = Year(d)
        .wMinute = Minute(t)
        .wHour = Hour(t)
        .wSecond = Second(t)
    End With
    
    'passe en filetime
    SystemTimeToFileTime ST, FT
    
    'passe en heure locale
    FileTimeToLocalFileTime FT, FT
    
    'passe en currency
    DateString2Currency = FT2Currency(FT)
        
End Function

'=======================================================
'enable ou non les fleches Signet suivant/précédent
'=======================================================
Public Sub RefreshBookMarkEnabled()
    With frmContent
        .Toolbar1.Buttons.Item(16).Enabled = (.ActiveForm.HW.NumberOfSignets > 0)
        .Toolbar1.Buttons.Item(17).Enabled = .Toolbar1.Buttons.Item(16).Enabled
        .mnuSignetNext.Enabled = .Toolbar1.Buttons.Item(16).Enabled
        .mnuSignetPrev.Enabled = .Toolbar1.Buttons.Item(16).Enabled
    End With
End Sub

'=======================================================
'création d'un fichier depuis la sélection dans la activeform
'=======================================================
Public Sub CreateFileFromCurrentSelection(ByVal lCreateNewFileOrNot As Long)
'créé un fichier depuis la sélection
Dim x As Long
Dim y As Long
Dim s2 As String
Dim s() As String
Dim sFile As String
Dim curBuf As Currency
Dim curPos2 As Currency
Dim curSize2 As Currency
Dim curSize As Currency
Dim curPos As Currency
Dim lLastBufSize As Long
Dim lSect As Long
Dim bOverWrite As Boolean

    'oui ou non ou fait à la suite d'un fichier
    bOverWrite = (lCreateNewFileOrNot = vbYes)
    
    If frmContent.ActiveForm Is Nothing Then Exit Sub
    
    On Error GoTo CancelPushed
    
    frmContent.Sb.Panels(1).Text = "Status=[Creating file from selection]"
    
    'détermine la taille
    curSize = frmContent.ActiveForm.HW.SecondSelectionItem.Offset + frmContent.ActiveForm.HW.SecondSelectionItem.Col - _
        frmContent.ActiveForm.HW.FirstSelectionItem.Offset - frmContent.ActiveForm.HW.FirstSelectionItem.Col + 1
    
    'détermine la position du premier offset
    curPos = frmContent.ActiveForm.HW.FirstSelectionItem.Offset + frmContent.ActiveForm.HW.FirstSelectionItem.Col - 1
        
    With frmContent.CMD
        .CancelError = True
        .DialogTitle = "Sélection du fichier à sauvegarder"
        .Filter = "Tous|*.*"
        .ShowSave
        sFile = .Filename
    End With
    
    If cFile.FileExists(sFile) And bOverWrite Then
        'fichier déjà existant
        If MsgBox("Le fichier existe déjà. Le remplacer ?", vbInformation + vbYesNo, "Attention") <> vbYes Then Exit Sub
    End If
    
    'ajoute du texte à la console
    Call AddTextToConsole("Création du fichier...")
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile et sauvegarde à la volée (buffers de 500Ko)
            
            If curSize <= 512000 Then
                'alors tout rentre dans un buffer
                'récupère la string
                s2 = GetBytesFromFile(frmContent.ActiveForm.Caption, curSize, curPos)
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
            curPos2 = ByND(curPos, frmContent.ActiveForm.GetDriveInfos.BytesPerSector)
            curSize2 = frmContent.ActiveForm.HW.SecondSelectionItem.Offset + frmContent.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, frmContent.ActiveForm.GetDriveInfos.BytesPerSector)
            
            If curSize2 <= frmContent.ActiveForm.GetDriveInfos.BytesPerSector Then
                'alors tout rentre dans un buffer (de la taille d'un secteur)
                'récupère la string
                DirectReadS frmContent.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / frmContent.ActiveForm.GetDriveInfos.BytesPerSector, CLng(curSize2), _
                    frmContent.ActiveForm.GetDriveInfos.BytesPerSector, s2
                    
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
    cFile.SaveDATAinFile sFile, s2, bOverWrite   'lance la sauvegarde

    'ajoute du texte à la console
    Call AddTextToConsole("Le fichier a été créé")
    
    GoTo CancelPushed
    
CreateMyFileFromBuffers:
    'sauvegarde le fichier (plusieurs buffers)
    
    'commence par créer un fichier vierge
    cFile.CreateEmptyFile sFile, bOverWrite
    
    Select Case TypeOfForm(frmContent.ActiveForm)
        Case "Fichier"
            'édition d'un fichier ==> va piocher avec ReadFile et sauvegarde à la volée (buffers de 500Ko)

            'détermine le nombre de buffers à utiliser
            curBuf = Int(curSize / 512000) + IIf(Mod2(curSize, 512000) = 0, 0, 1)
            
            'détermine la taille du dernier buffer
            lLastBufSize = curSize - (curBuf - 1) * 512000
            
            'récupère la string pour chaque buffer <> du dernier
            For x = 1 To curBuf - 1
                
                'récupère la string
                s2 = GetBytesFromFile(frmContent.ActiveForm.Caption, 512000, curPos + 512000 * (x - 1) + 1)
                
                'sauve le morceau à la fin du fichier
                WriteBytesToFileEnd sFile, s2
            Next x

            's'occupe du dernier buffer
            s2 = GetBytesFromFile(frmContent.ActiveForm.Caption, lLastBufSize, curPos + 512000 * (curBuf - 1) + 1)
            
            'sauvegarde la string
            WriteBytesToFileEnd sFile, s2
        
        Case "Processus"
            'sauvegarde avec un buffer de 50Ko
            
            'détermine le nombre de buffers à utiliser
            curBuf = Int(curSize / 512000) + IIf(Mod2(curSize, 512000) = 0, 0, 1)
            
            'détermine la taille du dernier buffer
            lLastBufSize = curSize - (curBuf - 1) * 512000
            
            'récupère la string pour chaque buffer <> du dernier
            For x = 1 To curBuf - 1
            
                'récupère la string
                s2 = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos + 512000 * (x - 1) + 1), CLng(512000))
                
                'sauve le morceau à la fin du fichier
                WriteBytesToFileEnd sFile, s2
            Next x

            's'occupe du dernier buffer
            s2 = cMem.ReadBytes(Val(frmContent.ActiveForm.Tag), CLng(curPos + 512000 * (curBuf - 1) + 1), CLng(512000))
            
            'sauvegarde la string
            WriteBytesToFileEnd sFile, s2
            
        Case "Disque"
            'sauvegarde avec un buffer de frmContent.ActiveForm.GetDriveInfos.BytesPerSector octets
            
            'bytes par secteur
            lSect = frmContent.ActiveForm.GetDriveInfos.BytesPerSector
            
            'redéfinit correctement la position et la taille (doivent être multiple du nombre
            'de bytes par secteur)
            curPos2 = ByND(curPos, lSect)
            curSize2 = frmContent.ActiveForm.HW.SecondSelectionItem.Offset + frmContent.ActiveForm.HW.SecondSelectionItem.Col - _
                curPos2  'recalcule la taille en partant du début du secteur
            curSize2 = ByN(curSize2, lSect)

            'détermine le nombre de buffers à utiliser
            curBuf = Int(curSize / (lSect * 1000)) + IIf(Mod2(curSize, (lSect * 1000)) = 0, 0, 1)
            
            'détermine la taille du dernier buffer
            lLastBufSize = curSize - (curBuf - 1) * (lSect * 1000)
            
            For x = 1 To curBuf - 1
                
                'récupère la string
                DirectReadS frmContent.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / lSect + (x - 1) * 1000, CLng(curSize2), lSect, s2
                
                'recoupe la string pour récupérer ce qui intéresse vraiment
                s2 = Mid$(s2, curPos - curPos2 + 1, (lSect * 1000))
            
                'écrit dans le fichier (à la fin)
                WriteBytesToFileEnd sFile, s2
            Next x
            
            'maintenant on s'occupe du dernier morceau de fichier
            DirectReadS frmContent.ActiveForm.GetDriveInfos.VolumeLetter & ":\", _
                    curPos2 / lSect + (curBuf - 1) * 1000, CLng(curSize2), lSect, s2
                    
            'recoupe la string pour récupérer ce qui intéresse vraiment
            s2 = Mid$(s2, curPos - curPos2 + 1, lLastBufSize)
            
            'écrit dans le fichier
            WriteBytesToFileEnd sFile, s2
            
    End Select

    'ajoute du texte à la console
    Call AddTextToConsole("Le fichier a été créé")
    
CancelPushed:
    
    frmContent.Sb.Panels(1).Text = "Status=[Ready]"
End Sub

'=======================================================
'récupère une string contenant l'erreur pointée par hError
'=======================================================
Public Function GetError(ByVal hError As Long) As String
Dim Buffer As String
    
    'buffer
    Buffer = Space$(1024)
    
    'récupère la string
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, hError, LANG_NEUTRAL, Buffer, Len(Buffer), ByVal 0&
    GetError = Trim$(Buffer)
    
End Function












'=======================================================
'FONCTIONS DE CONVERSION INTER-BASES
'=======================================================
Public Function Str2Hex(ByVal s As String) As String
    Str2Hex = Hex$(Str2Dec(s))
End Function
Public Function Str2Hex_(ByVal s As String) As String
    Str2Hex_ = Hex$(Str2Dec(s))
    If Len(Str2Hex_) = 1 Then Str2Hex_ = "0" & Str2Hex_
End Function
Public Function Str2Dec(ByVal s As String) As Long
    If s = vbNullString Then Exit Function
    Str2Dec = Asc(s)
End Function
Public Function Str2Oct(ByVal s As String) As String
    Str2Oct = Oct$(Str2Dec(s))
End Function
Public Function Hex2Dec(ByVal s As String) As Long
Dim x As Long
Dim l As Long

    For x = Len(s) To 1 Step -1
        l = l + HexVal(Mid$(s, Len(s) - x + 1, 1)) * AexpB(16, x - 1)
    Next x

    Hex2Dec = l
End Function
Public Function Hex2Str(ByVal s As String) As String
    Hex2Str = Byte2FormatedString(Hex2Dec(s))
End Function
Public Function Hex2Str_(ByVal s As String) As String
    Hex2Str_ = Chr$(Hex2Dec(s))
End Function
Public Function Hex2Oct(ByVal s As String) As String
    Hex2Oct = Oct$(Hex2Dec(s))
End Function
Public Function Dec2Bin(ByVal l As Long, Optional ByVal lSize As Long = 8) As String
Dim x As Long
Dim s As String

    s = vbNullString

    For x = lSize - 1 To 0 Step -1
        If l >= AexpB(2, x) Then
            l = l - AexpB(2, x)
            s = s & "1"
        Else
            s = s & "0"
        End If
    Next x
    
    Dec2Bin = s
        
End Function
Public Function Bin2Dec(ByVal s As String) As Long
Dim x As Long
Dim l As Long

    For x = Len(s) To 1 Step -1
        l = l + FormatedVal(Mid$(s, Len(s) - x + 1, 1)) * AexpB(2, x - 1)
    Next x

    Bin2Dec = l
End Function
Public Function Oct2Dec(ByVal s As String) As Long
Dim x As Long
Dim l As Long

    For x = Len(s) To 1 Step -1
        l = l + FormatedVal(Mid$(s, Len(s) - x + 1, 1)) * AexpB(8, x - 1)
    Next x

    Oct2Dec = l
End Function
Public Function FormatedVal(ByVal s As String) As Long
    On Error Resume Next
    FormatedVal = Abs(Int(Val(s)))
End Function
Public Function FormatedVal_(ByVal s As String) As Currency
    On Error Resume Next
    FormatedVal_ = Abs(Int(Val(s)))
End Function
Public Function HexVal(ByVal s As String) As Long
    If s = "0" Then
        HexVal = 0
    ElseIf s = "1" Then
        HexVal = 1
    ElseIf s = "2" Then
        HexVal = 2
    ElseIf s = "3" Then
        HexVal = 3
    ElseIf s = "4" Then
        HexVal = 4
    ElseIf s = "5" Then
        HexVal = 5
    ElseIf s = "6" Then
        HexVal = 6
    ElseIf s = "7" Then
        HexVal = 7
    ElseIf s = "8" Then
        HexVal = 8
    ElseIf s = "9" Then
        HexVal = 9
    ElseIf LCase(s) = "a" Then
        HexVal = 10
    ElseIf LCase(s) = "b" Then
        HexVal = 11
    ElseIf LCase(s) = "c" Then
        HexVal = 12
    ElseIf LCase(s) = "d" Then
        HexVal = 13
    ElseIf LCase(s) = "e" Then
        HexVal = 14
    ElseIf LCase(s) = "f" Then
        HexVal = 15
    End If
End Function
Public Function ExtendedHex(ByVal cVal As Currency) As String
Dim x As Long
Dim s As String
Dim table16(15) As Currency
Dim res(15) As Byte

    cVal = cVal + 1 'ajoute 1 pour que le résultat soit juste

    'contient la table des 16^n
    table16(0) = 1
    table16(1) = 16
    table16(2) = 256
    table16(3) = 4096
    table16(4) = 65536
    table16(5) = 1048576
    table16(6) = 16777216
    table16(7) = 268435456
    table16(8) = 4294967296#
    table16(9) = 68719476736#
    table16(10) = 1099511627776#
    table16(11) = 17592186044416#
    table16(12) = 281474976710656#

    'enlève, en partant des plus grosses valeurs, un maximum de fois un 16^x
    For x = 12 To 0 Step -1
        While cVal > table16(x)
            cVal = cVal - table16(x)
            res(x) = res(x) + 1 'ajoute 1 à l'occurence de table16(x)
        Wend
    Next x
    
    'créé la string
    For x = 12 To 0 Step -1
        s = s & Hex(res(x))
    Next
    
    ExtendedHex = s
End Function
'=======================================================
'fonction qui transforme une suite de valeur hexa en une string
'=======================================================
Public Function HexValues2String(ByVal sString As String) As String
Dim Sep As Boolean
Dim sRes As String
Dim x As Long

    Sep = True  'recherche un séparant entre les valeurs hexa (de longueur 2)

    While Sep
        If Len(sString) > 2 Then
            'alors on recherche un éventuel séparant
            If Val("&h" & Mid$(sString, 3, 1)) = 0 And Mid$(sString, 3, 1) <> "0" Then
                'alors le troisième caractère n'est pas un caractère qui compose une valeur hexa
                'donc c'est un séparant
                sString = Replace$(sString, Mid$(sString, 3, 1), vbNullString) 'vire tous les séparants
            Else
                'alors pas de séparant ==> on quitte la boucle
                Sep = False
            End If
        Else
            Sep = False
        End If
    Wend
    
    sRes = vbNullString
    'maintenant que la string ne comporte plus de séparants, on créé le résultat
    For x = 1 To Int(Len(sString) / 2)
        sRes = sRes & Chr$(Hex2Dec(Mid$(sString, 2 * x - 1, 2)))
    Next x
    
    HexValues2String = sRes
End Function
