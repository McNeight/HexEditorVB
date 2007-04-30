Attribute VB_Name = "mdlPass"
' =======================================================
'
' Hex Editor VB
' Coded by violent_ken (Alain Descotes)
'
' =======================================================
'
' A complete hexadecimal editor for Windows �
' (Editeur hexad�cimal complet pour Windows �)
'
' Copyright � 2006-2007 by Alain Descotes.
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
'//MODULE DE GESTION DU REMPLISSAGE PAR PASSES
'=======================================================


'=======================================================
'effectue le changement sur un fichier
'=======================================================
Public Sub ApplyPass_File(ByVal curPos1 As Currency, ByVal curPos2 As Currency, _
    ByVal HW As HexViewer, tP() As PASSE_TYPE, ByVal sFile As String)
    
Dim s1 As String
Dim y As Long
Dim x As Long
Dim z As Long
Dim s2 As String
Dim s As String
Dim nb As Long
Dim sUnik As String
Dim c2 As Currency
Dim c1 As Currency
Dim Sb() As String
Dim Frm As Form
Dim l1 As Long
Dim l2 As Long
Dim Frm2 As Form
    
    'on remplit pour chaque passe en temporaire, cad dans la liste des modifs du HW
    'on effectue les �critures par 16 bytes
    'on r�cup�re donc les premi�re et derni�re lignes de 16 pour les compl�ter
    'par la string actuelle

    'on d�termine l'offset (arrondi � 16 dessous) de la premi�re s�lection
    c1 = By16D(curPos1)
    'r�cup�re la 16-string de cet offset
    s1 = GetBytesFromFile(sFile, 16, c1)
            
    'de la derni�re
    c2 = By16D(curPos2)
    s2 = GetBytesFromFile(sFile, 16, c2)
    
    'd�termine le nombre de 16-string (sans premi�re et derni�re)
    nb = (c2 - c1) / 16
    
    'cr�� un pseudo hasard
    Call Randomize
    
    For x = 0 To UBound(tP()) - 1
        
        '//EFFECTUE LES REMLISSAGES
        If tP(x).tType = FixedByte Then
            
            'lance la sauvegarde dans le fichier
            Call WriteBytesToFile(sFile, String$(curPos2 - curPos1 + 1, _
                Hex2Dec(tP(x).sData1)), curPos1)
            
            'on ouvre ce nouveau fichier
            If x = UBound(tP()) - 1 Then
                'alors c'�tait la derni�re passe
            
                Set Frm = New Pfm
                Set Frm2 = frmContent.ActiveForm    'r�cup�re la form actuelle
                
                Call Frm.GetFile(sFile)
                Frm.Show
                lNbChildFrm = lNbChildFrm + 1
                DoEvents    '/!\ IMPORTANT DO NOT REMOVE
                frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
                With frmContent.ActiveForm
                    'r�cup�re les signets
                    For y = 1 To Frm2.lstSignets.ListItems.Count
                        .lstSignets.ListItems.Add _
                            Text:=Frm2.lstSignets.ListItems.Item(y)
                        .lstSignets.ListItems.Item(y).SubItems(1) = _
                            Frm2.lstSignets.ListItems.Item(y).SubItems(1)
                        .HW.AddSignet Val(Frm2.lstSignets.ListItems.Item(y))
                    Next y
                End With
                
                'd�charge l'autre form
                Set Frm2 = Nothing
                
                'refresh (signets)
                Call frmContent.ActiveForm.HW.TraceSignets
                
            End If
            
            DoEvents
            
        ElseIf tP(x).tType = ListByte Then
        
            'r�cup�re une liste des bytes possibles (en string)
            Sb() = Split(tP(x).sData1, " ", , vbBinaryCompare)
            
            For y = 0 To UBound(Sb())
                Sb(y) = Hex2Dec(Sb(y))
            Next y
            
            'effectue des calculs une bonne fois pour toutes
            l1 = (UBound(Sb()) + 1)
            
            'on cr�� une string al�atoire avec un byte compris dans la liste des bytes
            'possibles
            s = vbNullString
            For y = 1 To curPos2 - curPos1 + 1
                z = Int(Rnd * l1)
                s = s & Chr_(Val(Sb(z)))
                If (y Mod 50000) = 0 Then DoEvents  'rend la main de tps en tps
            Next y
        
            'lance la sauvegarde dans le fichier
            Call WriteBytesToFile(sFile, s, curPos1)
            
            'on ouvre ce nouveau fichier
            If x = UBound(tP()) - 1 Then
                'alors c'�tait la derni�re passe
            
                Set Frm = New Pfm
                Set Frm2 = frmContent.ActiveForm    'r�cup�re la form actuelle
                
                Call Frm.GetFile(sFile)
                Frm.Show
                lNbChildFrm = lNbChildFrm + 1
                DoEvents    '/!\ IMPORTANT DO NOT REMOVE
                frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
                'r�cup�re les signets
                With frmContent.ActiveForm
                    For y = 1 To Frm2.lstSignets.ListItems.Count
                        .lstSignets.ListItems.Add _
                            Text:=Frm2.lstSignets.ListItems.Item(y)
                        .lstSignets.ListItems.Item(y).SubItems(1) = _
                            Frm2.lstSignets.ListItems.Item(y).SubItems(1)
                        .HW.AddSignet Val(Frm2.lstSignets.ListItems.Item(y))
                    Next y
                End With
                
                'd�charge l'autre form
                Set Frm2 = Nothing
                
                'refresh (signets)
                Call frmContent.ActiveForm.HW.TraceSignets
                
            End If
            
            DoEvents
            
        ElseIf tP(x).tType = RandomByte Then
        
            'lance la sauvegarde dans le fichier
            
            'fait des calculs une bonne fois pour toutes
            l1 = 1 + Hex2Dec(tP(x).sData2) - Hex2Dec(tP(x).sData1)
            l2 = Hex2Dec(tP(x).sData1)
            
            'cr�� une string al�atoire (valeurs comprises entre deux bornes)
            s = CreateRandomString(Hex2Dec(tP(x).sData1), Hex2Dec(tP(x).sData2), _
                curPos2 - curPos1 + 1)
            
            Call WriteBytesToFile(sFile, s, curPos1)
            
            'on ouvre ce nouveau fichier
            If x = UBound(tP()) - 1 Then
                'alors c'�tait la derni�re passe
            
                Set Frm = New Pfm
                Set Frm2 = frmContent.ActiveForm    'r�cup�re la form actuelle
                
                Call Frm.GetFile(sFile)
                Frm.Show
                lNbChildFrm = lNbChildFrm + 1
                DoEvents    '/!\ IMPORTANT DO NOT REMOVE
                frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
                'r�cup�re les signets
                With frmContent.ActiveForm
                    For y = 1 To Frm2.lstSignets.ListItems.Count
                        .lstSignets.ListItems.Add _
                            Text:=Frm2.lstSignets.ListItems.Item(y)
                        .lstSignets.ListItems.Item(y).SubItems(1) = _
                            Frm2.lstSignets.ListItems.Item(y).SubItems(1)
                        .HW.AddSignet Val(Frm2.lstSignets.ListItems.Item(y))
                    Next y
                End With
                
                'd�charge l'autre form
                Set Frm2 = Nothing
                
                'refresh (signets)
                Call frmContent.ActiveForm.HW.TraceSignets
                
            End If
            
            DoEvents
            
        End If
    
    Next x
    
End Sub

'=======================================================
'r�cup�re une string al�atoire de 2 Mo de long
'=======================================================
Public Function GetRandom2MoString() As String
Dim pt As Long
Dim s As String
Dim cASM As CAsmProc
Dim Tbl(2097151) As Byte

    Set cASM = New CAsmProc

    'r�cup�re le pointeur sur la string
    Call cASM.bnAlloc2MoAlea(Tbl(0))
    
    'cr�� un buffer
    s = Space$(2097152)
    
    'copie les bytes du tableau sur la stirng
    Call CopyMemory(ByVal s, Tbl(0), 2097152)
    
    Set cASM = Nothing
    
    GetRandom2MoString = s
End Function

'=======================================================
'renvoie une string al�atoire
'contenant des bytes allant de lBG � lBD compris
'string de longueur lSize en r�sultat
'=======================================================
Public Function CreateRandomString(ByVal lBg As Long, ByVal lBD As Long, _
    ByVal lSize As Long) As String

Dim cString As clsString
Dim x As Long
Dim lByte As Long
Dim l1 As Long

    'cr�� un pseudo hasard
    Call Randomize

    'instancie la classe
    Set cString = New clsString
    
    'initialise la valeur de la string
    cString.Value = vbNullString
    
    'pr�-calcule cette addition
    l1 = 1 + lBD - lBg
    
    'pour chaque caract�re
    For x = 1 To lSize
    
        'tire une valeur au hasard dans l'intervalle
        lByte = Int(Rnd * l1) + lBg
        
        'concat�re avec la string
        Call cString.Append(Chr_(lByte))
    
    Next x
    
    'renvoie le r�sultat
    CreateRandomString = cString.Value
    
    'lib�re la classe
    Set cString = Nothing
    
End Function
