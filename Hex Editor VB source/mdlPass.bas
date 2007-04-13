Attribute VB_Name = "mdlPass"
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
'//MODULE DE GESTION DU REMPLISSAGE PAR PASSES
'=======================================================


'=======================================================
'effectue le changement sur un fichier
'=======================================================
Public Sub ApplyPass_File(ByVal curPos1 As Currency, ByVal curPos2 As Currency, _
    ByVal HW As HexViewer, tP() As PASSE_TYPE, ByVal sFile As String)
    
Dim s1 As String
Dim y As Long
Dim X As Long
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
    'on effectue les écritures par 16 bytes
    'on récupère donc les première et dernière lignes de 16 pour les compléter
    'par la string actuelle

    'on détermine l'offset (arrondi à 16 dessous) de la première sélection
    c1 = By16D(curPos1)
    'récupère la 16-string de cet offset
    s1 = GetBytesFromFile(sFile, 16, c1)
            
    'de la dernière
    c2 = By16D(curPos2)
    s2 = GetBytesFromFile(sFile, 16, c2)
    
    'détermine le nombre de 16-string (sans première et dernière)
    nb = (c2 - c1) / 16
    
    'créé un pseudo hasard
    Randomize
    
    For X = 0 To UBound(tP()) - 1
        
        '//EFFECTUE LES REMLISSAGES
        If tP(X).tType = FixedByte Then
            
            'lance la sauvegarde dans le fichier
            Call WriteBytesToFile(sFile, String$(curPos2 - curPos1 + 1, Hex2Dec(tP(X).sData1)), curPos1)
            
            'on ouvre ce nouveau fichier
            If X = UBound(tP()) - 1 Then
                'alors c'était la dernière passe
            
                Set Frm = New Pfm
                Set Frm2 = frmContent.ActiveForm    'récupère la form actuelle
                
                Call Frm.GetFile(sFile)
                Frm.Show
                lNbChildFrm = lNbChildFrm + 1
                DoEvents    '/!\ IMPORTANT DO NOT REMOVE
                frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
                'récupère les signets
                For y = 1 To Frm2.lstSignets.ListItems.Count
                    frmContent.ActiveForm.lstSignets.ListItems.Add _
                        Text:=Frm2.lstSignets.ListItems.Item(y)
                    frmContent.ActiveForm.lstSignets.ListItems.Item(y).SubItems(1) = _
                        Frm2.lstSignets.ListItems.Item(y).SubItems(1)
                    frmContent.ActiveForm.HW.AddSignet Val(Frm2.lstSignets.ListItems.Item(y))
                Next y
                
                'décharge l'autre form
                Set Frm2 = Nothing
                
                'refresh (signets)
                Call frmContent.ActiveForm.HW.TraceSignets
                
            End If
            
            DoEvents
            
        ElseIf tP(X).tType = ListByte Then
        
            'récupère une liste des bytes possibles (en string)
            Sb() = Split(tP(X).sData1, " ", , vbBinaryCompare)
            
            For y = 0 To UBound(Sb())
                Sb(y) = Hex2Dec(Sb(y))
            Next y
            
            'effectue des calculs une bonne fois pour toutes
            l1 = (UBound(Sb()) + 1)
            
            'on créé une string aléatoire avec un byte compris dans la liste des bytes
            'possibles
            s = vbNullString
            For y = 1 To curPos2 - curPos1 + 1
                z = Int(Rnd * l1)
                s = s & Chr$(Val(Sb(z)))
                If (y Mod 50000) = 0 Then DoEvents  'rend la main de tps en tps
            Next y
        
            'lance la sauvegarde dans le fichier
            Call WriteBytesToFile(sFile, s, curPos1)
            
            'on ouvre ce nouveau fichier
            If X = UBound(tP()) - 1 Then
                'alors c'était la dernière passe
            
                Set Frm = New Pfm
                Set Frm2 = frmContent.ActiveForm    'récupère la form actuelle
                
                Call Frm.GetFile(sFile)
                Frm.Show
                lNbChildFrm = lNbChildFrm + 1
                DoEvents    '/!\ IMPORTANT DO NOT REMOVE
                frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
                'récupère les signets
                For y = 1 To Frm2.lstSignets.ListItems.Count
                    frmContent.ActiveForm.lstSignets.ListItems.Add _
                        Text:=Frm2.lstSignets.ListItems.Item(y)
                    frmContent.ActiveForm.lstSignets.ListItems.Item(y).SubItems(1) = _
                        Frm2.lstSignets.ListItems.Item(y).SubItems(1)
                    frmContent.ActiveForm.HW.AddSignet Val(Frm2.lstSignets.ListItems.Item(y))
                Next y
                
                'décharge l'autre form
                Set Frm2 = Nothing
                
                'refresh (signets)
                Call frmContent.ActiveForm.HW.TraceSignets
                
            End If
            
            DoEvents
            
        ElseIf tP(X).tType = RandomByte Then
        
            'lance la sauvegarde dans le fichier
            
            'fait des calculs une bonne fois pour toutes
            l1 = 1 + Hex2Dec(tP(X).sData2) - Hex2Dec(tP(X).sData1)
            l2 = Hex2Dec(tP(X).sData1)
            
            'créé une string aléatoire (valeurs comprises entre deux bornes)
            s = CreateRandomString(Hex2Dec(tP(X).sData1), Hex2Dec(tP(X).sData2), _
                curPos2 - curPos1 + 1)
            
            Call WriteBytesToFile(sFile, s, curPos1)
            
            'on ouvre ce nouveau fichier
            If X = UBound(tP()) - 1 Then
                'alors c'était la dernière passe
            
                Set Frm = New Pfm
                Set Frm2 = frmContent.ActiveForm    'récupère la form actuelle
                
                Call Frm.GetFile(sFile)
                Frm.Show
                lNbChildFrm = lNbChildFrm + 1
                DoEvents    '/!\ IMPORTANT DO NOT REMOVE
                frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"
                
                'récupère les signets
                For y = 1 To Frm2.lstSignets.ListItems.Count
                    frmContent.ActiveForm.lstSignets.ListItems.Add _
                        Text:=Frm2.lstSignets.ListItems.Item(y)
                    frmContent.ActiveForm.lstSignets.ListItems.Item(y).SubItems(1) = _
                        Frm2.lstSignets.ListItems.Item(y).SubItems(1)
                    frmContent.ActiveForm.HW.AddSignet Val(Frm2.lstSignets.ListItems.Item(y))
                Next y
                
                'décharge l'autre form
                Set Frm2 = Nothing
                
                'refresh (signets)
                Call frmContent.ActiveForm.HW.TraceSignets
                
            End If
            
            DoEvents
            
        End If
    
    Next X
    
End Sub

'=======================================================
'renvoie une string aléatoire
'contenant des bytes allant de lBG à lBD compris
'string de longueur lSize en résultat
'=======================================================
Public Function CreateRandomString(ByVal lBG As Long, ByVal lBD As Long, _
    ByVal lSize As Long) As String

Dim cString As clsString
Dim X As Long
Dim lByte As Long
Dim l1 As Long

    'créé un pseudo hasard
    Call Randomize

    'instancie la classe
    Set cString = New clsString
    
    'initialise la valeur de la string
    cString.Value = vbNullString
    
    'pré-calcule cette addition
    l1 = 1 + lBD - lBG
    
    'pour chaque caractère
    For X = 1 To lSize
    
        'tire une valeur au hasard dans l'intervalle
        lByte = Int(Rnd * l1) + lBG
        
        'concatère avec la string
        Call cString.Append(Chr$(lByte))
    
    Next X
    
    'renvoie le résultat
    CreateRandomString = cString.Value
    
    'libère la classe
    Set cString = Nothing
    
End Function
