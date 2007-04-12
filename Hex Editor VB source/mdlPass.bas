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
Dim x As Long
Dim z As Long
Dim s2 As String
Dim s As String
Dim nb As Long
Dim sUnik As String
Dim c2 As Currency
Dim c1 As Currency
    
    'on remplit pour chaque passe en temporaire, cad dans la liste des modifs du HW
    For x = 0 To UBound(tP()) - 1
    
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
        
        '//EFFECTUE LES REMLISSAGES
        If tP(x).tType = FixedByte Then
            
            'lance la sauvegarde dans le fichier
            Call WriteBytesToFile(sFile, String$(curPos2 - curPos1 + 1, Hex2Dec(tP(x).sData1)), curPos1)
            
            'on ouvre ce nouveau fichier
            Dim Frm As Form
            Set Frm = New Pfm
            Call Frm.GetFile(sFile)
            Frm.Show
            lNbChildFrm = lNbChildFrm + 1
            DoEvents    '/!\ IMPORTANT DO NOT REMOVE
            frmContent.Sb.Panels(2).Text = frmContent.Lang.GetString("_Openings") & CStr(lNbChildFrm) & "]"

            
            '//premiere 16-string
            'la nouvelle première string
           ' s1 = Left$(s1, 16 - (By16(curPos1) - curPos1)) & _
                String$(By16(curPos1) - curPos1, sUnik)
            
            'ajoute les 16 changements
            'For y = 1 To 16
            '    Call frmContent.ActiveForm.AddChange(c1, 1, s1)
            'Next y
            
            '//dernière 16-string
            's2 = String$(By16(curPos2) - curPos2, sUnik) & _
                Right$(s2, 16 - (By16(curPos2) - curPos2))
            
            'ajoute les 16 changements
            'For y = 1 To 16
             '   Call frmContent.ActiveForm.AddChange(c2, 1, s2)
            'Next y
            
            '//les nb autres
        '    For z = 1 To nb
                 
                'ajoute les 16 changements
                'For y = 1 To 16
            '        Call frmContent.ActiveForm.AddChange(c1 + 16 * z, 1, String$(16, sUnik))
                'Next y
           ' Next z
            
        ElseIf tP(x).tType = ListByte Then
        
        ElseIf tP(x).tType = RandomByte Then
        
        End If
    
    Next x
    
End Sub

