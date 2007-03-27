Attribute VB_Name = "mdlUndo"
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
'//MODULE DE GESTION DE L'HISTORIQUE
'CONTIENT LES PROCEDURES DE GESTION DE L'HISTORIQUE
'=======================================================

'=======================================================
'effectue un Undo, prend en entrée
'-la liste de l'historique
'=======================================================
Public Sub UndoMe(Undo As clsUndoItem, ByRef Histo() As clsUndoSubItem)
    
    With Histo(Undo.lRang)
        If Undo.tEditType = edtFile Then
            'un fichier
            Select Case .tUndoType
                Case actByteWritten
                    Call Undo.Frm.AddChange(.curData1, .bytData1, .sData1)
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = .curData1
                    Undo.Frm.HW.MaxOffset = .curData2
                    Undo.Frm.HW.Refresh
            End Select
        ElseIf Undo.tEditType = edtDisk Then
            'un disque
            Select Case .tUndoType
                Case actByteWritten
                '/////// A FAIRE
                    Call Undo.Frm.AddChange(.curData1, .bytData1, .sData1)
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = .curData1
                    Undo.Frm.HW.MaxOffset = .curData2
                    Undo.Frm.HW.Refresh
            End Select
        Else
            'un processus
            Select Case .tUndoType
                Case actByteWritten
                    'on récupère PID, Offset et String
                    'on travaille directement avec la classe cMem
                    cMem.WriteBytes .lngData1, CLng(.curData1), .sData2 'Long suffisant pour l'offset (plage de 2Go uniquement)
                    Call Undo.Frm.VS_Change(Undo.Frm.VS.Value)  'refresh
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = .curData1
                    Undo.Frm.HW.MaxOffset = .curData2
                    Undo.Frm.HW.Refresh
            End Select
        End If
    End With
End Sub

'=======================================================
'effectue un Redo, prend en entrée
'-la liste de l'historique
'=======================================================
Public Sub RedoMe(Undo As clsUndoItem, ByRef Histo() As clsUndoSubItem)
    
    With Histo(Undo.lRang)
        If Undo.tEditType = edtFile Then
            'alors c'est un fichier
            Select Case .tUndoType
                Case actByteWritten
                    Call Undo.Frm.AddChange(.curData1, .bytData1, .sData2)
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = Histo(Undo.lRang + 1).curData1
                    Undo.Frm.HW.MaxOffset = Histo(Undo.lRang + 1).curData2
                    Undo.Frm.HW.Refresh
            End Select
        ElseIf Undo.tEditType = edtDisk Then
            'alors c'est un disque
            Select Case .tUndoType
                Case actByteWritten
                '/////// A FAIRE
                    Call Undo.Frm.AddChange(.curData1, .bytData1, .sData2)
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = Histo(Undo.lRang + 1).curData1
                    Undo.Frm.HW.MaxOffset = Histo(Undo.lRang + 1).curData2
                    Undo.Frm.HW.Refresh
            End Select
        ElseIf Undo.tEditType = edtPhys Then
            'alors c'est un disque physique
            Select Case .tUndoType
                Case actByteWritten
                '/////// A FAIRE
                    Call Undo.Frm.AddChange(.curData1, .bytData1, .sData2)
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = Histo(Undo.lRang + 1).curData1
                    Undo.Frm.HW.MaxOffset = Histo(Undo.lRang + 1).curData2
                    Undo.Frm.HW.Refresh
            End Select
        Else
            'ben là c'est un processus
            Select Case .tUndoType
                Case actByteWritten
                    'on récupère PID, Offset et String
                    'on travaille directement avec la classe cMem
                    cMem.WriteBytes .lngData1, CLng(.curData1), .sData1 'Long suffisant pour l'offset (plage de 2Go uniquement)
                    Call Undo.Frm.VS_Change(Undo.Frm.VS.Value)  'refresh
                Case actRestArea
                '/////// A FAIRE
                    'alors il faut redimensionner la zone avec les anciens offsets
                    Undo.Frm.HW.FirstOffset = Histo(Undo.lRang + 1).curData1
                    Undo.Frm.HW.MaxOffset = Histo(Undo.lRang + 1).curData2
                    Undo.Frm.HW.Refresh
            End Select
        End If
    End With
End Sub

'=======================================================
'ajoute une entrée à l'historique, prend en paramètre
'-le rang de la nouvelle entrée (donc supprime tout ce qui est après)
'-l'historique
'-les datas
'-le type de nouvelle entrée
'=======================================================
Public Sub AddHisto(ByVal lRang As Long, ByVal Undo As clsUndoItem, ByRef Histo() As clsUndoSubItem, _
    ByVal tUndoType As UNDO_TYPE, Optional ByVal sData1 As String, Optional ByVal sData2 As String, _
    Optional ByVal curData1 As Currency, Optional ByVal curData2 As Currency, _
    Optional ByVal bytData1 As Byte, Optional ByVal bytData2 As Byte, Optional ByVal lngData1 As Long)
Dim x As Long
Dim y As Long
Dim s As String

    'procède à la suppression de tout ce qui est après lRang
    'lRang=1 ==> pas de suppression
    If lRang <> -1 Then
        ReDim Preserve Histo(lRang)
    End If
    
    'ajoute à la fin de l'historique le nouvel élément
    ReDim Preserve Histo(UBound(Histo()) + 1)
    Set Histo(UBound(Histo())) = New clsUndoSubItem
    With Histo(UBound(Histo()))
        .tUndoType = tUndoType
        .sData1 = sData1
        .sData2 = sData2
        .curData2 = curData2
        .curData1 = curData1
        .bytData2 = bytData2
        .bytData1 = bytData1
        .lngData1 = lngData1
    End With
    
    'ajoute au lv un nouvel item
    Undo.lvHisto.Visible = False
    With Undo.lvHisto.ListItems
        If Undo.tEditType = edtFile Then
            'alors c'est un fichier
            If tUndoType = actByteWritten Then
                s = "o=[" & LTrim$(Str$(curData1)) & "]c=[" & LTrim$(Str$(bytData1)) & "]s=[" & Formated16String(sData2) & "]"
            End If
        ElseIf Undo.tEditType = edtDisk Then
            'alors c'est un disque
            If tUndoType = actByteWritten Then
                s = "o=[" & LTrim$(Str$(curData1)) & "]c=[" & LTrim$(Str$(bytData1)) & "]s=[" & Formated16String(sData2) & "]"
            End If
        Else
            'alors c'est un processus
            If tUndoType = actByteWritten Then
                s = "o=[" & LTrim$(Str$(curData1)) & "]c=[" & LTrim$(Str$(bytData1)) & "]s=[" & Formated1String(sData2) & "]"
            End If
        End If
        
        .Add Text:=s
        .Item(Undo.lvHisto.ListItems.Count).SubItems(1) = Str$(Undo.lvHisto.ListItems.Count)
    End With
    Undo.lvHisto.Visible = True
        
End Sub

'=======================================================
'supprime des entrées de l'historique, prend en paramètre
'-le rang de la limite de suppression (tout ce qui est > est supprimé)
'-l'historique
'=======================================================
Public Sub DelHisto(ByVal lRang As Long, ByRef Histo() As clsUndoSubItem, ByVal Undo As clsUndoItem)
Dim x As Long

    'procède à la suppression de tout ce qui est après lRang
    'lRang=1 ==> pas de suppression
    If lRang <> -1 Then
        ReDim Preserve Histo(lRang)
        
        'supprime également les items en trop dans le lv
        With Undo.lvHisto
            For x = .ListItems.Count To lRang Step -1
                .ListItems.Remove x
            Next x
            If .ListItems.Count Then .ListItems.Item(.ListItems.Count).Selected = True 'sélectionne le dernier item
        End With
    End If
End Sub
