Attribute VB_Name = "mdlPrint"
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
'//MODULE D'IMPRESSION
'=======================================================


'=======================================================
'procedure qui permet de renvoyer une imprimante et ses propriétés
'affiche une boite de dialogue de choix d'imprimante et de configuration
'=======================================================
Public Sub GetPrinter(ByRef tPrinter As Printer)
Dim strPrinterName As String
Dim selectedPrinter As PRINTER_INFO
Dim addrstructDev As Long
Dim infoDevice As DEVMODE
Dim tPrint As Printer

    'initialisation de la variable selectedPrinter
    With selectedPrinter
        .lStructSize = Len(selectedPrinter) 'taille de la structure
        .hDevMode = 0&
        .hDevNames = 0&
        .FLAGS = PD_RETURNIC
    End With
    
    'affichage du Dialog Control
    If PrintDlg(selectedPrinter) <> 1 Then Exit Sub
    
    addrstructDev = GlobalLock(selectedPrinter.hDevMode)
    CopyMemory infoDevice, ByVal addrstructDev, Len(infoDevice)

    'formate le nom de l'imprimante
    strPrinterName = Left(infoDevice.dmDeviceName, InStr(1, infoDevice.dmDeviceName, vbNullChar) - 1)
    
    'affecte le printer choisi
    For Each tPrint In Printers
        If strPrinterName = tPrint.DeviceName Then
            'c'est l'imprimante choisie, on la retourne
            Set tPrinter = tPrint
            Exit For
        End If
    Next tPrint
    
End Sub
