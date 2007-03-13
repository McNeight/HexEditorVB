Attribute VB_Name = "mdlExport"
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
'MODULE CONTENANT LES PROCEDURES D'EXPORTS VERS LES DIFFERENTS FORMATS
'=======================================================


'=======================================================
'sauvegarde en HTML
'paramètres : sOutputFile (fichier de sortie)
'boolean pour les options
'sStringHex : contient la suite des valeurs hexa, ou le fichier d'entrée si fichier entier
'curFirstOffset : premier offset de la sélection (-1 si fichier entier)
'=======================================================
Public Sub SaveAsHTML(ByVal sOutputFile As String, ByVal bOffset As Boolean, _
    bString As Boolean, ByVal sStringHex As String, ByVal curFirstOffset As Currency, _
    Optional ByVal curSecondOffset As Currency)
    
    'exemple de string au format HTML (contient une ligne avec offset, hexa et ASCII)
    '<font face="Courier New"><font size="1">1248A1ED1 </font><font color="#0000ff" size="1"
    '>4D5A 0000 0000 0000 0000 0000 0000 0000 </font><font color="#000000" size="1"> M
    'Z..................<BR>
    
    
End Sub
