VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'//CLASSE DE CONVERSION INTER BASES
'=======================================================

'=======================================================
'ENUM
'=======================================================
Public Enum TYPEOFERROR
    [Valeur a convertir trop grande]
    [Base d'arrivee interdite]
    [Base de depart interdite]
    [Base de depart et string a convertir incompatibles]
End Enum


'=======================================================
'VARIABLES PRIVEES
'=======================================================
Private strString As String 'string actuelle
Private bytBase As Byte 'base actuelle
Private bConversionPossible As Long  'conversion possible ou pas


'=======================================================
'EVENT
'=======================================================
Public Event ErrorOccured(Error As TYPEOFERROR)


'=======================================================
'PROPERTIES
'=======================================================
Public Property Get CurrentString() As String: CurrentString = strString: End Property
Public Property Let CurrentString(ByVal NewString As String): strString = UCase$(NewString): End Property
Public Property Get CurrentBase() As Byte: CurrentBase = bytBase: End Property
Public Property Let CurrentBase(ByVal NewBase As Byte)
    If NewBase >= 2 And NewBase <= 36 Then
        bytBase = NewBase
    Else
        RaiseEvent ErrorOccured([Base de depart interdite])
    End If
End Property
Public Property Get StringLen() As Long: StringLen = Len(Me.CurrentString): End Property
Public Property Get ConversionFailed() As Boolean: ConversionFailed = Not (bConversionPossible = 1): End Property



'=======================================================
'PUBLIC PROCEDURE & FUNCTIONS
'=======================================================

'=======================================================
'lance la conversion
'=======================================================
Public Function Convert(ByVal NewBase As Byte, Optional ActualBase As Byte) As String
Dim lBase As Long

    bConversionPossible = 0
    If Not (NewBase >= 2 And NewBase <= 36) Then
        Convert = "-1"
        RaiseEvent ErrorOccured([Base d'arrivee interdite])
        Exit Function
    End If
    
    'récupère la base de départ
    If (ActualBase < 2 Or ActualBase > 36) Then
        lBase = Me.CurrentBase
    Else
        lBase = ActualBase
    End If
    
    If IsBaseOk(lBase, Me.CurrentString) = False Then
        Convert = "-1"
        RaiseEvent ErrorOccured([Base de depart et string a convertir incompatibles])
        Exit Function
    End If

    'lance la conversion
    Convert = ConvertBASES(lBase, NewBase, Me.CurrentString)
End Function





'=======================================================
'PRIVATE PROCEDURE & FUNCTIONS
'=======================================================

'=======================================================
'initialisation de la classe
'=======================================================
Private Sub Class_Initialize()
    Me.CurrentBase = 10
    Me.CurrentString = vbNullString
    bConversionPossible = 1
End Sub

'=======================================================
'convertion inter bases
'récupère la base de départ, celle d'arrivée et la string à convertir
'renvoie la string convertie
'=======================================================
Private Function ConvertBASES(ByVal BaseDep As Long, ByVal BaseArriv As Long, _
    ByVal sWord As String) As String
Dim x As Long
Dim lLen As Long
Dim s As String
Dim nb As Currency
Dim sRes As String
Dim valD As Currency
Dim val3 As Long
Dim val2 As Currency

    On Error GoTo ErrGestion
    
    'taille de la string
    lLen = Len(sWord)
    
    sRes = vbNullString
    nb = 0
    
    '//on récupère la valeur du nombre en décimal
    'pour chaque caractère de la string
    For x = lLen To 1 Step -1
        s = Mid$(sWord, x, 1)
        
        'récupère la valeur de ce caractère dans la base de départ
        valD = AexpB(BaseDep, lLen - x)
        
        If bConversionPossible = 2 Then GoTo ErrGestion
                
        valD = GetAssociatedVal(s) * valD
        
        nb = nb + valD
    Next x
    
    '//détermine la taille de la future string
    val2 = 0: x = 0
    While nb >= val2
    
        If bConversionPossible = 2 Then GoTo ErrGestion
    
        val2 = AexpB(BaseArriv, x)
        x = x + 1
    Wend
    
    sRes = String$(x - 1, "0")

    '//maintenant on créé la string dans la nouvelle base
    While Not (nb = 0)
        
        'recherche la plus grande puissance possible
        val2 = 0: x = 0
        While nb >= val2
            val2 = AexpB(BaseArriv, x)
            
            If bConversionPossible = 2 Then GoTo ErrGestion
            
            x = x + 1
        Wend
        
        'alors on a récupère le nombre de fois que cette puissance rentre dans le nombre
        val2 = AexpB(BaseArriv, x - 2)
        val3 = Int(nb / val2)
        
        'créé la string
        If Len(sRes) > (x - 1) Then
            sRes = Left$(sRes, Len(sRes) - x + 1) & GetAssociatedCar(val3) & Right$(sRes, Len(sRes) - Len(Left$(sRes, Len(sRes) - x + 1)) - 1)
        Else
            sRes = GetAssociatedCar(val3) & Right$(sRes, Len(sRes) - 1)
        End If
        
        'on retranche le nombre
        val2 = val2 * val3
        nb = nb - val2
    Wend
            
    ConvertBASES = sRes
    
    bConversionPossible = 1
    Exit Function
    
ErrGestion:
    ConvertBASES = "-1"
    RaiseEvent ErrorOccured([Valeur a convertir trop grande])
End Function

'=======================================================
'renvoie le caractère associé à l'élémént bytChar (de 0 à 35)
'ex : bytChar=10 ==> GetAssociatedCar=A
'=======================================================
Private Function GetAssociatedCar(ByVal bytChar As Long) As String
    
    If bytChar < 10 Then
        GetAssociatedCar = Trim$(Str$(bytChar))
    Else
        GetAssociatedCar = Chr_(55 + bytChar)
    End If
  
End Function

'=======================================================
'renvoie la valeur associée à l'élémént strChar (de 0 à 35)
'ex : strChar="A" ==> GetAssociatedVal=10
'=======================================================
Private Function GetAssociatedVal(ByVal strChar As String) As Long
Dim byt As Long
    
    byt = Asc(strChar)
    
    If byt <= 57 Then
        'alors c'est un chiffre
        GetAssociatedVal = byt - 48
    Else
        'alors c'est une lettre
        GetAssociatedVal = byt - 55
    End If
        
End Function

'=======================================================
'renvoie a^b (plus rapide que a^b et pas d'erreur)
'=======================================================
Private Function AexpB(ByVal a As Currency, ByVal b As Long) As Currency
Dim x As Long
Dim l As Currency

    On Error GoTo ErrGestion

    If b = 0 Then
        AexpB = 1
        Exit Function
    End If
    
    l = 1
    For x = 1 To b
        l = l * a
    Next x
    AexpB = l
    
    Exit Function
ErrGestion:
    bConversionPossible = 2
End Function

'=======================================================
'vérifie que la base est compatible avec la string
'=======================================================
Private Function IsBaseOk(ByVal lngBase As Long, ByVal strS As String) As Boolean
Dim x As Long
Dim l As Byte
Dim lmax As Byte
Dim lmin As Byte

    IsBaseOk = False
    
    lmax = 0: lmin = 255
    'il faut vérifier que la string ne comporte pas de valeur hors base
    For x = 1 To Len(strS)
        l = Asc(Mid$(strS, x, 1))
        If l > lmax Then lmax = l   'récupère le plus grand char
        If l < lmin Then lmin = l 'récupère le plus petit char
    Next x
    
    If lngBase <= 10 Then
        'alors il faut que des chiffres
        
        If Not (lmin >= 48 And lmin < (48 + lngBase) And lmax >= 48 And lmax < (48 + lngBase)) Then
            'alors c'est pas bon
            Exit Function
        End If
    Else
        'chiffres et lettres A-Z
        If lmin < 48 Or (lmin > 57 And lmin < 65) Or lmin >= (55 + lngBase) Or lmax < 48 Or lmax >= (55 + lngBase) Or _
            (lmax > 57 And lmax < 65) Then
            'pas bon
            Exit Function
        End If
    End If
    
    IsBaseOk = True
    
End Function


