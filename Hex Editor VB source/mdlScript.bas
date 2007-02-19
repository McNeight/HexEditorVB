Attribute VB_Name = "mdlScript"
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
'//MODULE DE GESTION DES SCRIPTS
'=======================================================

Private Const COM As String = "ADR|ADR+|ADR-|BEEP|CLOSE_PROCESS_NAME|CLOSE_PROCESS_PATH|CLOSE_PROCESS_PID|COUNT_ALL_HEX|COUNT_ALL_STR|COUNT_HEX|COUNT_STR|CREATE_FILE|CREATE_PROCESS|DEL|DEL_FILE|EXIT|FILE_EXISTS|FIND_HEX|FIND_STR|FOLDER_EXISTS|GET_HEX|GET_STR|GOTO|IF|INSERT|LABEL|MSG|PASTE_HEX|PASTE_STR|PROCESS_EXISTS_NAME|PROCESS_EXISTS_PATH|PROCESS_EXISTS_PID|REBOOT|REM|REPEAT|REPLACE_ALL_HEX|REPLACE_ALL_STR|REPLACE_HEX|REPLACE_STR|RESUME_PROCESS_NAME|RESUME_PROCESS_PATH|RESUME_PROCESS_PID|SHUTDOWN|STO_A|STO_B|STO_C|STO_D|SUSPEND_PROCESS_NAME|SUSPEND_PROCESS_PATH|SUSPEND_PROCESS_PID|UNTIL|USE_FILE|USE_PROCESS_NAME|USE_PROCESS_PATH|USE_PROCESS_PID|USE_VOLUME|WHILE|"
Private sCom() As String


'=======================================================
'vérifie la cohérence d'un script
'renvoie 0 si OK, renvoie le numéro de la ligne si pas OK
'=======================================================
Public Function IsScriptCorrect(ByVal sText As String) As Long
Dim sLine() As String   'contiendra la liste des commandes
Dim x As Long
Dim l As Long
    
    On Error GoTo ErrGestion
    
    IsScriptCorrect = 0
    
    '//COMMANDES EXISTANTES
    '#A                                 ==> variable A
    '#B                                 ==> variable B
    '#C                                 ==> variable C
    '#COMMAND_FILE                      ==> contient le path du fichier qui ouvre le script (argument Command)
    '#D                                 ==> variable D
    '#DESKTOP_DIRECTORY                 ==> contient le path du bureau
    '#FILE_LENGHT                       ==> cotient la taille du fichier utilisé
    '#FILE_PATH                         ==> contient le path du fichier utilisé
    '#PROCESS_PATH                      ==> contient le path du processus utilisé
    '#PROCESS_PID                       ==> contient le PID du processus utilisé
    '#PROGRAM_DIRECTORY                 ==> contient le path de Program Files
    '#STARTUP_DIRECTORY                 ==> contient le path du dossier Startup
    '#TEMP_DIRECTORY                    ==> contient le path Temp
    '#USER_DOCUMENTS_DIRECTORY          ==> contient le path des documents perso
    '#VOLUME_LETTER                     ==> contient la lettre du drive utilisé
    '#WINDOWS_DIRECTORY                 ==> contient le path du répertoire Windows
    'ADR <addr>                         ==> se positionne à l'adresse addr (bouge le curseur)
    'ADR+ <n>                           ==> incrémente l'adresse (bouge le curseur)
    'ADR- <n>                           ==> décremente l'adresse (bouge le curseur)
    'BEEP                               ==> BEEP !
    'CLOSE_PROCESS_NAME <name>          ==> kille un processus par nom
    'CLOSE_PROCESS_PATH <file>          ==>                    par path
    'CLOSE_PROCESS_PID <pid>            ==>                    par PID
    'COUNT_ALL_HEX <hex>                ==> fonction renvoyant le nombre d'occurences de hex
    'COUNT_ALL_STR <str>                ==>                                           de str
    'COUNT_HEX <hex>                    ==> fonction renvoyant le nombre d'occurences de hex à partir du curseur
    'COUNT_STR <str>                    ==>                                           de str
    'CREATE_FILE <file>                 ==> créé le fichier file (pas d'overwrite)
    'CREATE_PROCESS <file>              ==> créé le processus du fichier file
    'DEL <n>                            ==> supprime n caractère à partir du curseur
    'DEL_FILE <file>                    ==> supprime le fichier file
    'EXIT                               ==> termine le script
    'FILE_EXISTS <file>                 ==> fonction renvoyant 1 si le fichier existe ou non
    'FIND_HEX <hex>                     ==> fonction qui renvoie la prochaine occurence de hex à partir du curseur
    'FIND_STR <str>                     ==>                                                str
    'FOLDER_EXISTS <folder>             ==> fonction renvoyant 1 si le dossier existe
    'GET_HEX <size>                     ==> fonction renvoyant size valeurs hexa à partir du curseur
    'GET_STR <size>                     ==>                                 string
    'GOTO <n>                           ==> se rend au label n
    'IF <condition> THEN <action>       ==> test logique (action contient une procédure)
    'INSERT <n>                         ==> insère n bytes à l'emplacement du curseur
    'LABEL <n>                          ==> marqueur pour le goto
    'MSG <message>                      ==> affiche un message
    'PASTE_HEX <hex>                    ==> colle la valeur hex à partir du curseur
    'PASTE_STR <str>                    ==>                 str
    'PROCESS_EXISTS_NAME <name>         ==> fonction renvoyant 1 si le processus name existe
    'PROCESS_EXISTS_PATH <file>         ==>                                      file
    'PROCESS_EXISTS_PID <pid>           ==>                                      PID
    'REBOOT                             ==> redémarre le système
    'REM                                ==> affiche une remarque
    'REPEAT <n> <action>                ==> répète n fois l'action
    'REPLACE_ALL_HEX <from> BY <to>     ==> remplace toutes les valeurs hexa from par to
    'REPLACE_ALL_STR <from> BY <to>     ==>                             strings
    'REPLACE_HEX <from> BY <to>         ==> remplace la prochaine valeur hexa from par to
    'REPLACE_STR <from> BY <to>         ==>                              string
    'RESUME_PROCESS_NAME <name>         ==> débloque le processus name
    'RESUME_PROCESS_PATH <path>         ==>                       path
    'RESUME_PROCESS_PID <pid>           ==>                       PID
    'SHUTDOWN                           ==> éteind le PC
    'STO_A <function>                   ==> stocke une variant dans A
    'STO_B <function>                   ==>                         B
    'STO_C <function>                   ==>                         C
    'STO_D <function>                   ==>                         D
    'SUSPEND_PROCESS_NAME <name>        ==> suspend le processus name
    'SUSPEND_PROCESS_PATH <path>        ==>                      path
    'SUSPEND_PROCESS_PID <pid>          ==>                      PID
    'UNTIL <condition> DO <action>      ==> répète action tant que pas condition
    'USE_FILE <file>                    ==> utilise le fichier file
    'USE_PROCESS_NAME <name>            ==> utilise le processus name
    'USE_PROCESS_PATH <file>            ==>                      file
    'USE_PROCESS_PID <pid>              ==>                      PID
    'USE_VOLUME <letter>                ==> utilise le disque letter
    'WHILE <condition> DO <action>      ==> tant que condition faire action
    
    
        
    'procède à quelques vérifications de base du code
    'renvoie un message pour chaque bug trouvé
    
    'vérifications effectuées ==>
        '1)vérifications syntaxiques ==> vérifie la cohérence de la syntaxe
            '-respect de l'écriture
            '-respect des arguments (leur type)
        '2)vérifications sur les labels/goto ==> pas l'un sans l'autre
        '3)vérifications surles variables ==> bug si utilise #A et pas STO_A par exemple
        '4)vérification que l'on travaille bien sur un objet, et sur un bon objet (pas d'accès mémoire opur un fichier par exemple)
    
    
    
    '//DECOUPAGE du texte en plusieurs lignes de code (une par ligne effetive)
    sLine = Split(sText, vbNewLine)
    
    '//VERIFICATION de la syntaxe (un espace après chaque commande de début de ligne)
        For x = 0 To UBound(sLine())
            
            'contient la taille de la commande (len)
            l = Len(GetCommand(sLine(x)))
            
            If l = 0 Then
                If Len(sLine(x)) = 0 Then
                    'ligne vierge
                    MsgBox "La ligne " & CStr(x + 1) & " est vierge.", vbCritical + vbOKOnly, "Erreur détectée"
                    IsScriptCorrect = x + 1
                    Exit Function
                Else
                    'alors ce n'est pas une commande valide
                    MsgBox "La ligne " & CStr(x + 1) & " n'est pas valide.", vbCritical + vbOKOnly, "Erreur détectée"
                    IsScriptCorrect = x + 1
                    Exit Function
                End If
            End If
        
            If sLine(x) <> "EXIT" And sLine(x) <> "BEEP" And sLine(x) <> "REBOOT" And _
                sLine(x) <> "REM" And sLine(x) <> "SHUTDOWN" Then
                If Mid$(sLine(x), l + 1, 1) <> " " And Left$(sLine(x), 1) <> "#" Then
                    'manque un espace à la ligne x+1
                    MsgBox "La commande de la ligne " & CStr(x + 1) & " n'est pas reconnue.", vbCritical + vbOKOnly, "Erreur détectée"
                    IsScriptCorrect = x + 1
                    Exit Function
                End If
            End If
        Next x
    
    '//VERIFICATION de la cohérence des types attendus
    
    
    
    Exit Function
ErrGestion:
    clsERREUR.AddError "mdlScript.IsScriptCorrect", True
End Function

'=======================================================
'lance un script
'=======================================================
Public Function LauchScript(ByVal sText As String) As Boolean

End Function

'=======================================================
'obtient la commande de la ligne S
'=======================================================
Private Function GetCommand(ByVal s As String) As String
Dim x As Long
Dim y As Long
Dim l As Long

    l = Len(s)
    For x = l To 1 Step -1  'parcourt toute la string
        For y = 0 To UBound(sCom()) 'parcourt toutes les commandes dispo.
            If Mid$(s, 1, x) = sCom(y) Then
                'on a trouvé une commande
                GetCommand = sCom(y)
                Exit Function
            End If
        Next y
    Next x
    
End Function

'=======================================================
'transforme la string COM en tableau
'0 à ubound
'=======================================================
Public Sub GetSplit()
    sCom = Split(COM, "|")
End Sub
