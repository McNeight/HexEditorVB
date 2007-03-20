Attribute VB_Name = "modDecode"
Option Explicit

Public regESP As Long         'état de la pile après l'instruction
Public regEBP As Long         'facultatif : état du registre de base (EBP) après l'instruction (si utilisé pour un contexte de pile)

'infos sur une ionstruction
Public Type Instruction
    opclass  As Byte    '0 one byte instruction, &h0f two byte instruction
    regIP As Long       'adresse virtuelle relative de l'instruction
    regESP As Long
    'bPrefixes(0 To 3) As Byte
    operandSizeOverride As Byte 'facultatif : prefixe de changement de taille d'opérandes (32bits -> 16bits ou 16bits à 8bits)
    addressSizeOverride As Byte 'facultatif : prefixe de changement de taille d'adresses virtuelles (32bits -> 16bits ou 16bits à 8bits)
    segmentOverride As Byte     'facultatif : définit un autre registre de segment que le registre de segment implicit (DS pour les données est implicit par ex)
    LockAndRepeat As Byte       'facultatif : définit un vérouillage ou une répétition
    iOpcode As Byte             'opcode de l'instruction
    bModRm As Byte              'facultatif : octet de ModRM de l'instruction
    bSib As Byte                'facultatif : octet de Scale Index Base de l'instruction
    'valeurs immediates
    i_byte As Long              'facultatif : octet immediat de l'instruction
    i_dword As Long             'facultatif : dword (ou word) immediat de l'instruction
    'adresse mémoire
    m_byte As Long              'facultatif : octet d'adresse mémoire (virtuelle) de l'instruction
    m_dword As Long             'facultatif : dword d'adresse mémoire (virtuelle) de l'instruction
    bStop As Boolean            'indique que la procédure est finie
End Type

'type d'opérandes pour les instructions d'un octet
Private Enum OneByteOpcodeType
    'pas d'instruction
    oboNoInstruction = -1
    'rien derrière l'instruction
    oboNoByteFollow = 0
    'un octet de donnée immédiate derrière l'instruction
    oboImmediatByte = 1
    'un mot de donnée immédiate derrière l'instruction
    oboImmediatWord = 2
    'un mot puis un octet de donnée immédiate derrière l'instruction
    oboImmediatWordByte = 3
    'une donnée immédiate (dépendant de la taille d'un opérande) derrière l'instruction
    oboImmediatOperandSize = 4
    'une donnée immédiate (dépendant de la taille de l'adresse mémoire) derrière l'instruction
    oboImmediatAddressSize = 44
    'un mot suivit d'un double mot (constituant une adresse mot:double_mot) derrière l'instruction
    oboImmediatDirectAddressWordDword = 5
    'un octet de ModRM derrière l'instruction
    oboModRMFollow = 6
    'un octet de ModRM puis un octet de donnée derrière l'instruction
    oboModRMByte = 7
    'un octet de ModRM puis une donnée (dépendant de la taille d'un opérande) derrière l'instruction
    oboModRMOperandSize = 8
    'un octet de ModRM (en tant qu'extension d'opcode) derrière l'instruction
    oboOpExt = 9
    'un octet de ModRM (en tant qu'extension d'opcode) puis un octet de donnée derrière l'instruction
    oboOpExtByte = 10
    'un octet de ModRM (en tant qu'extension d'opcode)puis une donnée (dépendant de la taille d'un opérande) derrière l'instruction
    oboOpExtOperandSize = 11
    'un octet de ModRM (en tant qu'extension d'opcode pour les instructions Escape) derrière l'instruction
    oboOpExt2 = 12
    'l'instruction JUMP
    oboJUMP = 13
    'l'instruction TEST
    oboTEST = 14
    'les instructions WAIT
    oboWAIT = 15
    'les REP/REPNE
    oboREPEAT = 16
End Enum

'contient le tableau en commentaire : les informations sur les opérandes des instructions à un octet
'pour plus d'infos sur la signification des chiffres, se reporter à l'énumération OneByteOpcodeType
'les numéros de lignes et colonnes (combinées) représentent les opcodes des instructions d'un octet
Public opcodeTable()
'int opcodeTable[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    6,  6,  6,  6,  1,  4,  0,  0,  6,  6,  6,  6,  1,  4,  0, 98,
'/*10*/    6,  6,  6,  6,  1,  4,  0,  0,  6,  6,  6,  6,  1,  4,  0,  0,
'/*20*/    6,  6,  6,  6,  1,  4, 99,  0,  6,  6,  6,  6,  1,  4, 99,  0,
'/*30*/    6,  6,  6,  6,  1,  4, 99,  0,  6,  6,  6,  6,  1,  4, 99,  0,
'/*40*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*50*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*60*/    0,  0,  6,  6, 99, 99, 99, 99,  4,  8,  1,  7,  0,  0,  0,  0,
'/*70*/    1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,
'/*80*/   10, 11, -1, 10,  6,  6,  6,  6,  6,  6,  6,  6,  6,  6,  6,  9,
'/*90*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  5, 15,  0,  0,  0,  0,
'/*A0*/   44, 44, 44, 44,  0,  0,  0,  0,  1,  4,  0,  0,  0,  0,  0,  0,
'/*B0*/    1,  1,  1,  1,  1,  1,  1,  1,  4,  4,  4,  4,  4,  4,  4,  4,
'/*C0*/   10, 10,  2,  0,  6,  6, 10, 11,  3,  0,  2,  0,  0,  1,  0,  0,
'/*D0*/    9,  9,  9,  9,  1,  1, -1,  0, 12, 12, 12, 12, 12, 12, 12, 12,
'/*E0*/    1,  1,  1,  1,  1,  1,  1,  1,  4,  4,  5,  1,  0,  0,  0,  0,
'/*F0*/    0,  0, 16, 16,  0,  0, 14, 14,  0,  0,  0,  0,  0,  0,  9, 13};
'/* -----------------------------------------------------------------------*/

'type des opérandes des instructions sur deux octets
Private Enum TwoByteOpcodeType
    'pas d'instruction
    tboNoInstruction = -1
    'rien derrière l'instruction
    tboNoByte = 0
    'une adresse relative sur un mot ou un double mot derrière l'instruction
    tboAddressFollow = 1
    'un octet de ModRM derrière l'instruction
    tboModRM = 2
    'un octet de ModRM puis un octet de donnée derrière l'instruction
    tboModRMByte = 3
    'un octet de MpdRM (en tant qu'extension de l'opcode) derrière l'instruction
    tboOpExt = 4
    'un octet de MpdRM (en tant qu'extension de l'opcode) puis un octet de donnée derrière l'instruction
    tboOpExtByte = 5
End Enum


'contient le tableau en commentaire : les informations sur les opérandes des instructions à deux octets
'pour plus d'infos sur la signification des chiffres, se reporter à l'énumération TwoByteOpcodeType
'les numéros de lignes et colonnes (combinées) représentent les opcodes des instructions de deux octets
Public opcode2Table()
'int opcode2Table[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    4,  4,  2,  2, -1, -1,  0, -1,  0,  0, -1,  0, -1, -1, -1, -1,
'/*10*/   -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
'/*20*/    2,  2,  2,  2, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
'/*30*/    0,  0,  0,  0,  0,  0, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
'/*40*/    2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,
'/*50*/   -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1,
'/*60*/    2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2, -1, -1,  2,  2,
'/*70*/   -1,  5,  5,  5,  2,  2,  2,  0, -1, -1, -1, -1, -1, -1,  2,  2,
'/*80*/    1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,  1,
'/*90*/    2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,  2,
'/*A0*/    0,  0,  0,  2,  3,  2, -1, -1,  0,  0,  0,  2,  3,  2,  4,  2,
'/*B0*/    2,  2,  2,  2,  2,  2,  2,  2, -1, -1,  5,  2,  2,  2,  2,  2,
'/*C0*/    2,  2, -1, -1, -1, -1, -1,  4,  0,  0,  0,  0,  0,  0,  0,  0,
'/*D0*/   -1,  2,  2,  2, -1,  2, -1, -1,  2,  2, -1,  2,  2,  2, -1,  2,
'/*E0*/   -1,  2,  2, -1, -1,  2, -1, -1,  2,  2, -1,  2,  2,  2, -1,  2,
'/*F0*/   -1,  2,  2,  2, -1,  2, -1, -1,  2,  2,  2, -1,  2,  2,  2, -1};
'/* -----------------------------------------------------------------------*/

'type de répétitions possibles
Private Enum RepeatGroup
    'pas de REP possible devant l'instruction
    REPNotAllowed = 0
    'REPNE et REPE possibles devant l'instruction
    REPNEAllowed = 1
    'seulement REP possible devant l'instruction
    AllREPNotAllowed = 2
End Enum

'contient le tableau en commentaire : les informations sur les possibilités de REP/REPNE avant l'instruction
'pour plus d'infos sur la signification des chiffres, se reporter à l'énumération RepeatGroup
'les numéros de lignes et colonnes (combinées) représentent les opcodes des instructions d'un octet
Public repeatgroupTable()
'int repeatgroupTable[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*10*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*20*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*30*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*40*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*50*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*60*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  2,  2,  2,  2,
'/*70*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*80*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*90*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*A0*/    0,  0,  0,  0,  2,  2,  1,  1,  0,  0,  2,  2,  2,  2,  1,  1,
'/*B0*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*C0*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*D0*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*E0*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,
'/*F0*/    0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0,  0};
'/* -----------------------------------------------------------------------*/


'type de ModRM pour les instructions
Private Enum ModRMType
    'pas d'octet ModRM derrière l'instruction
    NoModRM = 0
    'seulement l'octet de ModRM derrière l'instruction
    OnlyModRM = 1
    'un octet ModRM et un octet SIB (utilisé) derrière l'instruction
    ModRMSib = 2
    'un octet ModRM, un octet SIB (index) et un double mot de déplacement derrière l'instruction
    ModRMSibDword = 3
    'un octet ModRM, un octet SIB (non utilisé : registre fixé) et un octet (signé) de déplacement derrière l'instruction
    ModRMByte = 4
    'un octet ModRM, un octet de SIB (utilisé) et un octet de donnée derrière l'instruction
    ModRMSibByte = 5
    'un octet ModRM, un octet SIB (non utilisé : registre fixé) et un double mot de déplacement derrière l'instruction
    ModRMDword = 6
    'un octet ModRM, un octet de SIB (utilisé) et un double mot de donnée derrière l'instruction
    ModRMSibDword2 = 7
    'un octet ModRM désigne un registre derrière l'instruction
    OnlyModRMReg = 8
End Enum

'contient le tableau en commentaire : les informations sur le ModRM et les opérandes  en fonction des valeurs de ModRM
'pour plus d'infos sur la signification des chiffres, se reporter à l'énumération ModRMType
'les numéros de lignes et colonnes (combinées) représentent les valeurs possibles de ModRM
Public modTable()
'voir page 36 vol2
'int modTable[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    1,  1,  1,  1,  2,  3,  1,  1,  1,  1,  1,  1,  2,  3,  1,  1,
'/*10*/    1,  1,  1,  1,  2,  3,  1,  1,  1,  1,  1,  1,  2,  3,  1,  1,
'/*20*/    1,  1,  1,  1,  2,  3,  1,  1,  1,  1,  1,  1,  2,  3,  1,  1,
'/*30*/    1,  1,  1,  1,  2,  3,  1,  1,  1,  1,  1,  1,  2,  3,  1,  1,
'/*40*/    4,  4,  4,  4,  5,  4,  4,  4,  4,  4,  4,  4,  5,  4,  4,  4,
'/*50*/    4,  4,  4,  4,  5,  4,  4,  4,  4,  4,  4,  4,  5,  4,  4,  4,
'/*60*/    4,  4,  4,  4,  5,  4,  4,  4,  4,  4,  4,  4,  5,  4,  4,  4,
'/*70*/    4,  4,  4,  4,  5,  4,  4,  4,  4,  4,  4,  4,  5,  4,  4,  4,
'/*80*/    6,  6,  6,  6,  7,  6,  6,  6,  6,  6,  6,  6,  7,  6,  6,  6,
'/*90*/    6,  6,  6,  6,  7,  6,  6,  6,  6,  6,  6,  6,  7,  6,  6,  6,
'/*A0*/    6,  6,  6,  6,  7,  6,  6,  6,  6,  6,  6,  6,  7,  6,  6,  6,
'/*B0*/    6,  6,  6,  6,  7,  6,  6,  6,  6,  6,  6,  6,  7,  6,  6,  6,
'/*C0*/    8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,
'/*D0*/    8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,
'/*E0*/    8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,
'/*F0*/    8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8,  8};
'/* -----------------------------------------------------------------------*/


'type de ModRM 16 bits pour les instructions
Private Enum ModRM16Type
    'pas de ModRM derrière l'instruction
    NoModRM = 0
    'un octet de ModRM derrière l'instruction
    OnlyModRM = 1
    'un octet de ModRM et un mot de déplacement derrière l'instruction
    ModRMWord = 2
    'un octet de ModRM et un octet de déplacement (signé) derrière l'instruction
    ModRMByte = 3
    'un octet de ModRM et un mot de déplacement (ajouté à des registres) derrière l'instruction
    RegModRMWord = 4
    'un octet de ModRM désigne un registre derrière l'instruction
    OnlyRegModRM = 5
End Enum

'contient le tableau en commentaire : les informations sur le ModRM16 et les opérandes en fonction des valeurs de ModRM16
'pour plus d'infos sur la signification des chiffres, se reporter à l'énumération ModRM16Type
'les numéros de lignes et colonnes (combinées) représentent les valeurs possibles de ModRM
Public mod16Table()
'int mod16Table[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    1,  1,  1,  1,  1,  1,  2,  1,  1,  1,  1,  1,  1,  1,  2,  1,
'/*10*/    1,  1,  1,  1,  1,  1,  2,  1,  1,  1,  1,  1,  1,  1,  2,  1,
'/*20*/    1,  1,  1,  1,  1,  1,  2,  1,  1,  1,  1,  1,  1,  1,  2,  1,
'/*30*/    1,  1,  1,  1,  1,  1,  2,  1,  1,  1,  1,  1,  1,  1,  2,  1,
'/*40*/    3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,
'/*50*/    3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,
'/*60*/    3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,
'/*70*/    3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,  3,
'/*80*/    4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,
'/*90*/    4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,
'/*A0*/    4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,
'/*B0*/    4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,  4,
'/*C0*/    5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,
'/*D0*/    5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,
'/*E0*/    5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,
'/*F0*/    5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5,  5};
'/* -----------------------------------------------------------------------*/


'contient le tableau en commentaire : les informations sur le SIB en fonction des valeurs de cet octet
'pour plus d'infos sur la signification des chiffres :
'0 : pas de SIB
'1 : registres fixés pour Base Index (et Scale fixée)
'2 : un double mot de déplacement suit l'instruction, pas de base si Mod = 00 et EBP si Mod <> 0
'les numéros de lignes et colonnes (combinées) représentent les valeurs possibles de SIB
Public sibTable()
'int sibTable[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*10*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*20*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*30*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*40*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*50*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*60*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*70*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*80*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*90*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*A0*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*B0*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*C0*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*D0*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*E0*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2,
'/*F0*/    2,  2,  2,  2,  2,  1,  2,  2,  2,  2,  2,  2,  2,  1,  2,  2};
'/* -----------------------------------------------------------------------*/

'contient le tableau en commentaire : les informations sur les registres utilisés en fonction des valeurs de SIB
'pour plus d'infos sur la signification des chiffres :
'0 : EAX
'1 : ECX
'2 : EDX
'3 : EBX
'4 : ESP
'5 : EBP
'6 : ESI
'7 : EDI
'les numéros de lignes et colonnes (combinées) représentent les valeurs possibles de SIB ou ModRM
Public regTable()
'voir page 35 vol2
'int regTable[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    0,  0,  0,  0,  0,  0,  0,  0,  1,  1,  1,  1,  1,  1,  1,  1,
'/*10*/    2,  2,  2,  2,  2,  2,  2,  2,  3,  3,  3,  3,  3,  3,  3,  3,
'/*20*/    4,  4,  4,  4,  4,  4,  4,  4,  5,  5,  5,  5,  5,  5,  5,  5,
'/*30*/    6,  6,  6,  6,  6,  6,  6,  6,  7,  7,  7,  7,  7,  7,  7,  7,
'/*40*/    0,  0,  0,  0,  0,  0,  0,  0,  1,  1,  1,  1,  1,  1,  1,  1,
'/*50*/    2,  2,  2,  2,  2,  2,  2,  2,  3,  3,  3,  3,  3,  3,  3,  3,
'/*60*/    4,  4,  4,  4,  4,  4,  4,  4,  5,  5,  5,  5,  5,  5,  5,  5,
'/*70*/    6,  6,  6,  6,  6,  6,  6,  6,  7,  7,  7,  7,  7,  7,  7,  7,
'/*80*/    0,  0,  0,  0,  0,  0,  0,  0,  1,  1,  1,  1,  1,  1,  1,  1,
'/*90*/    2,  2,  2,  2,  2,  2,  2,  2,  3,  3,  3,  3,  3,  3,  3,  3,
'/*A0*/    4,  4,  4,  4,  4,  4,  4,  4,  5,  5,  5,  5,  5,  5,  5,  5,
'/*B0*/    6,  6,  6,  6,  6,  6,  6,  6,  7,  7,  7,  7,  7,  7,  7,  7,
'/*C0*/    0,  0,  0,  0,  0,  0,  0,  0,  1,  1,  1,  1,  1,  1,  1,  1,
'/*D0*/    2,  2,  2,  2,  2,  2,  2,  2,  3,  3,  3,  3,  3,  3,  3,  3,
'/*E0*/    4,  4,  4,  4,  4,  4,  4,  4,  5,  5,  5,  5,  5,  5,  5,  5,
'/*F0*/    6,  6,  6,  6,  6,  6,  6,  6,  7,  7,  7,  7,  7,  7,  7,  7};
'/* -----------------------------------------------------------------------*/

'contient le tableau en commentaire : les informations sur les registres utilisés (comme base) en fonction des valeurs de SIB
'pour plus d'infos sur la signification des chiffres :
'0 : EAX
'1 : ECX
'2 : EDX
'3 : EBX
'4 : ESP
'5 : EBP
'6 : ESI
'7 : EDI
'les numéros de lignes et colonnes (combinées) représentent les valeurs possibles de SIB
Public rmTable()
'int rmTable[] = {
'/*        0   1   2   3   4   5   6   7   8   9   A   B   C   D   E   F   */
'/* -----------------------------------------------------------------------*/
'/*00*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*10*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*20*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*30*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*40*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*50*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*60*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*70*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*80*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*90*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*A0*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*B0*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*C0*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*D0*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*E0*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7,
'/*F0*/    0,  1,  2,  3,  4,  5,  6,  7,  0,  1,  2,  3,  4,  5,  6,  7};
'/* -----------------------------------------------------------------------*/

'contient les adresses des CALLs
Dim callCol As Collection
'contient les adresses des sauts
Dim jCol As Collection
Public tryCallCol As Collection

Public Sub Init2()
InitDesasm
InitPrint

InitCOFF
InitNames

dwImageBase = 0
Erase retSectionTables

Set callCol = New Collection
Set jCol = New Collection
Set tryCallCol = New Collection
Set argCol = New Collection
Set varCol = New Collection
End Sub

Public Sub Init()
InitDesasm
InitPrint

InitCOFF
InitNames

dwImageBase = 0
Erase retSectionTables

Set callCol = New Collection
Set jCol = New Collection
Set tryCallCol = New Collection
Set argCol = New Collection
Set varCol = New Collection

ResetTimer
StartTimer
End Sub

'initialisation des tables
Public Sub InitDesasm()
'remplissages des tables
opcodeTable = Array( _
    6, 6, 6, 6, 1, 4, 0, 0, 6, 6, 6, 6, 1, 4, 0, 98, _
    6, 6, 6, 6, 1, 4, 0, 0, 6, 6, 6, 6, 1, 4, 0, 0, _
    6, 6, 6, 6, 1, 4, 99, 0, 6, 6, 6, 6, 1, 4, 99, 0, _
    6, 6, 6, 6, 1, 4, 99, 0, 6, 6, 6, 6, 1, 4, 99, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 6, 6, 99, 99, 99, 99, 4, 8, 1, 7, 0, 0, 0, 0, _
    1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
   10, 11, 10, 10, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 9, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 5, 15, 0, 0, 0, 0, _
   44, 44, 44, 44, 0, 0, 0, 0, 1, 4, 0, 0, 0, 0, 0, 0, _
    1, 1, 1, 1, 1, 1, 1, 1, 4, 4, 4, 4, 4, 4, 4, 4, _
   10, 10, 2, 0, 6, 6, 10, 11, 3, 0, 2, 0, 0, 1, 0, 0, _
    9, 9, 9, 9, 1, 1, -1, 0, 12, 12, 12, 12, 12, 12, 12, 12, _
    1, 1, 1, 1, 1, 1, 1, 1, 4, 4, 5, 1, 0, 0, 0, 0, _
    0, 0, 16, 16, 0, 0, 14, 14, 0, 0, 0, 0, 0, 0, 9, 13)

opcode2Table = Array( _
    4, 4, 2, 2, -1, -1, 0, -1, 0, 0, -1, 0, -1, -1, -1, -1, _
   -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    2, 2, 2, 2, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    0, 0, 0, 0, 0, 0, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
   -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, -1, -1, 2, 2, _
   -1, 5, 5, 5, 2, 2, 2, 0, -1, -1, -1, -1, -1, -1, 2, 2, _
    1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
    2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
    0, 0, 0, 2, 3, 2, -1, -1, 0, 0, 0, 2, 3, 2, 4, 2, _
    2, 2, 2, 2, 2, 2, 2, 2, -1, -1, 5, 2, 2, 2, 2, 2, _
    2, 2, -1, -1, -1, -1, -1, 4, 0, 0, 0, 0, 0, 0, 0, 0, _
   -1, 2, 2, 2, -1, 2, -1, -1, 2, 2, -1, 2, 2, 2, -1, 2, _
   -1, 2, 2, -1, -1, 2, -1, -1, 2, 2, -1, 2, 2, 2, -1, 2, _
   -1, 2, 2, 2, -1, 2, -1, -1, 2, 2, 2, -1, 2, 2, 2, -1)

repeatgroupTable = Array( _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 2, 2, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 2, 2, 1, 1, 0, 0, 2, 2, 2, 2, 1, 1, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

modTable = Array( _
    1, 1, 1, 1, 2, 3, 1, 1, 1, 1, 1, 1, 2, 3, 1, 1, _
    1, 1, 1, 1, 2, 3, 1, 1, 1, 1, 1, 1, 2, 3, 1, 1, _
    1, 1, 1, 1, 2, 3, 1, 1, 1, 1, 1, 1, 2, 3, 1, 1, _
    1, 1, 1, 1, 2, 3, 1, 1, 1, 1, 1, 1, 2, 3, 1, 1, _
    4, 4, 4, 4, 5, 4, 4, 4, 4, 4, 4, 4, 5, 4, 4, 4, _
    4, 4, 4, 4, 5, 4, 4, 4, 4, 4, 4, 4, 5, 4, 4, 4, _
    4, 4, 4, 4, 5, 4, 4, 4, 4, 4, 4, 4, 5, 4, 4, 4, _
    4, 4, 4, 4, 5, 4, 4, 4, 4, 4, 4, 4, 5, 4, 4, 4, _
    6, 6, 6, 6, 7, 6, 6, 6, 6, 6, 6, 6, 7, 6, 6, 6, _
    6, 6, 6, 6, 7, 6, 6, 6, 6, 6, 6, 6, 7, 6, 6, 6, _
    6, 6, 6, 6, 7, 6, 6, 6, 6, 6, 6, 6, 7, 6, 6, 6, _
    6, 6, 6, 6, 7, 6, 6, 6, 6, 6, 6, 6, 7, 6, 6, 6, _
    8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, _
    8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, _
    8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, _
    8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8)

mod16Table = Array( _
    1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 2, 1, _
    1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 2, 1, _
    1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 2, 1, _
    1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 2, 1, _
    3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
    3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
    3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
    3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
    4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, _
    4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, _
    4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, _
    4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, _
    5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, _
    5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, _
    5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, _
    5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5)


sibTable = Array( _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2, _
    2, 2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 1, 2, 2)

regTable = Array( _
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, _
    2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, _
    4, 4, 4, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, _
    6, 6, 6, 6, 6, 6, 6, 6, 7, 7, 7, 7, 7, 7, 7, 7, _
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, _
    2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, _
    4, 4, 4, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, _
    6, 6, 6, 6, 6, 6, 6, 6, 7, 7, 7, 7, 7, 7, 7, 7, _
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, _
    2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, _
    4, 4, 4, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, _
    6, 6, 6, 6, 6, 6, 6, 6, 7, 7, 7, 7, 7, 7, 7, 7, _
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, _
    2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, _
    4, 4, 4, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, _
    6, 6, 6, 6, 6, 6, 6, 6, 7, 7, 7, 7, 7, 7, 7, 7)

rmTable = Array( _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7, _
    0, 1, 2, 3, 4, 5, 6, 7, 0, 1, 2, 3, 4, 5, 6, 7)
End Sub

'désassemble une instruction d'un octet
Public Function onebyteinstr(i As Instruction) As Instruction
'un octet à la position courante (opcode) et à la suivante (ModRM)
Dim b As Long, b2 As Long

    'sauvegarde de la valeur de l'instruction précédente
    'sauvegarde des registres EBP, ESP et (E)IP
    onebyteinstr.regIP = i.regIP
    'sauvegarde les préfixes
    onebyteinstr.operandSizeOverride = i.operandSizeOverride
    onebyteinstr.addressSizeOverride = i.addressSizeOverride
    onebyteinstr.LockAndRepeat = i.LockAndRepeat
    onebyteinstr.segmentOverride = i.segmentOverride
        
    'récupère l'octet de l'opcode
    b = getByte
    'suivante le type de l'opcode
    Select Case opcodeTable(b)
        Case 0
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
        Case 1
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'on récupère l'octet qui suit
            onebyteinstr.i_byte = getByte
            'calcule sa valeur signée
            If onebyteinstr.i_byte > 127& Then onebyteinstr.i_byte = -256& + onebyteinstr.i_byte
            If b = &HCD Then
                If onebyteinstr.i_byte = &H20 Then
                    onebyteinstr.i_dword = getDword(1)
                End If
            End If
        Case 2
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'on récupère le mot qui suit
            onebyteinstr.i_dword = getWord
        Case 3
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'le double mot
            onebyteinstr.i_dword = getWord
            'l'octet
            onebyteinstr.i_byte = getByte
            'signé
            If onebyteinstr.i_byte > 127& Then onebyteinstr.i_byte = -256& + onebyteinstr.i_byte
        Case 4
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'si operandOverride
            If onebyteinstr.operandSizeOverride = bOperandSizeOverride Then
                'le mot
                onebyteinstr.i_dword = getWord
            Else
                'sinon le double mot
                onebyteinstr.i_dword = getDword
            End If
        Case 44
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'si adressOverride
            If onebyteinstr.addressSizeOverride = bAddressSizeOverride Then
                'le mot
                onebyteinstr.i_dword = getWord
            Else
                'le double mot
                onebyteinstr.i_dword = getDword
            End If
        Case 5
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'le premier mot
            'TODO vérifier
            onebyteinstr.m_dword = getWord + regDS * 16
            'le double mot
            onebyteinstr.i_dword = getDword
        Case 6
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
        Case 7
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
            'un octet
            onebyteinstr.i_byte = getByte
            'signé
            If onebyteinstr.i_byte > 127& Then onebyteinstr.i_byte = -256& + onebyteinstr.i_byte
        Case 8
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
            'suivant la taille de l'opérande
            If onebyteinstr.operandSizeOverride = bOperandSizeOverride Then
                'un mot
                'TODO vérifier
                onebyteinstr.i_dword = getWord
            Else
                'ou un double mot
                onebyteinstr.i_dword = getDword
            End If
        Case 9
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
        Case 10
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
            'un octet
            onebyteinstr.i_byte = getByte
            'signé
            If onebyteinstr.i_byte > 127& Then onebyteinstr.i_byte = -256& + onebyteinstr.i_byte
        Case 11
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
            'suivant la taille de l'opérande
            If onebyteinstr.operandSizeOverride = bOperandSizeOverride Then
                'on mot
                'TODO vérifier
                onebyteinstr.i_dword = getWord
            Else
                'ou un double mot
                onebyteinstr.i_dword = getDword
            End If
        Case 12
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet ModRM
            getModRM onebyteinstr
        Case 13 'JUMP
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'l'octet qui suit : ModRM
            b = peekByte
            'si il y a un SIB
            If (b = 36) Then
                'on demande la valeur du SIB
                b2 = peekByte2
                'si EBP comme base
                If (rmTable(b2) = 5) Then
                    'on récupère l'octet ModRM
                    getModRM onebyteinstr 'opext
                    'onebyteinstr.bSib = getByte
                End If
            'sion pas de SIB
            Else
                If (regTable(b) < 7) Then
                    'l'octet ModRM
                    getModRM onebyteinstr
                End If
            End If
        Case 14 'TEST
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            'si F6 : octet
            If (b = 246) Then
                'ModRM
                b2 = peekByte
                'si registre EAX
                If (regTable(b2) = 0) Then
                    'l'octet ModRM
                    getModRM onebyteinstr 'opext
                    'un octet qui suit
                    onebyteinstr.i_byte = getByte
                    'signé
                    If onebyteinstr.i_byte > 127& Then onebyteinstr.i_byte = -256& + onebyteinstr.i_byte
                ElseIf (regTable(b2) > 1) Then
                    'l'octet ModRM
                    getModRM onebyteinstr 'opext
                End If
            'sinon F7 : dépend de la taille de l'opérande
            Else
                'ModRM
                b2 = peekByte
                'si reg EAX
                If (regTable(b2) = 0) Then
                    'l'octet ModRM
                    getModRM onebyteinstr 'opext
                    'suivant la taille de l'opérande
                    If onebyteinstr.operandSizeOverride = bOperandSizeOverride Then
                        'un mot
                        onebyteinstr.i_dword = getWord
                    Else
                        'un double mot
                        onebyteinstr.i_dword = getDword
                    End If
                'si autre que EAX
                ElseIf (regTable(b2) > 1) Then
                    'l'octet ModRM
                    getModRM onebyteinstr 'opext
                End If
            End If
        Case 15 'WAIT
            'on récupère l'opcode
            onebyteinstr.iOpcode = b
            b = peekByte
            If (b = 217) Then
                b2 = peekByte2
                If (regTable(b2) = 6) Or (regTable(b2) = 7) Then
                    onebyteinstr.iOpcode = getByte
                    getByte
                End If
            ElseIf (b = 219) Then
                b2 = peekByte2
                If (b2 = 226) Or (b2 = 227) Then
                    onebyteinstr.iOpcode = getByte
                    getModRM onebyteinstr
                End If
            ElseIf (b = 221) Then
                b2 = peekByte2
                If (regTable(b2) = 6) Or (regTable(b2) = 7) Then
                    onebyteinstr.iOpcode = getByte
                    getByte
                End If
            ElseIf (b = 223) Then
                b2 = peekByte2
                If (b2 = 224) Then
                    onebyteinstr.iOpcode = getByte
                    getModRM onebyteinstr
                End If
            End If
        Case 16 'REPEAT
            'dans les prefixes
            onebyteinstr.LockAndRepeat = b
            'REPNE
            If (b = 242) Then
                Do
                    'octet suivant
                    b2 = getByte
                    'si operand size override
                    If b2 = 102 Then
                        onebyteinstr.operandSizeOverride = b2
                    'si adress size override
                    ElseIf b2 = 103 Then
                        onebyteinstr.addressSizeOverride = b2
                    End If
                'tant que pas une instruction à part entière
                Loop While opcodeTable(b2) = 99
                'si REPNE autorisé
                If (repeatgroupTable(b2) = REPNEAllowed) Then
                    'on récupère l'opcode
                    onebyteinstr.iOpcode = b2
                End If
            'REPE/REP
            Else
                Do
                    'octet suivant
                    b2 = getByte
                    'si operand size override
                    If b2 = 102 Then
                        onebyteinstr.operandSizeOverride = b2
                    'si adress size override
                    ElseIf b2 = 103 Then
                        onebyteinstr.addressSizeOverride = b2
                    End If
                'tant que pas une instruction à part entière
                Loop While opcodeTable(b2) = 99
                'si REP ou REPNE autorisés
                If (repeatgroupTable(b2) > REPNotAllowed) Then
                    'on récupère l'opcode
                    onebyteinstr.iOpcode = b2
                End If
            End If
        'two byte instruction with overrides
        Case 98
            onebyteinstr = twobyteinstr(onebyteinstr)
        'Overrides(operand, segment address), LOCK et REPEAT
        Case 99
            Select Case b
                'operand override
                Case &H66
                    onebyteinstr.operandSizeOverride = &H66
                    onebyteinstr = onebyteinstr(onebyteinstr)
                'address override
                Case &H67
                    onebyteinstr.addressSizeOverride = &H67
                    onebyteinstr = onebyteinstr(onebyteinstr)
                Case &HF0, &HF2, &HF3 'LOCK and REPEAT
                    onebyteinstr.LockAndRepeat = b
                    onebyteinstr = onebyteinstr(onebyteinstr)
                'segment override
                Case &H2E, &H36, &H3E, &H26, &H64, &H65 'SEG
                    onebyteinstr.segmentOverride = b
                    onebyteinstr = onebyteinstr(onebyteinstr)
            End Select
    End Select
End Function

'désassemble les instructions sur deux octets
Public Function twobyteinstr(i As Instruction) As Instruction
    '
    Dim r As Long, b As Long, X As Long, Y As Long
    
    twobyteinstr.addressSizeOverride = i.addressSizeOverride
    twobyteinstr.LockAndRepeat = i.LockAndRepeat
    twobyteinstr.operandSizeOverride = i.operandSizeOverride
    twobyteinstr.segmentOverride = i.segmentOverride
    
    twobyteinstr.regIP = i.regIP
    
    'opcode : deuxième octet
    b = getByte
    Select Case opcode2Table(b)
        Case 0
            'on récupère l'opcode
            twobyteinstr.iOpcode = b
        Case 1
            'on récupère l'opcode
            twobyteinstr.iOpcode = b
            'suivant la taille des adresses
            If i.addressSizeOverride = bAddressSizeOverride Then
                'un mot
                twobyteinstr.i_dword = getWord
            Else
                'ou un double mot
                twobyteinstr.i_dword = getDword
            End If
        Case 2
            'on récupère l'opcode
            twobyteinstr.iOpcode = b
            'un ModRM
            getModRM twobyteinstr
        Case 3
            'on récupère l'opcode
            twobyteinstr.iOpcode = b
            'un ModRM
            getModRM twobyteinstr
            'un octet
            twobyteinstr.i_byte = getByte
            'signé
            If twobyteinstr.i_byte > 127& Then twobyteinstr.i_byte = -256& + twobyteinstr.i_byte
        Case 4
            'on récupère l'opcode
            twobyteinstr.iOpcode = b
            'un ModRM
            getModRM twobyteinstr 'opext
        Case 5
            'on récupère l'opcode
            twobyteinstr.iOpcode = b
            getModRM twobyteinstr 'opext
            'un octet
            twobyteinstr.i_byte = getByte
            'signé
            If twobyteinstr.i_byte > 127& Then twobyteinstr.i_byte = -256& + twobyteinstr.i_byte
    End Select
End Function

'met dans l'instruction, le ModRM qui convient
Private Sub getModRM(Ins As Instruction)
    'si pas de changement d'adresse
    If Ins.addressSizeOverride = bAddressSizeOverride Then
        'ModRM16
        modrm2 Ins
    Else
        'ModRM32
        modrm1 Ins
    End If
End Sub

'ModRM32
Public Sub modrm1(Ins As Instruction)
    Dim b As Long
    'on demande l'octet de ModRM
    b = peekByte
    'suivant son type
    Select Case modTable(b)
        Case 1
            'ModRM seul
            Ins.bModRm = getByte
        Case 2
            'ModRM
            Ins.bModRm = getByte
            'SIB
            Ins.bSib = getByte
            'éventuellement un double mot
            If (sibTable(Ins.bSib) = 1) Then
                 Ins.m_dword = getDword
            End If
        Case 3
            'ModRM
            Ins.bModRm = getByte
            'un double mot
            Ins.m_dword = getDword
        Case 4
            'ModRm
            Ins.bModRm = getByte
            'un octet
            Ins.m_byte = getByte
            'signé
            If Ins.m_byte > 127& Then Ins.m_byte = -256& + Ins.m_byte
        Case 5
            'ModRM
            Ins.bModRm = getByte
            'SIB
            Ins.bSib = getByte
            'un octet
            Ins.m_byte = getByte
            'signé
            If Ins.m_byte > 127& Then Ins.m_byte = -256& + Ins.m_byte
        Case 6
            'ModRM
            Ins.bModRm = getByte
            'double mot
            Ins.m_dword = getDword
        Case 7
            'ModRM
            Ins.bModRm = getByte
            'SIB
            Ins.bSib = getByte
            'double mot
            Ins.m_dword = getDword
        Case 8
            'ModRM
            Ins.bModRm = getByte
    End Select
End Sub

'ModRM16
Public Sub modrm2(Ins As Instruction)
    Dim b As Long
    'on demande le ModRM
    b = peekByte
    'suivant son type
    Select Case mod16Table(b)
        Case 1
            'ModRM seul
            Ins.bModRm = getByte
        Case 2
            'ModRM
            Ins.bModRm = getByte
            'double mot
            Ins.m_dword = getWord + regDS * 16
        Case 3
            'ModRM
            Ins.bModRm = getByte
            'octet
            Ins.m_byte = getByte
            'signé
            If Ins.m_byte > 127& Then Ins.m_byte = -256& + Ins.m_byte
        Case 4
            'ModRM
            Ins.bModRm = getByte
            'double mot
            Ins.m_dword = getWord + regDS * 16
        Case 5
            'ModRM seul
            Ins.bModRm = getByte
   End Select
End Sub

'renvoie la représentation numérique d'un opérande en fonction de l'instruction
'===================================================================================
'Ins : instruction de l'opérande
'OperandType : type de l'opérande
Public Function GetOperandNumber(Ins As Instruction, ByVal OperandType As Long)
    Select Case OperandType
        Case 2
            Ins.i_byte = Ins.i_byte + Ins.regIP
            GetOperandNumber = Ins.i_byte
        Case 3
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                If mod16Table(Ins.bModRm) = 8 Then
                    GetOperandNumber = rmTable(Ins.bModRm)
                End If
            Else
                If modTable(Ins.bModRm) = 8 Then
                    GetOperandNumber = rmTable(Ins.bModRm)
                End If
            End If
        Case 7, 11
            GetOperandNumber = Ins.i_byte
        Case 8, 12, 13, 15
            GetOperandNumber = Ins.i_dword
        Case 17
            Ins.i_dword = Ins.i_dword + Ins.regIP
            GetOperandNumber = Ins.i_dword
        Case 20, 4
            GetOperandNumber = rmTable(Ins.bModRm)
        Case 24, 5, 14
            GetOperandNumber = regTable(Ins.bModRm)
    End Select
End Function

'renvoie le texte de ModRM en fonction d'une instruction et d'une taille d'opérande
'==================================================================================
'Ins : instruction dont on veut le ModRM
Private Function getModRMAddress(Ins As Instruction) As Long
    If Ins.addressSizeOverride = bAddressSizeOverride Then
        If mod16Table(Ins.bModRm) = 2 Then
                getModRMAddress = Ins.m_dword
        End If
    Else
        Select Case modTable(Ins.bModRm)
            Case 2
                If sibTable(Ins.bSib) = 1 Then
                    If regTable(Ins.bSib) <= 4 Then
                        getModRMAddress = Ins.m_dword
                    End If
                End If
            Case 3, 6
                getModRMAddress = Ins.m_dword
            Case 7
                getModRMAddress = Ins.m_dword
        End Select
    End If
End Function

'renvoie la représentation numérique d'un opérande de type adresse en fonction de l'instruction
'==============================================================================================
'Ins : instruction de l'opérande
'OperandType : type de l'opérande
'dwSizePtr : renvoie une éventuelle taille de donnée de l'adresse contenue dans l'opérande
Public Function GetOperandAddress(Ins As Instruction, ByVal OperandType As Long, ByRef dwSizePtr As Long) As Long
    Select Case OperandType
        Case 1
            dwSizePtr = 8
            GetOperandAddress = getModRMAddress(Ins)
        Case 3
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                dwSizePtr = 16
            Else
                dwSizePtr = 32
            End If
            GetOperandAddress = getModRMAddress(Ins)
        Case 6
            dwSizePtr = 16
            GetOperandAddress = getModRMAddress(Ins)
        Case 7
            dwSizePtr = 8
            GetOperandAddress = Ins.i_dword
        Case 8
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                dwSizePtr = 16
                GetOperandAddress = Ins.i_dword
            Else
                dwSizePtr = 32
                GetOperandAddress = Ins.i_dword
            End If
'        Case 11
'            GetOperandAddress = Ins.i_byte
        Case 12
            dwSizePtr = 0
            GetOperandAddress = Ins.i_dword
        Case 13
            dwSizePtr = 0
            GetOperandAddress = Ins.i_dword
        Case 19
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                dwSizePtr = 16
            Else
                dwSizePtr = 32
            End If
            GetOperandAddress = getModRMAddress(Ins)
        Case 21, 23
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                dwSizePtr = 16
            Else
                dwSizePtr = 32
            End If
            GetOperandAddress = getModRMAddress(Ins)
        'two bytes opcode
        Case 30
            dwSizePtr = 32
            GetOperandAddress = getModRMAddress(Ins)
        Case Else
            dwSizePtr = 0
            GetOperandAddress = 0
    End Select
End Function

'décode une instruction
'Ins : instruction précédente
Public Function Decode(Ins As Instruction) As Instruction
Dim b As Long 'opcode
Dim i As Instruction 'instruction
Dim opext As Long, Size As Long, strIns As String
Dim lpdata1 As Long, lpdata2 As Long
Dim s1 As Long, s2 As Long
Dim dw As Long

'on demande l'opcode
b = peekByte
'si instruction sur deux octets
If b = &HF& Then
    'on supprime le premier octet
    getByte
    'on decode l'instruction en deux octets
    Decode = twobyteinstr(i)
    'on indique : instruction sur deux octets
    Decode.opclass = &HF
Else
    'sinon instruction sur un octet
    Decode = onebyteinstr(i)
End If

With Decode
    'conservation des registres IP, EBP, ESP
    .regIP = getPointerVA
    .regESP = regESP

    'si changement de taille d'opérande
    If .operandSizeOverride = bOperandSizeOverride Then
        'mot
        Size = 2
    Else
        'double mot
        Size = 4
    End If
    
    
    'gestion des sauts et appels de procédures
    Dim bData As Boolean
    
    'si instruction sur deux octets
    If .opclass = &HF Then
        lpdata1 = GetOperandAddress(Decode, firstOperandTwoType(.iOpcode), s1)
        lpdata2 = GetOperandAddress(Decode, secondOperandTwoType(.iOpcode), s2)
        
        'choix parmi les opcodes recherchés
        Select Case .iOpcode
            'les sauts
            Case &H80 To &H8F 'Jcc
                jCol.Add GetOperandNumber(Decode, firstOperandTwoType(.iOpcode)) & ":" & regESP & ":" & regEBP
                bData = False
            'instruction de pile
            Case &HA1, &HA9 'POP
                regESP = regESP - 2
                bData = True
            Case &HA0, &HA8 'PUSH
                regESP = regESP + 2
                bData = True
            Case Else
                bData = True
        End Select
    'sinon instruction sur un octet
    Else
        lpdata1 = GetOperandAddress(Decode, firstOperandType(.iOpcode), s1)
        lpdata2 = GetOperandAddress(Decode, secondOperandType(.iOpcode), s2)
        
        'choix parmi les opcodes recherchés
        Select Case .iOpcode
            Case &HCD 'INT 20h
                .bStop = (.i_byte = &H20) And ((.i_dword = &H180BE) Or (.i_dword = &H180BF))
            Case &HE8 'CALL
                dw = GetOperandNumber(Decode, firstOperandType(.iOpcode))  'CLng("&H" & Mid$(strIns, 1, Len(strIns) - 1))
                callCol.Add dw
                AddSubName dw
                bData = False
            Case &H9A 'CALL
                dw = GetOperandNumber(Decode, firstOperandType(.iOpcode)) 'CLng("&H" & Mid$(strIns, 1, Len(strIns) - 1))
                callCol.Add dw
                AddSubName dw
                bData = False
            Case &H70 To &H7F, &HE3 'Jcc
                dw = GetOperandNumber(Decode, firstOperandType(.iOpcode)) 'CLng("&H" & Mid$(strIns, 1, Len(strIns) - 1))
                jCol.Add dw & ":" & regESP & ":" & regEBP
                bData = False
            Case &HE0 To &HE2 'LOOPcc
                dw = GetOperandNumber(Decode, firstOperandType(.iOpcode))
                jCol.Add dw & ":" & regESP & ":" & regEBP
                bData = False
            Case &HEB 'JUMP 8
                dw = GetOperandNumber(Decode, firstOperandType(.iOpcode))
                setPointerVA dw
                jCol.Add dw & ":" & regESP & ":" & regEBP
                bData = False
            Case &HE9 'JUMP 16/32
                dw = GetOperandNumber(Decode, firstOperandType(.iOpcode))
                setPointerVA dw
                jCol.Add dw & ":" & regESP & ":" & regEBP
                bData = False
            Case &HFF 'PUSH
                opext = (.bModRm And 56) / 8
                If opext = 6 Then
                    regESP = regESP + Size
                ElseIf opext = 5 Then 'JMPF
                    'TODO
                ElseIf opext = 4 Then 'JMPN
                    .bStop = ProcessPointer(.m_dword, True)
                    bData = False
                ElseIf opext = 3 Then 'CALLF
                    'TODO
                ElseIf opext = 2 Then 'CALLN
                    'TODO
                End If
                bData = True
            Case &H50 To &H57, &H68   'PUSH
                regESP = regESP + Size
                bData = True
            Case &HE, &H16, &H1E, &H6 'PUSH reg16
                regESP = regESP + 2
                bData = True
            Case &H6A 'PUSH imm8
                regESP = regESP + Size
                bData = True
            Case &H60 'PUSHA / PUSHAD
                regESP = regESP + 8 * Size
                bData = False
            Case &H9C 'PUSHF / PUSHFD
                regESP = regESP + Size
                bData = False
            Case &H58 To &H5F, &H8F   'POP
                regESP = regESP - Size
                bData = True
            Case &H17, &H1F, &H7  'POP reg16
                regESP = regESP - 2
                bData = True
            Case &H61 'POPA / POPAD
                regESP = regESP - 8 * Size
                bData = False
            Case &H9D 'POPF / POPFD
                regESP = regESP - Size
                bData = False
            Case &H80, &H81, &H83  ' ADD
                bData = True
                b = GetOperandNumber(Decode, firstOperandType(.iOpcode))
                If b = 4 Then 'ESP
                    b = GetOperandNumber(Decode, secondOperandType(.iOpcode))
                    Select Case (.bModRm And &H38) \ 8
                        Case 0
                            regESP = regESP - b
                        Case 1
                            regESP = regESP Or b
                        Case 2
                            regESP = regESP - b
                        Case 3
                            regESP = regESP + b
                        Case 4
                            regESP = regESP And b
                        Case 5
                            regESP = regESP + b
                        Case 6
                            regESP = regESP Xor b
                    End Select
                ElseIf b = 5 Then 'EBP
                    Select Case (.bModRm And &H38) \ 8
                        Case 0
                            regEBP = regEBP - b
                        Case 1
                            regEBP = regEBP Or b
                        Case 2
                            regEBP = regEBP - b
                        Case 3
                            regEBP = regEBP + b
                        Case 4
                            regEBP = regEBP And b
                        Case 5
                            regEBP = regEBP + b
                        Case 6
                            regEBP = regEBP Xor b
                    End Select
                End If
            Case &H8B 'MOV
                bData = True
                If .bModRm = &HEC Then
                    regEBP = regESP
                ElseIf .bModRm = &HE5 Then
                    regESP = regEBP
                End If
            Case &HB0 'MOV AL,Ib
                regEAX = regEAX And &HFFFFFF00
                regEAX = regEAX Or .i_byte
            Case &HB4 'MOV AH,Ib
                regEAX = regEAX And &HFFFF00FF
                regEAX = regEAX Or CLng(.i_byte * &H100)
            Case &HB8 'MOV eAX,Iv
                If .operandSizeOverride = bOperandSizeOverride Then
                    regEAX = regEAX And &HFFFF0000
                    regEAX = regEAX Or .i_dword
                Else
                    regEAX = .i_dword
                End If
            Case &H8E
                If (regTable(.bModRm) = 3) And (mod16Table(.bModRm) = 5) And (rmTable(.bModRm) = 0) And (bOperandSizeOverride = 0) Then
                    regDS = regEAX + &H1000
                End If
            Case &HCD
                If ((regEAX And &H4C00&) = &H4C00&) And (.i_byte = &H21) Then
                    .bStop = True
                End If
            'TODO LEA et ESP et EBP
            'Case &H8D 'LEA
            Case Else
                bData = True
        End Select
    End If
    .regIP = getPointerVA
End With

'gestion des données
If bData Then
    Dim dt As Long
    If CheckVA(lpdata1) Then
        Select Case s1
            Case 0 'offset
                dt = GetDataType(lpdata1, 0)
                Select Case dt
                    Case 0
                        tryCallCol.Add lpdata1
                    Case 3 'numérique taille inconnue
                        setMapVA lpdata1, 3
                    Case 4 'pointeur 4 octets
                        ProcessPointer lpdata1
                    Case 5 'SZ
                        setMapVA lpdata1, 5
                    Case 7 'PASCAL
                        setMapVA lpdata1, 7
                    Case 10 'UNICODE
                        setMapVA lpdata1, 10
                End Select
            Case 8
                setMapVA lpdata1, 30
            Case 16
                setMapVA lpdata1, 31
            Case 32
                setMapVA lpdata1, 32
                If IsCodeVA(lpdata1) Then
                    ProcessPointer lpdata1
                End If
            Case 64
                setMapVA lpdata1, 33
        End Select
    End If
    If CheckVA(lpdata2) Then
        Select Case s2
            Case 0 'offset
                dt = GetDataType(lpdata2, 0)
                Select Case dt
                    Case 0
                        tryCallCol.Add lpdata2
                        'ProcessPointer lpdata2
                    Case 3 'numérique taille inconnue
                        setMapVA lpdata2, 3
                    Case 4 'pointeur 4 octets
                        ProcessPointer lpdata2
                    Case 5 'SZ
                        setMapVA lpdata2, 5
                    Case 7 'PASCAL
                        setMapVA lpdata2, 7
                    Case 10 'UNICODE
                        setMapVA lpdata2, 10
                End Select
            Case 8
                setMapVA lpdata2, 30
            Case 16
                setMapVA lpdata2, 31
            Case 32
                setMapVA lpdata2, 32
                If IsCodeVA(lpdata2) Then
                    ProcessPointer lpdata2
                End If
            Case 64
                setMapVA lpdata2, 33
        End Select
    End If
End If
End Function

'parcourt un tableau de pointeur
Public Function ProcessPointer(ByVal va As Long, Optional ByVal bSetFirstOK As Boolean = False) As Boolean
    Dim dw As Long, cnt As Long, bVA As Boolean, off As Long
    
    off = VA2Offset(va)
    dw = getDwordOffset(off)
    'cnt = 0
    ProcessPointer = True
    Do While CheckVA(dw) 'Or (cnt = 0)
        If CheckVA(dw) Then
            setMapOffset off, 4
            If IsCodeVA(dw) Then
                If bSetFirstOK Then
                    setPointerVA dw
                    ProcessPointer = False
                    bSetFirstOK = False
                End If
                tryCallCol.Add dw
                
                'If getMapVA(dw) = 0 Then setMapVA dw, 254
                
                'cnt = 1
                
                dw = getDwordVA(dw)
                If CheckVA(dw) And (dw <> va) Then
                    ProcessPointer dw, False
                End If
            End If
        End If
        
        off = off + 4
        dw = getDwordOffset(off)
    Loop
End Function

'désassemble le code à l'emplacement indiqué
'===========================================
'iFileNum : numéro de fichier du listing ASM
'dwStartingAddress : adresse de départ du désassemblage
'dwVABase : indique la base des adresses virtuelles relatives de l'exécutable
'bProcessCall :  indique s'il faut descendre dans les procédures rencontrées
Public Sub DysCode(ByVal iFileNum As Integer, ByVal dwStartingAddress As Long, Optional bProcessCall As Boolean = False, Optional strProcName As String = "")
Dim X As Long, ipret As Long, addr As Long, b As Byte, spec() As String
Dim strName As String, cnt As Long

'on regarde si l'on a déjà désassembler
b = getMapVA(dwStartingAddress)
'si on a déjà été dans cette procédure, on ne recommence pas
If b Then Exit Sub

'imprime le titre de la procédure dans le fichier
Print #iFileNum, getNumber(dwStartingAddress, 8); ":0"
If Len(strProcName) Then
    'sous la forme : VA nom_procédure
    Print #iFileNum, getNumber(dwStartingAddress, 8); ":0", strProcName, "proc"
Else
    'sous la forme : VA sub_VA
    Print #iFileNum, getNumber(dwStartingAddress, 8); ":0", "sub_"; getNumber(dwStartingAddress, 8), "proc"
End If

'désassemble la procédure
ipret = DysassembleSub(iFileNum, dwStartingAddress, True)

'lit le liste des arguments trouvés
cnt = argCol.Count
For X = 1 To cnt
    'pour chaque argument trouvé
    'on sépare la structure mise dans la collection
    'spec(0) = numéro de l'argument : offset par rapport au registre de base de la pile pour la procédure (ESP ou EBP)
    'spec(1) = spécificateur de taille de donnée de l'argument (byte, word ou dword ptr)
    'spec(2) = offset par rapport au début de la pile en entrant dans la procédure : registre ESP
    spec = Split(argCol(X), ":")
    'on imprime l'argument dans le fichier
    'sous la forme : VA arg_numéro = taille ptr offset
    Print #iFileNum, getNumber(dwStartingAddress, 8); ":1", "arg_"; Hex$(CLng(spec(0))), " = "; getSpecifier(CLng(spec(1)) * 8, 0); spec(2); "H"
Next

'lit le liste des var locales trouvées
cnt = varCol.Count
For X = 1 To cnt
    'pour chaque variable trouvée
    'on sépare la structure mise dans la collection
    'spec(0) = numéro de la variable : offset par rapport au registre de base de la pile pour la procédure (ESP ou EBP)
    'spec(1) = spécificateur de taille de donnée de la variable (byte, word ou dword ptr)
    'spec(2) = offset par rapport au début de la pile en entrant dans la procédure : registre ESP
    spec = Split(varCol(X), ":")
    'on imprime l'argument dans le fichier
    'sous la forme : VA var_numéro = taille ptr offset
    Print #iFileNum, getNumber(dwStartingAddress, 8); ":1", "var_"; Hex$(CLng(spec(0))), " = "; getSpecifier(CLng(spec(1)) * 8, 0); "-" & spec(2); "H"
Next

'on imprime la finde procédure dans le fichier
If Len(strProcName) Then
    'sous la forme : VAfin nom_procédure endp
    Print #iFileNum, getNumber(ipret, 8); ":4", strProcName, "endp"
Else
    'sous la forme : VAfin sub_VAdeb endp
    Print #iFileNum, getNumber(ipret, 8); ":4", "sub_"; getNumber(dwStartingAddress, 8), "endp"
End If
Print #iFileNum, getNumber(ipret, 8); ":5"

'si l'on doit parcourir les procédures recontrées
If bProcessCall Then
    'pour chaque CALL
    cnt = callCol.Count
    For X = 1 To cnt
        Set argCol = New Collection
        Set varCol = New Collection
        'on récupère l'adresse
        addr = CLng(callCol(X))
        'on demande le type de donnée à l'adresse pointée par le call
        b = getMapVA(addr)
        'si on n'a pas encore été dans cette procédure
        If (b = 0) Then
            strName = GetSubName(addr)
            'on la désassemble par récurrence
            DysCode iFileNum, addr, True, strName
        End If
    Next
End If
End Sub

'désassemble une procédure
'=============================================================================
'iFileNum : numéro de fichier pour le fichier ASM
'dwStartingAddress : adresse de départ du désassemblage
'dwVABase : indique le base des adresses virtuelles relatives de l'exécutable
'bClear : indique s'il est nécessaire de réinitialiser la liste des sauts
'renvoie l'adresse IP à la fin de la procédure
Public Function DysassembleSub(ByVal iFileNum As Integer, ByVal dwStartingAddress As Long, Optional bClear As Boolean = False) As Long
Dim X As Long 'var de controle
Dim j() As String
Dim addr As Long 'adresse
Dim b As Byte 'un octet
Dim ip As Long 'adresse d'instruction
Static ipret As Long 'adresse de retour
Dim i As Instruction 'instruction
Dim jcount As Long, ptr As Long
Dim szJName As String

'si l'on doit effacer les sauts
If bClear Then
    'et le retour
    ipret = 0
    'et les sauts
    Set jCol = New Collection
    regESP = dwInitESP
    regEBP = 0
End If

'on se place au début de la procédure
setPointerVA dwStartingAddress

'on démarre dans la procédure avec ESP = 4 (juste, sur la pile, la place de l'adresse de retour de la procédure
i.regIP = dwStartingAddress

b = getMapVA(dwStartingAddress)
Do
    'on récupère le pointeur d'instruction suivante
    ip = i.regIP
    'reinit
    i.addressSizeOverride = 0
    i.LockAndRepeat = 0
    i.operandSizeOverride = 0
    i.segmentOverride = 0
    'on decode l'instruction suivante avec une trace la dernière
    i = Decode(i)
    'on imprime l'instruction dans le fichier ASM
    'sous la forme : VA opcode instruction
    Print #iFileNum, getNumber(ip, 8); ":3", , GetTextInstruction(i)
    If ip > ipret Then ipret = ip
'tant que l'on ne trouve pas de RET ou RETN ou que l'on n'atteint pas une zone déjà parcourue
Loop Until (i.iOpcode = &HC3) Or (i.iOpcode = &HC2) Or (i.iOpcode = &HCC) Or i.bStop Or getMap
If b = 2 Then setMapVA dwStartingAddress, b

'si on a atteint un RET ou RETN et que l'on a dépassé la valeur de fin de procédure précédemment stockée
'on stocke la nouvelle valeur de fin de procédure
'cela permet de connaitre la fin de la procédure même si elle possède plusieurs RET ou RETN
'If ((i.iOpcode = &HC3) Or (i.iOpcode = &HC2)) And (i.regIP > ipret) Then ipret = ip

'pour chaque saut
jcount = jCol.Count
For X = 1 To jcount
    'on récupère son adresse
    j = Split(jCol(X), ":")
    addr = j(0)
    regESP = j(1)
    regEBP = j(2)
    'on regarde le type de donnée à cet emplacement
    b = getMapVA(addr)
    'si c'est du code déjà parcourut ou pas encore parcouru (mais pas des données)
    If b = 0 Then
        szJName = GetName(addr)
        If Len(szJName) Then
            Print #iFileNum, getNumber(addr, 8); ":2", szJName; ":"
        Else
            'on imprime une étiquette pour l'emplacement (dans le fichier ASM)
            'sous la forme VA loc_VA:
            Print #iFileNum, getNumber(addr, 8); ":2", "loc_"; getNumber(addr, 8); ":"
        End If
        'on fixe le type de code à 2 = étiquette
        setMapVA addr, 2
        'on désassemble à partir de l'endroit du saut
        DysassembleSub iFileNum, addr
        'si on a encore dépassé la fin précédente de la procédure : cela est la nouvelle fin
        'If getPointer > ipret Then ipret = getPointer
    ElseIf b = 1 Then
        szJName = GetName(addr)
        If Len(szJName) Then
            Print #iFileNum, getNumber(addr, 8); ":2", szJName; ":"
        Else
            'on imprime une étiquette pour l'emplacement (dans le fichier ASM)
            'sous la forme VA loc_VA:
            Print #iFileNum, getNumber(addr, 8); ":2", "loc_"; getNumber(addr, 8); ":"
        End If
        'on fixe le type de code à 2 = étiquette
        setMapVA addr, 2
    ElseIf b = 4 Then
        setMapVA addr, 0
        b = 0
    End If
Next
'on renvoie la fin effective de la procédure
DysassembleSub = ipret
End Function

'Public Function ProcessDeadcode()
'    Dim x As Long, addrdeb As Long, addrfin As Long, addr As Long, b As Byte
'
'    For x = 0 To UBound(retSectionTables)
'        With retSectionTables(x)
'            If ((.Characteristics And IMAGE_SCN_CNT_CODE) = IMAGE_SCN_CNT_CODE) And _
'               ((.Characteristics And IMAGE_SCN_MEM_EXECUTE) = IMAGE_SCN_MEM_EXECUTE) Then
'                addrdeb = .PointerToRawData
'                addrfin = .PointerToRawData + .VirtualSize
'                For addr = addrdeb To addrfin
'                    b = getMap(addr)
'                    If b = 0 Then
'                Next
'            End If
'        End With
'    Next
'End Function
