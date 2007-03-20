Attribute VB_Name = "modPrint"
Option Explicit

'tables des représentations textuelles
'=====================================
'des instructions sur un octet
Public strOneByteInstruction()
'des instructions sur deux octets
Dim strTwoByteInstruction()
'des groupes d'instructions des extensions d'opcode
Dim strGroupExtensions()
'des groupes d'instruction Escape : instructions mathématiques
Dim strEscapeGroup0()
Dim strEscapeGroup1()
Dim strEscapeGroup2()
Dim strEscapeGroup3()
Dim strEscapeGroup4()
Dim strEscapeGroup5()
Dim strEscapeGroup6()
Dim strEscapeGroup7()
'regroupement des groupes précédents
Dim strEscape()

'type des opérandes pour les instructions sur un octet
Public firstOperandType()
Public secondOperandType()
'type des opérandes pour les instructions sur deux octets
Public firstOperandTwoType()
Public secondOperandTwoType()

'collection des arguments d'une procédure
Public argCol As Collection
Public varCol As Collection

'initialisation des tables
Public Function InitPrint()
strOneByteInstruction = Array( _
    "ADD *", "ADD *", "ADD *", "ADD *", "ADD AL,*", "ADD eAX,*", "PUSH ES", "POP ES", "OR *", "OR *", "OR *", "OR *", "OR AL,*", "OR eAX,*", "PUSH CS", "", _
    "ADC *", "ADC *", "ADC *", "ADC *", "ADC AL,*", "ADC eAX,*", "PUSH SS", "POP SS", "SBB *", "SBB *", "SBB *", "SBB *", "SBB AL,*", "SBB eAX,*", "PUSH DS", "POP DS", _
    "AND *", "AND *", "AND *", "AND *", "AND AL,*", "AND eAX,*", "SEG=ES", "DAA", "SUB *", "SUB *", "SUB *", "SUB *", "SUB AL,*", "SUB eAX,*", "SEG=CS", "DAS", _
    "XOR *", "XOR *", "XOR *", "XOR *", "XOR AL,*", "XOR eAX,*", "SEG=SS", "AAA", "CMP *", "CMP *", "CMP *", "CMP *", "CMP AL,*", "CMP eAX,*", "SEG=DS", "AAS", _
    "INC eAX", "INC eCX", "INC eDX", "INC eBX", "INC eSP", "INC eBP", "INC eSI", "INC eDI", "DEC eAX", "DEC eCX", "DEC eDX", "DEC eBX", "DEC eSP", "DEC eBP", "DEC eSI", "DEC eDI", _
    "PUSH eAX", "PUSH eCX", "PUSH eDX", "PUSH eBX", "PUSH eSP", "PUSH eBP", "PUSH eSI", "PUSH eDI", "POP eAX", "POP eCX", "POP eDX", "POP eBX", "POP eSP", "POP eBP", "POP eSI", "POP eDI", _
    "PUSHAD", "POPAD", "BOUND *", "ARPL *", "SEG=FS", "SEG=GS", "Opd Size", "Addr Size", "PUSH *", "IMUL *", "PUSH *", "IMUL *", "INSB *,DX", "INSD *,DX", "OUTSB DX,*", "OUTSD DX,*", _
    "JO *", "JNO *", "JB *", "JAE *", "JE *", "JNE *", "JBE *", "JA *", "JS *", "JNS *", "JPE *", "JPO *", "JL *", "JGE *", "JLE *", "JG *", _
    "", "", "", "", "TEST *", "TEST *", "XCHG *", "XCHG *", "MOV *", "MOV *", "MOV *", "MOV *", "MOV *", "LEA *", "MOV *", "POP *", _
    "NOP", "XCHG eAX,eCX", "XCHG eAX,eDX", "XCHG eAX,eBX", "XCHG eAX,eXP", "XCHG eAX,eBP", "XCHG eAX,eSI", "XCHG eAX,eDI", "CBW", "CWD", "CALLF *", "WAIT *", "PUSHFD", "POPFD", "SAHF", "LAHF", _
    "MOV AL,*", "MOV eAX,*", "MOV *,AL", "MOV *,eAX", "MOVSB *", "MOVS *", "CMPSB *", "CMPS *", "TEST AL,*", "TEST eAX,*", "STOSB *,AL", "STOS *,eAX", "LODSB AL,*", "LODS eAX,*", "SCASB AL,*", "SCAS eAX,*", _
    "MOV AL,*", "MOV CL,*", "MOV DL,*", "MOV BL,*", "MOV AH,*", "MOV CH,*", "MOV DH,*", "MOV BH,*", "MOV eAX,*", "MOV eCX,*", "MOV eDX,*", "MOV eBX,*", "MOV eSP,*", "MOV eBP,*", "MOV eSI,*", "MOV eDI,*", _
    "", "", "RETN *", "RETN", "LES *", "LDS *", "", "", "ENTER *", "LEAVE", "RETF *", "RETF", "INT 3", "INT *", "INTO", "IRET", _
    "", "", "*,CL", "*,CL", "AAM *", "AAD *", "", "XLAT", "", "", "", "", "", "", "", "", _
    "LOOPNE *", "LOOPE *", "LOOP *", "JCXZ *", "IN AL,*", "IN eAX,*", "OUT *,AL", "OUT *,eAX", "CALL *", "JMP *", "JMP *", "JMP *", "IN AL,DX", "IN eAX,DX", "OUT DX,AL", "OUT DX,eAX", _
    "LOCK", "", "REPNE", "REP", "HLT", "CMC", "", "", "CLC", "STC", "CLI", "STI", "CLD", "STD", "", "")

strTwoByteInstruction = Array( _
    "", "", "LAR *", "LSL *", "", "", "CLTS", "", "INVD", "WBINVD", "", "", "", "", "", "", _
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
    "MOV *", "MOV *", "MOV *", "MOV *", "", "", "", "", "", "", "", "", "", "", "", "", _
    "WRMSR *", "RDTSC", "RDMSR", "RDPMC", "SYSENTER", "SYSEXIT", "", "", "", "", "", "", "", "", "", "", _
    "CMOVO *", "CMOVNO *", "CMOVB *", "CMOVAE *", "CMOVE *", "CMOVNE *", "CMOVBE *", "CMOVA *", "CMOVS *", "CMOVNS *", "CMOVPE *", "CMOVPO *", "CMOVL *", "CMOVGE *", "CMOVLE *", "CMOVG *", _
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
    "punpcklbw *", "punpcklwd *", "punpckldq *", "packsswb *", "pcmpgtb *", "pcmpgtw *", "pcmpgtd *", "packuswb *", "punpckhbw *", "punpckhwd *", "punpckhdq *", "packssdw *", "", "", "movd *", "movq *", _
    "", "pshimw *", "pshimd *", "pshimq *", "pcmpeqb *", "pcmpeqw *", "pcmpeqd *", "emms *", "", "", "", "", "", "", "movd *", "movq *", _
    "JO *", "JNO *", "JB *", "JAE *", "JE *", "JNE *", "JBE *", "JA *", "JS *", "JNS *", "JPE *", "JPO *", "JL *", "JGE *", "JLE *", "JG *", _
    "SETO *", "SETNO *", "SETB *", "SETAE *", "SETE *", "SETNE *", "SETBE *", "SETA *", "SETS *", "SETNS *", "SETPE *", "SETPO *", "SETL *", "SETGE *", "SETLE *", "SETG *", _
    "PUSH FS", "POP FS", "CPUID", "BT *", "SHLD *", "SHLD *,CL", "", "", "PUSH GS", "POP GS", "RSM", "BTS *", "SHRD *", "SHRD *,CL", "", "IMUL *", _
    "CMPXCHG *", "CMPXCHG *", "LSS *", "BTR *", "LFS *", "LGS *", "MOVZX *", "MOVZX *", "", "", "", "BTC *", "BSF *", "BSR *", "MOVSX *", "MOVSX *", _
    "XADD *", "XADD *", "", "", "", "", "", "", "BSWAP EAX", "BSWAP ECX", "BSWAP EDX", "BSWAP EBX", "BSWAP ESP", "BSWAP EBP", "BSWAP ESI", "BSWAP EDI", _
    "", "psrlw *", "psrld *", "psrlq *", "", "pmullw", "", "", "psubusb *", "psubusw *", "", "pand *", "paddusb *", "paddusw *", "", "pandn *", _
    "", "psraw *", "psrad *", "", "", "psmulhw *", "", "", "psubsb *", "psubsw *", "", "por *", "paddsb *", "paddsw *", "", "pxor *", _
    "", "psllw *", "pslld *", "psllq *", "", "pmaddwd *", "", "", "psubb *", "psubw *", "psubd *", "", "paddb *", "paddw *", "paddd *", "")

strGroupExtensions = Array( _
    Array("ADD *", "OR *", "ADC *", "SBB *", "AND *", "SUB *", "XOR *", "CMP *"), _
    Array("ROL *", "ROR *", "RCL *", "RCR *", "SHL *", "SHR *", "", "SAR *"), _
    Array("TEST *", "", "NOT *", "NEG *", "MUL *,", "IMUL *,", "DIV *,", "IDIV *,"), _
    Array("INC *", "DEC *", "", "", "", "", "", ""), _
    Array("INC *", "DEC *", "CALLN *", "CALLF *", "JMPN *", "JMPF *", "PUSH *", ""), _
    Array("SLDT *", "STR *", "LLDT *", "LTR *", "VERR *", "VERW *", "", ""), _
    Array("SGDT *", "SIDT *", "LGDT *", "LIDT *", "SMSW *", "", "LMSW *", "INVLPG "), _
    Array("", "", "", "", "BT *", "BTS *", "BTR *", "BTC "), _
    Array("", "CMPXCH8 *", "", "", "", "", "", ""), _
    Array("", "", "", "", "", "", "", ""), _
    Array("MOV *", "", "", "", "", "", "", ""), _
    Array("", "", "psrtw *", "", "psraw *", "", "pslld *", ""), _
    Array("", "", "psrld *", "", "psrad *", "", "pslld *", ""), _
    Array("", "", "psrlq *", "", "", "", "psllq *", ""), _
    Array("fxsave", "fxrstor *", "ldmxcsr *", "stmxcsr *", "", "", "", "sfence "), _
    Array("prefetch NTA", "prefetch T0", "prefetch T1", "prefetch T2", "", "", "", ""))

'pour les types d'opérandes :
'-1 : aucun
'0 : fixé
'1: Eb
'2: rel8/Jb
'3: Ev
'4: Gb
'5: Gv
'6: Ew
'7: Ob
'8: Ov
'9: Xb
'10: Xv
'11: Ib
'12: Iw
'13: Iv
'14: Sw
'15: Ap
'16: Fv
'17: Jv
'18: Grp7
'19: Ma
'20: Gw
'21: Mp
'22: 1
'23: M
'24: Rd
'25: Cd
'26: Dd
'27: Pq
'28: Mq
'29: Pd
'30: Ed
'31: Qq
'32: Qd
'99: Escape
'90: Yb
'100: Yv
firstOperandType = Array( _
    1, 3, 4, 5, 0, 0, -1, -1, 1, 3, 4, 5, 0, 0, -1, -1, _
    1, 3, 4, 5, 0, 0, -1, -1, 1, 3, 4, 5, 0, 0, -1, -1, _
    1, 3, 4, 5, 0, 0, -1, -1, 1, 3, 4, 5, 0, 0, -1, -1, _
    1, 3, 4, 5, 0, 0, -1, -1, 1, 3, 4, 5, 0, 0, -1, -1, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    -1, -1, 5, 6, -1, -1, -1, -1, 13, 5, 11, 5, 90, 100, 0, 0, _
    2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
    1, 3, 3, 3, 1, 3, 1, 3, 1, 3, 4, 5, 6, 5, 14, 3, _
    -1, 0, 0, 0, 0, 0, 0, 0, -1, -1, 15, -1, 16, 16, -1, -1, _
    0, 0, 7, 8, 9, 10, 9, 10, 0, 0, 90, 100, 0, 0, 0, 0, _
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
    1, 3, 12, -1, 5, 5, 1, 3, 12, -1, 12, -1, -1, 11, -1, -1, _
    1, 3, 1, 3, 11, 11, -1, -1, 99, 99, 99, 99, 99, 99, 99, 99, _
    2, 2, 2, 2, 0, 0, 11, 11, 17, 17, 15, 2, 0, 0, 0, 0, _
    -1, -1, -1, -1, -1, -1, 1, 3, -1, -1, -1, -1, -1, -1, 1, 3)

secondOperandType = Array( _
    4, 5, 1, 3, 11, 13, -1, -1, 4, 5, 1, 3, 11, 13, -1, -1, _
    4, 5, 1, 3, 11, 13, -1, -1, 4, 5, 1, 3, 11, 13, -1, -1, _
    4, 5, 1, 3, 11, 13, -1, -1, 4, 5, 1, 3, 11, 13, -1, -1, _
    4, 5, 1, 3, 11, 13, -1, -1, 4, 5, 1, 3, 11, 13, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, 19, 20, -1, -1, -1, -1, -1, 3, -1, 3, 0, 0, 9, 10, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    11, 13, 11, 11, 4, 5, 4, 5, 4, 5, 1, 3, 14, 23, 6, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    7, 8, 0, 0, 90, 100, 90, 100, 11, 13, 0, 0, 9, 10, 90, 10, _
    11, 11, 11, 11, 11, 11, 11, 11, 13, 13, 13, 13, 13, 13, 13, 13, _
    11, 11, -1, -1, 21, 21, 11, 13, 11, -1, -1, -1, -1, -1, -1, -1, _
    22, 22, 0, 0, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, 11, 11, 0, 0, -1, -1, -1, -1, 0, 0, 0, 0, _
    -1, -1, -1, -1, -1, -1, 0, 0, -1, -1, -1, -1, -1, -1, -1, -1)

firstOperandTwoType = Array( _
    6, 18, 5, 5, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    24, 24, 25, 26, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, 27, -1, -1, 29, 27, _
    27, 27, 27, 27, 27, 27, 27, -1, -1, -1, -1, -1, -1, -1, 30, 31, _
    17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17, _
    1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
    -1, -1, -1, 3, 3, 3, -1, -1, -1, -1, -1, 3, 3, 3, -1, 5, _
    1, 3, 21, 3, 21, 21, 5, 5, -1, -1, 3, 3, 5, 5, 5, 5, _
    1, 3, -1, -1, -1, -1, -1, 28, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, 27, 27, 27, -1, 27, -1, -1, 27, 27, 27, 27, 27, 27, 27, 27, _
    -1, 27, 27, 27, -1, 27, -1, -1, 27, 27, 27, 27, 27, 27, 27, 27, _
    -1, 27, 27, 27, -1, 27, -1, -1, 27, 27, 27, -1, 27, 27, 27, -1)

secondOperandTwoType = Array( _
    -1, -1, 6, 6, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    25, 26, 24, 24, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    32, 32, 32, 31, 31, 31, 31, 31, 32, 32, 32, 32, -1, -1, 30, 31, _
    31, 31, 31, 31, 31, 31, 31, -1, -1, -1, -1, -1, -1, -1, 29, 27, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, -1, -1, 5, 5, 5, -1, -1, -1, -1, -1, 5, 5, 5, -1, 3, _
    4, 5, 21, 5, -1, -1, 1, 6, -1, -1, 11, 5, 3, 3, 1, 6, _
    4, 5, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, _
    -1, 31, 31, 31, -1, 31, -1, -1, 31, 31, 31, 31, 31, 31, 31, 31, _
    -1, 31, 31, 31, -1, 31, -1, -1, 31, 31, 31, 31, 31, 31, 31, 31, _
    -1, 31, 31, 31, -1, 31, -1, -1, 31, 31, 31, -1, 31, 31, 31, -1)
    
strEscapeGroup0 = Array( _
        "FADD 32real ptr ", "FMUL 32real ptr ", "FCOM 32real ptr ", "FCOMP 32real ptr ", "FSUB 32real ptr ", "FSUBR 32real ptr ", "FDIVR 32real ptr ", "", "", "", "", "", "", "", "", "", _
        "FADD ST(0),*", "FADD ST(0),*", "FADD ST(0),*", "FADD ST(0),*", "FADD ST(0),*", "FADD ST(0),*", "FADD ST(0),*", "FADD ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", "FMUL ST(0),*", _
        "FCOM ST(0),*", "FCOM ST(0),*", "FCOM ST(0),*", "FCOM ST(0),*", "FCOM ST(0),*", "FCOM ST(0),*", "FCOM ST(0),*", "FCOM ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", "FCOMP ST(0),*", _
        "FSUB ST(0),*", "FSUB ST(0),*", "FSUB ST(0),*", "FSUB ST(0),*", "FSUB ST(0),*", "FSUB ST(0),*", "FSUB ST(0),*", "FSUB ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", "FSUBR ST(0),*", _
        "FDIV ST(0),*", "FDIV ST(0),*", "FDIV ST(0),*", "FDIV ST(0),*", "FDIV ST(0),*", "FDIV ST(0),*", "FDIV ST(0),*", "FDIV ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*", "FDIVR ST(0),*")

strEscapeGroup1 = Array( _
        "FLD 32real ptr ", "", "FST 32real ptr ", "FSTP 32real ptr ", "FLDENV 14/28bytes ptr ", "FLDCW 2bytes ptr ", "FSTENV 14/28bytes ptr ", "FSTCW 2bytes ptr ", "", "", "", "", "", "", "", "", _
        "FLD ST(0),*", "FLD ST(0),*", "FLD ST(0),*", "FLD ST(0),*", "FLD ST(0),*", "FLD ST(0),*", "FLD ST(0),*", "FLD ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", "FXCH ST(0),*", _
        "FNOP", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "FCHS", "FABS", "", "", "FTST", "FXAM", "", "", "FLD1", "FLDL2T", "FLDL2E", "FLDPI", "FLDLG2", "FLDLN2", "FLD2", "", _
        "F2XM1", "FYL2X", "FPTAN", "FPATAN", "FXTRACT", "FPREM1", "FDECSTP", "FINCSTP", "FPREM", "FYL2XP1", "FSQRT", "FSINCOS", "FRNDINT", "FSCALE", "FSIN", "FCOS")

strEscapeGroup2 = Array( _
        "FIADD ", "FIMUL ", "FICOM ", "FICOMP ", "FISUB ", "FISUBR ", "FIDIV ", "FIDIVR ", "", "", "", "", "", "", "", "", _
        "FCMOVB ST(0),*", "", "", "", "", "", "", "", "FCMOVE ST(0),*", "", "", "", "", "", "", "", _
        "FCMOVBE ST(0),*", "", "", "", "", "", "", "", "FCMOVU ST(0),*", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "FUCOMPP", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

strEscapeGroup3 = Array( _
        "FILD ", "", "FIST ", "FISTP ", "", "FLD 80real ptr ", " FSTP 80real ptr ", "", "", "", "", "", "", "", "", "", _
        "FCMOVNB ST(0),*", "", "", "", "", "", "", "", "FCMOVNE ST(0),*", "", "", "", "", "", "", "", _
        "FCMOVNBE ST(0),*", "", "", "", "", "", "", "", "FCMOVNU ST(0),*", "", "", "", "", "", "", "", _
        "", "", "FCLEX", "FINIT", "", "", "", "", "FUCOMI ST(0),*", "", "", "", "", "", "", "", _
        "FCOMI ST(0),*", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

strEscapeGroup4 = Array( _
        "FADD 64real ptr ", "FMUL 64real ptr ", "FCOM 64real ptr ", "FCOMP 64real ptr ", "FSUB 64real ptr ", "FSUBR 64real ptr ", "FDIVR 64real ptr ", "", "", "", "", "", "", "", "", "", _
        "FADD *,ST(0)", "FADD *,ST(0)", "FADD *,ST(0)", "FADD *,ST(0)", "FADD *,ST(0)", "FADD *,ST(0)", "FADD *,ST(0)", "FADD *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", "FMUL *,ST(0)", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUBR *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", "FSUB *,ST(0)", _
        "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIVR *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)", "FDIV *,ST(0)")

strEscapeGroup5 = Array( _
        "FLD 64real ptr ", "", "FST 64real ptr ", "FSTP 64real ptr ", "FRSTOR 98/108bytes ptr ", "", "FSAVE 98/108bytes ptr ", "FSTSW 2bytes ptr ", "", "", "", "", "", "", "", "", _
        "FFREE *", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "FST *", "", "", "", "", "", "", "", "FSTP *", "", "", "", "", "", "", "", _
        "FUCOM *,ST(0)", "", "", "", "", "", "", "", "FUCOMP *", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

strEscapeGroup6 = Array( _
        "FIADD ", "FIMUL ", "FICOM ", "FICOMP ", "FISUB ", "FISUBR ", "FIDIV ", "FIDIVR ", "", "", "", "", "", "", "", "", _
        "FADDP *,ST(0)", "", "", "", "", "", "", "", "FMULP *,ST(0)", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "FCOMPP", "", "", "", "", "", "", _
        "FSUBRP *,ST(0)", "", "", "", "", "", "", "", "FSUBP *,ST(0)", "", "", "", "", "", "", "", _
        "FDIVRP *,ST(0)", "", "", "", "", "", "", "", "FDIVP *,ST(0)", "", "", "", "", "", "", "")

strEscapeGroup7 = Array( _
        "FILD ", "", "FIST ", "FISTP ", "FBLD 80BCD ptr ", "FILD ", "FBSTP 80BCD ptr ", "FISTP ", "", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        "FSTSWAX", "", "", "", "", "", "", "", "FUCOMIP ST(0),*", "", "", "", "", "", "", "", _
        "FCOMIP ST(0),*", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

strEscape = Array(strEscapeGroup0, strEscapeGroup1, strEscapeGroup2, strEscapeGroup3, strEscapeGroup4, strEscapeGroup5, strEscapeGroup6, strEscapeGroup7)
End Function

'renvoie le spécifier de taille de donnée en fonction de cette taille et du segment de base
'sous la forme taille ptr segment: (segment est facultatif)
'==========================================================================================
'dwSize : taille de la donnée
'dwSegment : valeur de segmentOverride
Public Function getSpecifier(ByVal dwSize As Long, ByVal dwSegment As Long) As String
    'en fonction de la taille
    Select Case dwSize
        Case 8
            getSpecifier = "byte ptr " & getPrefixe(dwSegment)
        Case 16
            getSpecifier = "word ptr " & getPrefixe(dwSegment)
        Case 32
            getSpecifier = "dword ptr " & getPrefixe(dwSegment)
        Case 32
            getSpecifier = "qword ptr " & getPrefixe(dwSegment)
    End Select
End Function

'renvoie le préfixe textuel en fonction de ce prefixe
'====================================================
'bPref : prefixe d'une instruction
Private Function getPrefixe(ByVal bPref As Long) As String
    'en fonction du préfixe
    Select Case bPref
        Case &HF0
            getPrefixe = "LOCK "
        Case &HF2
            getPrefixe = "REPNE "
        Case &HF3
            getPrefixe = "REP "
        Case &H2E
            getPrefixe = "cs:"
        Case &H36
            getPrefixe = "ss:"
        Case &H3E
            getPrefixe = "ds:"
        Case &H26
            getPrefixe = "es:"
        Case &H64
            getPrefixe = "fs:"
        Case &H65
            getPrefixe = "gs:"
    End Select
End Function

'renvoie la base d'un SIB
'========================
'bSib : octet du SIB
Private Function getBase(bSib As Byte) As String
    'en fonction de du SIB
    Select Case rmTable(bSib)
        Case 0
            getBase = "EAX"
        Case 1
            getBase = "ECX"
        Case 2
            getBase = "EDX"
        Case 3
            getBase = "EBX"
        Case 4
            getBase = "ESP"
        Case 5
            getBase = "EBP"
        Case 6
            getBase = "ESI"
        Case 7
            getBase = "EDI"
    End Select
End Function

'renvoie le registre d'index en fonction du SIB
'==============================================
'bSib : SIB
Private Function getRegS(bSib As Byte) As String
    'en fonction du SIB
    Select Case regTable(bSib)
        Case 0
            getRegS = "EAX"
        Case 1
            getRegS = "ECX"
        Case 2
            getRegS = "EDX"
        Case 3
            getRegS = "EBX"
        Case 4
            getRegS = "ESP"
        Case 5
            getRegS = "EBP"
        Case 6
            getRegS = "ESI"
        Case 7
            getRegS = "EDI"
    End Select
End Function

'renvoie le Index*Scale en fonction d'un SIB
'===========================================
'bSib : SIB
Private Function getScaleIndex(bSib As Byte) As String
    Dim SS As Byte
    SS = modTable(bSib)
    SS = SS / 2
    If SS > 0 Then SS = SS - 1
    getScaleIndex = getRegS(bSib)
    Select Case SS
        Case 1
            getScaleIndex = getScaleIndex & "*2"
        Case 2
            getScaleIndex = getScaleIndex & "*4"
        Case 3
            getScaleIndex = getScaleIndex & "*8"
    End Select
End Function

'renvoie le texte de ModRM en fonction d'une instruction et d'une taille d'opérande
'==================================================================================
'Ins : instruction dont on veut le ModRM
'dwSize : taille en bits des opérandes
Private Function getModRM16(Ins As Instruction, ByVal dwSize As Long) As String
    'on fonction du type de ModRM
    Select Case mod16Table(Ins.bModRm)
        Case 1
            getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getRegister16(rmTable(Ins.bModRm)) & "]"
        Case 2
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                If CheckVA(Ins.m_dword) Then
                    getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & getAddrName(Ins.m_dword, dwSize, 5)
                Else
                    getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getNumber(Ins.m_dword, 6, False) & "]"
                End If
            Else
                If CheckVA(Ins.m_dword) Then
                    getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & getAddrName(Ins.m_dword, dwSize, 8)
                Else
                    getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getNumber(Ins.m_dword, 8, False) & "]"
                End If
            End If
        Case 3
            getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getRegister16(rmTable(Ins.bModRm))
            getModRM16 = getModRM16 & GetArgVar(Ins.regESP, regEBP, Ins.m_byte, rmTable(Ins.bModRm), 2, 2, dwSize)
            getModRM16 = getModRM16 & "]"
        Case 4
            If CheckVA(Ins.m_dword) Then
                If Ins.addressSizeOverride = bAddressSizeOverride Then
                    getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & getAddrName(Ins.m_dword, dwSize, 5) & "[" & getRegister16(rmTable(Ins.bModRm)) & "]"
                Else
                    getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & getAddrName(Ins.m_dword, dwSize, 5) & "[" & getRegister16(rmTable(Ins.bModRm)) & "]"
                End If
            Else
                getModRM16 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getRegister16(rmTable(Ins.bModRm))
                If Ins.addressSizeOverride = bAddressSizeOverride Then
                    getModRM16 = getModRM16 & GetArgVar(Ins.regESP, regEBP, Ins.m_dword, rmTable(Ins.bModRm), 5, 2, dwSize)
                Else
                    getModRM16 = getModRM16 & GetArgVar(Ins.regESP, regEBP, Ins.m_dword, rmTable(Ins.bModRm), 8, 2, dwSize)
                End If
                getModRM16 = getModRM16 & "]"
            End If
        Case 5
            getModRM16 = getRegister(rmTable(Ins.bModRm), dwSize)
    End Select
End Function

'renvoie le texte de ModRM en fonction d'une instruction et d'une taille d'opérande
'==================================================================================
'Ins : instruction dont on veut le ModRM
'dwSize : taille en bits des opérandes
Private Function getModRM32(Ins As Instruction, ByVal dwSize As Long) As String
    'en fonction du type de ModRM
    Select Case modTable(Ins.bModRm) 'TODO vérifier
        Case 1
            getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getRegister32(rmTable(Ins.bModRm)) & "]"
        Case 2
            If sibTable(Ins.bSib) = 1 Then
                If regTable(Ins.bSib) <> 4 Then
                    If CheckVA(Ins.m_dword) Then
                        getModRM32 = getSpecifier(dwSize, Ins.segmentOverride)
                        If Ins.addressSizeOverride = bAddressSizeOverride Then
                            getModRM32 = getModRM32 & getAddrName(Ins.m_dword, dwSize, 4)
                        Else
                            getModRM32 = getModRM32 & getAddrName(Ins.m_dword, dwSize, 8)
                        End If
                        getModRM32 = getModRM32 & "[" & getScaleIndex(Ins.bSib) & "]"
                    Else
                        getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "["
                        getModRM32 = getModRM32 & getScaleIndex(Ins.bSib)
                        If Ins.addressSizeOverride = bAddressSizeOverride Then
                            getModRM32 = getModRM32 & getNumber(Ins.m_dword, 4, True) & "]"
                        Else
                            getModRM32 = getModRM32 & getNumber(Ins.m_dword, 8, True) & "]"
                        End If
                    End If
                End If
            Else
                getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getBase(Ins.bSib) & "+" & getScaleIndex(Ins.bSib) & "]"
            End If
        Case 3
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                If CheckVA(Ins.m_dword) Then
                    getModRM32 = getAddrName(Ins.m_dword, dwSize, 4)
                Else
                    getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getNumber(Ins.m_dword, 4, False) & "]"
                End If
            Else
                If CheckVA(Ins.m_dword) Then
                    getModRM32 = getAddrName(Ins.m_dword, dwSize, 8)
                Else
                    getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getNumber(Ins.m_dword, 8, False) & "]"
                End If
            End If
        Case 4
            getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getRegister32(rmTable(Ins.bModRm))
            getModRM32 = getModRM32 & GetArgVar(Ins.regESP, regEBP, Ins.m_byte, rmTable(Ins.bModRm), 2, 4, dwSize)
            getModRM32 = getModRM32 & "]"
        Case 5
            getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getBase(Ins.bSib)
            If regTable(Ins.bSib) <> 4 Then
                getModRM32 = getModRM32 & "+" & getScaleIndex(Ins.bSib)
                getModRM32 = getModRM32 & getNumber(Ins.m_byte, 2, True)
            Else
                getModRM32 = getModRM32 & GetArgVar(Ins.regESP, regEBP, Ins.m_byte, rmTable(Ins.bModRm), 2, 4, dwSize)
            End If
            getModRM32 = getModRM32 & "]"
        Case 6
            If CheckVA(Ins.m_dword) Then
                If Ins.addressSizeOverride = bAddressSizeOverride Then
                    getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & getAddrName(Ins.m_dword, dwSize, 4) & "[" & getRegister32(rmTable(Ins.bModRm)) & "]"
                Else
                    getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & getAddrName(Ins.m_dword, dwSize, 8) & "[" & getRegister32(rmTable(Ins.bModRm)) & "]"
                End If
            Else
                getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getRegister32(rmTable(Ins.bModRm))
                If Ins.addressSizeOverride = bAddressSizeOverride Then
                    getModRM32 = getModRM32 & GetArgVar(Ins.regESP, regEBP, Ins.m_dword, rmTable(Ins.bModRm), 4, 4, dwSize)
                Else
                    getModRM32 = getModRM32 & GetArgVar(Ins.regESP, regEBP, Ins.m_dword, rmTable(Ins.bModRm), 8, 4, dwSize)
                End If
                getModRM32 = getModRM32 & "]"
            End If
        Case 7
            If CheckVA(Ins.m_dword) Then
                If regTable(Ins.bSib) <> 4 Then
                    getModRM32 = getSpecifier(dwSize, Ins.segmentOverride)
                    If Ins.addressSizeOverride = bAddressSizeOverride Then
                        getModRM32 = getModRM32 & getAddrName(Ins.m_dword, dwSize, 4)
                    Else
                        getModRM32 = getModRM32 & getAddrName(Ins.m_dword, dwSize, 8)
                    End If
                    getModRM32 = getModRM32 & "[" & getBase(Ins.bSib) & "+" & getScaleIndex(Ins.bSib)
                Else
                    getModRM32 = getSpecifier(dwSize, Ins.segmentOverride)
                    If Ins.addressSizeOverride = bAddressSizeOverride Then
                        getModRM32 = getModRM32 & getAddrName(Ins.m_dword, dwSize, 4)
                    Else
                        getModRM32 = getModRM32 & getAddrName(Ins.m_dword, dwSize, 8)
                    End If
                    getModRM32 = getModRM32 & "[" & getBase(Ins.bSib)
                End If
                getModRM32 = getModRM32 & "]"
            Else
                getModRM32 = getSpecifier(dwSize, Ins.segmentOverride) & "[" & getBase(Ins.bSib)
                If regTable(Ins.bSib) <> 4 Then
                    getModRM32 = getModRM32 & "+" & getScaleIndex(Ins.bSib)
                    If Ins.addressSizeOverride = bAddressSizeOverride Then
                        getModRM32 = getModRM32 & getNumber(Ins.m_dword, 4, True)
                    Else
                        getModRM32 = getModRM32 & getNumber(Ins.m_dword, 8, True)
                    End If
                Else
                    If Ins.addressSizeOverride = bAddressSizeOverride Then
                        getModRM32 = getModRM32 & GetArgVar(Ins.regESP, regEBP, Ins.m_dword, rmTable(Ins.bModRm), 4, 4, dwSize)
                    Else
                        getModRM32 = getModRM32 & GetArgVar(Ins.regESP, regEBP, Ins.m_dword, rmTable(Ins.bModRm), 8, 4, dwSize)
                    End If
                End If
                getModRM32 = getModRM32 & "]"
            End If
        Case 8
            getModRM32 = getRegister(rmTable(Ins.bModRm), dwSize)
    End Select
End Function

'renvoie le texte de ModRM en fonction d'une instruction et d'une taille d'opérande
'==================================================================================
'Ins : instruction dont on veut le ModRM
'dwSize : taille en bits des opérandes
Private Function getModRM(Ins As Instruction, ByVal dwSize As Long) As String
    If Ins.addressSizeOverride = bAddressSizeOverride Then
        getModRM = getModRM16(Ins, dwSize)
    Else
        getModRM = getModRM32(Ins, dwSize)
    End If
End Function

'renvoie la représentation textuelle d'un registre
'=================================================
'bReg : numéro de registre
'dwSize : taille en bits du registre
Private Function getRegister(ByVal bReg As Byte, ByVal dwSize As Long) As String
    'suivant la taille
    Select Case dwSize
        Case 8
            getRegister = getRegister8(bReg)
        Case 16
            getRegister = getRegister16(bReg)
        Case 32
            getRegister = getRegister32(bReg)
        Case 64
            getRegister = getRegister64(bReg)
    End Select
End Function

'renvoie la représentation textuelle d'un registre 8bits
'=======================================================
'bReg : numéro de registre
Private Function getRegister8(ByVal bReg As Byte) As String
    Select Case bReg
        Case 0
            getRegister8 = "AL"
        Case 1
            getRegister8 = "CL"
        Case 2
            getRegister8 = "DL"
        Case 3
            getRegister8 = "BL"
        Case 4
            getRegister8 = "AH"
        Case 5
            getRegister8 = "CH"
        Case 6
            getRegister8 = "DH"
        Case 7
            getRegister8 = "BH"
    End Select
End Function

'renvoie la représentation textuelle d'un registre 16bits
'========================================================
'bReg : numéro de registre
Private Function getRegister16(ByVal bReg As Byte) As String
    Select Case bReg
        Case 0
            getRegister16 = "AX"
        Case 1
            getRegister16 = "CX"
        Case 2
            getRegister16 = "DX"
        Case 3
            getRegister16 = "BX"
        Case 4
            getRegister16 = "SP"
        Case 5
            getRegister16 = "BP"
        Case 6
            getRegister16 = "SI"
        Case 7
            getRegister16 = "DI"
    End Select
End Function

'renvoie la représentation textuelle d'un registre 16bits
'========================================================
'bReg : numéro de registre
Private Function getSegmentRegister(ByVal bReg As Byte) As String
    Select Case bReg
        Case 0
            getSegmentRegister = "ES"
        Case 1
            getSegmentRegister = "CS"
        Case 2
            getSegmentRegister = "SS"
        Case 3
            getSegmentRegister = "DS"
        Case 4
            getSegmentRegister = "FS"
        Case 5
            getSegmentRegister = "GS"
    End Select
End Function

'renvoie la représentation textuelle d'un registre 32bits
'========================================================
'bReg : numéro de registre
Private Function getRegister32(ByVal bReg As Byte) As String
    Select Case bReg
        Case 0
            getRegister32 = "EAX"
        Case 1
            getRegister32 = "ECX"
        Case 2
            getRegister32 = "EDX"
        Case 3
            getRegister32 = "EBX"
        Case 4
            getRegister32 = "ESP"
        Case 5
            getRegister32 = "EBP"
        Case 6
            getRegister32 = "ESI"
        Case 7
            getRegister32 = "EDI"
    End Select
End Function

'renvoie la représentation textuelle d'un registre 64bits
'========================================================
'bReg : numéro de registre
Private Function getRegister64(ByVal bReg As Byte) As String
    getRegister64 = "mm" & Str(bReg)
End Function

'met en forme un nombre
'========================================================
'dwNumber : nombre à formater
'dwNumberOfZero : nombre de zéro avant la virgule
'bSigne : indique s'il faut afficher le signe du nombre ou pas
Public Function getNumber(ByVal dwNumber As Long, ByVal dwNumberOfZero As Long, Optional bSigne As Boolean = False)
On Error Resume Next
    If bSigne Then
        If dwNumber >= 0 Then
            getNumber = "+"
        Else
            getNumber = "-"
            dwNumber = -dwNumber
        End If
    'Else
        'If dwNumber < 0 Then
        '    getNumber = "-"
        '    dwNumber = -dwNumber
        'End If
    End If
    getNumber = getNumber & Right$(String$(dwNumberOfZero, "0") & Hex$(dwNumber), dwNumberOfZero) & "H"
End Function

'renvoie la représentation textuelle des offsets relatifs au registre de base de la pile pour la procédure (ESP ou EBP)
'permet de trouver le nombre et la taille des arguments d'une procédure
'============================================================================
'dwValue : valeur d'offset relatif au registre qui sert de base à la pile pour la procédure (ESP ou EBP)
'dwSize : taille en bits de l'argument (des opérandes)
'regBase : registre de base (utilisé) de la pile : ESP ou EBP
Private Function GetArg(ByVal dwValue As Long, ByVal dwSize As Long, ByVal regBase As Long) As String
Dim spec() As String, X As Long

dwSize = dwSize / 8
On Error GoTo Pas
    'arg_:ptr type
    spec = Split(argCol(CStr("arg_" & Hex$(dwValue))), ":")
    GetArg = Hex$(dwValue)
Exit Function
Pas:
    For X = 1 To argCol.Count
        spec = Split(argCol(X), ":")
        If (dwValue > CLng(spec(0))) And (dwValue < (CLng(spec(0)) + CLng(spec(1)))) Then
            GetArg = Hex$(CLng(spec(0))) & "+" & Hex$(dwValue - CLng(spec(0)))
            Exit Function
        End If
    Next
    GetArg = Hex$(dwValue)
    argCol.Add "&H" & Hex$(dwValue) & ":" & CStr(dwSize) & ":" & Hex$(regBase), "arg_" & Hex$(dwValue)
End Function

'renvoie la représentation textuelle des offsets relatifs au registre de base de la pile pour la procédure (ESP ou EBP)
'permet de trouver le nombre et la taille des arguments d'une procédure
'============================================================================
'dwValue : valeur d'offset relatif au registre qui sert de base à la pile pour la procédure (ESP ou EBP)
'dwSize : taille en bits de l'argument (des opérandes)
'regBase : registre de base (utilisé) de la pile : ESP ou EBP
Private Function GetVar(ByVal dwValue As Long, ByVal dwSize As Long, ByVal regBase As Long) As String
Dim spec() As String, X As Long

dwSize = dwSize / 8
On Error GoTo Pas
    'var_:ptr type
    spec = Split(varCol(CStr("var_" & Hex$(dwValue))), ":")
    GetVar = Hex$(dwValue)
Exit Function
Pas:
    For X = 1 To varCol.Count
        spec = Split(varCol(X), ":")
        If (dwValue > CLng(spec(0))) And (dwValue < (CLng(spec(0)) + CLng(spec(1)))) Then
            GetVar = Hex$(CLng(spec(0))) & "-" & Hex$(dwValue - CLng(spec(0)))
            Exit Function
        End If
    Next
    GetVar = Hex$(dwValue)
    varCol.Add "&H" & Hex$(dwValue) & ":" & CStr(dwSize) & ":" & Hex$(regBase), "var_" & Hex$(dwValue)
End Function

'renvoie la représentation textuelle d'un argument ou d'une variable locale (suivant le signe de l'offset)
'=========================================================================================================
'regESP : registre ESP pour l'instruction
'regEBP : registre EBP pour l'instruction
'dwOffset : décalage relatif au registre de base de la pile pour la procédure
'rm : rm de la table en fonction du ModRM de l'instruction
'cDigit : nombre de chiffre avant la virgule (taille des opérandes)
'regSize : taille des registres en octet
'dwOperandSize : taille des opérandes en bits
Private Function GetArgVar(ByVal regESP As Long, ByVal regEBP As Long, ByVal dwOffset As Long, ByVal rm As Long, ByVal cDigit As Long, ByVal regSize As Long, ByVal dwOperandSize) As String
If rm = 4 Then 'ESP
    If dwOffset < regESP Then
        If dwOffset > 0 Then '[ESP+n+var_y]
            GetArgVar = "+0" & Hex$(regESP - dwInitESP) & "H+var_" & GetVar(regESP - dwOffset - dwInitESP, dwOperandSize, regESP - dwOffset - dwInitESP)
        Else
            If CheckVA(dwOffset) Then
                GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
            Else
                GetArgVar = getNumber(dwOffset, cDigit, True)
            End If
        End If
    ElseIf dwOffset >= regESP Then
        If regEBP > 0 Then
            If regESP - regEBP = 0 Then '[ESP+arg_y]
                GetArgVar = "+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset)
            ElseIf regESP - regEBP > 0 Then '[ESP+x+arg_y]
                GetArgVar = "+0" & Hex$(regESP - regEBP) & "H+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset - regESP + regEBP)
            Else '[ESP+n]
                If CheckVA(dwOffset) Then
                    GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
                Else
                    GetArgVar = getNumber(dwOffset, cDigit, True)
                End If
            End If
        ElseIf regEBP = 0 Then
            If regESP > dwInitESP Then '[ESP+x+arg_y]
                GetArgVar = "+0" & Hex$(regESP - dwInitESP) & "H+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset - regESP + dwInitESP)
            ElseIf regESP = dwInitESP Then  '[ESP+arg_y]
                GetArgVar = "+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset)
            End If
        End If
'    Else
'        If CheckVA(dwOffset) Then
'            GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'        Else
'            GetArgVar = getNumber(dwOffset, cDigit, True)
'        End If
    End If
ElseIf rm = 5 Then 'EBP
    If regEBP = 2 * dwInitESP Then
        If dwOffset > 0 Then '[EBP+arg_x]
            GetArgVar = "+arg_" & GetArg(dwOffset - regEBP, dwOperandSize, dwOffset)
        ElseIf dwOffset < 0 Then '[EBP+var_z]
            GetArgVar = "+var_" & GetVar(-dwOffset, dwOperandSize, -dwOffset)
        End If
    Else
        If CheckVA(dwOffset) Then
            GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
        Else
            GetArgVar = getNumber(dwOffset, cDigit, True)
        End If
    End If
Else 'other
    If CheckVA(dwOffset) Then
        GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
    Else
        GetArgVar = getNumber(dwOffset, cDigit, True)
    End If
End If
'    If dwOffset < 0 Then 'var_
'        If rm = 4 Then 'ESP
'            GetArgVar = "+" & Hex$(regESP - regSize) & "H+var_" & GetVar(regESP - regSize - dwOffset, dwOperandSize, -regESP + regSize + dwOffset)
'        ElseIf rm = 5 Then 'EBP
'            GetArgVar = "+var_" & GetVar(-dwOffset, dwOperandSize, -dwOffset)
'        Else
'            If CheckVA(dwOffset) Then
'                GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'            Else
'                GetArgVar = getNumber(dwOffset, cDigit, True)
'            End If
'        End If
'    Else 'arg_
'        If rm = 4 Then 'ESP
'            If regESP >= dwOffset Then
'                If CheckVA(dwOffset) Then
'                    GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'                Else
'                    GetArgVar = getNumber(dwOffset, cDigit, True)
'                End If
'            Else
'                If regEBP > 0 Then
'                    If regESP - regEBP > 0 Then
'                        GetArgVar = "+" & Hex$(regESP - regEBP) & "H+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset - regESP + regEBP)
'                    ElseIf regESP - regEBP = 0 Then
'                        GetArgVar = "+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset)
'                    Else
'                        If CheckVA(dwOffset) Then
'                            GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'                        Else
'                            GetArgVar = getNumber(dwOffset, cDigit, True)
'                        End If
'                    End If
'                ElseIf regEBP = 0 Then
'                    If regESP > regSize Then
'                        GetArgVar = "+" & Hex$(regESP - regSize) & "H+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset - regESP + regSize)
'                    ElseIf regESP = regSize Then
'                        GetArgVar = "+arg_" & GetArg(dwOffset - regESP, dwOperandSize, dwOffset)
'                    Else
'                        If CheckVA(dwOffset) Then
'                            GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'                        Else
'                            GetArgVar = getNumber(dwOffset, cDigit, True)
'                        End If
'                    End If
'                Else
'                    If CheckVA(dwOffset) Then
'                        GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'                    Else
'                        GetArgVar = getNumber(dwOffset, cDigit, True)
'                    End If
'                End If
'            End If
'        ElseIf rm = 5 Then 'EBP
'            If regEBP = 2 * regSize Then
'                GetArgVar = "+arg_" & GetArg(dwOffset - regEBP, dwOperandSize, dwOffset)
'            Else
'                If CheckVA(dwOffset) Then
'                    GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'                Else
'                    GetArgVar = getNumber(dwOffset, cDigit, True)
'                End If
'            End If
'        Else
'            If CheckVA(dwOffset) Then
'                GetArgVar = "+" & getAddrName(dwOffset, dwOperandSize, cDigit)
'            Else
'                GetArgVar = getNumber(dwOffset, cDigit, True)
'            End If
'        End If
'    End If
End Function

'renvoie la représentation d'un opérande en fonction de son type et de l'instruction
'===================================================================================
'Ins : instruction de l'opérande
'OperandType : type de l'opérande
Private Function GetOperand(Ins As Instruction, ByVal OperandType As Long)
    Dim s As Long, st As Byte, nnn As Byte, opext As Long, dt As Long
    Dim addr As Long, subname As String
    Select Case OperandType
        Case 1
            GetOperand = getModRM(Ins, 8)
        Case 2
            addr = Ins.i_byte
            subname = GetSubName(addr)
            If Len(subname) Then
                GetOperand = "short " & subname
            Else
                GetOperand = "short loc_" & getNumber(Ins.i_byte, 8)
            End If
        Case 3
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                s = 16
            Else
                s = 32
            End If
            GetOperand = getModRM(Ins, s)
        Case 4
            GetOperand = getRegister8(rmTable(Ins.bModRm))
        Case 5
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                GetOperand = getRegister16(regTable(Ins.bModRm))
            Else
                GetOperand = getRegister32(regTable(Ins.bModRm))
            End If
        Case 14
            If (Ins.addressSizeOverride = bAddressSizeOverride) Or (Ins.operandSizeOverride = bOperandSizeOverride) Then
                GetOperand = getSegmentRegister(regTable(Ins.bModRm))
            Else
                GetOperand = getSegmentRegister(regTable(Ins.bModRm))
            End If
        Case 6
            GetOperand = getModRM(Ins, 16)
        Case 7
            If Ins.segmentOverride Then
                GetOperand = getSpecifier(8, Ins.segmentOverride) & "[" & getNumber(Ins.i_dword, 8) & "]"
            Else
                GetOperand = getAddrName(Ins.i_dword, 8)
            End If
        Case 8
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                If Ins.segmentOverride Then
                    GetOperand = getSpecifier(16, Ins.segmentOverride) & "[" & getNumber(Ins.i_dword, 8) & "]"
                Else
                    GetOperand = getAddrName(Ins.i_dword, 16)
                End If
            Else
                If Ins.segmentOverride Then
                    GetOperand = getSpecifier(32, Ins.segmentOverride) & "[" & getNumber(Ins.i_dword, 8) & "]"
                Else
                    GetOperand = getAddrName(Ins.i_dword, 32)
                End If
            End If
        Case 9
            GetOperand = "byte ptr DS:[SI]"
        Case 90
            GetOperand = "byte ptr ES:[DI]"
        Case 10
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                GetOperand = "word ptr DS:[SI]"
            Else
                GetOperand = "dword ptr DS:[SI]"
            End If
        Case 100
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                GetOperand = "word ptr ES:[DI]"
            Else
                GetOperand = "dword ptr ES:[DI]"
            End If
        Case 11
            GetOperand = getNumber(Ins.i_byte, 2)
        Case 12
            GetOperand = getNumber(Ins.i_dword, 4)
        Case 13
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                s = 4
            Else
                s = 8
            End If
            If CheckVA(Ins.i_dword) Then
'                dt = GetDataType(Ins.i_dword, 0)
'                Select Case dt
'                    Case 0
'                        GetOperand = "offset sub_" & getNumber(Ins.i_dword, s)
'                    Case 3
'                        GetOperand = "offset unk_" & getNumber(Ins.i_dword, s)
'                    Case 4
'                        GetOperand = "offset ptr_" & getNumber(Ins.i_dword, s)
'                    Case 5
'                        GetOperand = "offset sz_" & getNumber(Ins.i_dword, s)
'                    Case 7
'                        GetOperand = "offset pascal_" & getNumber(Ins.i_dword, s)
'                    Case 10
'                        GetOperand = "offset uni_" & getNumber(Ins.i_dword, s)
'                    Case Else
'                        GetOperand = getNumber(Ins.i_dword, s)
'                End Select
                GetOperand = getAddrName(Ins.i_dword, 0)
            Else
                GetOperand = getNumber(Ins.i_dword, s)
            End If
        Case 15
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                s = 4
            Else
                s = 8
            End If
            GetOperand = getNumber(Ins.i_dword, 8) & ":" & getNumber(Ins.m_dword, s)
        Case 16
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                GetOperand = "FLAGS"
            Else
                GetOperand = "EFLAGS"
            End If
        Case 17
            addr = Ins.i_dword
            subname = GetSubName(addr)
            If Len(subname) Then
                GetOperand = subname
            Else
                If Ins.addressSizeOverride = bAddressSizeOverride Then
                    GetOperand = "loc_" & getNumber(addr, 4)
                Else
                    GetOperand = "loc_" & getNumber(addr, 8)
                End If
            End If
        Case 19
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                s = 16
            Else
                s = 32
            End If
            GetOperand = getModRM(Ins, s)
        Case 20
            GetOperand = getRegister16(rmTable(Ins.bModRm))
        Case 21, 23
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                s = 16
            Else
                s = 32
            End If
            GetOperand = getModRM(Ins, s)
        Case 22
            GetOperand = "1"
        Case 99 'Escape
            If (Ins.bModRm >= 0) And (Ins.bModRm <= &HBF) Then
                nnn = (Ins.bModRm And &H38) \ 8
                GetOperand = strEscape(Ins.iOpcode - &HD8)(nnn)
                Select Case Ins.iOpcode
                    Case &HDA
                        s = 32
                    Case &HDB
                        If (nnn < 4) Then
                            s = 32
                        Else
                            s = 0
                        End If
                    Case &HDD, &HDC, &HD9, &HD8
                        s = 0
                    Case &HDE
                        s = 16
                    Case &HDF
                        If (nnn < 4) Then
                            s = 16
                        ElseIf (nnn = 5) Or (nnn = 7) Then
                            s = 64
                        Else
                            s = 0
                        End If
                End Select
                GetOperand = GetOperand & getModRM(Ins, s)
            Else
                st = (Ins.bModRm And &HF)
                If st > 7 Then st = st - 8
                GetOperand = strEscape(Ins.iOpcode - &HD8)(Ins.bModRm - &HB0)
                GetOperand = Replace(GetOperand, "*", "ST(" & Str(st) & ")")
            End If
        'two bytes opcode
        Case 18 'Grp 7
            opext = (Ins.bModRm And 56) / 8
            Select Case opext
                Case 0 To 3
                    If Ins.addressSizeOverride = bAddressSizeOverride Then
                        GetOperand = "[" & getNumber(Ins.m_dword, 4) & "]"
                    Else
                        GetOperand = "[" & getNumber(Ins.m_dword, 8) & "]"
                    End If
                Case 4, 6
                    GetOperand = getModRM(Ins, 16)
                Case 7
                    GetOperand = getModRM(Ins, 8)
            End Select
        Case 24
            GetOperand = getRegister32(regTable(Ins.bModRm))
        Case 25
            GetOperand = "cr" & CStr(rmTable(Ins.bModRm))
        Case 26
            GetOperand = "dr" & CStr(rmTable(Ins.bModRm))
        Case 27, 29
            GetOperand = getRegister64(regTable(Ins.bModRm))
        Case 28, 31, 32
            GetOperand = getModRM(Ins, 64)
        Case 30
            GetOperand = getModRM(Ins, 32)
    End Select
End Function

'renvoie les opérandes d'une instructions
'========================================
'Ins : instruction
Private Function GetOperands(Ins As Instruction) As String
    Dim f, s As Long
    
    If Ins.opclass = &HF Then
        f = firstOperandTwoType(Ins.iOpcode)
    Else
        f = firstOperandType(Ins.iOpcode)
    End If
    
    If f > 0 Then GetOperands = GetOperand(Ins, f)

    If Ins.opclass = &HF Then
        s = secondOperandTwoType(Ins.iOpcode)
    Else
        s = secondOperandType(Ins.iOpcode)
    End If
    If (f > 0) And (s > 0) Then GetOperands = GetOperands & ","
    If s > 0 Then GetOperands = GetOperands & GetOperand(Ins, s)
End Function

'renvoie la représentation de l'instruction
'==========================================
'Ins : instruction
Public Function GetTextInstruction(Ins As Instruction) As String
    If Ins.opclass = &HF Then
        GetTextInstruction = getPrefixe(Ins.LockAndRepeat) & GetTwoByteInstruction(Ins)
    Else
        GetTextInstruction = getPrefixe(Ins.LockAndRepeat) & GetOneByteInstruction(Ins)
    End If
End Function

'renvoie la représentation textuelle d'une instruction sur un octet
'==================================================================
'Ins : instruction
Private Function GetOneByteInstruction(Ins As Instruction) As String
Dim opext As Byte
    Select Case opcodeTable(Ins.iOpcode)
        Case 0
            GetOneByteInstruction = strOneByteInstruction(Ins.iOpcode)
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                GetOneByteInstruction = Replace(GetOneByteInstruction, "e", vbNullString)
            Else
                GetOneByteInstruction = UCase$(GetOneByteInstruction)
            End If
            If (firstOperandType(Ins.iOpcode) = 10) Or (firstOperandType(Ins.iOpcode) = 100) Or (firstOperandType(Ins.iOpcode) = 9) Or (firstOperandType(Ins.iOpcode) = 90) Then
                GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
            End If
        Case 1, 2, 3, 5, 6, 7, 8
            If (Ins.iOpcode = &HCD) And (Ins.i_byte = &H20) Then
                If (Ins.i_dword And &H10000) = &H10000 Then
                    GetOneByteInstruction = "VMMCall " & GetVxDCalls(Ins.i_dword)
                Else
                    GetOneByteInstruction = "VxDCall " & GetVxDCalls(Ins.i_dword)
                End If
            Else
                GetOneByteInstruction = strOneByteInstruction(Ins.iOpcode)
                GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
            End If
        Case 4
            GetOneByteInstruction = strOneByteInstruction(Ins.iOpcode)
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                GetOneByteInstruction = Replace(GetOneByteInstruction, "e", vbNullString)
            Else
                GetOneByteInstruction = UCase$(GetOneByteInstruction)
            End If
            GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
        Case 44
            GetOneByteInstruction = strOneByteInstruction(Ins.iOpcode)
            If Ins.addressSizeOverride = bAddressSizeOverride Then
                GetOneByteInstruction = Replace(GetOneByteInstruction, "e", vbNullString)
            Else
                GetOneByteInstruction = UCase$(GetOneByteInstruction)
            End If
            GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
        Case 9, 10, 11 'opext
            opext = (Ins.bModRm And 56) / 8
            Select Case Ins.iOpcode
                Case &HFE 'grp 4
                    GetOneByteInstruction = strGroupExtensions(3)(opext)
                Case &HD0, &HD1, &HC0, &HC1 'grp 2
                    GetOneByteInstruction = strGroupExtensions(1)(opext)
                Case &HD2, &HD3 'grp 2
                    GetOneByteInstruction = strGroupExtensions(1)(opext) & ",CL"
                Case &H80, &H82, &H83, &H81 'grp 1
                    GetOneByteInstruction = strGroupExtensions(0)(opext)
                Case &HC6, &HC7 'grp 11
                    GetOneByteInstruction = strGroupExtensions(10)(opext)
            End Select
            GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
        Case 12 'Escape
            GetOneByteInstruction = GetOperands(Ins)
        Case 13 'JUMP
            opext = (Ins.bModRm And 56) / 8
            GetOneByteInstruction = strGroupExtensions(4)(opext)
            GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
        Case 14 'TEST
            opext = (Ins.bModRm And 56) / 8
            GetOneByteInstruction = strGroupExtensions(2)(opext)
            If Ins.iOpcode = &HF6 Then
                If opext = 0 Then
                    secondOperandType(&HF6) = 11
                ElseIf opext > 3 Then
                    secondOperandType(&HF6) = 0
                    GetOneByteInstruction = GetOneByteInstruction & "AL"
                Else
                    secondOperandType(&HF6) = -1
                End If
            Else
                If opext = 0 Then
                    secondOperandType(Ins.iOpcode) = 13
                ElseIf opext > 3 Then
                    secondOperandType(Ins.iOpcode) = 0
                    GetOneByteInstruction = GetOneByteInstruction & "eAX"
                Else
                    secondOperandType(Ins.iOpcode) = -1
                End If
            End If
            If Ins.operandSizeOverride = bOperandSizeOverride Then
                GetOneByteInstruction = Replace(GetOneByteInstruction, "e", vbNullString)
            Else
                GetOneByteInstruction = UCase$(GetOneByteInstruction)
            End If
            GetOneByteInstruction = Replace(GetOneByteInstruction, "*", GetOperands(Ins))
        Case 15 'WAIT
            GetOneByteInstruction = strOneByteInstruction(Ins.iOpcode)
        Case 16 'REPEAT
            GetOneByteInstruction = strOneByteInstruction(Ins.iOpcode)
    End Select
End Function

'renvoie la représentation textuelle d'une instruction sur deux octet
'====================================================================
'Ins : instruction
Private Function GetTwoByteInstruction(Ins As Instruction) As String
Dim opext As Byte
    Select Case opcode2Table(Ins.iOpcode)
        Case 0
            GetTwoByteInstruction = strTwoByteInstruction(Ins.iOpcode)
        Case 1, 2
            GetTwoByteInstruction = strTwoByteInstruction(Ins.iOpcode)
            GetTwoByteInstruction = Replace(GetTwoByteInstruction, "*", GetOperands(Ins))
            If (Ins.iOpcode = &HA5) Or (Ins.iOpcode = &HAD) Then 'shld
                GetTwoByteInstruction = GetTwoByteInstruction & ",CL"
            End If
        Case 3 'shld / shrd Ib
            GetTwoByteInstruction = strTwoByteInstruction(Ins.iOpcode)
            GetTwoByteInstruction = Replace(GetTwoByteInstruction, "*", GetOperands(Ins))
            If (Ins.iOpcode = &HA4) Or (Ins.iOpcode = &HAC) Then 'shld
                GetTwoByteInstruction = GetTwoByteInstruction & "," & getNumber(Ins.i_byte, 2)
            End If
        Case 4 'grp 6,7,9,15
            opext = (Ins.bModRm And 56) / 8
            Select Case Ins.iOpcode
                Case &H0  'grp 6
                    GetTwoByteInstruction = strGroupExtensions(5)(opext)
                Case &H1  'grp 7
                    GetTwoByteInstruction = strGroupExtensions(6)(opext)
                Case &HC7 'grp 9
                    GetTwoByteInstruction = strGroupExtensions(8)(opext)
                Case &HAE 'grp 15
                    GetTwoByteInstruction = strGroupExtensions(14)(opext)
            End Select
            GetTwoByteInstruction = Replace(GetTwoByteInstruction, "*", GetOperands(Ins))
        Case 5 'grp 12,13,14,8
            opext = (Ins.bModRm And 56) / 8
            Select Case Ins.iOpcode
                Case &HBA  'grp 8
                    GetTwoByteInstruction = strGroupExtensions(7)(opext)
                Case &H71  'grp 12
                    GetTwoByteInstruction = strGroupExtensions(11)(opext)
                Case &H72 'grp 13
                    GetTwoByteInstruction = strGroupExtensions(12)(opext)
                Case &H73 'grp 14
                    GetTwoByteInstruction = strGroupExtensions(13)(opext)
            End Select
            GetTwoByteInstruction = Replace(GetTwoByteInstruction, "*", GetOperands(Ins))
    End Select
End Function
