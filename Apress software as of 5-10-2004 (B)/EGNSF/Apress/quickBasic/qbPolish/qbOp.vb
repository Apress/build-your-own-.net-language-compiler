Option Strict

' *******************************************************************************************
' *                                                                                         *
' * qbOp     QuickBasicEngine operators                                                     *
' *                                                                                         *
' * This stateless class identifies the operators supported by the non-CLR Nutty            *
' * Professor machine as a large enumerator, and it provides Shared conversion tools        *
' * for enumerator values (opCodeToString, opCodeFromString, opCodeToDescription).          *
' *                                                                                         *
' * The ops2Doc application can be used to convert this class to and from a                 *
' * Word-ready table.                                                                       *
' *                                                                                         *
' * C H A N G E   R E C O R D ------------------------------------------------------------- *
' *         DATE                  PROGRAMMER/GENERATOR               D E S C R I P T I O N  *
' * -------------------- --------------------------------------     ----------------------- *
' * 5/17/2003            Nilges                                     Version1                *
' * 6/20/2003 9:33:47 PM C:\egnsf\APress\quickBasic\ops2Doc\bin     Generated from op table *
' * 6/20/2003            Nilges                                     Numerous manual changes *
' *                                                                                         *
' *******************************************************************************************

Public Class qbOp

    Private Shared _OBJutilities As utilities.utilities
    
    Public Enum ENUop
        opAdd          ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)+stack(top)
        opAnd          ' Replaces stack(top) and stack(top-1) with stack(top-1) And
                       ' stack(top) (And is not short circuited)
        opAsc          ' Replaces stack(top) by its ASCII value
        opCeil         ' Replaces stack(top) with first integer n > stack(top)
        #If QUICKBASICENGINE_FMSEXTENSION Then
        opChoose       ' Returns the nth value of the Choose stack frame: expects
                       ' n, list and count
        #End If
        opChr          ' Replaces stack(top) with its ASCII character value
        opCircle       ' Draws a circle: stack(top - 2) is x coordinate, stack(top
                       ' - 1) is y coordinate, and stack(top) is radius
        opCls          ' Clears the simulated QuickBasic screen
        opCoGo         ' Computed GoTo/GoSub: translate string/number at stack(top)
                       ' to address
        opConcat       ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)&stack(top)
        opCos          ' Replaces stack(top) with its cosine
        opDivide       ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)/stack(top)
        opDuplicate    ' Copies stack(top-n) to stack top
        opEnd          ' Stops processing immediately
        opEval         ' Evaluates stack(top) as a Quickbasic expression using
                       ' lightweight evaluation: a new quickBasicEngine with default
                       ' options is used to evaluate stack(top)
        opEvaluate     ' Evaluates stack(top) as a Quickbasic expression using
                       ' heavyweight evaluation: a new quickBasicEngine with the same
                       ' options as the current engine is used to evaluate stack(top)
        opFloor        ' Replaces stack(top) with first integer n < stack(top)
        opForTest      ' Stacks stack(top-1)>=0 And c(stack(top))<stack(top-2) Or
                       ' stack(top-1)<0 And c(stack(top))>stack(top-2)
        opIif          ' Replaces stack(top-2)..stack(top) with stack(top-1) when
                       ' stack(top-2) is True, with stack(top) otherwise
        opIn           ' Searches the IN stack frame for a value. Expects the
                       ' search value, the frame size, and n values for the search at
                       ' the top of the stack QUICKBASICENGINE_FMSEXTENSION
        opInput        ' Reads a number or a string to stack(top)
        opInt          ' Replaces stack(top) with integer part
        opInvalid      ' Invalid marker op
        opIsNumeric    ' Replaces stack(top) with True when stack(top) is a number,
                       ' False otherwise
        opJump         ' Jumps to location
        opJumpIndirect ' Jumps to location identified at the top of the stack
        opJumpNZ       ' Jumps to location when stack(top) <> 0 (pop the stack top)
        opJumpZ        ' Jumps to location when stack(top) = 0 (pop the stack top)
        opLabel        ' Identifies position of a code label or statement number
        opLCase        ' Replaces the string at stack(top) with its lower case
                       ' translation
        opLen          ' Replace stack(top) by its length as a string
        opLike         ' Compare two strings at the stack top for a pattern match,
                       ' replacing them by True or False
        opLog          ' Replaces stack(top) by its logarithm
        opMax          ' Replaces stack(top) and stack(top-1) with the max
        opMid          ' Replaces stack(top-2)..stack(top) with the substring of
                       ' stack(top-2) starting at stack(top-1) for length at
                       ' stack(top)
        opMin          ' Replaces stack(top) and stack(top-1) with the min
        opMod          ' Replaces stack(top) and stack(top-1) with the integer
                       ' division remainder from stack(top-1) \ stack(top)
        opMultiply     ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)*stack(top)
        opNegate       ' Reverse sign at top of stack
        opNop          ' Does nothing (infinitely useful: cf. Tao t'eh Ch'ing)
        opNot          ' Replaces stack(top) with Not stack(top)
        opOr           ' Replaces stack(top) and stack(top-1) with stack(top-1) Or
                       ' stack(top) (Or is not short circuited)
        opPop          ' Sends stack(top) to a memory location
        opPopIndirect  ' Sends stack(top) to a memory location at stack(top-1):
                       ' removes stack(top), leaves stack(top-1) alone
        opPopOff       ' Removes stack(top) without sending it to a memory location
        opPower        ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)^stack(top)
        opPrint        ' Prints (and removes) value at top of the stack
        opPush         ' Pushes the contents of a memory location in the Case op
        #If QUICKBASICENGINE_FMSEXTENSION Then
        opPushDateDIFF ' Replaces stack(top-2), stack(top-1) and stack(top) by the
                       ' elapsed time from stack(top-2) to stack(top-1). stack(top)
                       ' should be the interval as seconds, minutes, hours, days,
                       ' months or years. If stack(top) is a null string then the
                       ' interval is days when stack(top-2) and stack(top-1) are both
                       ' single words, seconds otherwise.
        #End If
        #If QUICKBASICENGINE_FMSEXTENSION Then
        opPushDateEQ   ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)=stack(top) as a date, 0 otherwise
        #End If
        #If QUICKBASICENGINE_FMSEXTENSION Then
        opPushDateGT   ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)>stack(top) as a date, 0 otherwise
        #End If
        #If QUICKBASICENGINE_FMSEXTENSION Then
        opPushDateLT   ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)<stack(top) as a date, 0 otherwise
        #End If
        opPushEQ       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)=stack(top), 0 otherwise
        opPushGE       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)>=stack(top), 0 otherwise
        opPushGT       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)>stack(top), 0 otherwise
        opPushIndirect ' Pushes the contents of a memory location indexed at
                       ' stack(top)...replacing the index
        opPushLE       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)<=stack(top), 0 otherwise
        opPushLiteral  ' Pushes a literal string or number
        opPushLT       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)<stack(top), 0 otherwise
        opPushNE       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)<>stack(top), 0 otherwise
        opPushReturn   ' Pushes the subroutine's return address
        opRand         ' Seeds the random number generator to unpredictable values
        opRead         ' Reads from the data statements to stack(top)
        opRem          ' Equivalent to a Nop
        opReplace      ' Replaces all occurences of the string at stack(top-1) by
                       ' the string at stack(top) in the string at stack(top-2).
                       ' Replaces all entries by the translated string
        opRnd          ' Pushes an unseeded random number on the stack
        opRndSeed      ' Pushes a seeded random number on the stack (seed is
                       ' stack(top), and is replaced)
        opRotate       ' Exchanges stack(top) with stack(top-n)
        opRound        ' Rounds stack(top-1) to stack(top) digits
        opSgn          ' Replaces stack(top) with its signum (0 for 0: 1 for
                       ' positive: -1 for negative)
        opSin          ' Replaces stack(top) with its sine
        opSqr          ' Replaces stack(top) with its square root
        opString       ' Replaces stack(top) and stack(top-1) with n copies of the
                       ' character at stack top, where n is at stack(top-1)
        opSubtract     ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)-stack(top)
        opTrace        ' Changes trace settings
        opTracePop     ' Restores trace settings from a LIFO stack
        opTracePush    ' Saves trace settings in a LIFO stack
        opTrim         ' Replaces stack(top) with trimmed string (leading and
                       ' trailing blanks removed)
        opUCase        ' Replaces the string at stack(top) with its upper case
                       ' translation
        #If QUICKBASICENGINE_EXTENSION Then
        opUtility      ' Replaces stack(top-n)..stack(top) (where n=stack(top-1))
                       ' with the result of calling the utility named in stack(top)
        #End If
    End Enum
    
    Private Const ABOUTINFO As String = _
    "This stateless class returns the ENUop enumerator which " & _
    "identifies the operators of our non-CLR Nutty Professor " & _
    "interpreter for Quick Basic, and provides tools for " & _
    "converting operators to and from string names, and to their " & _
    "description." & _
    vbNewline & vbNewline & _
    "This class can be generated from the Word table in the document " & _
    "qbPolish.DOC which lists our operators, using qbOps2Doc.EXE."

    ' ----------------------------------------------------------------------
    ' Return Easter yEgg
    '
    '
    Public Shared ReadOnly Property About As String
        Get
            Return(ABOUTINFO)
        End Get        
    End Property    

    ' ----------------------------------------------------------------------
    ' Return classname
    '
    '
    Public Shared ReadOnly Property Name As String
        Get
            Return("qbOp")
        End Get        
    End Property    

    ' ----------------------------------------------------------------------
    ' Translates opcode name to its enumerator
    '
    '
    Public Shared Function opCodeFromString(ByVal strOpcode As String) As ENUop
        Dim strOpcodeWork As String = UCase(Trim(strOpcode))
        If Len(strOpcodeWork) < 2 OrElse Mid(strOpcodeWork, 1, 2) <> "OP" Then 
            strOpcodeWork = "OP" & strOpcodeWork
        End If
        Select Case strOpcodeWork
            Case "OPADD":               Return(ENUop.opAdd)
            Case "OPAND":               Return(ENUop.opAnd)
            Case "OPASC":               Return(ENUop.opAsc)
            Case "OPCEIL":              Return(ENUop.opCeil)
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case "OPCHOOSE":            Return(ENUop.opChoose)
            #End If
            Case "OPCHR":               Return(ENUop.opChr)
            Case "OPCIRCLE":            Return(ENUop.opCircle)
            Case "OPCLS":               Return(ENUop.opCls)
            Case "OPCOGO":              Return(ENUop.opCoGo)
            Case "OPCONCAT":            Return(ENUop.opConcat)
            Case "OPCOS":               Return(ENUop.opCos)
            Case "OPDIVIDE":            Return(ENUop.opDivide)
            Case "OPDUPLICATE":         Return(ENUop.opDuplicate)
            Case "OPEND":               Return(ENUop.opEnd)
            Case "OPEVAL":              Return(ENUop.opEval)
            Case "OPEVALUATE":          Return(ENUop.opEvaluate)
            Case "OPFLOOR":             Return(ENUop.opFloor)
            Case "OPFORTEST":           Return(ENUop.opForTest)
            Case "OPIIF":               Return(ENUop.opIif)
            Case "OPIN":                Return(ENUop.opIn)
            Case "OPINPUT":             Return(ENUop.opInput)
            Case "OPINT":               Return(ENUop.opInt)
            Case "OPINVALID":           Return(ENUop.opInvalid)
            Case "OPISNUMERIC":         Return(ENUop.opIsNumeric)
            Case "OPJUMP":              Return(ENUop.opJump)
            Case "OPJUMPINDIRECT":      Return(ENUop.opJumpIndirect)
            Case "OPJUMPNZ":            Return(ENUop.opJumpNZ)
            Case "OPJUMPZ":             Return(ENUop.opJumpZ)
            Case "OPLABEL":             Return(ENUop.opLabel)
            Case "OPLCASE":             Return(ENUop.opLCase)
            Case "OPLEN":               Return(ENUop.opLen)
            Case "OPLIKE":              Return(ENUop.opLike)
            Case "OPLOG":               Return(ENUop.opLog)
            Case "OPMAX":               Return(ENUop.opMax)
            Case "OPMID":               Return(ENUop.opMid)
            Case "OPMIN":               Return(ENUop.opMin)
            Case "OPMOD":               Return(ENUop.opMod)
            Case "OPMULTIPLY":          Return(ENUop.opMultiply)
            Case "OPNEGATE":            Return(ENUop.opNegate)
            Case "OPNOP":               Return(ENUop.opNop)
            Case "OPNOT":               Return(ENUop.opNot)
            Case "OPOR":                Return(ENUop.opOr)
            Case "OPPOP":               Return(ENUop.opPop)
            Case "OPPOPINDIRECT":       Return(ENUop.opPopIndirect)
            Case "OPPOPOFF":            Return(ENUop.opPopOff)
            Case "OPPOWER":             Return(ENUop.opPower)
            Case "OPPRINT":             Return(ENUop.opPrint)
            Case "OPPUSH":              Return(ENUop.opPush)
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case "OPPUSHDATEDIFF":      Return(ENUop.opPushDateDIFF)
            #End If
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case "OPPUSHDATEEQ":        Return(ENUop.opPushDateEQ)
            #End If
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case "OPPUSHDATEGT":        Return(ENUop.opPushDateGT)
            #End If
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case "OPPUSHDATELT":        Return(ENUop.opPushDateLT)
            #End If
            Case "OPPUSHEQ":            Return(ENUop.opPushEQ)
            Case "OPPUSHGE":            Return(ENUop.opPushGE)
            Case "OPPUSHGT":            Return(ENUop.opPushGT)
            Case "OPPUSHINDIRECT":      Return(ENUop.opPushIndirect)
            Case "OPPUSHLE":            Return(ENUop.opPushLE)
            Case "OPPUSHLITERAL":       Return(ENUop.opPushLiteral)
            Case "OPPUSHLT":            Return(ENUop.opPushLT)
            Case "OPPUSHNE":            Return(ENUop.opPushNE)
            Case "OPPUSHRETURN":        Return(ENUop.opPushReturn)
            Case "OPRAND":              Return(ENUop.opRand)
            Case "OPREAD":              Return(ENUop.opRead)
            Case "OPREM":               Return(ENUop.opRem)
            Case "OPREPLACE":           Return(ENUop.opReplace)
            Case "OPRND":               Return(ENUop.opRnd)
            Case "OPRNDSEED":           Return(ENUop.opRndSeed)
            Case "OPROTATE":            Return(ENUop.opRotate)
            Case "OPROUND":             Return(ENUop.opRound)
            Case "OPSGN":               Return(ENUop.opSgn)
            Case "OPSIN":               Return(ENUop.opSin)
            Case "OPSQR":               Return(ENUop.opSqr)
            Case "OPSTRING":            Return(ENUop.opString)
            Case "OPSUBTRACT":          Return(ENUop.opSubtract)
            Case "OPTRACE":             Return(ENUop.opTrace)
            Case "OPTRACEPOP":          Return(ENUop.opTracePop)
            Case "OPTRACEPUSH":         Return(ENUop.opTracePush)
            Case "OPTRIM":              Return(ENUop.opTrim)
            Case "OPUCASE":             Return(ENUop.opUCase)
            #If QUICKBASICENGINE_EXTENSION Then
            Case "OPUTILITY":           Return(ENUop.opUtility)
            #End If
            Case Else:                 
                _OBJutilities.errorHandler("Invalid opcode name " & _
                                           _OBJutilities.enquote(strOpcode), _
                                           Name, _
                                           "opcodeFromString", _
                                           "Returning the Invalid op code")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Translates opcode to extended description
    '
    '
    Public Shared Function opCodeToDescription(ByVal enuOpcode As ENUop) As String
        Select Case enuOpcode
            Case ENUop.opAdd:           Return("opAdd: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)+stack(top) ")
            Case ENUop.opAnd:           Return("opAnd: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1) And " & _
                                        "stack(top) (And is not " & _
                                        "short circuited) ")
            Case ENUop.opAsc:           Return("opAsc: Replaces " & _
                                        "stack(top) by its ASCII " & _
                                        "value ")
            Case ENUop.opCeil:          Return("opCeil: Replaces " & _
                                        "stack(top) with first " & _
                                        "integer n > stack(top) ")
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case ENUop.opChoose:        Return("opChoose: Returns the " & _
                                        "nth value of the Choose " & _
                                        "stack frame: expects n, " & _
                                        "list and count." & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True")
            #End If
            Case ENUop.opChr:           Return("opChr: Replaces " & _
                                        "stack(top) with its " & _
                                        "ASCII character value ")
            Case ENUop.opCircle:        Return("opCircle: Draws a " & _
                                        "circle: stack(top - 2) " & _
                                        "is x coordinate, " & _
                                        "stack(top - 1) is y " & _
                                        "coordinate, and " & _
                                        "stack(top) is radius ")
            Case ENUop.opCls:           Return("opCls: Clears the " & _
                                        "simulated QuickBasic " & _
                                        "screen ")
            Case ENUop.opCoGo:          Return("opCoGo: Computed " & _
                                        "GoTo/GoSub: translate " & _
                                        "string/number at " & _
                                        "stack(top) to address ")
            Case ENUop.opConcat:        Return("opConcat: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)&stack(top) ")
            Case ENUop.opCos:           Return("opCos: Replaces " & _
                                        "stack(top) with its " & _
                                        "cosine ")
            Case ENUop.opDivide:        Return("opDivide: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)/stack(top) ")
            Case ENUop.opDuplicate:     Return("opDuplicate: Copies " & _
                                        "stack(top-n) to stack " & _
                                        "top ")
            Case ENUop.opEnd:           Return("opEnd: Stops processing " & _
                                        "immediately ")
            Case ENUop.opEval:          Return("opEval: Evaluates " & _
                                        "stack(top) as a " & _
                                        "Quickbasic expression " & _
                                        "using lightweight " & _
                                        "evaluation: a new " & _
                                        "quickBasicEngine with " & _
                                        "default options is used " & _
                                        "to evaluate stack(top) ")
            Case ENUop.opEvaluate:      Return("opEvaluate: Evaluates " & _
                                        "stack(top) as a " & _
                                        "Quickbasic expression " & _
                                        "using heavyweight " & _
                                        "evaluation: a new " & _
                                        "quickBasicEngine with " & _
                                        "the same options as the " & _
                                        "current engine is used " & _
                                        "to evaluate stack(top) ")
            Case ENUop.opFloor:         Return("opFloor: Replaces " & _
                                        "stack(top) with first " & _
                                        "integer n < stack(top) ")
            Case ENUop.opForTest:       Return("opForTest: Stacks " & _
                                        "stack(top-1)>=0 And " & _
                                        "c(stack(top))<stack(top-2) " & _
                                        "Or stack(top-1)<0 And " & _
                                        "c(stack(top))>stack(top-2) ")
            Case ENUop.opIif:           Return("opIif: Replaces " & _
                                        "stack(top-2)..stack(top) " & _
                                        "with stack(top-1) when " & _
                                        "stack(top-2) is True, " & _
                                        "with stack(top) " & _
                                        "otherwise ")
            #If QUICKBASICENGINE_FMSEXTENSION Then                                        
            Case ENUop.opIn:            Return("opIn: Searches the IN " & _
                                        "stack frame for a " & _
                                        "value. Expects the " & _
                                        "search value, the frame " & _
                                        "size, and n values for " & _
                                        "the search at the top " & _
                                        "of the stack." & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True")
            #End If                                        
            Case ENUop.opInput:         Return("opInput: Reads a number " & _
                                        "or a string to " & _
                                        "stack(top) ")
            Case ENUop.opInt:           Return("opInt: Replaces " & _
                                        "stack(top) with integer " & _
                                        "part ")
            Case ENUop.opInvalid:       Return("opInvalid: Invalid " & _
                                        "marker op ")
            Case ENUop.opIsNumeric:     Return("opIsNumeric: Replaces " & _
                                        "stack(top) with True " & _
                                        "when stack(top) is a " & _
                                        "number, False otherwise ")
            Case ENUop.opJump:          Return("opJump: Jumps to " & _
                                        "location ")
            Case ENUop.opJumpIndirect:  Return("opJumpIndirect: Jumps " & _
                                        "to location identified " & _
                                        "at the top of the stack ")
            Case ENUop.opJumpNZ:        Return("opJumpNZ: Jumps to " & _
                                        "location when " & _
                                        "stack(top) <> 0 (pop " & _
                                        "the stack top) ")
            Case ENUop.opJumpZ:         Return("opJumpZ: Jumps to " & _
                                        "location when " & _
                                        "stack(top) = 0 (pop the " & _
                                        "stack top) ")
            Case ENUop.opLabel:         Return("opLabel: Identifies " & _
                                        "position of a code " & _
                                        "label or statement " & _
                                        "number ")
            Case ENUop.opLCase:         Return("opLCase: Replaces the " & _
                                        "string at stack(top) " & _
                                        "with its lower case " & _
                                        "translation ")
            Case ENUop.opLen:           Return("opLen: Replace " & _
                                        "stack(top) by its " & _
                                        "length as a string ")
            Case ENUop.opLike:          Return("opLike: Compare two " & _
                                        "strings at the stack " & _
                                        "top for a pattern " & _
                                        "match, replacing them " & _
                                        "by True or False ")
            Case ENUop.opLog:           Return("opLog: Replaces " & _
                                        "stack(top) by its " & _
                                        "logarithm ")
            Case ENUop.opMax:           Return("opMax: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with the " & _
                                        "max ")
            Case ENUop.opMid:           Return("opMid: Replaces " & _
                                        "stack(top-2)..stack(top) " & _
                                        "with the substring of " & _
                                        "stack(top-2) starting " & _
                                        "at stack(top-1) for " & _
                                        "length at stack(top) ")
            Case ENUop.opMin:           Return("opMin: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with the " & _
                                        "min ")
            Case ENUop.opMod:           Return("opMod: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with the " & _
                                        "integer division " & _
                                        "remainder from " & _
                                        "stack(top-1) \ " & _
                                        "stack(top) ")
            Case ENUop.opMultiply:      Return("opMultiply: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)*stack(top) ")
            Case ENUop.opNegate:        Return("opNegate: Reverse sign " & _
                                        "at top of stack ")
            Case ENUop.opNop:           Return("opNop: Does nothing " & _
                                        "(infinitely useful: cf. " & _
                                        "Tao t'eh Ch'ing) ")
            Case ENUop.opNot:           Return("opNot: Replaces " & _
                                        "stack(top) with Not " & _
                                        "stack(top) ")
            Case ENUop.opOr:            Return("opOr: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1) Or " & _
                                        "stack(top) (Or is not " & _
                                        "short circuited) ")
            Case ENUop.opPop:           Return("opPop: Sends stack(top) " & _
                                        "to a memory location ")
            Case ENUop.opPopIndirect:   Return("opPopIndirect: Sends " & _
                                        "stack(top) to a memory " & _
                                        "location at " & _
                                        "stack(top-1): removes " & _
                                        "stack(top), leaves " & _
                                        "stack(top-1) alone ")
            Case ENUop.opPopOff:        Return("opPopOff: Removes " & _
                                        "stack(top) without " & _
                                        "sending it to a memory " & _
                                        "location ")
            Case ENUop.opPower:         Return("opPower: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)^stack(top) ")
            Case ENUop.opPrint:         Return("opPrint: Prints (and " & _
                                        "removes) value at top " & _
                                        "of the stack ")
            Case ENUop.opPush:          Return("opPush: Pushes the " & _
                                        "contents of a memory " & _
                                        "location in the Case op ")
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case ENUop.opPushDateDIFF:  Return("opPushDateDIFF: " & _
                                        "Replaces stack(top-2), " & _
                                        "stack(top-1) and " & _
                                        "stack(top) by the " & _
                                        "elapsed time from " & _
                                        "stack(top-2) to " & _
                                        "stack(top-1). " & _
                                        "stack(top) should be " & _
                                        "the interval as " & _
                                        "seconds, minutes, " & _
                                        "hours, days, months or " & _
                                        "years. If stack(top) is " & _
                                        "a null string then the " & _
                                        "interval is days when " & _
                                        "stack(top-2) and " & _
                                        "stack(top-1) are both " & _
                                        "single words, seconds " & _
                                        "otherwise & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True")
            #End If
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case ENUop.opPushDateEQ:    Return("opPushDateEQ: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)=stack(top) " & _
                                        "as a date, 0 otherwise." & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True"
            #End If
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case ENUop.opPushDateGT:    Return("opPushDateGT: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)>stack(top) " & _
                                        "as a date, 0 otherwise." & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True"
            #End If
            #If QUICKBASICENGINE_FMSEXTENSION Then
            Case ENUop.opPushDateLT:    Return("opPushDateLT: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<stack(top) " & _
                                        "as a date, 0 otherwise." & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True"
            #End If
            Case ENUop.opPushEQ:        Return("opPushEQ: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)=stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushGE:        Return("opPushGE: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)>=stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushGT:        Return("opPushGT: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)>stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushIndirect:  Return("opPushIndirect: Pushes " & _
                                        "the contents of a " & _
                                        "memory location indexed " & _
                                        "at " & _
                                        "stack(top)...replacing " & _
                                        "the index ")
            Case ENUop.opPushLE:        Return("opPushLE: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<=stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushLiteral:   Return("opPushLiteral: Pushes a " & _
                                        "literal string or " & _
                                        "number ")
            Case ENUop.opPushLT:        Return("opPushLT: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushNE:        Return("opPushNE: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<>stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushReturn:    Return("opPushReturn: Pushes " & _
                                        "the subroutine's return " & _
                                        "address ")
            Case ENUop.opRand:          Return("opRand: Seeds the " & _
                                        "random number generator " & _
                                        "to unpredictable values ")
            Case ENUop.opRead:          Return("opRead: Reads from the " & _
                                        "data statements to " & _
                                        "stack(top) ")
            Case ENUop.opRem:           Return("opRem: Equivalent to a " & _
                                        "Nop ")
            Case ENUop.opReplace:       Return("opReplace: Replaces all " & _
                                        "occurences of the " & _
                                        "string at stack(top-1) " & _
                                        "by the string at " & _
                                        "stack(top) in the " & _
                                        "string at stack(top-2). " & _
                                        "Replaces all entries by " & _
                                        "the translated string ")
            Case ENUop.opRnd:           Return("opRnd: Pushes an " & _
                                        "unseeded random number " & _
                                        "on the stack ")
            Case ENUop.opRndSeed:       Return("opRndSeed: Pushes a " & _
                                        "seeded random number on " & _
                                        "the stack (seed is " & _
                                        "stack(top), and is " & _
                                        "replaced) ")
            Case ENUop.opRotate:        Return("opRotate: Exchanges " & _
                                        "stack(top) with " & _
                                        "stack(top-n) ")
            Case ENUop.opRound:         Return("opRound: Rounds " & _
                                        "stack(top-1) to " & _
                                        "stack(top) digits ")
            Case ENUop.opSgn:           Return("opSgn: Replaces " & _
                                        "stack(top) with its " & _
                                        "signum (0 for 0: 1 for " & _
                                        "positive: -1 for " & _
                                        "negative) ")
            Case ENUop.opSin:           Return("opSin: Replaces " & _
                                        "stack(top) with its " & _
                                        "sine ")
            Case ENUop.opSqr:           Return("opSqr: Replaces " & _
                                        "stack(top) with its " & _
                                        "square root ")
            Case ENUop.opString:        Return("opString: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with n " & _
                                        "copies of the character " & _
                                        "at stack top, where n " & _
                                        "is at stack(top-1) ")
            Case ENUop.opSubtract:      Return("opSubtract: Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)-stack(top) ")
            Case ENUop.opTrace:         Return("opTrace: Changes trace " & _
                                        "settings ")
            Case ENUop.opTracePop:      Return("opTracePop: Restores " & _
                                        "trace settings from a " & _
                                        "LIFO stack ")
            Case ENUop.opTracePush:     Return("opTracePush: Saves " & _
                                        "trace settings in a " & _
                                        "LIFO stack ")
            Case ENUop.opTrim:          Return("opTrim: Replaces " & _
                                        "stack(top) with trimmed " & _
                                        "string (leading and " & _
                                        "trailing blanks " & _
                                        "removed) ")
            Case ENUop.opUCase:         Return("opUCase: Replaces the " & _
                                        "string at stack(top) " & _
                                        "with its upper case " & _
                                        "translation ")
            #If QUICKBASICENGINE_EXTENSION Then
            Case ENUop.opUtility:       Return("opUtility: Replaces " & _
                                        "stack(top-n)..stack(top) " & _
                                        "(where n=stack(top-1)) " & _
                                        "with the result of " & _
                                        "calling the utility " & _
                                        "named in stack(top)." & _
                                        vbNewline & vbNewline & _
                                        "This op is generated " & _
                                        "only when the preprocessor " & _
                                        "variable " & _
                                        "QUICKBASICENGINE_FMSEXTENSION " & _
                                        "is True"
            #End If
            Case Else:
                _OBJutilities.errorHandler("Unexpected enumerator value", _
                                           Name, "opcode2Description", _
                                           "Returning error information")
                Return("Unexpected enumerator value")                                           
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Translates opcode to its name
    '
    '
    Public Shared Function opCodeToString(ByVal enuOpcode As ENUop) As String
        Return(_OBJutilities.item(opCodeToDescription(enuOpcode), 1, ":", False))
    End Function

End Class