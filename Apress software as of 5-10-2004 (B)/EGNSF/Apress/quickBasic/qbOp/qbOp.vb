Option Strict On

Imports System.Reflection.Emit

' *********************************************************************
' *                                                                   *
' * qbOp     QuickBasicEngine operators                               *
' *                                                                   *
' * This stateless class identifies the operators supported by the    *
' * non-CLR Nutty Professor machine as a large enumerator, and it     *
' * provides Shared conversion tools for enumerator values.           *
' *                                                                   *     
' * The rest of this document describes:                              *
' *                                                                   * 
' *                                                                   * 
' *      *  The properties and methods exposed by qbOp                *
' *      *  Procedure for adding a new operator                       *
' *      *  Change record                                             *
' *                                                                   * 
' *                                                                   * 
' * PROPERTIES AND METHODS ------------------------------------------ *                     
' *                                                                   *                      
' * Properties of this class start with an upper case letter; methods *
' * start with a lower case letter.                                   *
' *                                                                   *
' * Any method, which does not otherwise return a value, will return  *
' * True on success or False on failure.                              *
' *                                                                   * 
' *                                                                   * 
' *      *  About: this read-only Shared property returns information * 
' *         about the class                                           * 
' *                                                                   * 
' *      *  class2XML: this Shared method returns information about   * 
' *         the class in XML format                                   * 
' *                                                                   * 
' *      *  ClassName: this Shared property returns the class name    * 
' *         qbOp                                                      * 
' *                                                                   * 
' *      *  IsJumpOp: this read-only Shared property returns True when* 
' *         the operator is a jump operator, that the assembler needs * 
' *         to resolve, False otherwise                               * 
' *                                                                   * 
' *      *  opCodeFromString(c): where c is the case-independent op   * 
' *         name this Shared method returns the op enumerator         * 
' *                                                                   * 
' *      *  opCodeToDescription(x): where x is an op code specified   * 
' *         as a string or as an enumerator this Shared method        * 
' *         returns the op code's description. Note that the          * 
' *         description will be in this format:                       *
' *                                                                   *
' *         op(template): text                                        *
' *                                                                   *
' *      *  opCodeToStackTemplate(x): where x is an op code specified *
' *         as a string or enumerator this Shared method returns the  *
' *         template of expected operands for this opcode.            *
' *                                                                   *
' *         The template is a string and the comma-separated list of  *
' *         expected stack values, from lower down in the stack to the*
' *         top of the stack. The template is defined inside the op   * 
' *         description statement in the opCodeToDescription method.  *
' *                                                                   *
' *         Each stack value must be one of the following:            *
' *                                                                   *
' *         + x: any qbVariable is permitted at this position         *
' *                                                                   *
' *         + s: any scalar qbVariable is permitted                   *
' *                                                                   *
' *         + n: any numeric qbVariable is permitted                  *
' *                                                                   *
' *         + i: any numeric integer qbVariable is permitted          *
' *                                                                   *
' *         + u: operator expects the utility stack frame: stack(top) *
' *           is an operand count: stack(top+1) is the name of a      *
' *           utility: stack(top+n+1)..stack(top+2) are the operands. *     
' *                                                                   *
' *         + <name>: where name is the name of one of the values of  *
' *           the ENUvarType enumerator, this specifies that the      *
' *           stack value is restricted to the varType                *
' *                                                                   *
' *         + a: an array index frame is expected at this location,   *
' *           in the form i(1), i(2)...i(n), count, array, where:     *
' *                                                                   *
' *           - i(n) is the index at dimension n                      *
' *           - count is the number of preceding indexes              *
' *           - array is a qbVariable with the type array             *
' *                                                                   *
' *      *  opCodeToString(x): where x is an op code specified        * 
' *         as a string or as an enumerator this Shared method        * 
' *         returns the op code's name only                           * 
' *                                                                   * 
' *                                                                   * 
' * PROCEDURE FOR ADDING A NEW OPERATOR ----------------------------- *
' *                                                                   * 
' * To add a new operator:                                            * 
' *                                                                   * 
' *                                                                   * 
' *      1.  Add the operator name as a blank-delimited word to the   * 
' *          OPERATORS constant                                       *
' *                                                                   * 
' *      2.  Add operator as an ENUop value                           *
' *                                                                   * 
' *      3.  Add operator as a case to opcodeFromString               *
' *                                                                   * 
' *      4.  Add operator as a case to opcodeToDescription            *
' *                                                                   * 
' *      5.  In quickBasicEngine's interpreter_ method add the case   *
' *          that implements the new operator                         *
' *                                                                   * 
' *                                                                   * 
' *                                                                   * 
' * C H A N G E   R E C O R D --------------------------------------- *                      
' *   DATE     PROGRAMMER   D E S C R I P T I O N                     *
' * --------   ----------   ----------------------------------------- *                      
' * 5/17/03    Nilges       Version1                                  *
' * 6/20/03    Nilges       Numerous manual changes                   *
' * 08 19 03   Nilges       Added isJumpOp                            *
' * 09 07 03   Nilges       Added stack expected                      *
' * 09 07 03   Nilges       Added documentation                       *
' * 09 07 03   Nilges       Removed and archived FMS                  *
' * 09 07 03   Nilges       Removed test method                       * 
' * 12 14 03   Nilges       Changed desc of duplicate                 * 
' * 12 16 03   Nilges       ForTest is a jump op                      *
' * 12 16 03   Nilges       Added ForIncrement                        *
' * 01 02 04   Nilges       Added opcodeToMSIL                        *
' * 02 22 04   Nilges       Bug: forIncrement op not converted to     *
' *                         string: added conversion code             *
' * 04 14 04   Nilges       Updated documentation of pushIndirect     *
' * 04 15 04   Nilges       Added popToArrayElement                    *
' *                                                                   *
' *********************************************************************

Public Class qbOp

    ' ***** Shared utilities *****
    Private Shared _OBJutilities As utilities.utilities
    
    ' ***** Operators *****
    ' --- Operator list
    Private Const OPERATORS As String = _
    "add and asc ceil chr circle cls cogo concat cos divide duplicate end " & _
    "eval evaluate floor forTest forIncrement iif input int invalid isnumeric jump jumpIndirect " & _
    "jumpNZ jumpZ label lCase len like log max mid min mod multiply negate nop not " & _
    "or popToArrayElement pop popIndirect popOff power print push pushEQ pushGT pushIndirect " & _
    "pushLE pushLiteral pushLT pushNE pushReturn rand read rem replace " & _
    "rnd rndSeed rotate round sgn sin sqr string subtract trace tracePop tracePush " & _
    "trim uCase utility"
    ' --- Operator enumerator
    Public Enum ENUop
        opAdd          ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)+stack(top):
                       ' expects numeric,numeric
        opAnd          ' Replaces stack(top) and stack(top-1) with stack(top-1) And
                       ' stack(top) (And is not short circuited):
                       ' expects numeric,numeric
        opAsc          ' Replaces stack(top) by its ASCII value:
                       ' expects string 
        opCeil         ' Replaces stack(top) with first integer n > stack(top):
                       ' expects numeric
        opChr          ' Replaces stack(top) with its ASCII character value:
                       ' expects string
        opCircle       ' Draws a circle: stack(top - 2) is x coordinate, stack(top
                       ' - 1) is y coordinate, and stack(top) is radius:
                       ' expects numeric, numeric, numeric
        opCls          ' Clears the simulated QuickBasic screen: expects nothing
        opCoGo         ' Computed GoTo/GoSub: translate string/number at stack(top)
                       ' to address: expects string
        opConcat       ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)&stack(top): expects string,string
        opCos          ' Replaces stack(top) with its cosine: expects numeric
        opDivide       ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)/stack(top): expects numeric
        opDuplicate    ' Duplicates the stack at top-operand: expects something
        opEnd          ' Stops processing immediately: expects nothing
        opEval         ' Evaluates stack(top) as a Quickbasic expression using
                       ' lightweight evaluation: a new quickBasicEngine with default
                       ' options is used to evaluate stack(top): expects string
        opEvaluate     ' Evaluates stack(top) as a Quickbasic expression using
                       ' heavyweight evaluation: a new quickBasicEngine with the same
                       ' options as the current engine is used to evaluate stack(top):
                       ' expects string
        opFloor        ' Replaces stack(top) with first integer n < stack(top):
                       ' expects number
        opForIncrement ' Increments or decrements the For control value  
                       '
                       ' Expects number,number (step value and control variable 
                       ' location)
        opForTest      ' Jumps to for exit when contents of control variable
                       ' location are greater than final value (when step value is
                       ' positive): jumps to for exit when contents of 
                       ' control variable location are less.  
                       '
                       ' Expects number,number,number (final value, step value
                       ' and control variable location)
        opIif          ' Replaces stack(top-2)..stack(top) with stack(top-1) when
                       ' stack(top-2) is True, with stack(top) otherwise:
                       ' expects number,anything,anything
        opInput        ' Reads a number or a string to stack(top)
        opInt          ' Replaces stack(top) with integer part: expects
        opInvalid      ' Invalid marker op
        opIsNumeric    ' Replaces stack(top) with True when stack(top) is a number,
                       ' False otherwise: expects nothing
        opJump         ' Jumps to location: expects integer
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
        opNop          ' Does nothing  
        opNot          ' Replaces stack(top) with Not stack(top)
        opOr           ' Replaces stack(top) and stack(top-1) with stack(top-1) Or
                       ' stack(top) (Or is not short circuited)
        opPop          ' Sends stack(top) to a memory location
        opPopIndirect  ' Sends stack(top) to a memory location at stack(top-1):
                       ' removes stack(top), leaves stack(top-1) alone
        opPopOff       ' Removes stack(top) without sending it to a memory location
        opPopToArrayElement  ' Expecting the stack frame arrayReference, count,
                             ' subscript1, ... subscriptn, value, this op changes
                             ' the value at the subscripted array location to the
                             ' value on the stack, and destroys the stack frame
        opPower        ' Replaces stack(top) and stack(top-1) with
                       ' stack(top-1)^stack(top)
        opPrint        ' Prints (and removes) value at top of the stack
        opPush         ' Pushes the contents of a memory location in the Case op
        opPushArrayElement  ' Expecting the stack frame arrayReference, count,
                            ' subscript1, ... subscriptn, this op replaces this
                            ' stack frame by the subscripted value
        opPushEQ       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)=stack(top), 0 otherwise
        opPushGE       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)>=stack(top), 0 otherwise
        opPushGT       ' Replaces stack(top) and stack(top-1) by -1 when
                       ' stack(top-1)>stack(top), 0 otherwise
        opPushIndirect ' Pushes the contents of a memory location indexed at
                       ' stack(top)...replacing the index. If the location
                       ' contains a scalar (or a variant with scalar value), 
                       ' this value is cloned for copying
                       ' to the stack. If the location contains an array or
                       ' a UDT, the array or UDT itself is pushed on the stack.
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
        opRotate       ' Exchanges stack(top) with stack(top-n): when n=0 this is a NOP
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
        opUtility      ' Replaces stack(top-n-1)..stack(top) (where n=stack(top))
                       ' with the result of calling the utility named in stack(top-1)
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

    Private Const CLASS_NAME As String = "qbOp"

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
    ' Return name of the class
    '
    '
    Public Shared ReadOnly Property ClassName As String
        Get
            Return(CLASS_NAME)
        End Get        
    End Property    

    ' ----------------------------------------------------------------------
    ' Returns True for jump operations only (operations that have to be
    ' assembled)
    '
    '
    Public Shared Function isJumpOp(ByVal enuOpcode As ENUop) As Boolean
        Select Case enuOpcode
            Case enuOpcode.opJump:        Return(True) 
            Case enuOpcode.opJumpNZ:      Return(True)
            Case enuOpcode.opJumpZ:       Return(True)
            Case enuOpcode.opPushReturn : Return (True)
            Case enuOpcode.opForTest : Return (True)
        End Select
        Return (False)
    End Function

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
            Case "OPADD" : Return (ENUop.opAdd)
            Case "OPAND" : Return (ENUop.opAnd)
            Case "OPASC" : Return (ENUop.opAsc)
            Case "OPCEIL" : Return (ENUop.opCeil)
            Case "OPCHR" : Return (ENUop.opChr)
            Case "OPCIRCLE" : Return (ENUop.opCircle)
            Case "OPCLS" : Return (ENUop.opCls)
            Case "OPCOGO" : Return (ENUop.opCoGo)
            Case "OPCONCAT" : Return (ENUop.opConcat)
            Case "OPCOS" : Return (ENUop.opCos)
            Case "OPDIVIDE" : Return (ENUop.opDivide)
            Case "OPDUPLICATE" : Return (ENUop.opDuplicate)
            Case "OPEND" : Return (ENUop.opEnd)
            Case "OPEVAL" : Return (ENUop.opEval)
            Case "OPEVALUATE" : Return (ENUop.opEvaluate)
            Case "OPFLOOR" : Return (ENUop.opFloor)
            Case "OPFORTEST" : Return (ENUop.opForTest)
            Case "OPFORINCREMENT" : Return (ENUop.opForIncrement)
            Case "OPIIF" : Return (ENUop.opIif)
            Case "OPINPUT" : Return (ENUop.opInput)
            Case "OPINT" : Return (ENUop.opInt)
            Case "OPINVALID" : Return (ENUop.opInvalid)
            Case "OPISNUMERIC" : Return (ENUop.opIsNumeric)
            Case "OPJUMP" : Return (ENUop.opJump)
            Case "OPJUMPINDIRECT" : Return (ENUop.opJumpIndirect)
            Case "OPJUMPNZ" : Return (ENUop.opJumpNZ)
            Case "OPJUMPZ" : Return (ENUop.opJumpZ)
            Case "OPLABEL" : Return (ENUop.opLabel)
            Case "OPLCASE" : Return (ENUop.opLCase)
            Case "OPLEN" : Return (ENUop.opLen)
            Case "OPLIKE" : Return (ENUop.opLike)
            Case "OPLOG" : Return (ENUop.opLog)
            Case "OPMAX" : Return (ENUop.opMax)
            Case "OPMID" : Return (ENUop.opMid)
            Case "OPMIN" : Return (ENUop.opMin)
            Case "OPMOD" : Return (ENUop.opMod)
            Case "OPMULTIPLY" : Return (ENUop.opMultiply)
            Case "OPNEGATE" : Return (ENUop.opNegate)
            Case "OPNOP" : Return (ENUop.opNop)
            Case "OPNOT" : Return (ENUop.opNot)
            Case "OPOR" : Return (ENUop.opOr)
            Case "OPPOP" : Return (ENUop.opPop)
            Case "OPPOPINDIRECT" : Return (ENUop.opPopIndirect)
            Case "OPPOPOFF" : Return (ENUop.opPopOff)
            Case "OPPOPTOARRAYELEMENT" : Return (ENUop.opPopToArrayElement)
            Case "OPPOWER" : Return (ENUop.opPower)
            Case "OPPRINT" : Return (ENUop.opPrint)
            Case "OPPUSH" : Return (ENUop.opPush)
            Case "OPPUSHARRAYELEMENT" : Return (ENUop.opPushArrayElement)
            Case "OPPUSHEQ" : Return (ENUop.opPushEQ)
            Case "OPPUSHGE" : Return (ENUop.opPushGE)
            Case "OPPUSHGT" : Return (ENUop.opPushGT)
            Case "OPPUSHINDIRECT" : Return (ENUop.opPushIndirect)
            Case "OPPUSHLE" : Return (ENUop.opPushLE)
            Case "OPPUSHLITERAL" : Return (ENUop.opPushLiteral)
            Case "OPPUSHLT" : Return (ENUop.opPushLT)
            Case "OPPUSHNE" : Return (ENUop.opPushNE)
            Case "OPPUSHRETURN" : Return (ENUop.opPushReturn)
            Case "OPRAND" : Return (ENUop.opRand)
            Case "OPREAD" : Return (ENUop.opRead)
            Case "OPREM" : Return (ENUop.opRem)
            Case "OPREPLACE" : Return (ENUop.opReplace)
            Case "OPRND" : Return (ENUop.opRnd)
            Case "OPRNDSEED" : Return (ENUop.opRndSeed)
            Case "OPROTATE" : Return (ENUop.opRotate)
            Case "OPROUND" : Return (ENUop.opRound)
            Case "OPSGN" : Return (ENUop.opSgn)
            Case "OPSIN" : Return (ENUop.opSin)
            Case "OPSQR" : Return (ENUop.opSqr)
            Case "OPSTRING" : Return (ENUop.opString)
            Case "OPSUBTRACT" : Return (ENUop.opSubtract)
            Case "OPTRACE" : Return (ENUop.opTrace)
            Case "OPTRACEPOP" : Return (ENUop.opTracePop)
            Case "OPTRACEPUSH" : Return (ENUop.opTracePush)
            Case "OPTRIM" : Return (ENUop.opTrim)
            Case "OPUCASE" : Return (ENUop.opUCase)
            Case "OPUTILITY" : Return (ENUop.opUtility)
            Case Else
                _OBJutilities.errorHandler("Invalid opcode name " & _
                                           _OBJutilities.enquote(strOpcode), _
                                           ClassName, _
                                           "opcodeFromString", _
                                           "Returning the Invalid op code")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Translates opcode to extended description
    '
    '
    ' Note when adding or changing an op code that its description must have
    ' the format name(stack frame): description.
    '
    '
    ' --- String is arg
    Public Overloads Shared Function opCodeToDescription(ByVal strOpcode As String) As String
        Return (opCodeToDescription(opCodeFromString(strOpcode)))
    End Function
    ' --- Evaluation is arg
    Public Overloads Shared Function opCodeToDescription(ByVal enuOpcode As ENUop) As String
        Select Case enuOpcode
            Case ENUop.opAdd : Return ("opAdd(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)+stack(top) ")
            Case ENUop.opAnd : Return ("opAnd(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1) And " & _
                                        "stack(top) (And is not " & _
                                        "short circuited) ")
            Case ENUop.opAsc : Return ("opAsc(s): Replaces " & _
                                        "stack(top) by its ASCII " & _
                                        "value ")
            Case ENUop.opCeil : Return ("opCeil(n): Replaces " & _
                                        "stack(top) with first " & _
                                        "integer n > stack(top) ")
            Case ENUop.opChr : Return ("opChr(i): Replaces " & _
                                        "stack(top) with its " & _
                                        "value as an ASCII character ")
            Case ENUop.opCircle : Return ("opCircle(n,n,n): Draws a " & _
                                        "circle: stack(top - 2) " & _
                                        "is x coordinate, " & _
                                        "stack(top - 1) is y " & _
                                        "coordinate, and " & _
                                        "stack(top) is radius ")
            Case ENUop.opCls : Return ("opCls: Clears the " & _
                                        "simulated QuickBasic " & _
                                        "screen ")
            Case ENUop.opCoGo : Return ("opCoGo(s): Computed " & _
                                        "GoTo/GoSub: translate " & _
                                        "string/number at " & _
                                        "stack(top) to address ")
            Case ENUop.opConcat : Return ("opConcat(s,s): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)&stack(top) ")
            Case ENUop.opCos : Return ("opCos(s): Replaces " & _
                                        "stack(top) with its " & _
                                        "cosine ")
            Case ENUop.opDivide : Return ("opDivide(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)/stack(top) ")
            Case ENUop.opDuplicate : Return ("opDuplicate: Duplicates the " & _
                                        "top of the stack")
            Case ENUop.opEnd : Return ("opEnd: Stops processing " & _
                                        "immediately ")
            Case ENUop.opEval : Return ("opEval(s): Evaluates " & _
                                        "stack(top) as a " & _
                                        "Quickbasic expression " & _
                                        "using lightweight " & _
                                        "evaluation: a new " & _
                                        "quickBasicEngine with " & _
                                        "default options is used " & _
                                        "to evaluate stack(top): " & _
                                        "note that stack(top) is replaced by " & _
                                        "its value")
            Case ENUop.opEvaluate : Return ("opEvaluate(s): Evaluates " & _
                                        "stack(top) as a " & _
                                        "Quickbasic expression " & _
                                        "using heavyweight " & _
                                        "evaluation: a new " & _
                                        "quickBasicEngine with " & _
                                        "the same options as the " & _
                                        "current engine is used " & _
                                        "to evaluate stack(top): " & _
                                        "note that stack(top) is replaced by " & _
                                        "its value")
            Case ENUop.opFloor : Return ("opFloor(n): Replaces " & _
                                        "stack(top) with first " & _
                                        "integer n < stack(top) ")
            Case ENUop.opForIncrement : Return ("opForIncrement(n,n): Increments " & _
                                                "or decrements the control variable")
            Case ENUop.opForTest : Return ("opForTest(n,n,n): Exits " & _
                                            "from the for loop when contents " & _
                                            "of control variable location " & _
                                            "are greater than final value " & _
                                            "(when step value is positive): " & _
                                            "exits from the for loop when " & _
                                            "contents of control variable " & _
                                            "location are less. Otherwise " & _
                                            "changes contents of control " & _
                                            "variable by step value")
            Case ENUop.opIif : Return ("opIif(n,x,x): Replaces " & _
                                        "stack(top-2)..stack(top) " & _
                                        "with stack(top-1) when " & _
                                        "stack(top-2) is True, " & _
                                        "with stack(top) " & _
                                        "otherwise ")
            Case ENUop.opInput : Return ("opInput: Reads a number " & _
                                        "or a string to " & _
                                        "stack(top) ")
            Case ENUop.opInt : Return ("opInt(n): Replaces " & _
                                        "stack(top) with integer " & _
                                        "part ")
            Case ENUop.opInvalid : Return ("opInvalid: Invalid " & _
                                        "marker op ")
            Case ENUop.opIsNumeric : Return ("opIsNumeric(x): Replaces " & _
                                        "stack(top) with True " & _
                                        "when stack(top) is a " & _
                                        "number, False otherwise ")
            Case ENUop.opJump : Return ("opJump: Jumps to " & _
                                        "location ")
            Case ENUop.opJumpIndirect : Return ("opJumpIndirect(i): Jumps " & _
                                        "to location identified " & _
                                        "at the top of the stack ")
            Case ENUop.opJumpNZ : Return ("opJumpNZ(n): Jumps to " & _
                                        "location when " & _
                                        "stack(top) <> 0 (pop " & _
                                        "the stack top) ")
            Case ENUop.opJumpZ : Return ("opJumpZ(n): Jumps to " & _
                                        "location when " & _
                                        "stack(top) = 0 (pop the " & _
                                        "stack top) ")
            Case ENUop.opLabel : Return ("opLabel: Identifies " & _
                                        "position of a code " & _
                                        "label or statement " & _
                                        "number ")
            Case ENUop.opLCase : Return ("opLCase(s): Replaces the " & _
                                        "string at stack(top) " & _
                                        "with its lower case " & _
                                        "translation ")
            Case ENUop.opLen : Return ("opLen(s): Replace " & _
                                        "stack(top) by its " & _
                                        "length as a string ")
            Case ENUop.opLike : Return ("opLike(s): Compare two " & _
                                        "strings at the stack " & _
                                        "top for a pattern " & _
                                        "match, replacing them " & _
                                        "by True or False ")
            Case ENUop.opLog : Return ("opLog(n): Replaces " & _
                                        "stack(top) by its " & _
                                        "logarithm ")
            Case ENUop.opMax : Return ("opMax(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with the " & _
                                        "max ")
            Case ENUop.opMid : Return ("opMid(s,i,i): Replaces " & _
                                        "stack(top-2)..stack(top) " & _
                                        "with the substring of " & _
                                        "stack(top-2) starting " & _
                                        "at stack(top-1) for " & _
                                        "length at stack(top) ")
            Case ENUop.opMin : Return ("opMin(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with the " & _
                                        "min ")
            Case ENUop.opMod : Return ("opMod(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with the " & _
                                        "integer division " & _
                                        "remainder from " & _
                                        "stack(top-1) \ " & _
                                        "stack(top) ")
            Case ENUop.opMultiply : Return ("opMultiply(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)*stack(top) ")
            Case ENUop.opNegate : Return ("opNegate(n): Reverse sign " & _
                                        "at top of stack ")
            Case ENUop.opNop : Return ("opNop: Does nothing")
            Case ENUop.opNot : Return ("opNot(n): Replaces " & _
                                        "stack(top) with Not " & _
                                        "stack(top) ")
            Case ENUop.opOr : Return ("opOr(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1) Or " & _
                                        "stack(top) (Or is not " & _
                                        "short circuited) ")
            Case ENUop.opPop : Return ("opPop(x): Sends stack(top) " & _
                                        "to a memory location ")
            Case ENUop.opPopToArrayElement : Return ("opPopToArrayElement(a,x): " & _
                                                     "Pops value to an array location")
            Case ENUop.opPopIndirect : Return ("opPopIndirect(i,x): Sends " & _
                                        "stack(top) to a memory " & _
                                        "location at " & _
                                        "stack(top-1): removes " & _
                                        "stack(top), leaves " & _
                                        "stack(top-1) alone ")
            Case ENUop.opPopOff : Return ("opPopOff(x): Removes " & _
                                        "stack(top) without " & _
                                        "sending it to a memory " & _
                                        "location ")
            Case ENUop.opPower : Return ("opPower(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)^stack(top) ")
            Case ENUop.opPrint : Return ("opPrint(x): Prints (and " & _
                                        "removes) value at top " & _
                                        "of the stack ")
            Case ENUop.opPush : Return ("opPush[n]: Pushes the " & _
                                        "contents of a memory " & _
                                        "location in the Case op, where " & _
                                        "the array element is at the top " & _
                                        "of the stack")
            Case ENUop.opPushArrayElement : Return ("opPush(a): Resolves a list " & _
                                            "of subscripts to an array location " & _
                                            "and replaces the list with the " & _
                                            "array value: expects the list, " & _
                                            "followed by its size, followed by " & _
                                            "the array address, at the top of " & _
                                            "the stack")
            Case ENUop.opPushEQ : Return ("opPushEQ(x,x): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)=stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushGE : Return ("opPushGE(x,x): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)>=stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushGT : Return ("opPushGT(x,x): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)>stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushIndirect : Return ("opPushIndirect(i): Pushes " & _
                                        "the contents of a " & _
                                        "memory location indexed " & _
                                        "at " & _
                                        "stack(top)...replacing " & _
                                        "the index: " & _
                                        "push is by value unless the " & _
                                        "memory location contains an array " & _
                                        "or a UDT")
            Case ENUop.opPushLE : Return ("opPushLE(x,x): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<=stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushLiteral : Return ("opPushLiteral[x]: Pushes a " & _
                                        "literal string or " & _
                                        "number ")
            Case ENUop.opPushLT : Return ("opPushLT(x,x): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushNE : Return ("opPushNE(x,x): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) by -1 when " & _
                                        "stack(top-1)<>stack(top), " & _
                                        "0 otherwise ")
            Case ENUop.opPushReturn : Return ("opPushReturn: Pushes " & _
                                        "the subroutine's return " & _
                                        "address ")
            Case ENUop.opRand : Return ("opRand: Seeds the " & _
                                        "random number generator " & _
                                        "to unpredictable values ")
            Case ENUop.opRead : Return ("opRead: Reads from the " & _
                                        "data statements to " & _
                                        "stack(top) ")
            Case ENUop.opRem : Return ("opRem: Equivalent to a " & _
                                        "Nop ")
            Case ENUop.opReplace : Return ("opReplace(s,s,s): Replaces all " & _
                                        "occurences of the " & _
                                        "string at stack(top-1) " & _
                                        "by the string at " & _
                                        "stack(top) in the " & _
                                        "string at stack(top-2). " & _
                                        "Replaces all entries by " & _
                                        "the translated string ")
            Case ENUop.opRnd : Return ("opRnd: Pushes an " & _
                                        "unseeded random number " & _
                                        "on the stack ")
            Case ENUop.opRndSeed : Return ("opRndSeed(n): Pushes a " & _
                                        "seeded random number on " & _
                                        "the stack (seed is " & _
                                        "stack(top), and is " & _
                                        "replaced) ")
            Case ENUop.opRotate : Return ("opRotate n: Exchanges " & _
                                          "stack(top) with " & _
                                          "stack(top-n): this is a NOP " & _
                                          "when the operand n is 0")
            Case ENUop.opRound : Return ("opRound(n,i): Rounds " & _
                                        "stack(top-1) to " & _
                                        "stack(top) digits ")
            Case ENUop.opSgn : Return ("opSgn(n): Replaces " & _
                                        "stack(top) with its " & _
                                        "signum (0 for 0: 1 for " & _
                                        "positive: -1 for " & _
                                        "negative) ")
            Case ENUop.opSin : Return ("opSin(n): Replaces " & _
                                        "stack(top) with its " & _
                                        "sine ")
            Case ENUop.opSqr : Return ("opSqr(n): Replaces " & _
                                        "stack(top) with its " & _
                                        "square root ")
            Case ENUop.opString : Return ("opString(s,i): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with n " & _
                                        "copies of the character " & _
                                        "at stack top, where n " & _
                                        "is at stack(top-1) ")
            Case ENUop.opSubtract : Return ("opSubtract(n,n): Replaces " & _
                                        "stack(top) and " & _
                                        "stack(top-1) with " & _
                                        "stack(top-1)-stack(top) ")
            Case ENUop.opTrace : Return ("opTrace: Changes trace " & _
                                        "settings ")
            Case ENUop.opTracePop : Return ("opTracePop: Restores " & _
                                        "trace settings from a " & _
                                        "LIFO stack ")
            Case ENUop.opTracePush : Return ("opTracePush: Saves " & _
                                        "trace settings in a " & _
                                        "LIFO stack ")
            Case ENUop.opTrim : Return ("opTrim(s): Replaces " & _
                                        "stack(top) with trimmed " & _
                                        "string (leading and " & _
                                        "trailing blanks " & _
                                        "removed) ")
            Case ENUop.opUCase : Return ("opUCase(s): Replaces the " & _
                                        "string at stack(top) " & _
                                        "with its upper case " & _
                                        "translation ")
            Case ENUop.opUtility : Return ("opUtility(u): Replaces " & _
                                                "stack(top-n-1)..stack(top) " & _
                                                "(where n=stack(top)) " & _
                                                "with the result of " & _
                                                "calling the utility " & _
                                                "named in stack(top-1)." & _
                                                vbNewLine)
            Case Else
                _OBJutilities.errorHandler("Unexpected enumerator value", _
                                           ClassName, "opcode2Description", _
                                           "Returning error information")
                Return ("Unexpected enumerator value")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Translates opcode to its stack template
    '
    '
    Public Shared Function opCodeToStackTemplate(ByVal enuOpcode As ENUop) As String
        Dim strDescription As String = opCodeToDescription(enuOpcode)
        Dim intIndex1 As Integer = _OBJutilities.verify(strDescription, "( ", True)
        If intIndex1 = 0 Then Return ("")
        If Mid(strDescription, intIndex1, 1) = " " Then Return ("")
        Return (Mid(strDescription, _
                   intIndex1 + 1, _
                   InStr(intIndex1, strDescription & ")", ")") - intIndex1 - 1))
    End Function

    ' ----------------------------------------------------------------------
    ' Translates opcode to its name
    '
    '
    Public Shared Function opCodeToString(ByVal enuOpcode As ENUop) As String
        Return (_OBJutilities.item(enuOpcode.ToString, 1, "(", False))
    End Function

End Class