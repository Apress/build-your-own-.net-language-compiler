' ***** Common regular expressions ********************************
' *                                                               *
' * This module defines the COMMON_REGULAR_EXPRESSIONS constant.  *
' *                                                               *
' * The COMMON_REGULAR_EXPRESSIONS constant defines the regular   *
' * expressions and their test cases. It consists of a series of  *
' * blocks of four lines each, where a DBCS character is used to  *
' * separate the lines:                                           *
' *                                                               *
' *                                                               *
' *      *  Line 1 contains the re name and value: the re name    *
' *         must be a single word with no blanks.                 *
' *                                                               *
' *      *  Line 2 explains the re                                *
' *                                                               *
' *      *  Line 3 contains a test string for the re              *
' *                                                               *
' *      *  Line 4 contains the expected results from applying    *
' *         the regular expression to the test string, as a       *
' *         series of blank-delimited pairs of numbers.           *
' *                                                               *
' *         + The first number is the expected index of one       *
' *                                                               *
' *         + The second number is the expected length of one     *
' *           captured substring                                  *
' *                                                               *
' *                                                               *
' * The constant may contain substituted keywords, surrounded with*
' * DBCS characters.                                              *
' *                                                               *
' * The use of DBCS characters means that the test strings can    *
' * contain any ASCII characters; however, the                    *
' * COMMON_REGULAR_EXPRESSIONS constant cannot define tests for   *
' * DBCS data that includes these separators.                     *
' *                                                               *
' *                                                               *
' * C H A N G E   R E C O R D ----------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE               *
' * --------   ----------     ----------------------------------- *
' * 06 10 03   Nilges         1.  Added graphicalCharacterSet     *
' *                                                               *
' *****************************************************************

Module commonRegularExpressions
    ' --- DBCS separators
    Friend Const COMMON_REGULAR_EXPRESSIONS_SEP As String = ChrW(300)
    Friend Const COMMON_REGULAR_EXPRESSIONS_LD As String = ChrW(301)
    Friend Const COMMON_REGULAR_EXPRESSIONS_RD As String = ChrW(302)
    ' --- Common regular expressions for VB identifiers
    Private Const COMMON_REGULAR_EXPRESSIONS_IDENTIFIERS As String = _
        "identifierCOM [A-Za-z][A-Za-z0-9_]*" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Simple Visual Basic identifier (release 6 and before)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "identifier1 identifier2_ IDENTIFIER3_ 2notIdentifier" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "0 11 12 12 25 12" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "identifierNET [A-Za-z_][A-Za-z0-9_]*" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Simple Visual Basic identifier (.Net)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "_identifier identifier2_ IDENTIFIER3_ 2notIdentifier" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "0 11 12 12 25 12" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "compoundIdentifierCOM ([A-Za-z][A-Za-z0-9_]*)(\.([A-Za-z][A-Za-z0-9_]*))*" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Simple or compound Visual Basic identifier (release 6 and before)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        " identifier.identifier2_ IDENTIFIER3_ 2Identifier" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "1 23 25 12 39 10" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "compoundIdentifierNET ([A-Za-z_][A-Za-z0-9_]*)(\.([A-Za-z_][A-Za-z0-9_]*))*" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Simple or compound Visual Basic identifier (.Net)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        " __identifier.identifier2_ IDENTIFIER3_ 2notIdentifier" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "1 25 27 12 41 13" 
    ' --- Common regular expressions for character sets
    Private Const COMMON_REGULAR_EXPRESSIONS_CHARSETS As String = _
        "graphicalCharacterSet " & _
        "[ A-Za-z0-9\~\`\!\@\#\$\%\^\&\*\(\)_\-\+\=\{\[\}\]\|\\\:\;\""\'\<\,\>\.\?\/]" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Graphical character set (characters on most PC keyboards in the USA)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        ChrW(0) & "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
        ChrW(0) & "abcdefghijklmnopqrstuvwxyz" & _
        "0123456789" & Chr(19) & "~`!@#$%^&*()_-+={[}]|\:;""'<,>.?/ " & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "1 1 2 1 3 1 4 1 5 1 6 1 7 1 8 1 9 1 10 1 11 1 12 1 13 1 " & _
        "14 1 15 1 16 1 17 1 18 1 19 1 20 1 21 1 22 1 23 1 24 1 25 1 26 1 " & _
        "28 1 29 1 30 1 31 1 32 1 33 1 34 1 35 1 36 1 37 1 38 1 39 1 40 1 " & _
        "41 1 42 1 43 1 44 1 45 1 46 1 47 1 48 1 49 1 50 1 51 1 52 1 53 1 " & _
        "54 1 55 1 56 1 57 1 58 1 59 1 60 1 61 1 62 1 63 1 " & _
        "65 1 66 1 67 1 68 1 69 1 70 1 71 1 72 1 73 1 74 1 75 1 76 1 77 1 " & _
        "78 1 79 1 80 1 81 1 82 1 83 1 84 1 85 1 86 1 87 1 88 1 89 1 90 1 " & _
        "91 1 92 1 93 1 94 1 95 1 96 1 97 1"
    ' --- Common regular expressions for line breaks
    Private Const COMMON_REGULAR_EXPRESSIONS_NEWLINES As String = _
        "newlineWebWindows (\x0D\x0A)|\x0A" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Newline for Web (linefeed) or for Windows (carriage return and linefeed)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "     " & vbNewline & "    " & vbCrLf & " " & vbLf & ChrW(10) & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "5 2 11 2 14 1 15 1" 
    ' --- Common regular expressions for procedure headers
    Private Const COMMON_REGULAR_EXPRESSIONS_PROCEDUREHEADER As String = _
        "vbProcedureHeaderCOM " & _
        "((\x0D\x0A)|\x0A)[ ]*((Public )|(Private )|(Friend )){0,1}((Sub )|(Function )|(Property (Get|Set|Let) ))([A-Za-z_][A-Za-z0-9_]*)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic procedure header (release 6 and before: can't appear at start of file, includes preceding new line and text through procedure name: standard case and spacing: can't be broken between lines)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        COMMON_REGULAR_EXPRESSIONS_LD & "vbProcedureHeaderCOMtests" & COMMON_REGULAR_EXPRESSIONS_RD & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "vbProcedureHeaderNET " & _
        "((\x0D\x0A)|\x0A)[ ]*((Public )|(Private )|(Friend )){0,1}([A-Za-z_][A-Za-z0-9_]* )*((Sub )|(Function )|(Property ))([A-Za-z_][A-Za-z0-9_]*)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic procedure header (.Net: scope keyword is required: can't appear at start of file, includes preceding new line and text through procedure name: standard case and spacing: the header cannot be broken between lines)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        COMMON_REGULAR_EXPRESSIONS_LD & "vbProcedureHeaderNettests" & COMMON_REGULAR_EXPRESSIONS_RD & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "vbParameterDefCOM (((ByVal )|(ByRef )){0,1}([A-Za-z][A-Za-z0-9_]*)(\([,]*\)){0,1}([ ]+As ([A-Za-z][A-Za-z0-9_]*)){0,1})" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic parameter definition (VB-6)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "ByVal parm1 As Integer, ByRef parm2, parm3(,,) As Double" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "0 22 24 11 37 19" & _ 
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "vbParameterDefNet (((ByVal )|(ByRef )){0,1}([A-Za-z_][A-Za-z0-9_]*)(\([,]*\)){0,1}([ ]+As ([A-Za-z][A-Za-z0-9_]*)){0,1})" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic parameter definition (VB-6)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "ByVal parm1 As Integer, ByRef parm2, _parm3(,,) As Double" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "0 22 24 11 37 20"  
    ' --- Common regular expressions for VB comments
    Private Const COMMON_REGULAR_EXPRESSIONS_COMMENTS As String = _
        "vbComment '([^\x00-\x19])*" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic comment excluding trailing newline" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "    ' This is a comment " & vbNewline & "    " & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "4 20" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "vbCommentLine [ ]*'([^\x00-\x19])*((\x0D\x0A)|\x0A)" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic comment including trailing newline" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "    ' This is a comment " & vbLF & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "0 25" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "vbCommentBlock ([ ]*'([^\x00-\x19])*((\x0D\x0A)|\x0A))*([ ]*'([^\x00-\x19])*){0,1}" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "Visual Basic comment block" & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "    ' This is a comment " & vbCRLF & _
        "    ' This is a comment " & vbCRLF & _
        "    ' This is a comment " & _
        COMMON_REGULAR_EXPRESSIONS_SEP & _
        "0 76" 
    ' --- Common regular expressions 
    Friend Const COMMON_REGULAR_EXPRESSIONS As String = _
        COMMON_REGULAR_EXPRESSIONS_CHARSETS & _ 
        COMMON_REGULAR_EXPRESSIONS_SEP & _ 
        COMMON_REGULAR_EXPRESSIONS_COMMENTS & _ 
        COMMON_REGULAR_EXPRESSIONS_SEP & _ 
        COMMON_REGULAR_EXPRESSIONS_IDENTIFIERS & _ 
        COMMON_REGULAR_EXPRESSIONS_SEP & _ 
        COMMON_REGULAR_EXPRESSIONS_NEWLINES & _ 
        COMMON_REGULAR_EXPRESSIONS_SEP & _ 
        COMMON_REGULAR_EXPRESSIONS_PROCEDUREHEADER 
        
End Module
