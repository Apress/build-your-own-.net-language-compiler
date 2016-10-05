VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' **************************************************************************
' *                                                                        *
' * Shared utilities                                                       *
' *                                                                        *
' *                                                                        *
' *      addUniqueListItem: add item to list, suppressing duplicates       *
' *      append: append to a string                                        *
' *      align: align and fill a string                                    *
' *      attachPath: attaches fileid or path to a path                     *
' *      baseN2Long: convert base N number to Long                         *
' *      char2: translate DBCS/Unicode characters to their name            *
' *      character2Name: translate common characters to their name         *
' *      collection2XML: convert the collection to XML                     *
' *      collectionClear: clear a Collection                               *
' *      collectionDestroy: clean up a Collection                          *
' *      datatype: test string for data type                               *
' *      decode: decodes strings encoded by encode                         *
' *      driveServerChange: change drive or server                         *
' *      ellipsis: return truncated string, adding ... if necessary        *
' *      encode: performs a simple shift encode on a string                *
' *      enquote: return string, adding quotes if necessary                *
' *      errorHandler: error handling                                      *
' *      errorHandlerForm: returns the error handler form                  *
' *      char2Name: convert character to its name                          *
' *      deco2Var: convert decorated string back to Variant value          *
' *      deleteFile: delete a file                                         *
' *      enquote: return string, adding quotes if necessary                *
' *      errorHandler: error handling                                      *
' *      fileExists: tell caller whether file exists                       *
' *      findword: find the blank-delimited word                           *
' *      hex2String: converts 2 or five byte ascii back to string          *
' *      histogram: maps a range of values onto another range              *
' *      isLongInteger: returns True when input string is a valid long int *
' *      isQuoted: return True when string has quotes                      *
' *      item: return delimited item                                       *
' *      items: return count of delimited items                            *
' *      long2BaseN: convert Long integer to arbitrary base                *
' *      keywordChange: replace keyword values                             *
' *      max: return double-precision maximum                              *
' *      min: return double-precision mininum                              *
' *      mkXMLcomment: create an XML comment                               *
' *      mkXMLelement: create an XML element                               *
' *      mkXMLtag: create an XML tag                                       *
' *      mkXMLtagWithAttributes: create an XML tag with attributes         *
' *      name2Char: convert character name back to character               *
' *      nextPrime: return next prime number                               *
' *      parse: parse around delimiter                                     *
' *      hungarian2Prefix: get the Hungarian prefix                        *
' *      hungarian2Suffix: get the Hungarian suffix                        *
' *      hungarianParse: parse the Hungarian identifier                    *
' *      parse: split extender                                             *
' *      range2String: create a single string from a start/end range       *
' *      removeParentheses: remove parentheses                             *
' *      searchListBox: search list box                                    *
' *      signum: return 0 for 0, 1 for positive, -1 for negative number    *
' *      string2Box: wrap a string in a nerd view box                      *
' *      string2File: writes the string to a file                          *
' *      string2HardParagraph: create a fixed-width paragraph from string  *
' *      tempFileid: return a temporary file identifier                    *
' *      test: test the utility library (not implemented)                  *
' *      var2Deco: convert a Variant to a displayable (decorated) form     *
' *      varValue2VarType: convert a Variant to a specified type           *
' *      verify: scan a string for character sets                          *
' *      word: return blank-delimited word                                 *
' *      words: return count of blank-delimited words                      *
' *      xmlTag2endTag: convert the XML tag to its corresponding endtag    *
' *                                                                        *
' *                                                                        *
' * C H A N G E   R E C O R D -------------------------------------------- *
' *   DATE       PROGRAMMER   DESCRIPTION OF CHANGE                        *
' * --------     ----------   -------------------------------------------- *
' * 04 18 03     Nilges       Corrected "check" to click in error handler  *
' *                           form                                         *
' *                                                                        *
' * 04 19 03     Nilges       Added range2String                           *
' *                                                                        *
' * 04 20 03     Nilges       Changed errorHandler                         *
' *                                                                        *
' * 04 23 03     Nilges       Added collectionDestroy                      *
' *                                                                        *
' * 04 24 03     Nilges       Changed string2HardParagraph                 *
' *                                                                        *
' * 04 26 03     Nilges       Changed string2Box                           *
' *                                                                        *
' * 05 20 03     Nilges       1.  Added encode and decode                  *
' *                           2.  Added long2BaseN and baseN2Long          *
' *                           3.  Added string2Hex and hex2String          *
' *                           4.  Added parse                              *
' *                                                                        *
' * 05 31 03     Nilges       1.  Added datatype                           *
' *                           2.  Added isLongInteger                      *
' *                           3.  Added findword/finditem                  *
' *                                                                        *
' * 06 01 03     Nilges       1.  Added test (as unimplemented stub)       *
' *                                                                        *
' * 01 25 04     Nilges       1.  Added mkXMLelement, mkXMLcomment and     *
' *                               mkXMLtag.                                *
' *                           2.  Added name2Char and char2Name            *
' *                                                                        *
' * 01 26 04     Nilges       1.  Added collection2XML                     *
' *                           2.  Added xmlTag2EndTag                      *
' *                           3.  Changed mkXMLelement                     *
' *                                                                        *
' * 01 28 04     Nilges       1.  Changed xmlTag2EndTag                    *
' *                           2.  Changed mkXMLtagWithAttributes           *
' *                                                                        *
' **************************************************************************

' ***** Alignment codes ******
Public Enum ENUalign
    alignLeft = 1
    alignCenter = 2
    alignRight = 3
End Enum

' ***** Errors *****
Private Const ERR_BADTYPE_NUMBER = 11
Private Const ERR_BADTYPE = "Invalid parameter type"
Private Const ERR_ALIGN_LEN_NUMBER = 1
Private Const ERR_ALIGN_LEN = "Invalid alignment length &L"
Private Const ERR_ALIGN_FILL_NUMBER = 2
Private Const ERR_ALIGN_FILL = "Fill is not one character"
Private Const ERR_ALIGN_CODE_NUMBER = 3
Private Const ERR_ALIGN_CODE = "Align code is not valid"
Private Const ERR_APPEND_MAXLEN_NUMBER = 4
Private Const ERR_APPEND_MAXLEN = "Invalid maximum length &L"
Private Const ERR_APPEND_LEN_NUMBER = 5
Private Const ERR_APPEND_LEN = "Updated string length (&L) would exceed the maximum length (&M)"
Private Const ERR_TEMPFILEID_COL_NUMBER = 6
Private Const ERR_TEMPFILEID_COL = "Error in creating or extending a collection: &N &D"
Private Const ERR_VERIFY_STARTINDEX_NUMBER = 7
Private Const ERR_VERIFY_STARTINDEX_COL = "Start index &I is not valid  "
Private Const ERR_STRING2BOX_INVALID_BUILD_CHAR_NUMBER = 8
Private Const ERR_STRING2BOX_INVALID_BUILD_CHAR = "Build character is not a single character"
Private Const ERR_COLLECTIONCREATE_NUMBER = 9
Private Const ERR_COLLECTIONCREATE = "Could not create &C collection: &N &D"
Private Const ERR_CHARACTER2NAME_NUMBER = 10
Private Const ERR_CHARACTER2NAME = "strCharacter argument ""&C"" is not a single character"
Private Const ERR_COLLECTIONDESTROY_NUMBER = 11
Private Const ERR_COLLECTIONDESTROY = "Error in collection destruction: &D"


Public Function addUniqueListItem(ByRef objListBox As Object, _
                                  ByVal strNew As String) As Long
    '
    ' Add item to list box unless it is present
    '
    '
    ' This function returns the new entry index or the index of an existing entry.
    ' -1 is returned on an error.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 08 28 02   Nilges         Handle list-type controls other than
    '                           list boxes
    '
    '
    Dim booSorted As Boolean
    Dim lngIndex1 As Long
    Dim ctlHandle As Control
    addUniqueListItem = -1
    On Error GoTo addUniqueListItem_Lbl1_objectErrorHandler
        Set ctlHandle = objListBox
    On Error GoTo 0
    lngIndex1 = searchListBox(ctlHandle, strNew, booCaseSensitive:=False)
    If lngIndex1 <> -1 Then
        addUniqueListItem = lngIndex1: Exit Function
    End If
    On Error GoTo addUniqueListItem_Lbl1_objectErrorHandler
        ctlHandle.AddItem strNew
    On Error Resume Next
        booSorted = ctlHandle.Sorted
    On Error GoTo 0
    If booSorted Then
        addUniqueListItem = searchListBox(ctlHandle, strNew)
    Else
        On Error GoTo addUniqueListItem_Lbl1_objectErrorHandler
            addUniqueListItem = ctlHandle.ListCount - 1
        On Error GoTo 0
    End If
    Exit Function
addUniqueListItem_Lbl1_objectErrorHandler:
    errorHandler ERR_BADTYPE_NUMBER, _
                 ERR_BADTYPE, _
                 "addUniqueListItem", "clsUtilities"
End Function
Public Function align(ByVal strInstring As String, _
                      ByVal intAlignLength As Integer, _
                      Optional ByVal enuAlignment As ENUalign = alignLeft, _
                      Optional ByVal strFill As String = " ") As String
    '
    ' Align and fill a string
    '
    '
    Dim intQuotient As Integer
    Dim intSizeDelta As Integer
    Dim strAligned As String
    Dim strInstringTruncate As String
    Dim strPad As String
    align = ""
    If intAlignLength < 0 Then
        errorHandler ERR_ALIGN_LEN_NUMBER, ERR_ALIGN_LEN, "align", "clsUtilities", intAlignLength
    End If
    If Len(strFill) <> 1 Then
        errorHandler ERR_ALIGN_FILL_NUMBER, ERR_ALIGN_FILL, "align", "clsUtilities"
    End If
    intSizeDelta = max(0, intAlignLength - Len(strInstring))
    strInstringTruncate = Mid(strInstring, 1, intAlignLength)
    If enuAlignment = ENUalign.alignCenter Then
        intQuotient = intSizeDelta \ 2
        strAligned = String$(intQuotient, strFill) & _
                        strInstringTruncate & _
                        String$(intQuotient + intSizeDelta Mod 2, strFill)
    Else
        strPad = String$(intSizeDelta, strFill)
        Select Case enuAlignment
            Case ENUalign.alignLeft:
                strAligned = strInstringTruncate & strPad
            Case ENUalign.alignRight:
                strAligned = strPad & strInstringTruncate
            Case Else:
                errorHandler ERR_ALIGN_CODE_NUMBER, ERR_ALIGN_CODE, "align", "clsUtilities"
        End Select
    End If
    align = strAligned
End Function
Public Function append(ByRef strString As String, _
                       ByVal strSeparator As String, _
                       ByVal strAppend As String, _
                       Optional ByVal booAppendInPlace As Boolean = False, _
                       Optional ByVal booToStart As Boolean = False, _
                       Optional ByVal lngMaxLength As Long = -1) As Variant
        '
    ' Append string
    '
    '
    ' This method attaches strAppend to a string.  If the string is not
    ' null on entry to this code then strSeparator is placed between
    ' the string and strAppend.
    '
    ' If the booAppendInPlace optional parameter is present and True, strString is
    ' modified in place, and this function returns True on success or False
    ' on failure.  If booAppendInPlace is missing or False, strString is modified
    ' and returned as the value of this function.
    '
    ' If the optional booToStart parameter is present and True then the append
    ' string and separator are placed at the start of the main input string.
    '
    ' If the optional lngMaxLength parameter is present it may specify
    ' the maximum string length.
    '
    '
    Dim lngNewLength As Long
    Dim strSeparatorWork As String
    Dim varErrorReturn As Variant
    varErrorReturn = IIf(booAppendInPlace, False, strString)
    If lngMaxLength < -1 Then
        errorHandler ERR_APPEND_MAXLEN_NUMBER, ERR_APPEND_MAXLEN, "append", "clsUtilities", "L", lngMaxLength
        append = varErrorReturn
        Exit Function
    End If
    strSeparatorWork = ""
    If Len(strString) <> 0 Then strSeparatorWork = strSeparator
    If lngMaxLength <> -1 Then
        lngNewLength = Len(strString) + Len(strAppend) + Len(strSeparatorWork)
        If lngNewLength > lngMaxLength Then
            errorHandler ERR_APPEND_LEN_NUMBER, _
                         ERR_APPEND_LEN, _
                         "append", "clsUtilities", _
                         "L", lngNewLength, _
                         "M", lngMaxLength
            append = varErrorReturn
            Exit Function
        End If
    End If
    If booToStart Then
        If booAppendInPlace Then
            strString = strAppend & strSeparatorWork & strString
        Else
            append = strAppend & strSeparatorWork & strString
        End If
    Else
        If booAppendInPlace Then
            strString = strString & strSeparatorWork & strAppend
        Else
            append = strString & strSeparatorWork & strAppend
        End If
    End If
    If booAppendInPlace Then append = True
End Function
Public Function attachPath(ByVal strPath As String, _
                           ByVal strFileTitle As String) As String
    '
    ' Attach path
    '
    '
    ' This function takes care of a situation that arises when a path
    ' is just a drive (for example c:\) as opposed to a path that
    ' includes directories (as in c:\barf\kamunkle.)  If the path ends in
    ' a backslash, this function returns the mere concatenation of strPath
    ' and strFileTitle: if the path does not end in a back slash then one
    ' is added.
    '
    '
    Dim intLength As Integer
    intLength = Len(strPath)
    If intLength = 0 Then
        attachPath = strFileTitle
        Exit Function
    End If
    attachPath = strPath & _
                 IIf(Mid$(strPath, intLength) = "\", "", "\") & _
                 strFileTitle
End Function
Public Function baseN2Long(ByVal strBaseN As String, _
                           ByVal strDigits As String, _
                           Optional ByRef varSuccess As Variant, _
                           Optional ByVal bytWordSize As Byte) As Long
    '
    ' Convert base N number to implicit base 10
    '
    '
    ' This function converts a base N integer (where the base and the symbols the
    ' base uses are in strDigits) to its long value.
    '
    ' For example, if strBaseN is "1000" and strDigits is "01", then this function
    ' converts the binary value 1000 to 8.
    '
    ' If the optional reference parameter varSuccess is present, it is set to
    ' True when strBaseN is a valid base n number, that can be converted to a
    ' Long integer: otherwise varSuccess is set to False.  If varSuccess is
    ' missing this function will raise an errorHandler error when an error occurs.
    '
    ' By default, all numbers passed to this routine are expected to be unsigned
    ' integers, but signed integers can be passed when the optional bytWordSize
    ' parameter is present and nonzero, and there are exactly two characters in
    ' strDigits: these integers should represent signed quantities in twos-complement
    ' notation.  For details on twos-complement, see the baseTen2N function.
    '
    ' By default, this method will raise errors on parameter errors: but when the
    ' optional parameter varSuccess is present, it will be set to True on success or
    ' False on an error.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 12 30 98   Nilges         Version 1
    ' 04 28 00   Nilges         Convert to a Long
    ' 05 23 00   Nilges         Added varSuccess parameter
    '                           Handle twos complement
    ' 05 20 03   Nilges         Converted to the streamlined utilities
    '                           library
    '
    '
    Dim bytAdjustment As Byte
    Dim intBaseValue As Integer
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intSignum As Integer
    Dim lngBase10 As Long
    Dim lngPower As Long
    Dim strBaseNWork As String
    baseN2Long = 0
    If Not IsMissing(varSuccess) Then varSuccess = True
    intBaseValue = Len(strDigits)
    If intBaseValue < 2 Then
        If IsMissing(varSuccess) Then
            errorHandler 0, _
                         "Only zero or one digits were supplied for the output number", _
                         "baseN2Long", "clsUtilities"
        Else
            varSuccess = False
        End If
        Exit Function
    End If
    intSignum = 1
    strBaseNWork = strBaseN
    bytAdjustment = 0
    If intBaseValue = 2 And bytWordSize <> 0 Then
        If Mid$(strBaseN, 1, 1) = Mid$(strDigits, 2, 1) Then
            intSignum = -1: bytAdjustment = 1
        End If
        strBaseNWork = translate(Mid$(strBaseN, 2), strDigits, StrReverse(strDigits))
    End If
    lngPower = 1
    For intIndex1 = Len(strBaseNWork) To 1 Step -1
        intIndex2 = InStr(strDigits, Mid$(strBaseNWork, intIndex1, 1))
        If intIndex2 < 1 Then
            If IsMissing(varSuccess) Then
                errorHandler 0, _
                             "Invalid digits in base " & intBaseValue & " number " & strBaseN, _
                             "baseN2Long", "clsUtilities"
            Else
                varSuccess = False
            End If
            Exit Function
        End If
        On Error GoTo baseN2Long_Lbl1_overflowErrorHandler
            lngBase10 = lngBase10 + (intIndex2 - 1) * lngPower
        On Error GoTo 0
        lngPower = lngPower * intBaseValue
    Next intIndex1
    baseN2Long = lngBase10 * intSignum - bytAdjustment
    Exit Function
baseN2Long_Lbl1_overflowErrorHandler:
    errorHandler 0, "Overflow", "baseN2Long", "clsUtilities"
End Function

    ' ---------------------------------------------------------------------
    ' Convert character to its name
    '
    '
    ' This method converts Unicode, DBCS and ASCII characters to their
    ' names.
    '
    ' Digits and alphabetic characters are converted to their actual value:
    ' special characters available on common USA keyboards are converted to
    ' their names.
    '
    ' The NUL character is converted to Nul: Tab is converted to Tab:
    ' carriage return is converted to carriageReturn: line feed is
    ' converted to lineFeed: and blank is converted to blank.
    '
    ' All other characters convert to Chrnnnnn where nnnnn is the value of
    ' the character.
    '
    ' Note: see also character2Name.
    '
    '
    Public Function char2Name(ByVal strChar As String) As String
        If Len(strChar) <> 1 Then
            errorHandler 0, _
                         "strChar is not a single character", _
                         "char2Name", _
                         "utilities"
            Exit Function
        End If
        If strChar >= "0" And strChar <= "9" Then
            char2Name = strChar: Exit Function
        End If
        If strChar >= "A" And strChar <= "Z" Then
            char2Name = strChar: Exit Function
        End If
        If strChar >= "a" And strChar <= "z" Then
            char2Name = strChar: Exit Function
        End If
        Select Case strChar
            Case "~": char2Name = "tilde"
            Case "`": char2Name = "acuteAccent"
            Case "!": char2Name = "exclamationPoint"
            Case "@": char2Name = "atSign"
            Case "#": char2Name = "poundSign"
            Case "$": char2Name = "dollarSign"
            Case "%": char2Name = "percent"
            Case "^": char2Name = "carat"
            Case "&": char2Name = "ampersand"
            Case "*": char2Name = "asterisk"
            Case "(": char2Name = "leftParenthesis"
            Case ")": char2Name = "rightParenthesis"
            Case "_": char2Name = "underscore"
            Case "-": char2Name = "dash"
            Case "+": char2Name = "plus"
            Case "=": char2Name = "equals"
            Case "{": char2Name = "leftBrace"
            Case "[": char2Name = "leftBracket"
            Case "}": char2Name = "rightBrace"
            Case "]": char2Name = "rightBracket"
            Case "|": char2Name = "verticalStroke"
            Case "\": char2Name = "backslash"
            Case ":": char2Name = "colon"
            Case ";": char2Name = "semicolon"
            Case "~": char2Name = "tilde"
            Case """": char2Name = "doubleQuote"
            Case "'": char2Name = "singleQuote"
            Case "<": char2Name = "lessThan"
            Case ",": char2Name = "comma"
            Case ">": char2Name = "greaterThan"
            Case ".": char2Name = "period"
            Case "?": char2Name = "questionMark"
            Case "/": char2Name = "slash"
            Case Chr(0): char2Name = "NUL"
            Case vbTab: char2Name = "Tab"
            Case Chr(10): char2Name = "LineFeed"
            Case Chr(13): char2Name = "carriageReturn"
            Case " ": char2Name = "Blank"
            Case Else
                char2Name = "Chr" & align(CStr(AscW(strChar)), 5, strFill:="0", enuAlignment:=alignRight)
        End Select
    End Function
Public Function character2Name(ByVal strCharacter As String) As String
    '
    ' Translate common characters to their name
    '
    '
    ' This method returns, for special characters available on standard PC keyboards,
    ' the name of the character.  For alphabetic characters and for digits, the
    ' character's original value is returned.  CONTROL is returned for anything else.
    '
    ' Note: see also char2Name for a different version of this functionality that is
    ' matched with its inverse function name2Char.
    '
    '
    Static colSpecial As Collection     ' Key: character: data: name
    Dim lngCharacter As Long
    character2Name = ""
    If (colSpecial Is Nothing) Then
        On Error GoTo character2Name__Lbl1_collectionCreateErrorHandler
            Set colSpecial = New Collection
            With colSpecial
                .Add "tilde", "~"
                .Add "acute accent", "`"
                .Add "exclamation point", "!"
                .Add "at sign", "@"
                .Add "pound sign", "#"
                .Add "dollar sign", "$"
                .Add "percent", "%"
                .Add "carat", "^"
                .Add "ampersand", "&"
                .Add "asterisk", "*"
                .Add "left parenthesis", "("
                .Add "right parenthesis", ")"
                .Add "underscore", "_"
                .Add "dash", "-"
                .Add "plus sign", "+"
                .Add "equals sign", "="
                .Add "left brace", "{"
                .Add "right brace", "}"
                .Add "left bracket", "["
                .Add "right bracket", "]"
                .Add "colon", ":"
                .Add "semicolon", ";"
                .Add "quote", """"
                .Add "single quote", "'"
                .Add "less than", "<"
                .Add "greater than", ">"
                .Add "comma", ","
                .Add "period", "."
                .Add "question mark", "?"
                .Add "forward slash", "/"
                .Add "backslash", "\"
                .Add "vertical bar", "|"
            End With
        On Error GoTo 0
    End If
    If Len(strCharacter) <> 1 Then
        errorHandler ERR_CHARACTER2NAME_NUMBER, _
                     ERR_CHARACTER2NAME, _
                     "character2Name", _
                     "clsUtilities", _
                     "&S", strCharacter
        Exit Function
    End If
    lngCharacter = AscW(strCharacter)
    If lngCharacter >= AscW("0") And lngCharacter <= AscW("9") Then
        character2Name = strCharacter
    Else
        character2Name = "CONTROL"
        On Error Resume Next
            character2Name = colSpecial(strCharacter)
        On Error GoTo 0
    End If
    Exit Function
character2Name__Lbl1_collectionCreateErrorHandler:
    errorHandler ERR_COLLECTIONCREATE_NUMBER, _
                 ERR_COLLECTIONCREATE, _
                 "character2Name", _
                 "clsUtilities", _
                 "&C", "special character", _
                 "&N", Err.Number, _
                 "&D", Err.Description
End Function
Public Function collection2XML(ByVal colCollection As Collection, _
                               ByVal strTagName As String, _
                               ByVal booFormat As Boolean) _
       As String
    '
    ' Convert collection to XML
    '
    '
    ' This method converts  a collection to an eXtended Markup Language
    ' (XML) tag. The strTagName parameter supplies the main tag name.                                         *
    '
    ' Each item is converted to a subtag with the name Itemnn where nn is
    ' the item number; the item number is expanded to the size of the largest
    ' item number and zero-filled.
    '
    ' Each item will be in the format <Itemnn>value</Itemnn> where value
    ' is the string form of the item variant as converted by var2deco.
    '
    ' Each subcollection in the main collection is converted to a subtag in
    ' the same format as the main tag but with the name subcollectionnn, \
    ' where nn is a sequence number.
    '
    ' If the collection is Nothing, then the tag <name>Nothing</name>
    ' is produced.
    '
    ' Suppose a collection contains an integer, a string, and a sub-
    ' collection containing one integer. collection2XML will produce
    ' the following tag.
    '
    '              <anonymousCollection>
    '                   <item1>System.Int32(32767)</item1>
    '                   <item2>System.String("Example")</item2>
    '                   <item3>
    '                       <subcollection1>
    '                           <item1>System.Int32(14)</item1>
    '                       </subcollection1>
    '                   </item3>
    '              </anonymousCollection>
    '
    '
    ' C H A N G E   R E C O R D -------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -------------------------------------------
    ' 01 26 04   Nilges         Converted from .Net version
    '
    '
    collection2XML = collection2XML_(colCollection, strTagName, booFormat, 0)
End Function
Private Function collection2XML_(ByRef colCollection As Collection, _
                                 ByVal strTagName As String, _
                                 ByVal booFormat As Boolean, _
                                 ByVal intLevel As Integer) As String
    '
    ' Collection to XML: recursion step
    '
    '
    Dim colNext As Collection
    Dim intIndex1 As Integer
    Dim intItemSeq As Integer
    Dim intLength As Integer
    Dim intSubcollectionSeq As Integer
    Dim strIndent As String
    Dim strItemTagName As String
    Dim strResult As String
    Dim strNewline As String
    Dim strNext As String
    Dim strOutstring As String
    strOutstring = mkXMLTag(strTagName)
    If (colCollection Is Nothing) Then
        collection2XML_ = strOutstring & "Nothing" & mkXMLTag(strTagName, True)
        Exit Function
    End If
    intItemSeq = 1
    intSubcollectionSeq = 1
    strIndent = copies(" ", (intLevel + 1) * 4)
    strNewline = CStr(IIf(booFormat, vbNewLine, ""))
    With colCollection
        If .Count <> 0 Then
            strResult = Log(.Count) / Log(10) + 1
            intLength = Mid$(strResult, 1, InStr(strResult & ".", ".") - 1)
        End If
        For intIndex1 = 1 To .Count
            strItemTagName = "Item" & align(CStr(intItemSeq), _
                                            intLength, _
                                            strFill:="0", _
                                            enuAlignment:=alignRight)
            If (TypeOf .item(intIndex1) Is Collection) Then
                Set colNext = .item(intIndex1)
                strNext = strNewline & _
                            collection2XML_(colNext, _
                                            "subcollection" & intSubcollectionSeq, _
                                            booFormat, _
                                            intLevel + 1) & _
                            strNewline & strIndent
                intSubcollectionSeq = intSubcollectionSeq + 1
            Else
                strNext = var2Deco(.item(intIndex1))
            End If
            intItemSeq = intItemSeq + 1
            If booFormat Then
                strOutstring = strOutstring & strNewline & strIndent
            End If
            strOutstring = strOutstring & mkXMLElement(strItemTagName, strNext, 0)
        Next intIndex1
        strOutstring = strOutstring & strNewline & mkXMLTag(strTagName, booEndTag:=True)
        If booFormat And intLevel > 0 Then
            strIndent = copies(" ", intLevel * 4)
            strOutstring = strIndent & _
                           Replace(strOutstring, strNewline, strNewline & strIndent)
        End If
        collection2XML_ = strOutstring
    End With
End Function
Public Function collectionClear(ByRef colCollection As Collection) As Boolean
    '
    ' Clear the collection
    '
    '
    Dim lngIndex1 As Long
    With colCollection
        For lngIndex1 = .Count To 1 Step -1
            colCollection.Remove (lngIndex1)
        Next lngIndex1
    End With
End Function
Public Function collectionDestroy(ByRef colCollection As Collection, _
                                  Optional ByVal booErrIfNothing As Boolean = True) _
       As Boolean
    '
    ' Destroy the collection
    '
    '
    ' This method examines colCollection recursively for all contained
    ' collections. It removes each item from the main collection and all
    ' contained collections. Each item that is a collection or an Object
    ' is set to Nothing, including colCollection itself.
    '
    ' By default this method will raise or display an error when the
    ' collection is Nothing. To suppress this and return True for a
    ' nonexistent collection, use booErrIfNothing:=False.
    '
    ' Using this method ensures that storage isn't cluttered with dead
    ' objects.
    '
    '
    ' C H A N G E   R E C O R D ----------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ----------------------------------------
    ' 04 23 03   Nilges         Version 1
    '
    '
    Dim colInnerHandle As Collection
    Dim lngIndex1 As Long
    Dim objItem As Object
    collectionDestroy = False
    If (colCollection Is Nothing) Then
        If booErrIfNothing Then
            errorHandler ERR_COLLECTIONDESTROY_NUMBER, _
                         ERR_COLLECTIONDESTROY, _
                         "collectionDestroy", _
                         "clsUtilities", _
                         "&D", "collection doesn't exist"
            Exit Function
        End If
        collectionDestroy = True
    End If
    On Error GoTo collectionDestroy_Lbl1_errorHandler
        With colCollection
                For lngIndex1 = .Count To 1 Step -1
                    If IsObject(.item(lngIndex1)) Then
                        Set objItem = .item(lngIndex1)
                        If (TypeOf objItem Is Collection) Then
                            Set colInnerHandle = objItem
                            If Not collectionDestroy(colInnerHandle) Then Exit Function
                        End If
                        Set objItem = Nothing
                    End If
                    .Remove lngIndex1
                Next lngIndex1
            On Error GoTo 0
        End With
        Set colCollection = Nothing
    On Error GoTo 0
    collectionDestroy = True
    Exit Function
collectionDestroy_Lbl1_errorHandler:
    errorHandler ERR_COLLECTIONDESTROY_NUMBER, _
                 ERR_COLLECTIONDESTROY, _
                 "collectionDestroy", _
                 "clsUtilities", _
                 "&D", Err.Number & " " & Err.Description
End Function
Public Function copies(ByVal strInstring As String, ByVal lngCopies As Long) As String
    '
    ' Make zero, one or more copies of a string
    '
    '
    ' This function returns a null string when its lngCopies parameter is 0: it returns strInstring when
    ' lngCopies is 1.  If lngCopies is greater than one, this function returns strInstring concatenated
    ' to itself n times.
    '
    ' Note the algorithm used.  We return a null string when lngCopies is zero.  Otherwise we set
    ' the output string to the input string, and then double this string by concatenating it to itself
    ' until we have the highest power of two that is less than lngCopies.  Then we call ourselves
    ' recursively.  Note that this applies the speed advantage of doubling strings (as opposed to
    ' mindlessly concatenating strings one by one) in a recursive fashion.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 08 13 99   Nilges         Version 1
    '
    ' 02 16 01   Nilges         Efficiency improvement: double the
    '                           size of the string until close to
    '                           target copy count and then call yourself
    '                           recursively.  For 10,000 copies,
    '                           copies took 42 seconds before this
    '                           change.  It took close to 0 seconds
    '                           after this modification.
    '
    ' 02 21 01   Nilges         Algorithm improvements
    '
    ' 05 21 01   Nilges         Transferred to streamlined utilities
    '                           library
    '
    '
    Dim lngIndex1 As Long
    Dim strOutstring As String
    copies = ""
    If lngCopies < 0 Then
        errorHandler 0, _
                     "Copies parameter " & _
                     "(" & lngCopies & ") " & _
                     "is invalid", _
                     "copies", "clsUtilities"
        Exit Function
    End If
    If lngCopies > 0 Then
        strOutstring = strInstring
        For lngIndex1 = 1 To Int(Log(lngCopies) / Log(2))
            strOutstring = strOutstring & strOutstring
        Next lngIndex1
        strOutstring = strOutstring & copies(strInstring, lngCopies - 2 ^ (lngIndex1 - 1))
    End If
    copies = strOutstring
End Function
Public Function datatype(ByVal strInstring As String, _
                         Optional ByVal strExpectedDatatype As String) As Variant
    '
    ' Test string for data type
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 05 31 03   Nilges         1.  Converted to streamlined utilities
    '                               library
    '                           2.  Removed FILEID and LIST
    '
    '
    Static bytAlphabeticIndex As Byte
    Static bytByteIndex As Byte
    Static bytDateIndex As Byte
    Static bytDoubleIndex As Byte
    Static bytStringIndex As Byte
    Static bytLowerIndex As Byte
    Static bytUnsignedSerialIndex As Byte
    Dim bytValue As Byte
    Dim dblRealValue As Double
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intIndex3 As Integer
    Dim lngIntegerValue As Long
    Dim LNGstringLength As Long
    Dim strInstringWork As String
    Dim varEndIndex As Variant
    Dim varExpectedDatatypeWork As Variant
    Dim varStartIndex As Variant
    datatype = False
    Select Case UCase(Trim(strExpectedDatatype))
        Case "ALPHABETIC":
            datatype = (verify(UCase$(strInstring), "ABCDEFGHIJKLMNOPQRSTUVWXYZ") = 0)
        Case "BYTE":
            If isLongInteger(strInstring) Then
                lngIntegerValue = strInstring
                datatype = (lngIntegerValue >= 0 And lngIntegerValue <= 255)
            End If
        Case "UNSIGNEDINTEGER":
            If isLongInteger(strInstring) Then
                lngIntegerValue = strInstring
                datatype = (lngIntegerValue >= 0 And lngIntegerValue <= 32767)
            End If
        Case "INTEGER":
            If isLongInteger(strInstring) Then
                lngIntegerValue = strInstring
                datatype = (lngIntegerValue >= -32768 And lngIntegerValue <= 32767)
            End If
        Case "UNSIGNEDLONG":
            If isLongInteger(strInstring) Then
                lngIntegerValue = strInstring
                datatype = (lngIntegerValue >= 0)
            End If
        Case "LONG":
            If isLongInteger(strInstring) Then
                lngIntegerValue = strInstring
                datatype = True
            End If
        Case "SINGLE":
            If IsNumeric(strInstring) Then
                On Error Resume Next
                    dblRealValue = CSng(strInstring)
                    datatype = (Err.Number = 0)
                On Error GoTo 0
            End If
        Case "DOUBLE":
            datatype = IsNumeric(strInstring)
        Case "UNSIGNEDSERIAL":
            If isLongInteger(strInstring) Then
                datatype = Mid$(strInstring, 1, 1) <> "-"
            End If
        Case "DATE":
            datatype = (IsDate(strInstring))
        Case "STRING":
            datatype = True
        Case "UPPER":
            datatype = verify(strInstring, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") = 0
        Case "LOWER":
            datatype = verify(strInstring, "abcdefghijklmnopqrstuvwxyz") = 0
        Case "LABEL":
            datatype = InStr(strInstring, " ") = 0
        Case "SERIAL":
            datatype = (verify(strInstring, "0123456789") = 0)
        Case "SIGNUM":
            If datatype(strInstring, "UNSIGNEDINTEGER") Then
                lngIntegerValue = CInt(strInstring)
                datatype = (lngIntegerValue <= 1 And lngIntegerValue >= -1)
            End If
        Case "VBID":
            If Len(strInstring) > 0 Then
                If datatype(Mid$(strInstring, 1, 1), "ALPHABETIC") Then
                    datatype = (verify(strInstring, _
                                       "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                       "abcdefghijklmnopqrstuvwxyz0123456789_") _
                                = _
                                0) _
                                And _
                                Len(strInstring) <= "VBID_MAX_LENGTH"
                End If
            End If
        Case "SQLID":
            If datatype(strInstring, "VBID") Then
                datatype = True
            Else
                LNGstringLength = Len(strInstring)
                If LNGstringLength > 2 Then
                   If Mid$(strInstring, 1, 1) = "[" _
                      And _
                      Mid$(strInstring, LNGstringLength, 1) = "]" Then
                       datatype = datatype(Replace(strInstring, " ", ""), "VBID")
                    End If
                End If
            End If
        Case "PERCENTVALUE":
            If datatype(strInstring, "BYTE") Then
                bytValue = CByte(strInstring)
                datatype = (bytValue <= 100)
            End If
        Case "NATURALNUMBER":
            If datatype(strInstring, "UNSIGNEDLONG") Then
                datatype = (CLng(strInstring) <> 0)
            End If
        Case Else:
            errorHandler 32767, _
                         "Invalid datatype code " & enquote(strExpectedDatatype), _
                         "datatype", _
                         "clsUtilities"
    End Select
Datatype_Lbl9_exit:
    Exit Function
End Function
Public Function deco2Var(ByVal strDeco As String, _
                         Optional ByRef varDefault As Variant) As Variant
    '
    ' Convert decorated variant in the form type(value) back to variant form
    '
    '
    Dim strType As String
    Dim strValue As String
    Dim varValue As Variant
    If Not isDeco(strDeco, strType, strValue) Then
        errorHandler 1, _
                     "strDeco " & enquote(strDeco) & " is invalid", _
                     "deco2Var", _
                     "clsUtilities"
        Exit Function
    End If
    On Error Resume Next
        Select Case UCase(strType)
            Case "VBBOOLEAN": varValue = CBool(strValue)
            Case "VBBYTE": varValue = CByte(strValue)
            Case "VBINTEGER": varValue = CInt(strValue)
            Case "VBLONG": varValue = CLng(strValue)
            Case "VBSINGLE": varValue = CSng(strValue)
            Case "VBDOUBLE": varValue = CDbl(strValue)
            Case "VBSTRING": varValue = dequote(CStr(strValue))
            Case Else:
                If Not IsMissing(varDefault) Then
                    varValue = varDefault
                Else
                    errorHandler 1, _
                                 "Cannot convert decorated variant with type " & _
                                 strType, _
                                 "deco2Var", _
                                 "clsUtilities"
                End If
        End Select
    On Error GoTo 0
    deco2Var = varValue
End Function
Public Function decode(ByVal strInstring As String, _
                       ByVal strKey As String, _
                       Optional ByVal booAscii As Boolean = True) As String
    '
    ' Perform simple decoding of the output of the encode method
    '
    '
    ' strInstring is converted to its decoded (original) form by subtracting the Ascii
    ' value in strKey from the corresponding character (the character in the same position) in
    ' strInstring.
    '
    ' The subtraction is modulus 256 unless the optional booAscii parameter is present
    ' and False: if this parameter is False, the subtraction is modulus 32767. Note that
    ' the booAscii parameter should be False if the characters encoded can contain
    ' Unicode or DBCS characters.
    '
    ' If the key is shorter than the input string, then the decoding restarts at the
    ' beginning of the key.
    '
    ' See also encode.
    '
    '
    Dim intDecode As Integer
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intKey As Integer
    Dim intNext As Integer
    Dim lngModulus As Long
    Dim strOutstring As String
    lngModulus = IIf(booAscii, 256, 32767)
    intIndex2 = 1
    For intIndex1 = 1 To Len(strInstring)
        intNext = AscW(Mid(strInstring, intIndex1, 1))
        If booAscii And intNext > 255 Then
            errorHandler 0, _
                         "Invalid unicode/DBCS character in data with integer value " & _
                         intNext & " " & _
                         "can't be encoded because booAscii is True", _
                         "encode", "clsUtilities"
        End If
        intKey = AscW(Mid(strKey, intIndex2, 1))
        intIndex2 = intIndex2 + 1
        If intIndex2 > Len(strKey) Then intIndex2 = 1
        If booAscii And intKey > 255 Then
            errorHandler 0, _
                         "Invalid unicode/DBCS character in key with integer value " & _
                         intKey & " " & _
                         "can't be used because booAscii is True", _
                         "encode", "clsUtilities"
        End If
        intDecode = intNext - intKey
        If intKey > intNext Then intDecode = 256 + intDecode
        strOutstring = strOutstring & ChrW$(intDecode)
    Next intIndex1
    decode = strOutstring
End Function
Public Function deleteFile(ByVal strFileid As String, _
                           Optional ByRef varErrNumber As Variant, _
                           Optional ByRef varErrDesc As Variant) As Boolean
    '
    ' Delete a file and return True on success or False on failure
    '
    '
    ' One reason for this method is that it's just barbaric to have to code "kill."
    '
    '
    On Error Resume Next
        Kill strFileid
        deleteFile = (Err.Number = 0)
        If Not IsMissing(varErrNumber) Then varErrNumber = Err.Number
        If Not IsMissing(varErrDesc) Then varErrDesc = Err.Description
    On Error GoTo 0
End Function
Public Function driveServerChange(ByRef objDrive As Object, _
                                  ByRef objDirectory As Object, _
                                  Optional ByVal booAdvice As Boolean = True) _
                As String
    '
    ' Drive or server change event handler
    '
    '
    ' This method is a generalized handler for the change to a drive or
    ' server name in a Drive list box (or control with the appropriate Drive
    ' property):
    '
    '
    '      *  It optionally ensures that the drive or server name (the first or only
    '         blank-delimited word in the Drive property of the control)
    '         is assigned to a linked directory control's Path property.
    '
    '      *  It checks for any error (such as an unavailable drive) in the selection
    '         of the new drive, and, when the new drive cannot be accessed this
    '         subroutine restores the previous setting.
    '
    '         The previous setting is restored either from the Path property of the
    '         objDirectory linked directory parameter.
    '
    '
    ' If the optional booAdvice parameter is absent or present and True the error
    ' description is extended with advice, usually "insert readable media in drive."
    '
    ' When called as a function this routine returns a null string (indicating success)
    ' or a string in the format <errorNumber> <errorDescription> (indicating failure.)
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 05 26 00   Nilges         Version 1.0
    ' 08 07 02   Nilges         Added to utilities class
    '
    '
    Dim intIndex1 As Integer
    Dim lngErrorNumber As Long
    Dim strErrorDescription As String
    Dim strPrevious As String
    driveServerChange = "0 Unspecified error"
    If Not (TypeOf objDrive Is Control) Or Not (TypeOf objDirectory Is Control) Then
        errorHandler 1, "Drive and/or directory parameters aren't controls", _
                     "driveServerChange", "clsUtilities"
        Exit Function
    End If
    On Error Resume Next
        strPrevious = objDirectory.Path
    On Error GoTo 0
    If strPrevious = "" Then
        errorHandler 1, "Directory object does not expose Path", _
                     "driveServerChange", "clsUtilities"
        Exit Function
    End If
    intIndex1 = InStr(strPrevious, ":")
    If intIndex1 = 0 Then
        intIndex1 = InStr(strPrevious, "\\")
    End If
    If intIndex1 < 1 Then
        errorHandler 1, "Path of directory does not contain drive or server", _
                     "driveServerChange", "clsUtilities"
        Exit Function
    End If
    strPrevious = Mid$(strPrevious, 1, intIndex1 - 1)
    On Error Resume Next
        objDirectory.Path = Mid$(objDrive.Drive, 1, InStr(objDrive.Drive & " ", " ") - 1)
        lngErrorNumber = Err.Number: strErrorDescription = Err.Description
    On Error GoTo 0
    If lngErrorNumber = 0 Then
        driveServerChange = ""
    Else
        driveServerChange = lngErrorNumber & " " & _
                            strErrorDescription & _
                            IIf(booAdvice _
                                And _
                                UCase$(strErrorDescription) = "DEVICE UNAVAILABLE", _
                                ": Insert readable " & _
                                IIf(UCase$(Mid$(objDrive.Drive, 1, 1)) = "A", _
                                    "floppy disk or other media", _
                                    "media") & " " & _
                                "in drive", _
                                "")
        If strPrevious <> "" Then
            On Error Resume Next
                objDrive.Drive = strPrevious
            On Error GoTo 0
        End If
    End If
End Function
Public Function dequote(ByVal strInstring As String) As String
    '
    ' Remove double quotes from string
    '
    '
    ' Doubled quotes are replaced by single occurences.
    '
    '
    Dim intLength As Integer
    dequote = strInstring
    intLength = Len(strInstring)
    If intLength < 2 Then Exit Function
    dequote = Replace(Mid$(strInstring, 2, intLength - 2), ChrW(34) & ChrW(34), ChrW(34))
End Function
Public Function encode(ByVal strInstring As String, _
                       ByVal strKey As String, _
                       Optional ByVal booAscii As Boolean = True, _
                       Optional ByVal booInspect As Boolean = True) As String
    '
    ' Perform simple encoding
    '
    '
    ' strInstring is converted to a code by adding the Ascii value in strKey to
    ' the corresponding character (the character in the same position) in strInstring.
    '
    ' The addition is modulus 256 unless the optional booAscii parameter is present
    ' and False: if this parameter is False, the addition is modulus 32767. Note that
    ' the booAscii parameter should be False if the characters encoded can contain
    ' Unicode or DBCS characters.
    '
    ' If the key is shorter than the input string, then the encoding restarts at the
    ' beginning of the key.
    '
    ' By default, this method inspects its result. The encoded string is converted
    ' back to the original string using the decode method and the key. If the result
    ' is not identifical to strInstring then an error occurs. This inspection can be
    ' bypassed by passing booInspect:=False.
    '
    ' See also decode.
    '
    '
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intKey As Integer
    Dim lngModulus As Long
    Dim intNext As Integer
    Dim strDecoded As String
    Dim strOutstring As String
    lngModulus = IIf(booAscii, 256, 32767)
    intIndex2 = 1
    For intIndex1 = 1 To Len(strInstring)
        intNext = AscW(Mid(strInstring, intIndex1, 1))
        If booAscii And intNext > 255 Then
            errorHandler 0, _
                         "Invalid unicode/DBCS character in data with integer value " & _
                         intNext & " " & _
                         "can't be encoded because booAscii is True", _
                         "encode", "clsUtilities"
        End If
        intKey = AscW(Mid(strKey, intIndex2, 1))
        intIndex2 = intIndex2 + 1
        If intIndex2 > Len(strKey) Then intIndex2 = 1
        If booAscii And intKey > 255 Then
            errorHandler 0, _
                         "Invalid unicode/DBCS character in key with integer value " & _
                         intKey & " " & _
                         "can't be used because booAscii is True", _
                         "encode", "clsUtilities"
        End If
        strOutstring = strOutstring & ChrW$(CLng(intNext + intKey) Mod lngModulus)
    Next intIndex1
    If booInspect Then
        strDecoded = decode(strOutstring, strKey)
        If strDecoded <> strInstring Then
            errorHandler 0, _
                         "Inspection failure: decoded string " & _
                         enquote(strDecoded) & " " & _
                         "does not match input string " & _
                         enquote(strInstring) & ", " & _
                         "when the input string was encoded as " & _
                         string2Hex(strOutstring), _
                         "encode", "clsUtilities"
        End If
    End If
    encode = strOutstring
End Function
Public Function enquote(ByVal strInstring As String) As String
    '
    ' Return string, adding quotes when string is not already in double quotes
    '
    '
    ' Internal double quotes are replaced by sequences of two internal double quotes.
    ' For example, for the string " we will return """".
    '
    '
    enquote = strInstring
    If isQuoted(strInstring) Then Exit Function
    enquote = Chr(34) & Replace(strInstring, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function
Public Function ellipsis(ByVal strInstring As String, _
                         ByVal lngMaxLength As Long, _
                         Optional strEllipsis As String = "...", _
                         Optional booLeftEllipsis As Boolean) _
                As String
    '
    ' Replaces string by ellipsis
    '
    '
    ' This function is passed a string in strInstring and a max length: it returns the string
    ' contents in a string that does not exceed lngMaxLength.  If any characters have to be
    ' hidden to do this, the string returned ends in the strEllipsis value.
    '
    ' The ellipsis can be placed in FRONT of the string.  To do this when the string is overlong,
    ' pass booLeftEllipsis:=True.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 07 17 99   Nilges         Version 1
    ' 10 26 99   Nilges         Added booLeftEllipsis
    '
    '
    Dim lngEllipsisLength As Long
    Dim LNGstringLength As Long
    ellipsis = ""
    LNGstringLength = Len(strInstring): lngEllipsisLength = Len(strEllipsis)
    If lngEllipsisLength > lngMaxLength Then
        errorHandler 1, "Ellipsis string length exceeds max length", _
                     "ellipsis", "clsUtilities"
        Exit Function
    End If
    If LNGstringLength < lngMaxLength Then
        ellipsis = strInstring
    Else
        If booLeftEllipsis Then
            ellipsis = strEllipsis & Mid$(strInstring, 1, max(0, lngMaxLength - lngEllipsisLength))
        Else
            ellipsis = Mid$(strInstring, 1, max(0, lngMaxLength - lngEllipsisLength)) & strEllipsis
        End If
    End If
End Function
Public Sub errorHandler(ByVal intNumber As Integer, _
                        ByVal strMessage As String, _
                        ByVal strProcedure As String, _
                        ByVal strComponent As String, _
                        ParamArray varSubstitution() As Variant)
    '
    ' Error handler
    '
    '
    ' This method handles the error with number intNumber and description
    ' strMessage, by Raising the error and/or displaying it in the error handler
    ' form.
    '
    ' strProcedure and strComponent should identify the property or method as the
    ' procedure, and the form, object or other component.
    '
    ' varSubstitution normally provides keywords that are used for variable entries
    ' in strMessage and their replacement values as a series of strings.
    '
    ' However, when varSubstitution contains two Boolean values, the first may
    ' enable or suppress the error message display and the second may enable or
    ' suppress the Raise operation.  These values override the defaults that are
    ' stored in the error handler form.
    '
    '
    ' C H A N G E   R E C O R D ----------------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     ----------------------------------------------------
    ' 04 20 03   Nilges         Truncate the error message gracefully when out of
    '                           memory
    '
    '
    Dim booDisplay As Boolean
    Dim booOK As Boolean
    Dim booRaise As Boolean
    Dim lngErr As Long
    Dim strErr As String
    Dim strMessageWork As String
    strMessageWork = strMessage
    If UBound(varSubstitution) = 1 Then
        If VarType(varSubstitution(0)) = vbBoolean _
           And _
           VarType(varSubstitution(1)) = vbBoolean Then
           booDisplay = varSubstitution(0): booRaise = varSubstitution(1)
           booOK = True
        End If
    End If
    If Not booOK Then
        booDisplay = frmErrorHandler.ShowForm
        booRaise = frmErrorHandler.RaiseError
        Select Case UBound(varSubstitution)
            Case 1:
                strMessageWork = keywordChange(strMessageWork, _
                                               "&", _
                                               varSubstitution(0), _
                                               varSubstitution(1))
            Case 3:
                strMessageWork = keywordChange(strMessageWork, _
                                               "&", _
                                               varSubstitution(0), _
                                               varSubstitution(1), _
                                               varSubstitution(2), _
                                               varSubstitution(3))
            Case 5:
                strMessageWork = keywordChange(strMessageWork, _
                                               "&", _
                                               varSubstitution(0), _
                                               varSubstitution(1), _
                                               varSubstitution(2), _
                                               varSubstitution(3), _
                                               varSubstitution(4), _
                                               varSubstitution(5))
            Case 7:
                strMessageWork = keywordChange(strMessageWork, _
                                               "&", _
                                               varSubstitution(0), _
                                               varSubstitution(1), _
                                               varSubstitution(2), _
                                               varSubstitution(3), _
                                               varSubstitution(4), _
                                               varSubstitution(5), _
                                               varSubstitution(6), _
                                               varSubstitution(7))
            Case Else:
        End Select
    End If
    strMessageWork = "Message at " & Now & " from the following component: " & _
                        vbNewLine & vbNewLine & _
                        strComponent & "." & strProcedure & _
                        vbNewLine & vbNewLine & _
                        strMessageWork
    If booDisplay Then
        With frmErrorHandler
            .txtError = "Message cannot be displayed for an unknown reason"
            Do
                On Error Resume Next
                    .txtError = strMessageWork
                    lngErr = Err.Number: strErr = Err.Description
                On Error GoTo 0
                If lngErr = 0 Then Exit Do
                If Not errorHandler_outOfMemory_(lngErr, strErr) Then
                    Exit Do
                End If
                strMessageWork = errorHandler_truncate_(strMessageWork)
                If strMessageWork = "" Then Exit Do
            Loop
            .Show vbModal
        End With
    End If
    If booRaise Then
        Err.Raise 32768 + intNumber, _
                  strComponent & "." & strProcedure, _
                  strMessageWork
    End If
End Sub
Private Function errorHandler_outOfMemory_(ByVal lngErr As Long, _
                                           ByVal strErr As String) As Boolean
    '
    ' Return True if error appears to be an out of memory condition (allowing for
    ' some platform differences), False otherwise
    '
    '
    errorHandler_outOfMemory_ = False
    If lngErr = 7 Or UCase$(Trim$(strErr)) = "OUT OF MEMORY" Then
        errorHandler_outOfMemory_ = True
    End If
End Function
Private Function errorHandler_truncate_(strMessage) As String
    '
    ' Gracefully truncate the error message
    '
    '
    Dim lngLength As Long
    Dim lngTruncate As Long
    errorHandler_truncate_ = ""
    lngLength = Len(strMessage)
    lngTruncate = lngLength * 0.1
    If lngTruncate = 0 Then Exit Function
    lngLength = lngLength - lngTruncate
    errorHandler_truncate_ = Mid$(strMessage, 1, lngLength) & _
                             vbNewLine & vbNewLine & _
                             "The above information is incomplete since it " & _
                             "exceeds the limits of the text box" & _
                             vbNewLine & vbNewLine & _
                             "The original message was " & _
                             Len(strMessage) & " " & _
                             "characters long: " & _
                             lngLength & " " & _
                             "characters are displayed"
End Function
Public Function errorHandlerForm() As Object
    '
    ' Returns the error handler form
    '
    '
    Set errorHandlerForm = frmErrorHandler
End Function
Public Function file2String(ByVal strFileid As String) As String
    '
    ' File to string
    '
    '
    Dim bytFileNumber As Byte
    Dim strInputArea As String
    On Error GoTo file2String_Lbl1_freeFileError
        bytFileNumber = FreeFile()
    On Error GoTo 0
    On Error GoTo file2String_Lbl2_openError
        Open strFileid _
             For Binary _
             Access Read _
             Shared _
             As bytFileNumber
    On Error GoTo 0
    On Error GoTo file2String_Lbl3_storageError
        strInputArea = String$(FileLen(strFileid), " ")
    On Error GoTo 0
    On Error GoTo file2String_Lbl4_getError
        Get bytFileNumber, , strInputArea
    On Error GoTo 0
    file2String = strInputArea
    On Error GoTo file2String_Lbl5_closeError
        Close bytFileNumber
    On Error GoTo 0
file2String_Lbl9_exit:
    Exit Function
file2String_Lbl1_freeFileError:
    MsgBox "file2String error: No free file numbers are available"
    GoTo file2String_Lbl9_exit
file2String_Lbl2_openError:
    MsgBox "file2String error: can 't open file " & strFileid
    GoTo file2String_Lbl9_exit
file2String_Lbl3_storageError:
    MsgBox "file2String error: Insufficient storage"
    GoTo file2String_Lbl9_exit
file2String_Lbl4_getError:
    MsgBox "file2String error: Error in Get: " & _
           Err.Number & " " & Err.Description
    GoTo file2String_Lbl9_exit
file2String_Lbl5_closeError:
    MsgBox "file2String error: Error in Close: " & _
           Err.Number & " " & Err.Description
    GoTo file2String_Lbl9_exit
End Function
Public Function fileExists(ByVal strFileid As String) As Boolean
    '
    ' Return True if file exists: return False otherwise
    '
    '
    Dim lngFileLen As Long
    lngFileLen = -1
    On Error Resume Next
        lngFileLen = FileLen(strFileid)
    On Error GoTo 0
    fileExists = (lngFileLen <> -1)
End Function
Public Function findItem(ByVal strInstring As String, _
                                ByVal strTarget As String, _
                                ByVal strDelimiter As String, _
                                ByVal booSetDelimiter As Boolean, _
                                Optional ByVal booIgnoreCase As Boolean = False) _
       As Integer
    '
    ' Find item in string
    '
    '
    Dim intDelimiterLength As Integer
    Dim strSplitArray() As String
    Dim strTargetWork As String
    Dim intIndex1 As Integer
    Dim strDelimiter1 As String
    Dim strInstringWork As String
    intDelimiterLength = Len(strDelimiter)
    strTargetWork = strTarget
    If intDelimiterLength = 0 Then
        errorHandler 0, "The delimiter can not be a null string", "findItem", "clsUtilities"
    End If
    If booIgnoreCase Then
        strTargetWork = UCase(strTargetWork)
    End If
    ' --- Try two shortcuts to avoid a search
    If Not booSetDelimiter _
       Or _
       booSetDelimiter And Len(strDelimiter) = 1 Then
        ' Shortcut 1: check for common case where string starts with target
        Dim strWork As String
        strWork = Mid(strInstring, 1, Len(strTarget))
        If booIgnoreCase Then strWork = UCase(strWork)
        If strWork = strTargetWork Then
            findItem = 1: Exit Function
        End If
        ' Shortcut 2: use Instr and split.
        intIndex1 = InStr(strDelimiter & strInstring & strDelimiter, _
                          strDelimiter & strTarget & strDelimiter, _
                          IIf(booIgnoreCase, vbTextCompare, vbBinaryCompare))
        If intIndex1 = 0 Then
            findItem = 0: Exit Function
        End If
        strSplitArray = Split(Mid(strInstring, 1, intIndex1 - Len(strDelimiter)), _
                              strDelimiter, _
                              Compare:=IIf(booIgnoreCase, vbTextCompare, vbBinaryCompare))
        findItem = UBound(strSplitArray) + 1
        Exit Function
    End If
    ' --- Brute-force search
    strDelimiter1 = strDelimiter
    If booSetDelimiter Then strDelimiter1 = Mid(strDelimiter, 1, 1)
    strInstringWork = strInstring
    If booIgnoreCase Then strInstringWork = UCase(strInstringWork)
    strSplitArray = Split(translate(strInstringWork, _
                                    strDelimiter, _
                                    copies(intDelimiterLength, strDelimiter)), _
                          strDelimiter1, _
                          Compare:=IIf(booIgnoreCase, vbTextCompare, vbBinaryCompare))
    Dim intNullEntries As Integer
    For intIndex1 = 0 To UBound(strSplitArray)
        If strSplitArray(intIndex1) = "" Then
            intNullEntries = intNullEntries + 1
        Else
            If strSplitArray(intIndex1) = strTargetWork Then
                findItem = intIndex1 + 1 - intNullEntries
                Exit Function
            End If
        End If
    Next intIndex1
End Function

' -----------------------------------------------------------------------
' Find blank-delimited word in string
'
'
Public Function findWord(ByVal strInstring As String, _
                         ByVal strTarget As String, _
                        Optional ByVal booIgnoreCase As Boolean = False) _
       As Integer
    findWord = findItem(strInstring, strTarget, " ", True, booIgnoreCase:=booIgnoreCase)
End Function
Public Function hex2String(ByVal strHex As String, _
                           Optional ByVal bytCharsPerByte As Byte = 2)
    '
    ' Convert hexadecimal to string
    '
    '
    ' This method converts the hex information in strHex to its string value. It expects
    ' to find two characters per byte in strHex unless its optional parameter bytCharsPerByte
    ' specifies a different byte width.
    '
    ' See also string2Hex.
    '
    '
    Dim lngIndex1 As Long
    Dim strHexWork As String
    Dim strOutstring As String
    strHexWork = UCase(Trim(strHex))
    If bytCharsPerByte = 0 Then
        errorHandler 0, _
                     "bytCharsPerByte cannot be zero", _
                     "hex2String", "clsUtilities"
        Exit Function
    End If
    If verify(strHexWork, "0123456789ABCDEF") <> 0 Then
        errorHandler 0, _
                     "Input string " & _
                     enquote(ellipsis(strHex, 32)) & " " & _
                     "does not consist, exclusively, of hex characters", _
                     "hex2String", "clsUtilities"
        Exit Function
    End If
    For lngIndex1 = 1 To Len(strHex) Step bytCharsPerByte
        strOutstring = strOutstring & _
                       ChrW(baseN2Long(Mid(strHex, lngIndex1, bytCharsPerByte), _
                                       "0123456789ABCDEF"))
    Next lngIndex1
    hex2String = strOutstring
End Function
Public Function histogram(ByVal dblValue As Double, _
                          Optional ByVal dblRangeMin As Double = 0, _
                          Optional ByVal dblRangeMax As Double = 32767, _
                          Optional ByVal dblValueMin As Double = 0, _
                          Optional ByVal dblValueMax As Double = 32767) As Double
    '
    ' Calculate "histogram" (maps dblValue to limited range)
    '
    '
    ' This method solves the equation
    '
    '
    '      (value - minValue) / (valueMax - valueMin) = (x - rangeMin) / (rangeMax - rangeMin)
    '
    '
    ' Multiplying both sides by (rangeMax - rangeMin) and adding the rangeMin we obtain the solution
    ' for x, which is the needed range unknown.
    '
    '
    '      x = (rangeMax - rangeMin) * ((value - minValue) / (valueMax - valueMin)) + rangeMin
    '
    '
    histogram = 0
    If dblValueMin = dblValueMax Then Exit Function
    histogram = Abs(dblRangeMax - dblRangeMin) _
                * _
                ((dblValue - dblValueMin) / Abs(dblValueMax - dblValueMin)) _
                + _
                dblRangeMin
End Function
Public Function hungarian2Prefix(ByVal strHungarianID As String) As String
    '
    ' Get the Hungarian prefix (cf. parseHungarian for details)
    '
    '
    Dim strPrefixWork As String
    Dim strSuffixWork As String
    If Not hungarianParse(strHungarianID, strPrefixWork, strSuffixWork) Then Exit Function
    hungarian2Prefix = strPrefixWork
End Function
Public Function hungarian2Suffix(ByVal strHungarianID As String) As String
    '
    ' Get the Hungarian suffix (cf. parseHungarian for details)
    '
    '
    Dim strPrefixWork As String
    Dim strSuffixWork As String
    If Not hungarianParse(strHungarianID, strPrefixWork, strSuffixWork) Then Exit Function
    hungarian2Suffix = strSuffixWork
End Function
Public Function hungarianParse(ByVal strHungarianID As String, _
                               ByRef strPrefix As String, _
                               ByRef strSuffix As String) As Boolean
    '
    ' Parse the Hungarian identifier
    '
    '
    ' This method parses an Hungarian ID into its type prefix and "meaningful" suffix.
    '
    ' The strHungarianID must be valid, according to these rules.
    '
    '
    '      1.  It must be at least four characters long
    '      2.  The first three characters must be lower case alphabetic
    '      3.  The fourth character must be upper case and alphabetic
    '
    '
    ' This method returns True (when called as a function) only when the ID is valid.
    ' Also when it is valid this method places the prefix in strPrefix and the suffix
    ' in strSuffix.
    '
    '
    Dim strPrefixWork As String
    Dim strWork As String
    If Len(strHungarianID) < 4 Then Exit Function
    strPrefixWork = Mid$(strHungarianID, 1, 3)
    If verify(strPrefixWork, "abcdefghijklmnopqrstuvwxyz") <> 0 Then Exit Function
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid$(strHungarianID, 4, 1)) = 0 Then Exit Function
    strPrefix = strPrefixWork: strSuffix = Mid$(strHungarianID, 4)
    hungarianParse = True
End Function
Public Function isDeco(ByVal strDeco As String, _
                       ByRef strType As String, _
                       ByRef strValue As String) As Boolean
    '
    ' Parse decorated variant in the form type(value)
    '
    '
    Dim intIndex1 As Integer
    Dim intLength As Integer
    Dim strDecoTrim As String
    strDecoTrim = Trim$(strDeco)
    intLength = Len(strDecoTrim)
    If Mid$(strDecoTrim, intLength) <> ")" Then Exit Function
    intIndex1 = InStr(strDecoTrim, "(")
    If intIndex1 < 2 Then Exit Function
    strType = Mid$(strDecoTrim, 1, intIndex1 - 1)
    strValue = Mid$(strDecoTrim, intIndex1 + 1, intLength - intIndex1 - 1)
    isDeco = True
End Function
Public Function isLongInteger(ByVal strInstring As String) As Boolean
    '
    ' Return True (string is a long integer) or False
    '
    '
    Dim bytIndex1 As Byte
    Dim lngTest As Long
    bytIndex1 = 1
    Select Case Mid(strInstring, bytIndex1, 1)
        Case "+": bytIndex1 = bytIndex1 + 1
        Case "-": bytIndex1 = bytIndex1 + 1
        Case Else:
    End Select
    If bytIndex1 > Len(strInstring) Then Exit Function
    If verify(strInstring, "0123456789") <> 0 Then Exit Function
    On Error GoTo isLongInteger_Lbl1_errorHandler
        lngTest = CLng(strInstring)
    On Error GoTo 0
    isLongInteger = True
    Exit Function
isLongInteger_Lbl1_errorHandler:
End Function
Public Function isQuoted(ByVal strInstring As String) As Boolean
    '
    ' Return True (string is in quotes) or False
    '
    '
    Dim lngLength As Long
    isQuoted = False
    lngLength = Len(strInstring)
    If lngLength < 2 Then Exit Function
    isQuoted = (Mid$(strInstring, 1, 1) = Chr(34) And Mid$(strInstring, lngLength) = Chr(34))
End Function
Public Function item(ByVal strInstring As String, _
                     ByVal intIndex As Integer, _
                     Optional ByVal strDelimiter As String = " ", _
                     Optional ByVal booSetDelimiter As Boolean = True) _
       As String
    '
    ' Return the nth delimited item from strInstring
    '
    '
    ' This method extracts an "item" from a string, where an item is a
    ' sequence of characters surrounded by a delimiter.
    '
    ' Two types of delimiters are supported:
    '
    '
    '      *  String delimiters are fixed characters or multiple-character
    '         strings which separate the items. Note that strings with string
    '         delimiters can contain null (zero length) items.
    '
    '      *  Set delimiters are sets of single characters such that any
    '         character in the set can be an item delimiter. Note that
    '         runs of adjacent set delimiters are considered a single delimiter,
    '         and this implies that strings with set delimiters cannot contain
    '         null items.
    '
    '
    ' This method has the following optional parameters.
    '
    '
    '      *  The optional strDelimiter parameter specifies the character or
    '         string delimiter; strDelimiter defaults to a space.
    '
    '      *  The optional booSetDelimiter specifies the mode of parsing.
    '
    '         + When booSetDelimiter is True, the item is assumed to be
    '           delimited by "set" delimiters such that no item is null
    '           and delimiter is a set of alternative characters
    '
    '         + When booSetDelimiter is False, the item is assumed to be
    '           delimited by "string" delimiters such that no item is null
    '           and the delimiter is a fixed string or character.
    '
    '
    ' C H A N G E   R E C O R D -------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -------------------------------------------
    ' 05 13 03   Nilges         1. Added documentation
    '                           2. Replaced message box by error handling
    ' 05 25 03   Nilges         1. Bug in parsing for set mode delimiter
    ' 05 31 03   Nilges         1. Converted back to VB-6 for streamlined
    '                              VB-6 library
    '
    Dim intCount As Integer
    Dim intIndex1 As Integer
    Dim intUBound As Integer
    Dim strArray() As String
    Dim strItem As String
    If intIndex < 1 Then
        errorHandler 0, _
                     "intIndex " & intIndex & " is not valid", _
                     "utilities", _
                     "item", _
                     "Returning a null item"
        Exit Function
    End If
    If booSetDelimiter Then
        If Len(strDelimiter) = 1 Then
            ' Use split, then ignore null entries
            If item_string2Array_(strInstring, _
                                  strDelimiter, _
                                  strArray) Then
                For intIndex1 = LBound(strArray) To UBound(strArray)
                    If strArray(intIndex1) <> "" Then
                        intCount = intCount + 1
                        If intCount = intIndex Then Exit For
                    End If
                Next intIndex1
                If intIndex1 <= UBound(strArray) Then
                    strItem = strArray(intIndex1)
                End If
            End If
        Else
            ' Scan the string
            intCount = 1
            intIndex1 = 1
            Do While intIndex1 <= Len(strInstring)
                intIndex1 = verify(strInstring, _
                                   strDelimiter, _
                                   intStartIndex:=intIndex1)
                If intIndex1 = 0 Then Exit Do
                strItem = Mid(strInstring & strDelimiter, _
                              intIndex1, _
                              verify(strInstring, _
                                     strDelimiter, _
                                     intStartIndex:=intIndex1 + 1, _
                                     booMatch:=True) _
                              - _
                              intIndex1)
                If intIndex = intCount Then Exit Do
                intCount = intCount + 1
                intIndex1 = intIndex1 + Len(strItem)
            Loop
        End If
    Else
        ' Use Split to find item
        If item_string2Array_(strInstring, strDelimiter, strArray) Then
            If intIndex <= UBound(strArray) + 1 Then strItem = strArray(intIndex - 1)
        End If
    End If
    item = strItem
End Function
Private Function item_string2Array_(ByVal strInstring As String, _
                                    ByVal strDelimiter As String, _
                                    ByRef strArray() As String) As Boolean
    '
    ' Parse the input string using Split on behalf of item: return True on
    ' success or False on failure
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 11 23 01     Nilges       1.  Bug: incorrect success test
    '
    '
    Dim intUBound As Integer
    strArray = Split(strInstring, strDelimiter)
    intUBound = 0
    On Error Resume Next
        intUBound = UBound(strArray)
    On Error GoTo 0
    item_string2Array_ = (intUBound <> -1)
End Function
Public Function items(ByVal strInstring As String, _
                      Optional ByVal strDelimiter As String = " ", _
                      Optional ByVal booSetDelimiter As Boolean = True) _
       As Integer
    '
    ' Return the count of delimited items from strInstring
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 12 13 01   Nilges         Improved efficiency: use split in all cases,
    '                           counting nulls as needed
    '
    ' 05 31 03   Nilges         Converted back to VB-6 for streamlined
    '                           utility library
    '
    '
    Dim intCount As Integer
    Dim intIndex1 As Integer
    Dim intItems As Integer
    Dim intLBound As Integer
    Dim strArray() As String
    ' Use Split to find raw item count
    strArray = Split(strInstring, strDelimiter)
    intCount = 0
    On Error Resume Next
        intCount = UBound(strArray) + 1
    On Error GoTo 0
    If booSetDelimiter Then
        ' Don't count null items
        For intIndex1 = 0 To intCount - 1
            If strArray(intIndex1) = "" Then intCount = intCount - 1
        Next intIndex1
    End If
    items = intCount
End Function
Public Function keywordChange(ByVal strInstring As String, _
                              ByVal strControlString As String, _
                              ParamArray varSubstitution() As Variant) As String
    '
    ' Replace keywords
    '
    '
    Dim intIndex1 As Integer
    Dim strInstringWork As String
    Dim strNextKeyword As String
    Dim strNextValue As String
    strInstringWork = strInstring
    For intIndex1 = LBound(varSubstitution) To UBound(varSubstitution) Step 2
        strNextKeyword = varSubstitution(intIndex1)
        If Mid$(strNextKeyword, 1, Len(strControlString)) <> strControlString Then
            strNextKeyword = strControlString & strNextKeyword
        End If
        strNextValue = varSubstitution(intIndex1 + 1)
        strInstringWork = Replace(strInstringWork, strNextKeyword, strNextValue)
    Next intIndex1
    On Error GoTo 0
    keywordChange = strInstringWork
End Function
Public Function long2BaseN(ByVal lngBase10 As Long, _
                           ByVal strDigits As String, _
                           Optional ByRef varSuccess As Variant, _
                           Optional ByVal bytWordSize As Byte) As String
    '
    ' Convert base 10 number to base N
    '
    '
    ' This function converts a base 10 integer to its representation that uses the
    ' N distinct characters in strBaseN.  The first character in strBaseN
    ' represents zero: the second represents one, and so on.
    '
    ' For example, if lngBase10 is 8 and strBaseN is "01", then this function converts
    ' 8 to the binary value 1000.
    '
    ' If the input number is negative, an error occurs UNLESS:
    '
    '
    '      *  The new base is 2 (binary)
    '      *  A nonzero value is specified in bytWordSize as described below
    '
    '
    ' The base n (binary) representation of negative numbers is returned in
    ' twos-complement notation (which more or less requires a fixed word length.)
    ' In this notation, the lead digit is always one, and the remaining
    ' digits contain the absolute value of the number MINUS ONE, with ones changed to
    ' zeroes and zeroes to ones.
    '
    ' For example, if the input number is -1, the new base is 2 and the word size
    ' is 16 this function will return the string 11111111111111111: if the input
    ' number is -5, the new base is 2 and the word size is 8 this function will
    ' return the string 11111011.
    '
    ' If the optional reference parameter varSuccess is present, it is set to
    ' True when strBaseN is a valid base n number, that can be converted to a
    ' Long integer: otherwise varSuccess is set to False.  If varSuccess is
    ' missing this function will raise an errorHandler error when an error occurs.
    '
    ' If the optional parameter bytWordSize is present, it may specify the
    ' fixed-length word size of the result: otherwise exactly the needed number
    ' of significant digits are returned.  Note that the word size must be specified
    ' when the input number is negative.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 12 30 98   Nilges         Version 1
    ' 05 23 00   Nilges         Added varSuccess parameter
    ' 05 23 00   Nilges         Handle negative numbers in twos-
    '                           complement notation
    ' 05 20 03   Nilges         Added to streamlined utilities library
    '
    '
    Dim intBaseValue As Integer
    Dim lngBase10Work As Long
    Dim lngNext As Long
    Dim strBaseN As String
    long2BaseN = ""
    If Not IsMissing(varSuccess) Then varSuccess = True
    If Len(strDigits) < 2 Then
        If IsMissing(varSuccess) Then
            errorHandler 0, _
                         "Only zero or one digits were supplied for the output number", _
                         "long2BaseN", _
                         "clsUtilities"
        Else
            varSuccess = False
        End If
        Exit Function
    End If
    If lngBase10 = 0 Then
        strBaseN = Mid$(strDigits, 1, 1)
    Else
        intBaseValue = Len(strDigits)
        If lngBase10 < 0 Then
            If intBaseValue <> 2 Or bytWordSize = 0 Then
                If IsMissing(varSuccess) Then
                        errorHandler 0, _
                                     "Input number " & lngBase10 & " is negative, " & _
                                     "but base is not 2 or word size is not specified", _
                                     "long2BaseN", _
                                     "clsUtilities"
                Else
                    varSuccess = False
                End If
                Exit Function
            End If
            lngBase10Work = Abs(lngBase10) - 1
        Else
            lngBase10Work = lngBase10
        End If
        strBaseN = ""
        Do While lngBase10Work <> 0
            strBaseN = Mid$(strDigits, lngBase10Work Mod intBaseValue + 1, 1) & strBaseN
            lngBase10Work = lngBase10Work \ intBaseValue
        Loop
    End If
    If bytWordSize <> 0 Then
        If Len(strBaseN) <= bytWordSize Then
            strBaseN = align(strBaseN, _
                             bytWordSize, _
                             enuAlignment:=alignRight, _
                             strFill:=Mid$(strDigits, 1, 1))
        Else
            If IsMissing(varSuccess) Then
                errorHandler 0, _
                             "Overflow: input number " & lngBase10 & " " & _
                             "cannot be represented in " & _
                             bytWordSize & " digits", _
                             "long2BaseN", "clsUtilities"
            Else
                varSuccess = False
            End If
        End If
    End If
    If lngBase10 < 0 Then strBaseN = translate(strBaseN, _
                                               strDigits, _
                                               Mid$(strDigits, 2, 1) & _
                                               Mid$(strDigits, 1, 1))
    long2BaseN = strBaseN
End Function
Public Function max(ByVal dblX As Double, ByVal dblY As Double) As Double
    '
    ' Return double-precision maximum
    '
    '
    max = IIf(dblX < dblY, dblY, dblX)
End Function
Public Function min(ByVal dblX As Double, ByVal dblY As Double) As Double
    '
    ' Return double-precision mininum
    '
    '
    min = IIf(dblX > dblY, dblY, dblX)
End Function
Public Function mkXMLComment(ByVal strComment As String, _
                             Optional ByVal booMultipleLineEdit As Boolean = False) As String
    '
    ' Make an XML comment
    '
    '
    ' If the booMultipleLineEdit parameter is present and True, and the
    ' comment strComment contains newlines then the comment is broken into
    ' lines and each line is decorated with the XML commenting characters.
    ' If the booMultipleLineEdit parameter is absent or False then a
    ' multiple-line XML comment is returned when the strComment contains
    ' newlines.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------------
    ' 01 25 04     Nilges       Converted from .Net version
    '
    '
    Dim intIndex1 As Integer
    Dim strArray() As String
    Dim strOutstring As String
    If InStr(strComment, vbNewLine) <> 0 And booMultipleLineEdit Then
        On Error GoTo mkXMLComment_LBL1_splitErrorHandler
            strArray = Split(strComment, vbNewLine)
        On Error GoTo 0
        For intIndex1 = LBound(strArray) To UBound(strArray)
            append strOutstring, _
                   vbNewLine, _
                   mkXMLComment(strArray(intIndex1)), _
                   booAppendInPlace:=True
        Next intIndex1
        strOutstring = strOutstring
    Else
        strOutstring = strComment
    End If
    mkXMLComment = ("<!-- " & strOutstring & " -->")
    Exit Function
mkXMLComment_LBL1_splitErrorHandler:
    errorHandler Err.Number, Err.Description, "mkXMLComment", ""
End Function
Public Function mkXMLElement(ByVal strTag As String, _
                              ByVal strValue As String, _
                              ByVal intLineLen As Integer, _
                              ParamArray varAttributes() As Variant) As String
    '
    ' Make an XML element
    '
    '
    ' This method, for a tag name and value creates the tagged XML
    ' element including a start tag, end tag and value.
    '
    ' If the value contains any XML meta-characters (less than,
    ' greater than, ampersand, apostrophe or quote), AND the value
    ' is not an XML tag itself, these characters are converted to their
    ' names &lt;, &gt;, &amp;, &apos;, and &quot; (cf.
    ' http://www.mulberrytech.com/quickref/XMLquickref.pdf).
    '
    ' If the value is a null string then the element is returned as one tag
    ' with the slash at the end to indicate it is self-contained, as in
    ' <tag/>.
    '
    ' Note that this method uses the following test to see whether the
    ' value is an XML tag: if the value starts with a less than
    ' and ends with a greater than it is assumed to be a tag, and the
    ' XML meta-characters in the tag are not changed. Otherwise the
    ' XML meta-characters are changed.
    '
    ' White space is removed from the element to make the above test.
    '
    ' strTag should be the tag name. strValue should be the tag value. intLineLen
    ' should be 0 or the line length limit such that breaking spaces will be
    ' inserted.
    '
    ' The strAttributes parameter array may be passed in the form
    '
    '
    '      attributeName1, attributeValue1, ...
    '
    '
    ' For every attribute its name and its value is passed.  If the value is not
    ' a null string the attribute in the tag has the form name=value: if the value
    ' is null, the attribute is its name alone.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------------
    ' 10 21 02     Nilges       Added attributes parameter array
    ' 10 18 03     Nilges       Convert meta-characters to names
    ' 01 25 04     Nilges       Converted to VB-6 from .Net
    ' 01 26 04     Nilges       Bug: <> was considered probable XML.
    '
    '
    '
    Dim intIndex1 As Integer
    Dim intLineLenWork As Integer
    Dim strElement As String
    Dim strNewline As String
    If intLineLen < 0 Then
        errorHandler 0, _
                     "Internal programming error: intLineLen parameter " & intLineLen & " " & _
                     "is not valid", _
                     "mkXMLElement", "", _
                     "This parameter must be zero or positive.  A null string has been returned."
        Exit Function
    End If
    strNewline = CStr(IIf(intLineLen = 0, "", vbNewLine))
    intLineLenWork = CInt(IIf(intLineLen = 0, Len(strValue), intLineLen))
    For intIndex1 = 1 To Len(strValue) Step intLineLenWork
        strElement = strElement & _
                      strNewline & _
                      Mid(strValue, intIndex1, intLineLenWork)
    Next intIndex1
    If Not mkXMLElement_probableXML_(strElement) Then
        strElement = mkXMLelement_replaceXMLmeta_(strElement)
    End If
    If strElement = "" Then
        mkXMLElement = mkXMLTagWithAttributes(strTag, _
                                              True, _
                                              varAttributes)
        Exit Function
    End If
    mkXMLElement = (mkXMLTagWithAttributes(strTag, False, varAttributes) & _
                    strElement & _
                    mkXMLTag(strTag, booEndTag:=True))
End Function
Private Function mkXMLElement_probableXML_(ByVal strElement As String) _
        As Boolean
    '
    ' Return True on behalf of mkXMLElement if string is probably XML
    '
    '
    ' The string is probably XML when it starts with a less than and ends with
    ' a greater than and is at least three characters in length.
    '
    '
    Dim strWS As String
    Dim strWSgone As String
    If Len(strElement) <= 2 Then Exit Function
    strWS = range2String(0, 32)
    strWSgone = Trim(translate(strElement, strWS, copies(" ", Len(strWS))))
    mkXMLElement_probableXML_ = (Mid(strWSgone, 1, 1) = "<" _
                                 And _
                                 Mid(strWSgone, Len(strWSgone)) = ">")
End Function
Private Function mkXMLelement_replaceXMLmeta_(ByVal strElement As String) As String
    '
    ' Replacement of the XML meta-characters
    '
    '
    Dim strElementWork As String
    strElementWork = strElement
    strElementWork = Replace(strElementWork, "<", "&lt;")
    strElementWork = Replace(strElementWork, ">", "&gt;")
    strElementWork = Replace(strElementWork, "&", "&amp;")
    strElementWork = Replace(strElementWork, "'", "&apos;")
    strElementWork = Replace(strElementWork, """", "&quot;")
    mkXMLelement_replaceXMLmeta_ = strElementWork
End Function
Public Function mkXMLTag(ByVal strTagName As String, _
                         Optional ByVal booEndTag As Boolean = False, _
                         Optional ByVal booSelfContained As Boolean = False) As String
    '
    ' Make an XML tag
    '
    '
    ' This method, for a tag name, returns the XML tag:
    '
    '
    '      *  If booEndTag and booSelfContained are both False, this method
    '         returns <strTagName>.
    '
    '      *  If booEndTag is True and booSelfContained is False, this method
    '         returns </strTagName>.
    '
    '      *  If booEndTag is False and booSelfContained is True, this method
    '         returns <strTagName/>.
    '
    '      *  If booEndTag is True and booSelfContained is True, an error occurs
    '
    '
    ' Note that the name strTagName will be forced to conform to XML standards.
    ' Each nonalpha/nondigit will be converted to its name (using char2name)
    ' surrounded by underscores. If the char2name is more than one character,
    ' it is surrounded by underscores. In addition, if the strTagName starts with a
    ' nonletter or nondigit, the entire XML name will be prefixed with X.
    '
    ' For example, when the tag name is an asterisk, the XML tag name will be
    ' X_asterisk.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------------
    ' 10 21 02     Nilges       Added attributes parameter array: made booEndTag
    '                           an overloaded parameter rather than an Optional
    '                           parameter
    '
    ' 01 24 04     Nilges       Convert name to XML standards
    '
    ' 01 25 04     Nilges       1.  Converted to VB-6
    '                           2.  Added support for self-contained tags
    '
    ' 01 26 04     Nilges       Bug: produced an invalid XML name when the name
    '                           commenced with an underscore. Start such names
    '                           with an extra X
    '
    '
    If booEndTag And booSelfContained Then
        errorHandler 0, _
                     "Internal programming error: endtag and self-contained tag " & _
                     "both specified", _
                     "mkXMLTag", _
                     ""
        Exit Function
    End If
    mkXMLTag = ("<" & _
                CStr(IIf(booEndTag, "/", "")) & _
                mkXMLtag_name2XML_(strTagName) & _
                CStr(IIf(booSelfContained, "/", "")) & _
                ">")
End Function
Public Function mkXMLTagWithAttributes(ByVal strTagName As String, _
                                       ByVal booSelfContained As Boolean, _
                                       ParamArray varAttributes() As Variant) As String
    '
    ' Make an XML tag along with attributes
    '
    '
    ' This method creates a start tag with the name in strTagName. The tag
    ' will be "self-contained", with a slash at the end as in <tag/>, if
    ' the booSelfContained parameter is True.
    '
    ' The parameter varAttributes should be in the format:
    '
    '
    '      attrName1, attrValue1, ...
    '
    '
    ' attrNamen should be an attribute name; attrValuen should be its value.
    ' If attrValue1 is not a null string then the attribute attrNamen=attrValuen
    ' is inserted in the tag.  If the value is null then the attribute name
    ' alone is inserted in the tag.
    '
    ' C H A N G E   R E C O R D --------------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------------
    ' 01 24 04     Nilges       Convert name to XML standards
    ' 01 28 04     Nilges       1.  Bug: attribute handling incorrect
    '                           2.  Need to quote attribute values as well as
    '                               convert their meta-characters
    '
    '
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim strTag As String
    Dim strAttributesWork As String
    strTag = mkXMLTag(strTagName, False)
    If UBound(varAttributes) > LBound(varAttributes) Then
        For intIndex1 = LBound(varAttributes) To UBound(varAttributes) Step 2
            intIndex2 = intIndex1 + 1
            If Not append(strAttributesWork, " ", _
                          varAttributes(intIndex1) & _
                          CStr(IIf(varAttributes(intIndex2) = "", _
                                   "", _
                                   "=" & _
                                   enquote(mkXMLelement_replaceXMLmeta_(varAttributes(intIndex2))))), _
                          booAppendInPlace:=True) Then
                errorHandler Err.Number, _
                             Err.Description, _
                             "mkXMLTag_attributes2String_", ""
            End If
        Next intIndex1
    End If
    If strAttributesWork <> "" Then strAttributesWork = " " & strAttributesWork
    mkXMLTagWithAttributes = Mid(strTag, 1, Len(strTag) - 1) & _
                             strAttributesWork & _
                             CStr(IIf(booSelfContained, "/", "")) & _
                             ">"
End Function
Private Function mkXMLtag_name2XML_(ByVal strName As String) As String
    '
    ' Convert XML tag name to a string on behalf of mkXMLTag
    '
    '
    Dim intIndex1 As Integer
    Dim strNext As String
    Dim strOutstring As String
    For intIndex1 = 1 To Len(strName)
        strNext = char2Name(Mid(strName, intIndex1, 1))
        If Len(strNext) > 1 Then
            strNext = "_" & strNext & "_"
            If intIndex1 = 1 Then strNext = "X" & strNext
        End If
        strOutstring = strOutstring & strNext
    Next intIndex1
    mkXMLtag_name2XML_ = (strOutstring)
End Function
Public Function name2Char(ByVal strName As String) As String
    '
    ' Name to character
    '
    '
    ' This method converts a character name (as returned by char2Name)
    ' back to the original character
    '
    '
    Dim strNameWork As String
    strNameWork = Trim(strName)
    If Len(strNameWork) = 1 Then
        If strNameWork >= "0" And strNameWork <= "9" Then
            name2Char = strNameWork: Exit Function
        End If
        If strNameWork >= "A" And strNameWork <= "Z" Then
            name2Char = strNameWork: Exit Function
        End If
        If strNameWork >= "a" And strNameWork <= "z" Then
            name2Char = strNameWork: Exit Function
        End If
    End If
    strNameWork = UCase(strNameWork)
    Select Case strNameWork
        Case "TILDE": name2Char = "~"
        Case "ACUTEACCENT": name2Char = "`"
        Case "EXCLAMATIONPOINT": name2Char = "!"
        Case "ATSIGN": name2Char = "@"
        Case "POUNDSIGN": name2Char = "#"
        Case "DOLLARSIGN": name2Char = "$"
        Case "PERCENT": name2Char = "%"
        Case "CARAT": name2Char = "^"
        Case "AMPERSAND": name2Char = "&"
        Case "ASTERISK": name2Char = "*"
        Case "LEFTPARENTHESIS": name2Char = "("
        Case "RIGHTPARENTHESIS": name2Char = ")"
        Case "UNDERSCORE": name2Char = "_"
        Case "DASH": name2Char = "-"
        Case "PLUS": name2Char = "+"
        Case "EQUALS": name2Char = "="
        Case "LEFTBRACE": name2Char = "{"
        Case "LEFTBRACKET": name2Char = "["
        Case "RIGHTBRACE": name2Char = "}"
        Case "RIGHTBRACKET": name2Char = "]"
        Case "VERTICALSTROKE": name2Char = "|"
        Case "BACKSLASH": name2Char = "\"
        Case "COLON": name2Char = ":"
        Case "SEMICOLON": name2Char = ";"
        Case "DOUBLEQUOTE": name2Char = """"
        Case "SINGLEQUOTE": name2Char = "'"
        Case "LESSTHAN": name2Char = "<"
        Case "COMMA": name2Char = ","
        Case "GREATERTHAN": name2Char = ">"
        Case "PERIOD": name2Char = "."
        Case "QUESTIONMARK": name2Char = "?"
        Case "SLASH": name2Char = "/"
        Case "NUL": name2Char = Chr(0)
        Case "TAB": name2Char = vbTab
        Case "LINEFEED": name2Char = Chr(10)
        Case "CARRIAGERETURN": name2Char = Chr(13)
        Case "BLANK": name2Char = " "
        Case Else
            If Len(strNameWork) = 8 _
               And _
               Mid(strNameWork, 1, 3) = "CHR" Then
                strNameWork = Mid(strNameWork, 4)
                If verify(strNameWork, "0123456789") = 0 Then
                    name2Char = ChrW(CInt(strNameWork))
                End If
                Exit Function
            End If
            errorHandler 0, _
                         "Invalid character name " & _
                         enquote(strNameWork), _
                         "", _
                         "name2Char", _
                         "Returning a null string"
    End Select
End Function
Public Function range2String(ByVal varStart As Variant, _
                             ByVal varEnd As Variant) As String
    '
    ' Character range to string
    '
    '
    ' This method returns the string of characters, starting with varStart
    ' and ending with varEnd.  Note that the character values can lie
    ' outside the ASCII character range.
    '
    ' If varStart and varEnd are both Integers then the range returned starts with
    ' the dbcs/unicode value of varStart and it ends with the dbcs/unicode value
    ' of varEnd. This is a range of sequentially ascending characters unless
    ' varStart is greater than varEnd, where it is a range of sequentially
    ' descending characters.
    '
    ' If varStart and varEnd are both single-character Strings then the range
    ' returned starts with the character in varStart and it ends with the character
    ' in varEnd. This is a range of sequentially ascending characters unless
    ' varStart is greater than varEnd, where it is a range of sequentially
    ' descending characters.
    '
    ' An error results if varStart and varEnd have other Variant types.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   ---------------------------------------
    ' 04 19 03     Nilges       Version 1
    '
    '
    Dim intEnd As Integer
    Dim intStart As Integer
    Dim intIndex1 As Integer
    Dim strOutstring As String
    range2String = ""
    intStart = -1
    If VarType(varStart) = vbInteger And VarType(varEnd) = vbInteger Then
        If varStart >= 0 And varEnd >= 0 Then
            intStart = varStart: intEnd = varEnd
        End If
    ElseIf VarType(varStart) = vbString And VarType(varEnd) = vbString Then
        If Len(varStart) = 1 And Len(varEnd) = 1 Then
            intStart = AscW(varStart): intEnd = AscW(varEnd)
        End If
    Else
        errorHandler 1000, _
                     "Invalid varStart/varEnd: " & _
                     "varStart = " & var2Deco(varStart) & ": " & _
                     "varEnd = " & var2Deco(varEnd), _
                     "range2String", "clsUtilities"
        Exit Function
    End If
    For intIndex1 = intStart To intEnd Step CInt(IIf(intStart <= intEnd, 1, -1))
        strOutstring = strOutstring & ChrW(intIndex1)
    Next intIndex1
    range2String = strOutstring
End Function
Public Function removeParentheses(ByVal strInstring As String) As String
    '
    ' Remove balanced parentheses when present
    '
    '
    Dim lngLength As Long
    removeParentheses = strInstring
    lngLength = Len(strInstring)
    If lngLength < 2 Then Exit Function
    If Mid$(strInstring, 1, 1) <> "(" Or Mid$(strInstring, lngLength) <> ")" Then Exit Function
    removeParentheses = Mid$(strInstring, 2, lngLength - 2)
End Function
Public Function searchListBox(ByRef objListBox As Object, _
                              ByVal strTarget As String, _
                              Optional ByVal booCaseSensitive As Boolean = False, _
                              Optional ByVal booTrim As Boolean = False) As Long
    '
    ' Search list box
    '
    '
    ' This method returns the index of the target or -1 on not found.  The search is case
    ' insensitive, unless the optional parameter booCaseSensitive is present and True:
    ' the search ignores leading and trailing spaces unless the optional parameter
    ' booTrim is present and False.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 08 28 02   Nilges         Handle list-type controls other than
    '                           list boxes
    '
    '
    Dim ctlHandle As Control
    Dim intListCount As Integer
    Dim lngIndex1 As Long
    Dim strNext As String
    Dim strTargetWork As String
    searchListBox = -1
    On Error GoTo searchListBox_Lbl1_objectErrorHandler
        Set ctlHandle = objListBox
        intListCount = ctlHandle.ListCount
    On Error GoTo 0
    strTargetWork = strTarget
    If Not booCaseSensitive Then strTargetWork = UCase$(strTargetWork)
    If booTrim Then strTargetWork = Trim$(strTargetWork)
    With ctlHandle
        For lngIndex1 = 0 To intListCount - 1
            On Error GoTo searchListBox_Lbl1_objectErrorHandler
                strNext = ctlHandle.List(lngIndex1)
            On Error Resume Next
            If Not booCaseSensitive Then strNext = UCase$(strNext)
            If booTrim Then strNext = Trim$(strNext)
            If strTargetWork = strNext Then Exit For
        Next lngIndex1
        If lngIndex1 < intListCount Then
            searchListBox = lngIndex1
        End If
    End With
    Exit Function
searchListBox_Lbl1_objectErrorHandler:
    errorHandler ERR_BADTYPE_NUMBER, _
                 ERR_BADTYPE_NUMBER, _
                 "searchListBox", "clsUtilities"
End Function
Public Function Signum(ByVal dblValue As Double) As Integer
    '
    ' Return 0 for dblValue=0: return 1 for dblValue>0: else return -1
    '
    '
    If dblValue > 0 Then
        Signum = 1: Exit Function
    End If
    If dblValue = 0 Then
        Signum = 0: Exit Function
    End If
    Signum = -1
End Function
Public Function parse(ByVal strInstring As String, _
                      Optional ByVal strDelimiter As String = " ", _
                      Optional ByVal intLimit As Long = -1, _
                      Optional ByVal booBinaryCompare As Boolean = True, _
                      Optional ByVal lngStartIndex As Long = 1) As Collection
    '
    ' Parse delimited tokens in a string
    '
    '
    ' This method is similar to split, with the following differences:
    '
    '
    '      *  It produces a new collection in place of the array returned
    '         by split, allowing it more flexible output
    '
    '      *  It provides the start index and the length of each token
    '
    '
    ' strInstring should be the string for parsing.
    '
    ' The optional strDelimiter parameter should be the token separator. If this
    ' parameter is not specified, it defaults to a single space. Note that runs or
    ' series of adjacent delimiters will separate null items.
    '
    ' The optional intLimit parameter can specify the split limit as the number of
    ' tokens to be returned. If intLimit is not specified then all tokens are
    ' parsed to the end of the string.
    '
    ' By default, comparision is exact binary, and case sensitive. Use the optional
    ' parameter booBinaryCompare:=False to obtain a case insensitive text compare.
    '
    ' By default, the parse starts at position 1 of the strInstring, but this value
    ' may be overriden using the optional lngStartIndex parameter.
    '
    ' The new collection returned as the function value is an unkeyed series
    ' of token entries, and, each token entry is a three element subcollection,
    ' containing the following items:
    '
    '
    '      *  Item(1) is the value of the token
    '      *  Item(2) is the start index of the token (from 1)
    '      *  Item(3) is the length of the token
    '
    '
    ' On most errors, Nothing will be returned by this method.
    '
    '
    Dim colHandle(1) As Collection
    Dim lngCount As Long
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim lngLength As Long
    Set parse = Nothing
    If strDelimiter = "" Then
        errorHandler 0, _
                     "Delimiter cannot be blank", _
                     "parse", "clsUtilities"
        Exit Function
    End If
    If intLimit < -1 Then
        errorHandler 0, _
                     "Invalid intLimit parameter " & intLimit, _
                     "parse", "clsUtilities"
        Exit Function
    End If
    If lngStartIndex < 1 Then
        errorHandler 0, _
                     "Invalid lngStartIndex parameter " & lngStartIndex, _
                     "parse", "clsUtilities"
        Exit Function
    End If
    On Error GoTo parse_Lbl1_collectionCreateErrorHandler
        Set colHandle(0) = New Collection
    On Error GoTo 0
    lngIndex1 = lngStartIndex
    Do While lngIndex1 <= Len(strInstring)
        lngIndex2 = InStr(lngIndex1, _
                          strInstring & strDelimiter, _
                          strDelimiter, _
                          IIf(booBinaryCompare, _
                              VbCompareMethod.vbBinaryCompare, _
                              VbCompareMethod.vbTextCompare))
        If lngIndex2 > Len(strInstring) Then Exit Do
        lngCount = lngCount + 1
        If intLimit <> -1 And lngCount > intLimit Then Exit Do
        lngLength = lngIndex2 - lngIndex1
        On Error GoTo parse_Lbl1_collectionCreateErrorHandler
            Set colHandle(1) = New Collection
            With colHandle(1)
                .Add Mid(strInstring, lngIndex1, lngLength)
                .Add lngIndex1
                .Add lngLength
            End With
            colHandle(0).Add colHandle(1)
        On Error GoTo 0
    Loop
    Set parse = colHandle(0)
    Exit Function
parse_Lbl1_collectionCreateErrorHandler:
    errorHandler ERR_COLLECTIONCREATE_NUMBER, _
                 ERR_COLLECTIONCREATE, _
                 "parse", _
                 "clsUtilities", _
                 "&C", "special character", _
                 "&N", Err.Number, _
                 "&D", Err.Description
End Function
Public Function SQLenquote(ByVal strInstring As String) As String
    '
    ' Enquote a string using the SQL standard
    '
    '
    SQLenquote = "'" & Replace(strInstring, "'", "''") & "'"
End Function
Public Function string2Box(ByVal strInstring As String, _
                           Optional ByVal strBoxLabel As String = "", _
                           Optional ByVal strBuildChar As String = "*") As String
    '
    ' Wrap some lines in a nice box that is suitable for viewing by nerds
    '
    '
    ' This function returns strInstring, edited to fit inside a box synthesized
    ' from asterisks (or the character specified in the optional strBuildChar
    ' parameter.)  The result is best viewed inside a monospaced font by a nerd.
    ' It allows logs and other displays to be crudely formatted.
    '
    ' If the optional strBoxLabel parameter is present and not null, it is
    ' inserted into the top line of the box.
    '
    '
    ' C H A N G E   R E C O R D -----------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   -----------------------------------------------
    ' 04 26 03     Nilges       Changed integer to Long arithmetic
    '
    '
    Dim colLines As New Collection
    Dim lngIndex1 As Integer
    Dim lngIndex2 As Integer
    Dim intMaxLength As Integer
    Dim intNextLength As Integer
    Dim strFence As String
    Dim strOutstring As String
    string2Box = ""
    ' --- Build character must be a character
    If Len(strBuildChar) <> 1 Then
        errorHandler ERR_STRING2BOX_INVALID_BUILD_CHAR_NUMBER, _
                     ERR_STRING2BOX_INVALID_BUILD_CHAR, _
                     "string2Box", "clsUtilities"
        Exit Function
    End If
    ' --- Find length of longest line, while parsing lines
    intMaxLength = 0: lngIndex1 = 1
    colLines.Add ""
    Do While lngIndex1 <= Len(strInstring)
        lngIndex2 = InStr(lngIndex1, strInstring & vbNewLine, vbNewLine)
        intNextLength = lngIndex2 - lngIndex1
        colLines.Add (Mid(strInstring, lngIndex1, intNextLength))
        If intNextLength > intMaxLength Then intMaxLength = intNextLength
        lngIndex1 = lngIndex2 + Len(vbNewLine)
    Loop
    colLines.Add ""
    intMaxLength = intMaxLength + 4
    intMaxLength = CInt(max(intMaxLength, _
                            CInt(IIf(strBoxLabel = "", 10, Len(strBoxLabel) + 15))))
    ' --- Produce the text box
    strFence = String$(intMaxLength, strBuildChar)
    If strBoxLabel = "" Then
        strOutstring = strFence
    Else
        strOutstring = String$(5, strBuildChar) & _
                       " " & strBoxLabel & " " & _
                       String$(CInt(max(0, intMaxLength - 5 - Len(strBoxLabel) - 2)), _
                               strBuildChar)
    End If
    For lngIndex1 = 1 To colLines.Count
        strOutstring = strOutstring & _
                       vbNewLine & _
                       strBuildChar & " " & _
                       CStr(colLines(lngIndex1)) & _
                       String$(intMaxLength - Len(colLines(lngIndex1)) - 4, " ") & _
                       " " & strBuildChar
    Next lngIndex1
    string2Box = strOutstring & vbNewLine & strFence
End Function
Public Function string2File(ByVal strOutstring As String, _
                            ByVal varFile As Variant, _
                            Optional ByRef varFileNumber As Variant, _
                            Optional ByVal booClose As Boolean = True) As Boolean
    '
    ' String to file: return True (file written ok) or False (file not
    ' written ok)
    '
    '
    ' This function writes strOutstring to a file.
    '
    ' varFile may be a string or a byte:
    '
    '
    '      *  If varFile is a string then the identified file is opened, replaced by
    '         strOutstring, and closed
    '
    '      *  If varFile is a byte then strOutstring is appended to the end of the
    '         open file identified by varFile.  The file is neither opened nor closed.
    '         Note that the out string is appended to the last line of the file
    '         when it is not prefixed by a carriage return and a line feed
    '
    '
    ' If the optional reference parameter varFileNumber is present it is set to the
    ' file number used to open the file (when varFile is a string) or it is set
    ' to varFile (when varFile is a byte.)
    '
    ' If the optional parameter booClose is present and True, or absent, then the
    ' file is CLOSED after output: if booClose is present and False, then the file
    ' remains open.  Note that if you use varFileNumber to obtain the file number
    ' you normally should pass booClose:=False so that the file number as returned
    ' is useful the next time you call this routine.
    '
    '
    Dim booOpenClose As Boolean
    Dim bytFileNumber As Byte
    string2File = False
    Select Case VarType(varFile)
        Case vbString: booOpenClose = True
        Case vbByte: booOpenClose = False
        Case vbInteger: booOpenClose = False
        Case vbLong: booOpenClose = False
        Case Else:
            MsgBox "Error in string2File: varFile parameter " & _
                   "(" & var2Deco(varFile) & ") " & _
                   "has an unexpected varType"
            GoTo string2File_Lbl9_exit
    End Select
    If booOpenClose Then
        On Error GoTo string2File_Lbl1_freeFileError
            bytFileNumber = FreeFile()
        On Error GoTo 0
        If fileExists(varFile) Then
            On Error GoTo string2File_Lbl6_killFileError
                Kill varFile
            On Error GoTo 0
        End If
        On Error GoTo string2File_Lbl2_openError
            Open varFile _
                 For Binary _
                 Access Write _
                 Lock Write _
                 As bytFileNumber
        On Error GoTo 0
    Else
        bytFileNumber = varFile
    End If
    On Error GoTo string2File_Lbl4_putError
        Put bytFileNumber, , strOutstring
    On Error GoTo 0
    If booOpenClose And booClose Then
        On Error GoTo string2File_Lbl5_closeError
            Close bytFileNumber
        On Error GoTo 0
    End If
    If Not IsMissing(varFileNumber) Then varFileNumber = bytFileNumber
    string2File = True
string2File_Lbl9_exit:
    Exit Function
string2File_Lbl1_freeFileError:
    MsgBox "Error in string2File: No free file numbers are available"
    GoTo string2File_Lbl9_exit
string2File_Lbl2_openError:
    MsgBox "Error in string2File: Can't open file " & varFile & ": " & _
                                       Err.Number & " " & Err.Description
    GoTo string2File_Lbl9_exit
string2File_Lbl4_putError:
    MsgBox "Error in string2File: Error in Put: " & Err.Number & " " & Err.Description
    GoTo string2File_Lbl9_exit
string2File_Lbl5_closeError:
    MsgBox "Error in string2File: Error in Close: " & Err.Number & " " & Err.Description
    GoTo string2File_Lbl9_exit
string2File_Lbl6_killFileError:
    MsgBox "Error in string2File: Error in Kill: " & Err.Number & " " & Err.Description
    GoTo string2File_Lbl9_exit
End Function
Public Function string2HardParagraph(ByVal strInstring As String, _
                                     Optional ByVal intWidth As Integer = 72, _
                                     Optional ByVal strSpace As String = " ", _
                                     Optional ByVal enuAlignment As ENUalign = alignLeft) As String
    '
    ' String to hard paragraph
    '
    '
    ' This method returns strInstring formatted into "hard" paragraphs.
    '
    ' A hard paragraph is one in which newlines separate lines, and double newlines
    ' separate paragraphs.
    '
    ' Each newline in strInstring is translated to a double newline.  Each line in
    ' its output string is no longer than 72 characters, and no word is split between
    ' lines.
    '
    ' By default, the word separator is assumed to be the blank; the optional strSpace
    ' parameter may specify another value.
    '
    ' The optional enuAlignment parameter may be used to left-align, center, or
    ' right-align the hard paragraph
    '
    '
    ' C H A N G E   R E C O R D -------------------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -------------------------------------------------------
    ' 04 24 03   Nilges         Added strSpace and enuAlignment
    '
    '
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strOutline As String
    Dim strOutstring As String
    Dim strSplit() As String
    Dim strSplit2() As String
    strSplit = Split(Replace(strInstring, vbNewLine, vbNewLine & vbNewLine), strSpace)
    strOutstring = ""
    For lngIndex1 = LBound(strSplit) To UBound(strSplit)
        strSplit2 = Split(strSplit(lngIndex1), vbNewLine)
        For lngIndex2 = LBound(strSplit2) To UBound(strSplit2)
            If Len(strOutline) _
               + _
               Len(strSplit2(lngIndex2)) _
               + _
               IIf(strOutline = "", 0, 1) _
               > _
               intWidth _
               Or _
               lngIndex2 > LBound(strSplit2) Then
                append strOutstring, _
                       vbNewLine, _
                       strOutline, _
                       booAppendInPlace:=True
                strOutline = ""
            End If
            If strSplit2(lngIndex2) <> "" Then
                append strOutline, _
                       strSpace, _
                       strSplit2(lngIndex2), _
                       booAppendInPlace:=True
            End If
        Next lngIndex2
    Next lngIndex1
    If strOutline <> "" Then
        append strOutstring, _
               vbNewLine, _
               align(strOutline, intWidth, enuAlignment), _
               booAppendInPlace:=True
    End If
    string2HardParagraph = strOutstring
End Function
Public Function string2Hex(ByVal strInstring As String, _
                           Optional ByVal bytCharacterLength As Byte = 2) As String
    '
    ' Convert string to hexadecimals
    '
    '
    ' This function converts strInstring to its character-hexadecimal representation.
    ' By default, each character is converted to its two-character hex value; however,
    ' the length used for each character can be specified as the value of
    ' bytCharacterLength.
    '
    ' See also hex2String.
    '
    '
    Dim intIndex1 As Integer
    Dim strOutstring As String
    For intIndex1 = 1 To Len(strInstring)
        strOutstring = strOutstring & _
                       align(long2BaseN(AscW(Mid(strInstring, intIndex1, 1)), _
                                        "0123456789ABCDEF"), _
                             bytCharacterLength, _
                             enuAlignment:=alignRight, _
                             strFill:="0")
    Next intIndex1
    string2Hex = strOutstring
End Function
Public Function tempFileid(ByVal strPath As String, _
                           Optional ByVal strFileType As String = "") As String
    '
    ' Return a temporary file identifier
    '
    '
    ' This method returns an available file identifier in the path identified by strPath.
    ' Its type will be TMP unless the optional parameter strFileType is present and not
    ' null or blank.  Its name will be TEMPnnnn where nnnn is a four-digit sequence number.
    '
    ' If a tempFileid cannot be provided a null string is returned.
    '
    '
    Dim colDir As Collection
    Dim lngIndex1 As Long
    Dim lngFileNumber As Long
    Dim lngNext As Long
    Dim strDir As String
    Dim strFileTypeWork As String
    tempFileid = ""
    strFileTypeWork = Trim$(strFileType)
    If strFileTypeWork = "" Then strFileTypeWork = "TMP"
    On Error GoTo tempFileid_Lbl1_collectionErrorHandler
        Set colDir = New Collection
    On Error GoTo 0
    strDir = Dir$(attachPath(strPath, "*." & strFileType))
    Do While strDir <> ""
        On Error GoTo tempFileid_Lbl1_collectionErrorHandler
            colDir.Add strDir
        On Error GoTo 0
        strDir = Dir$
    Loop
    lngFileNumber = 1
    For lngIndex1 = 1 To colDir.Count
        If tempFileid_isTempFileid_(colDir.item(lngIndex1), lngNext, strFileType:=strFileType) Then
            lngFileNumber = max(lngFileNumber, lngNext)
        End If
    Next lngIndex1
    lngFileNumber = lngFileNumber + 1
    If lngFileNumber < 10000 Then
        tempFileid = attachPath(strPath, _
                                "temp" & _
                                align(lngFileNumber, 4, enuAlignment:=alignRight, strFill:="0") & _
                                "." & _
                                strFileTypeWork)
    End If
    Exit Function
tempFileid_Lbl1_collectionErrorHandler:
    errorHandler ERR_TEMPFILEID_COL_NUMBER, ERR_TEMPFILEID_COL, _
                 "tempFileid", "clsUtilities", _
                 Err.Number, Err.Description
End Function
Private Function tempFileid_isTempFileid_(ByVal strFileid As String, _
                                          ByRef lngNumber As Long, _
                                          Optional ByVal strFileType As String = "") As Boolean
    '
    ' Return True (and set lngNumber) when file id is in the form tempnnnn.TMP (or tempnnnn.typ
    ' when strFileType is present and not null or blank)
    '
    '
    Dim strFileTypeWork As String
    Dim strNumber As String
    Dim strSplit() As String
    tempFileid_isTempFileid_ = False
    strSplit = Split(strFileid, ".")
    If UBound(strSplit) <> 1 Then Exit Function
    If Len(strSplit(0)) <> 8 Then Exit Function
    strFileTypeWork = Trim$(UCase$(strFileType))
    If strFileTypeWork = "" Then strFileTypeWork = "TMP"
    If UCase$(strSplit(1)) <> strFileTypeWork Then Exit Function
    If UCase$(Mid$(strSplit(0), 1, 4)) <> "TEMP" Then Exit Function
    strNumber = Mid$(strSplit(0), 5)
    If verify(strNumber, "0123456789") = 0 Then
        lngNumber = strNumber
        tempFileid_isTempFileid_ = True
    End If
End Function
Public Function test(ByRef strReport As String) As Boolean
    '
    ' Test the library
    '
    '
    strReport = "Test functionality hasn't been implemented"
    test = True
End Function
Public Function translate(ByVal strInstring As String, _
                          ByVal strSourceCharacterSet As String, _
                          ByVal strTargetCharacterSet As String) As String
    '
    ' Translate source to target characters
    '
    '
    ' This function changes each occurence of a character from strSourceCharacterSet
    ' in strInstring to the corresponding character in strTargetCharacterSet, and
    ' returns the translated string as its value.
    '
    ' For each source character, the target character (to which it is translated) is
    ' taken from the corresponding position in strTargetCharacterSet.  If the target
    ' character set is shorter than the source character set then excess source
    ' characters are not translated.
    '
    '
    Dim intIndex1 As Integer
    Dim intTargetLength As Integer
    Dim lngIndex1 As Long
    Dim strNext As String
    Dim strOutstring As String
    strOutstring = ""
    intTargetLength = Len(strTargetCharacterSet)
    For lngIndex1 = 1 To Len(strInstring)
        strNext = Mid$(strInstring, lngIndex1, 1)
        intIndex1 = InStr(strSourceCharacterSet, strNext)
        If intIndex1 <> 0 And intIndex1 <= intTargetLength Then
            strOutstring = strOutstring & Mid$(strTargetCharacterSet, intIndex1, 1)
        Else
            strOutstring = strOutstring & strNext
        End If
    Next lngIndex1
    translate = strOutstring
End Function
Public Function var2Deco(ByVal varInputVariant As Variant, _
                         Optional ByVal varProgress As Variant, _
                         Optional ByVal varProgressBar As Variant, _
                         Optional ByVal booSpacesAreGraphic As Boolean = False) As String
    '
    ' Convert Variant to decorated format: typeName(stringValue)
    '
    '
    ' This function changes a Visual Basic Variant variable's value to a
    ' "decorated" format consisting of the VarType and where possible the value.
    '
    ' The purpose of variant decoration is to be able to store common variants
    ' in environments such as communications where control characters can cause
    ' confusion when they occur in the data.
    '
    ' This function returns a string in the form VarType(value) where
    ' VarType is usually the name (such as vbLong for a Long integer, and vbString
    ' for a string) of the variable's VarType, and value is its value.
    '
    ' The value is represented in one of these ways.
    '
    '
    '      *  If the original variant was numeric, value is its numeric value.
    '         For example, when passed an Integer with the value 4, this
    '         routine returns vbInteger(4).
    '
    '      *  If the original variant is a date value is its date value surrounded by
    '         pound signs.
    '
    '      *  If the original variant was a string that does not contain
    '         internal double quotes, spaces or non-graphic characters
    '         (characters not on common PC keyboards), value is its quoted value.
    '
    '      *  If the original variant was a string containing spaces, non-
    '         graphic characters, single or double quotes, or international characters,
    '         the VarType is vbExpression, and value is an expression in a
    '         variation of VB notation:
    '
    '              +  Spaces, non-graphic characters and internal single double quotes are
    '                 represented by calls on the ChrW function (such as ChrW(32) for
    '                 space, ChrW(10) for a line feed, and ChrW(34) for a double quote.)
    '
    '                 Use the optional booSpacesAreGraphic parameter to specify the treatment
    '                 of spaces as graphic characters.  If booSpacesAreGraphic is present and
    '                 True, then spaces are NOT converted to ChrW(32).
    '
    '              +  Substrings containing graphic characters and no double/single
    '                 quotes, spaces or international characters are quoted using
    '                 SINGLE quotes.
    '
    '              +  Subexpressions are connected using the & (ampersand) as
    '                 a concatenation operator
    '
    '         For example, when passed a String with the value "A", this routine
    '         returns vbString("A").  When passed a string consisting of the
    '         letter A, followed by a blank and a tab character, followed by a (quoted)
    '         double quote, this routine returns:
    '
    '              vbExpression("'A' & ChrW(32) & ChrW(9)")
    '
    '      *  If the original variant was a collection, value is a comma-separated
    '         list of each of the collection entries.  Each entry in this list
    '         is in turn the decorated form of the corresponding collection
    '         entry.
    '
    '         Suppose the collection colTest contains the string "ABC" and the
    '         integer 5; var2Deco(colTest) will return
    '
    '              vbCollection(vbString("ABC"), vbInteger(5))
    '
    '         Suppose the collection colTest2 contains the string "ABC" and the
    '         above collection colTest; var2Deco(colTest2) will return
    '
    '              vbCollection(vbString("ABC"),
    '                           vbCollection(vbString("ABC"), vbInteger(5)))
    '
    '      *  If the original variant was Nothing then Nothing is returned
    '
    '      *  If the original variant was anything else, the value depends on
    '         whether the original variant was an object exposing the object2XML
    '         method:
    '
    '              + If the original variant was an object, which exposed the
    '                object2XML method, and this method returned a string value,
    '                then this function returns typeName(xml), where typeName
    '                is the name of the object's type and xml is the quoted XML
    '                (or other) string returned by the object's XML method.
    '
    '                Note that this will return a "safe" value for all objects
    '                that (1) expose an object2XML value, and (2) use this
    '                method, var2Deco, to return the values of variables.  This
    '                value will be "safe" in the sense that it will not contain
    '                quotes, non-graphic characters, or international characters
    '                that have not been converted to XML.
    '
    '              + If the original variant was not an object, or it does not
    '                expose the object2XML method, or its object2XML method does
    '                not return a string, then var2Deco returns typeName() (the
    '                object's type name followed by empty parentheses.)
    '
    '
    Dim lngIndex1 As Long
    Dim lngIndex2 As Long
    Dim strDateDelimiter As String
    Dim strDeco As String
    Dim strInputVariant As String
    Static strGraphic As String
    Dim strTypeName As String
    If strGraphic = "" Then
        strGraphic = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                     "abcdefghijklmnopqrstuvwxyz" & _
                     "0123456789" & _
                     "`~!@#$%^&*()_-+={[}]|\:;<,>.?/"
    End If
    strTypeName = "vb" & TypeName(varInputVariant)
    strDeco = "()"
    strInputVariant = ""
    If IsObject(varInputVariant) Then
        If (varInputVariant Is Nothing) Then
            var2Deco = "Nothing": Exit Function
        End If
    End If
    If Not (TypeOf varInputVariant Is Collection) Then
        If Not IsObject(varInputVariant) Then
            On Error Resume Next
                strInputVariant = CStr(varInputVariant)
            On Error GoTo 0
            If VarType(varInputVariant) = vbString Then
                lngIndex1 = 1: strInputVariant = ""
                Do While lngIndex1 <= Len(varInputVariant)
                    lngIndex2 = verify(varInputVariant & ChrW(0), _
                                       strGraphic & _
                                       IIf(booSpacesAreGraphic, " ", ""), _
                                       intStartIndex:=lngIndex1)
                    If lngIndex2 > lngIndex1 Then
                        append strInputVariant, _
                               " & ", _
                               enquote(Mid$(varInputVariant, lngIndex1, lngIndex2 - lngIndex1)), _
                               booAppendInPlace:=True
                    End If
                    If lngIndex2 > Len(varInputVariant) Then Exit Do
                    append strInputVariant, _
                           " & ", _
                           "ChrW(" & AscW(Mid$(varInputVariant, lngIndex2, 1)) & ")", _
                           booAppendInPlace:=True
                    lngIndex1 = lngIndex2 + 1
                Loop
            End If
        Else
            On Error Resume Next
                strInputVariant = varInputVariant.object2XML
            On Error GoTo 0
        End If
    Else
        For lngIndex1 = 1 To varInputVariant.Count
            strInputVariant = strInputVariant & _
                              var2Deco(varInputVariant.item(lngIndex1)) & _
                              IIf(lngIndex1 < varInputVariant.Count, ", ", "")
        Next lngIndex1
    End If
    If VarType(varInputVariant) = vbDate Then
        strDateDelimiter = "#"
    Else
        strDateDelimiter = ""
    End If
    var2Deco = strTypeName & "(" & _
                             strDateDelimiter & _
                             strInputVariant & _
                             strDateDelimiter & _
                             ")"
End Function
Public Function varValueInVarType(ByVal varValue As Variant, _
                                   ByVal intType As Integer) As Boolean
    '
    ' Return True (variant's value converts safely to specified type) or False
    '
    '
    Dim varConverted As Variant
    On Error GoTo varValueInVarType_Lbl1_errorHandler
        Select Case intType
            Case vbBoolean: varConverted = CBool(varValue)
            Case vbByte: varConverted = CByte(varValue)
            Case vbInteger: varConverted = CInt(varValue)
            Case vbLong: varConverted = CLng(varValue)
            Case vbSingle: varConverted = CSng(varValue)
            Case vbDouble: varConverted = CDbl(varValue)
            Case vbString: varConverted = CStr(varValue)
            Case Else:
                errorHandler 32768, _
                             "Invalid type " & intType, _
                             "varValueInVarType", _
                             "clsUtilities"
                Exit Function
        End Select
    On Error GoTo 0
    varValueInVarType = True
varValueInVarType_Lbl1_errorHandler:
    Exit Function
End Function
Public Function verify(ByVal strInstring As String, _
                       ByVal strCharacterSet As String, _
                       Optional ByVal intStartIndex As Integer = 1, _
                       Optional ByVal booMatch As Boolean) _
            As Integer
    '
    ' Verify string
    '
    '
    Dim bytIndex1 As Byte
    Dim bytIndex2 As Byte
    Dim intCharacterSetLength As Integer
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
    Dim intVerify As Integer
    intCharacterSetLength = Len(strCharacterSet)
    verify = 0
    ' --- Error checking
    If intStartIndex <= 0 Then
        MsgBox ("Error in verify function: intStartIndex " & intStartIndex & " is not valid")
        Exit Function
    End If
    ' --- Degenerate cases: null input string, unity  set
    If strInstring = "" Then Exit Function
    If booMatch And intCharacterSetLength = 1 Then
        verify = InStr(intStartIndex, strInstring, strCharacterSet)
        Exit Function
    End If
    ' --- Scan the string
    intVerify = 0
    If booMatch Then
        If Len(strCharacterSet) <= 26 And Len(strInstring) > 1024 Then
            ' The character set is small (it is probably digits or
            ' the alphabet), and the input string is large.  The
            ' character set should drive verification.
            intIndex2 = Len(strInstring) + 1
            For bytIndex1 = 1 To CByte(intCharacterSetLength)
                intIndex1 = InStr(intStartIndex, _
                                  strInstring, _
                                  Mid(strCharacterSet, CInt(bytIndex1), 1))
                If intIndex1 <> 0 Then
                    If intIndex1 < intIndex2 Then
                        intIndex2 = intIndex1
                        bytIndex2 = bytIndex1
                        If intIndex1 = intStartIndex Then Exit For
                    End If
                End If
            Next bytIndex1
            intIndex1 = intIndex2
        Else
            For intIndex1 = intStartIndex To Len(strInstring)
                If InStr(strCharacterSet, CStr(Mid(strInstring, intIndex1, 1))) <> 0 Then Exit For
            Next intIndex1
        End If
        intVerify = CInt(IIf(intIndex1 <= Len(strInstring), intIndex1, 0))
    Else
        For intIndex1 = intStartIndex To Len(strInstring)
            intIndex2 = InStr(strCharacterSet, _
                              Mid$(strInstring, intIndex1, 1))
            If intIndex2 = 0 Then
                intVerify = intIndex1
                Exit For
            End If
        Next intIndex1
    End If
    verify = intVerify
End Function
Public Function word(ByVal strInstring As String, ByVal intIndex As Integer) As String
    '
    ' Parse a blank-delimited word
    '
    '
    word = item(strInstring, intIndex)
End Function
Public Function words(ByVal strInstring As String) As Integer
    '
    ' Count blank-delimited words
    '
    '
    words = items(strInstring)
End Function
Public Function xmlTag2EndTag(ByVal strInstring As String) As String
    '
    ' Convert <tag> to </tag>: remove attributes
    '
    '
    ' C H A N G E   R E C O R D -------------------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -------------------------------------------------------
    ' 01 27 04   Nilges         Need to remove attributes
    '
    '
    xmlTag2EndTag = Mid(strInstring, 1, 1) & "/" & Mid(word(strInstring, 1), 2) & ">"
End Function




