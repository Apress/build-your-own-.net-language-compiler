Option Strict

Imports System
Imports System.Collections
Imports Microsoft.VisualBasic
Imports DotNetEnum = System.Enum
Imports System.Drawing
Imports System.Threading

' **************************************************************************
' *                                                                        *
' * A collection of utility methods                                        *
' *                                                                        *
' *                                                                        *
' * This Shared and stateless object exposes a number of utility methods as*
' * well as constant-valued properties.                                    *
' *                                                                        *
' * The rest of this header block addresses the following topics.          *
' *                                                                        *
' *                                                                        *
' *      *  Methods (and some properties) exposed by this object           *
' *      *  Compile-time symbols                                           *
' *      *  Multithreading considerations                                  *
' *      *  References                                                     *
' *      *  Change record                                                  *
' *                                                                        *
' *                                                                        *
' * METHODS (AND SOME PROPERTIES) EXPOSED BY THIS OBJECT ----------------- *
' *                                                                        *
' *                                                                        *
' *      *  abbrev: return True when string1 is an abbreviation of string2 *
' *      *  append: append a string with or without a separator            *
' *      *  appendPath: append to a Windows path                           *
' *      *  align: align and fill a string                                 *
' *      *  asciiCharsetEnum2String: obtains ASCII character sets          *
' *      *  baseN2Long: convert nondecimal base values to Long values      *
' *      *  bg2fgColor: convert a background to a foreground (font) color  *
' *      *  canonicalTypeCast: convert an object to a standard type        *
' *      *  canonicalTypeEnum2Name: convert the type enum to name          *
' *      *  canonicalTypeName2Enum: convert the type name to enum          *
' *      *  charset: return international character sets                   *
' *      *  collection2Report: convert a collection to a columnwise report *
' *      *  collection2String: convert a collection to a string            *
' *      *  copies: make many copies of a string                           *
' *      *  datatype: test a string for a data type                        *
' *      *  datatypeEnum2Name: data type enumerator to name                *
' *      *  datatypeName2Enum: data type name to enumerator                *
' *      *  dequote: remove quotes from a string                           *
' *      *  directoryExists: check for Windows directory                   *
' *      *  display2String: convert displayable string to value            *
' *      *  ellipsis: abbreviates string to ellipsis...                    *
' *      *  enquote: intelligently enquote a string                        *
' *      *  errorHandler: trivial error handler                            *
' *      *  extendTextBox: extend a text box (Windows)                     *
' *      *  file2String: read a file, return its contents as a string      *
' *      *  fileExists: return True when file exists                       *
' *      *  findAbbrev: find potentially abbreviated string                *
' *      *  findItem: locate item in string                                *
' *      *  findXMLtag: locate eXtensible Markup Language tags             *
' *      *  hasReferenceType: determines if object is a reference object   *
' *      *  histogram: models data on a value range                        *
' *      *  inspectionAppend: append inspection and test reports           *
' *      *  int2Digits: convert positive integer to width                  *
' *      *  isQuoted: return True when string is quoted                    *
' *      *  isXMLcomment: return True when a string is an XML comment      *
' *      *  isXMLname: return True when a string is an XML name            *
' *      *  item: return the nth delimited item from a string              *
' *      *  itemPhrase: return the nth through mth adjacent items          *
' *      *  items: return the count of items                               *
' *      *  joinLines: join two multiple-line lists                        *
' *      *  listBox2Registry: save the list box in the Registry (Windows)  *
' *      *  long2BaseN: convert Long integers to nondecimal bases          *
' *      *  mkXMLcomment: make an XML comment                              *
' *      *  mkXMLelement: make an XML element                              *
' *      *  mkXMLtag: make an XML tag                                      *
' *      *  object2String: convert an object to a string                   *
' *      *  parseXMLtag: parse an XML tag                                  *
' *      *  phrase: return the nth through mth adjacent words              *
' *      *  progress: tells whether object supports progress reports       *
' *      *  range2String: create a single string from a start/end range    *
' *      *  registry2Listbox: restore list box from the Registry (Windows) *
' *      *  replaceXMLmetaChars: replace XML special chars <, > and &      *
' *      *  searchListBox: searches list box for a string value (Windows)  *
' *      *  soft2HardParagraph: format paragraphs                          *
' *      *  spinlock: wait for a locked resource, then lock it             *
' *      *  string2Box: place a string in a box of asterisks               *
' *      *  string2Collection: convert a string to a displayable form      *
' *      *  string2Display: convert a string to a displayable form         *
' *      *  string2File: write a string to a file                          *
' *      *  string2Object: convert many strings to an object               *
' *      *  string2Percent: convert string to percent                      *
' *      *  string2Range: convert a string to a range of strings           *
' *      *  string2ValueObject: convert a string to a value object         *
' *      *  test: run stress and smoke tests on the utilities library      *
' *      *  translate: translate source to target characters               *
' *      *  trimContainer: adjust form/container width and height (Windows)*
' *      *  utility: runs one of these methods from the quickBasic engine  *
' *      *  verify: scan a string                                          *
' *      *  windowsUtilities: tells whether object exposes Windows tools   *
' *      *  word: return the nth blank-delimited word from a string        *
' *      *  xy2Point: creates the Point object from coordinates            *
' *      *  xy2Size: creates the Size object from height and width         *
' *                                                                        *
' *                                                                        *
' * COMPILE-TIME SYMBOLS ------------------------------------------------- *
' *                                                                        *
' * The following compile-time symbols may be set to create different ver- *
' * sions of this software:                                                *
' *                                                                        *
' *                                                                        *
' *      *  UTILITIES_PROGRESS to control references to the progress       *
' *         report control.                                                *
' *                                                                        *
' *      *  WINDOWS_LOGGING to control writing of errorHandler messages to *
' *         the Windows 2000 log                                           *
' *                                                                        *
' *                                                                        *
' * The UTILITIES_PROGRESS compile-time symbol should be True to reference *
' * the progressReport DLL in some methods for visual progress reporting.  *
' * If this DLL or its project is not available, set this symbol to False  *
' * in order to be able to compile this object.                            *
' *                                                                        *
' * Note that the progress method indicates whether a specific utilities   *
' * version supports progress reports.                                     *
' *                                                                        *
' * The WINDOWS_LOGGING compile-time symbol should be True, to write all   *
' * messages sent to utilities.errorHandler to the Windows log named       *
' * utilities_errorHandler.                                                *
' *                                                                        *
' * When running on Windows 2000 when WINDOWS_LOGGING is in effect, you can*
' * examine errorHandler messages by accessing the MMC display for Computer*
' * Management.  In Event Viewer within System Tools, view the log named   *
' * Application.                                                           *
' *                                                                        *
' * Errors containing utilities_errorHandler in the column labeled Source  *
' * will be from the utilities error handler.                              *
' *                                                                        *
' * MULTITHREADING CONSIDERATIONS ---------------------------------------- *
' *                                                                        *
' * This object is stateless, containing no Shared or class-level variables*
' * other than enumerators, constants and structures.  Therefore this      *
' * object may be used in multiple threading applications including        *
' * multiple copies running simultaneously and a single copy, with multiple*
' * methods executing simultaneously.                                      *
' *                                                                        *
' *                                                                        *
' * REFERENCES ----------------------------------------------------------- *
' *                                                                        *
' * W3C 1998: World Wide Web Consortium REC-xml-19980210: Extensible       *
' *      Markup Language (XML) 1.0: W3C Recommendation 10-February-1998    *
' *                                                                        *
' *                                                                        *
' * C H A N G E   R E C O R D -------------------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE                        *
' * --------   ----------     -------------------------------------------- *
' * 11 23 01   Nilges         1.  Changed item                             *
' *                           2.  Changed string2Box                       *
' *                           3.  Changed copies                           *
' * 11 24 01   Nilges         1.  Changed verify                           *
' * 12 01 01   Nilges         1.  Changed extendTextBox                    *
' * 12 01 01   Nilges         1.  Added int2Digits                         *
' * 12 04 01   Nilges         1.  Changed quickBasicEval                   *
' * 12 04 01   Nilges         1.  Added string2Object and dequote          *
' *                           2.  Changed object2String                    *
' *                           3.  Added string2Object                      *
' *                           4.  Added canonicalTypeCast                  *
' *                           5.  Added canonicalTypeName2Enum             *
' *                           6.  Added canonicalTypeEnum2Name             *
' *                           7.  Added dequote                            *
' *                           8.  Added isQuoted                           *
' *                           9.  Added string2ValueObject                 *
' * 12 05 01   Nilges         1.  Add GRID                                 *    
' * 12 08 01   Nilges         1.  Add joinlines                            *    
' * 12 09 01   Nilges         1.  Add histogram                            *   
' * 12 13 01   Nilges         1.  Changed items                            * 
' * 02 07 02   Nilges         1.  Added finditem/findword/translate        * 
' * 02 07 02   Nilges         1.  Added baseN2Long and long2BaseN          * 
' * 03 03 02   Nilges         1.  Changed MAXTEXTBOXLENGTH from 3e5 to 3e4,* 
' *                               made it Public                           *
' * 03 06 02   Nilges         1.  Added datatype                           *
' * 03 16 02   Nilges         1.  Changed Module to shared Class           *
' *                               1.1 Changed errorHandler to only Throw:  *
' *                                   removed screen-related functionality *
' *                           2.  Removed Max and Min                      *
' * 03 17 02   Nilges         1.  Changed extendTextBox                    *
' * 03 18 02   Nilges         1.  Added isXMLcomment                       *
' *                           2.  Added errorHandler parameters            *
' *                               2.1 Object                               *
' *                               2.2 Procedure                            *
' *                               2.3 Help                                 * 
' *                           3.  Added indent                             *
' *                           4.  Added mkXMLcomment                       *
' *                           5.  Added mkXMLtag                           *
' *                           6.  Added mkXMLelement                       *
' *                           7.  Added asciiCharsetEnum2String            *
' *                           8.  Added range2String                       *
' *                           9.  Added string2Display and display2String  *
' * 03 19 02   Nilges         1.  Added isCollection                       *
' * 03 20 02   Nilges         1.  Added string interface to display2String *
' *                               and string2Display                       *
' * 03 20 02   Nilges         1.  Added evaluate                           *
' * 03 25 02   Nilges         1.  Added utility                            *
' * 03 25 02   Nilges         1.  Added alignLeft, alignRight, alignCenter *
' * 03 25 02   Nilges         1.  Changed datatype                         *
' * 03 25 02   Nilges         1.  Changed string2Display                   *
' * 10 16 02   Nilges         1.  Removed evaluate                         *
' * 10 20 02   Nilges         1.  Removed doc ref to quickBasic            *
' *                           2.  Changed enquote, dequote, object2String  *
' *                               and string2Object                        *
' *                           3.  Changed mkXMLelement/mkXMLtag            *
' * 10 25 02   Nilges         1.  Added trimContainer, controlRight and    *
' *                               controlBottom                            *
' *                           2.  Added UTILITIES_PROGRESS                 *
' * 11 15 02   Nilges         1.  Changed trimContainer                    *
' * 11 18 02   Nilges         1.  Added soft2HardParagraph                 *
' * 11 20 02   Nilges         1.  Added collection2String                  *
' * 11 21 02   Nilges         1.  Added replaceXMLMetaChars                *
' * 11 23 02   Nilges         1.  Added findXMLtag                         *
' *                           2.  Added isXMLname                          *
' *                           3.  Added parseXMLtag                        *
' * 11 25 02   Nilges         1.  Removed unused declarations              *
' *                           2.  Added mkTempFile                         *
' * 11 26 02   Nilges         1.  Changed file2String                      *
' * 12 13 02   Nilges         1.  Added findAbbrev (added also to utility  *
' *                               interface)                               *
' *                           2.  Added abbrev (added also to utility      *
' *                               interface)                               *
' * 12 27 02   Nilges         1.  Added hasReferenceType                   *
' * 01 03 03   Nilges         1.  Added bg2fgColor                         *
' * 01 06 03   Nilges         1.  Added xy2Point and xy2Size               *
' *                           2.  Added string2Percent                     *
' * 01 07 03   Nilges         1.  Moved Windows utilities to a separate DLL*
' *                           2.  Corrected eXtended Markup Language to    *
' *                               eXtensible Markup Language               *
' *                           3.  Changed baseN2Long and long2BaseN        *
' *                           4.  Added string2Collection                  *
' *                           5.  Changed string2Object                    *
' *                           6.  Changed string2Display                   *
' *                           7.  Added spinlock                           *
' *                           8.  Changed display2String                   *
' *                           9.  Changed errorHandler                     *
' *                           10. Added collection2Report                  *
' *                           11. Added inspectionAppend                   *
' *                           12. Changed itemPhrase                       *
' *                           13. Added WINDOWS_LOGGING                    *
' * 02 17 03   Nilges         1.  Changed collection2String                *
' *                           2.  Changed string2Object                    *
' * 02 19 03   Nilges         1.  Added directoryExists                    *
' * 03 17 03   Nilges         1.  Made unshared functions shared           *
' **************************************************************************

Public Class utilities

    ' ***** Alignment *****
    Public Enum ENUalign
        alignLeft = 1
        alignCenter = 2
        alignRight = 3
    End Enum
    
    ' **** The character sets *****
    Public Enum ENUasciiCharset
        allAscii
        alphabetic            
        alphanumeric
        commonPunctuation
        digits        
        doubleQuote
        graphic
        graphicNonnumeric
        graphicSpecial
        hexDigits
        lower
        null
        upper
        unknown
        XMLmeta
    End Enum
    
    ' ***** The standard types *****
    Public Enum ENUcanonicalType
        ctUnknown = 0
        ctBoolean = 1
        ctByte = 2
        ctChar = 3
        ctShort = 4
        ctInteger = 5
        ctLong = 6
        ctSingle = 7
        ctDouble = 8
        ctString = 9
    End Enum
    
    ' ***** String data types *****
    Public Enum ENUdatatype 
        AlphabeticDatatype                ' Alphabetic string: returned by 
                                          ' datatype(string)
        ByteDatatype                      ' Integer in the range 0..255: returned by 
                                          ' datatype(string)
        DateDatatype                      ' Valid date 
        DoubleDatatype                    ' Real number in the Double range: 
                                          ' returned by datatype(string)
        FalseDatatype                     ' Return value only
        IntegerDatatype                   ' Integer in the range -2^31..2^31-1: 
                                          ' returned by datatype(string)
        LabelDatatype                     ' Word that contains no blanks
        LongDatatype                      ' Integer in the range -2^63..2^63-1: 
                                          ' returned by datatype(string)
        LowerDatatype                     ' Alphabetic lower-case string: returned by 
                                          ' datatype(string)
        PercentValueDatatype              ' Integer percent in range 0..100
        SerialDatatype                    ' Any integer value 
        SignumDatatype                    ' Integer in the range -1..1 
        ShortDatatype                     ' 16-bit integer: returned by 
                                          ' datatype(string)
        SingleDatatype                    ' Real number in the Single range: returned 
                                          ' by datatype(string)
        StringDatatype                    ' Any string
        TrueDatatype                      ' Return value only
        UnknownDatatype                   ' Special value
        UnsignedIntegerDatatype           ' Integer in the range 0..2^31-1
        UnsignedLongDatatype              ' Integer in the range 0..2^63-1
        UnsignedShortDatatype             ' Integer number in range 0..2^31-1
        UnspecifiedDatatype               ' Not normally used
        UpperDatatype                     ' Alphabetic upper-case string: returned by 
                                          ' datatype(string)
        VBIdentifierDatatype              ' Visual Basic identifier 
    End Enum
    
    ' ***** Display to string syntax *****
    Public Enum ENUdisplay2StringSyntax
        C
        XML
        VB
        VBExpression
        VBExpressionCondensed
        Determine
        Invalid
    End Enum
        
    ' ***** bg2fgColor *****        
    Public Const BG2FGCOLOR_DEFAULT As String = "red white green white blue white yellow black 75"
    
    ' ***** spinlock *****
    Private Const SPINLOCK_TIMEOUT_DEFAULT As Integer = 10     ' Secs
    Private Const SPINLOCK_SLEEPWAIT_DEFAULT As Integer = 1    ' Secs
    
    ' -----------------------------------------------------------------------
    ' Return True (strCheck abbreviates strMaster) or False
    '
    ' This method returns True when strCheck (after removal of leading and 
    ' trailing blanks, and conversion to upper case) is identical to
    ' 1 or more characters at the beginning of strMaster (after removal of 
    ' leading and trailing blanks, and conversion to upper case), False otherwise.
    '
    ' See also findAbbrev.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------------
    ' 12 13 02   Nilges         Version 1
    '
    '
    Public Shared Function abbrev(ByVal strCheck As String, _
                                  ByVal strMaster As String) As Boolean
        Dim intLength As Integer = Len(strCheck)
        If intLength > Len(strMaster) Then Return(False)                           
        Return(Mid(Ucase(Trim(strMaster)), 1, Len(strCheck)) _
               = _
               UCase(Trim(strCheck))) 
    End Function                           

    ' -----------------------------------------------------------------------
    ' Align and fill a string
    '
    '
    Public Shared Overloads Function align(ByVal strInstring As String, _
                                           ByVal intAlignLength As Integer) As String
        Return(align(strInstring, intAlignLength, ENUalign.alignLeft, " "))
    End Function
    Public Shared Overloads Function align(ByVal strInstring As String, _
                                           ByVal intAlignLength As Integer, _
                                           ByVal enuAlignment As ENUalign) As String
        Return(align(strInstring, intAlignLength, enuAlignment, " "))
    End Function
    Public Shared Overloads Function align(ByVal strInstring As String, _
                                           ByVal intAlignLength As Integer, _
                                           ByVal strFill As String) As String
        Return(align(strInstring, intAlignLength, ENUalign.alignLeft, strFill))
    End Function
    Public Shared Overloads Function align(ByVal strInstring As String, _
                                           ByVal intAlignLength As Integer, _
                                           ByVal enuAlignment As ENUalign, _
                                           ByVal strFill As String) As String
        Dim intSizeDelta As Integer
        Dim strAligned As String
        Dim strInstringTruncate As String
        Dim strPad As String
        If intAlignLength < 0 Then
            errorHandler("Invalid alignment length " & intAlignLength)
        End If
        If Len(strFill) <> 1 Then
            errorHandler("Fill is not one character")
        End If
        intSizeDelta = Math.max(0, intAlignLength - Len(strInstring))
        strInstringTruncate = Mid(strInstring, 1, intAlignLength)
        If enuAlignment = ENUalign.alignCenter Then
            Dim intQuotient As Integer = intSizeDelta \ 2
            strAligned = copies(intQuotient, strFill) & _
                            strInstringTruncate & _
                            copies(intQuotient + intSizeDelta Mod 2, strFill)
        Else
            strPad = copies(intSizeDelta, strFill)
            Select Case enuAlignment
                Case ENUalign.alignLeft:
                    strAligned = strInstringTruncate & strPad
                Case ENUalign.alignRight:
                    strAligned = strPad & strInstringTruncate
                Case Else:
                    errorHandler("Invalid align code " & enuAlignment)
            End Select
        End If
        Return(strAligned)
    End Function

    ' ------------------------------------------------------------------
    ' Center align a string, padding with blanks or other characters
    '
    '
    ' --- Align with blanks
    Public Shared Overloads Function alignCenter(ByVal strInstring As String, _
                                                 ByVal intAlignLength As Integer) As String
        Return(alignCenter(strInstring, intAlignLength, " "))
    End Function           
    ' --- Align with specified fill                                    
    Public Shared Overloads Function alignCenter(ByVal strInstring As String, _
                                                 ByVal intAlignLength As Integer, _
                                                 ByVal strFill As String) As String
        Return(align(strInstring, intAlignLength, ENUalign.alignCenter, strFill))
    End Function                                           

    ' ------------------------------------------------------------------
    ' Left align a string, padding with blanks or other characters
    '
    '
    ' --- Align with blanks
    Public Shared Overloads Function alignLeft(ByVal strInstring As String, _
                                               ByVal intAlignLength As Integer) As String
        Return(alignLeft(strInstring, intAlignLength, " "))
    End Function           
    ' --- Align with specified fill                                    
    Public Shared Overloads Function alignLeft(ByVal strInstring As String, _
                                               ByVal intAlignLength As Integer, _
                                               ByVal strFill As String) As String
        Return(align(strInstring, intAlignLength, ENUalign.alignLeft, strFill))
    End Function                                           

    ' ------------------------------------------------------------------
    ' Right align a string, padding with blanks or other characters
    '
    '
    ' --- Align with blanks
    Public Shared Overloads Function alignRight(ByVal strInstring As String, _
                                                ByVal intAlignLength As Integer) As String
        Return(alignRight(strInstring, intAlignLength, " "))
    End Function           
    ' --- Align with specified fill                                    
    Public Shared Overloads Function alignRight(ByVal strInstring As String, _
                                               ByVal intAlignLength As Integer, _
                                               ByVal strFill As String) As String
        Return(align(strInstring, intAlignLength, ENUalign.alignRight, strFill))
    End Function                                           

    ' ---------------------------------------------------------------------
    ' Append string
    '
    '
    ' This method attaches strAppend to a string.  If the string is not
    ' null on entry to this code then strSeparator is placed between
    ' the string and strAppend.
    '
    ' The modified string may be a string or a StringBuilder but strAppend and
    ' strSeparator must be strings.
    '
    ' This function returns the appended string when passed a string 
    ' (objString is a string and booIsStringBuilder is False).
    '
    ' If the optional booStart parameter is present and True then the append
    ' string and separator are placed at the start of the main input string.
    '
    ' If the optional intMaxLength parameter is present it may specify
    ' the maximum string length.
    '
    ' When passed a StringBuilder object the StringBuilder is modified
    ' in place, and this function returns True on success and False on 
    ' failure.  A failed append occurs when a maximum length is specified 
    ' using the optional intMaxLength parameter and the string cannot be 
    ' appended without exceeding this length.
    '
    '
    ' --- Append a string and return result
    Public Shared Overloads Function append(ByRef strInstring As String, _
                                            ByVal strSeparator As String, _
                                            ByVal strAppend As String, _
                                            Optional ByVal booToStart As Boolean = False, _
                                            Optional ByVal intMaxLength As Integer = -1) _
           As String
        Return(CStr(append_(strInstring, _
                            Nothing, _
                            strSeparator, _
                            strAppend, _
                            booToStart, _
                            intMaxLength)))
    End Function
    ' --- Append a string to a string builder in place
    Public Shared Overloads Function append(ByRef objStringBuilder As System.Text.StringBuilder, _
                                            ByVal strSeparator As String, _
                                            ByVal strAppend As String, _
                                            Optional ByVal booToStart As Boolean = False, _
                                            Optional ByVal intMaxLength As Integer = -1) _
           As Boolean  
        Dim strString As String
        Return(CBool(append_(strString, _
                             objStringBuilder, _
                             strSeparator, _
                             strAppend, _
                             booToStart, _
                             intMaxLength)))
    End Function
    ' --- Common functionality
    Private Shared Overloads Function append_(ByRef strString As String, _
                                                ByVal objStringBuilder As System.Text.StringBuilder, _
                                                ByVal strSeparator As String, _
                                                ByVal strAppend As String, _
                                                ByVal booToStart As Boolean, _
                                                ByVal intMaxLength As Integer) As Object
        Dim objErrorReturn As Object = Iif(objStringBuilder Is Nothing, "", False)
        Dim objStringBuilderHandle As System.Text.StringBuilder
        If intMaxLength < -1 Then
            MsgBox("Invalid intMaxLength = " & intMaxLength)
            Return(objErrorReturn)
        End If
        If objStringBuilder Is Nothing Then
            Dim objStringBuilderNew As New System.Text.StringBuilder(strString)
            objStringBuilderHandle = objStringBuilderNew
        Else
            objStringBuilderHandle = objStringBuilder
        End If
        With objStringBuilderHandle
            Dim strSeparatorWork As String
            If .Length <> 0 Then  
                strSeparatorWork = strSeparator
            End If
            If intMaxLength <> -1 _
               AndAlso _
               .Length + Len(strAppend) + len(strSeparatorWork) > intMaxLength Then
                Return(objErrorReturn)
            End If
            If booToStart Then
                .Insert(0, strAppend & strSeparatorWork)
            Else
                .Append(strSeparatorWork & strAppend)
            End If
        End With
        If objStringBuilder Is Nothing Then
            strString = objStringBuilderHandle.ToString
            objStringBuilderHandle = Nothing
            Return(strString)
        Else
            Return(True)
        End If
    End Function

    ' ---------------------------------------------------------------------
    ' Append the path to a file title
    '
    '
    ' This method takes care of a situation that arises when a path
    ' is just a drive (for example c:\) as opposed to a path that
    ' includes directories (as in c:\barf\kamunkle.)  If the path ends in
    ' a backslash, this method returns the mere concatenation of strPath
    ' and strFileTitle: if the path does not end in a back slash then one
    ' is added.
    '
    '
    Public Shared Function appendPath(ByVal strPath As String, _
                                      ByVal strFileTitle As String) As String
        Dim objPath As IO.Path
        Return(objPath.Combine(strPath, strFileTitle))
    End Function
    
    ' ------------------------------------------------------------------
    ' Convert the name of an ASCII character set to the string set
    '
    '
    Public Shared Function asciiCharsetEnum2String(ByVal enuCharset As ENUasciiCharset) As String
        Select Case enuCharset
            Case ENUasciiCharset.allAscii: 
                Return(range2String(0, 255))
            Case ENUasciiCharset.alphabetic: 
                Return(range2String("A", "Z") & range2String("a", "z"))
            Case ENUasciiCharset.alphanumeric: 
                Return(asciiCharsetEnum2String(ENUasciiCharset.alphabetic) & _
                       asciiCharsetEnum2String(ENUasciiCharset.digits))
            Case ENUasciiCharset.commonPunctuation: Return(".,;:?!")
            Case ENUasciiCharset.digits: Return("0123456789")
            Case ENUasciiCharset.graphic: Return(asciiCharsetEnum2String(ENUasciiCharset.alphanumeric) & _
                                                 asciiCharsetEnum2String(ENUasciiCharset.graphicSpecial))
            Case ENUasciiCharset.graphicSpecial: Return(" ~!@#$%^&*()_+`-={}[]:;'<>,./|\" & Chr(34))
            Case ENUasciiCharset.hexDigits: Return("0123456789ABCDEF")
            Case ENUasciiCharset.lower: Return(range2String("a", "z"))
            Case ENUasciiCharset.null: Return("")
            Case ENUasciiCharset.upper: Return(range2String("A", "Z"))
            Case ENUasciiCharset.XMLmeta: Return("&'<>" & Chr(34))
        End Select
    End Function

    ' ------------------------------------------------------------------
    ' Convert base N number (in string format) to a Long integer
    '
    '
    '
    ' This method has the following overloaded syntax:
    '
    '
    '      baseN2Long(strBaseN, intBase): Returns the Long value
    '           of the number using the specified base.  The symbols
    '           of the base are taken from the character set
    '           "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".  The base
    '           must be in the range 2..36.
    '
    '      baseN2Long(strBaseN, intBase, bytWordsize): wordsize is the 
    '           maximum length of the base N number.  
    '
    '      baseN2Long(strBaseN, strDigits): Returns the Long value
    '           of the number using the specified set of digits to 
    '           define the base.  The base can be any (reasonable)
    '           range beginning with 2, up to the length of strDigits.
    '
    '           The first digit represents 0, the second digit, 1, and so on,
    '           and the length of digits is the base N.  
    '
    '      baseN2Long(strBaseN, strDigits, bytWordsize): wordsize is the 
    '           maximum length of the base N number.  
    '
    '
    ' In addition to the above four overloads, the optional parameter 
    ' booIgnoreCase is available in each.  booIgnoreCase:=True will cause
    ' this method to ignore case in strDigits when letters are digits.  For
    ' example, when strDigits are the hex digits "0123456789ABCDEF", passing
    ' booIgnoreCase:=True will convert hex values such as "aa12d".
    '
    '
    ' C H A N G E   R E C O R D ----------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     ----------------------------------------
    ' 01 07 03   Nilges         Added overloads to allow the base to be
    '                           specified as a number, with conventional
    '                           digits
    '
    '
    ' --- Base is a number
    Public Shared Overloads Function baseN2Long(ByVal strBaseN As String, _
                                                ByVal intBase As Integer, _
                                                Optional ByVal booIgnoreCase _
                                                         As Boolean = False) As Long
        Return(baseN2Long(strBaseN, intBase, 0, booIgnoreCase))
    End Function
    ' --- Base is a number and a word size is present
    Public Shared Overloads Function baseN2Long(ByVal strBaseN As String, _
                                                ByVal intBase As Integer, _
                                                ByVal bytWordSize As Byte, _
                                                Optional ByVal booIgnoreCase _
                                                         As Boolean = False) As Long
        Return(baseN2Long(strBaseN, _
                          baseN2Long_base2Digits_(intBase), _
                          booIgnoreCase:=booIgnoreCase))                                   
    End Function       
    ' --- Base is specified as the digits                                                    
    Public Shared Overloads Function baseN2Long(ByVal strBaseN As String, _
                                                ByVal strDigits As String, _
                                                Optional ByVal booIgnoreCase _
                                                         As Boolean = False) As Long
        Return baseN2Long(strBaseN, strDigits, 0, booIgnoreCase)
    End Function
    ' --- Base is specified as the digits and a word size is present 
    Public Shared Overloads Function baseN2Long(ByVal strBaseN As String, _
                                                ByVal strDigits As String, _
                                                ByVal bytWordSize As Byte, _
                                                Optional ByVal booIgnoreCase _
                                                         As Boolean = False) As Long
        Dim bytAdjustment As Byte
        Dim intBaseValue As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intSignum As Integer
        Dim lngBase10 As Long
        Dim lngPower As Long
        Dim strBaseNWork As String = strBaseN
        Dim strDigitsWork As String = strDigits
        intBaseValue = Len(strDigits)
        If intBaseValue < 2 Then
            errorHandler("Only zero or one digits were supplied for the output " & _
                         "number")
            Return(0)
        End If
        If intBaseValue <> 2 And bytWordSize <> 0 Then
            errorHandler("Word size can only be defined for binary conversion")
            Return(0)
        End If
        intSignum = 1
        strBaseNWork = strBaseN
        If booIgnoreCase Then 
            strBaseNWork = UCase(strBaseNWork): strDigitsWork = UCase(strDigitsWork)
        End If
        bytAdjustment = 0
        If intBaseValue = 2 And bytWordSize <> 0 Then
            If Mid(strBaseN, 1, 1) = Mid(strDigits, 2, 1) Then
                intSignum = -1: bytAdjustment = 1
                strBaseNWork = translate(Mid(strBaseN, 2), _
                                             strDigits, _
                                             StrReverse(strDigits))
            End If
        End If
        lngPower = 1
        For intIndex1 = Len(strBaseNWork) To 1 Step -1
            intIndex2 = InStr(strDigits, Mid(strBaseNWork, intIndex1, 1))
            If intIndex2 < 1 Then
                errorHandler("Invalid digits in base " & intBaseValue & " number " & strBaseN) 
                Return(0)
            End If
            Try
                lngBase10 = lngBase10 + (intIndex2 - 1) * lngPower
                lngPower = lngPower * intBaseValue
            Catch
                errorHandler("Apparent numeric overflow while calculating base " & _
                             "10 value: " & _
                             Err.Number & " " & Err.Description & ": returning 0")
                Return(0)
            End Try
        Next intIndex1
        Return(lngBase10 * intSignum - bytAdjustment)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Convert the base to a digit string on behalf of baseN2Long and
    ' long2BaseN
    '
    '
    Private Shared Function baseN2Long_base2Digits_(ByVal intBase As Integer) As String
        Dim intBaseWork As Integer = intBase
        If intBaseWork < 2 OrElse intBaseWork > 36 Then
            errorHandler("Invalid intBase " & intBase, _
                         "utilities", "baseN2Long", _
                         "intBase must be greater than one and less than 36.  " & _
                         "Assuming hex conversion (for no good reason)")
            intBaseWork = 16                         
        End If   
        Return(Mid("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", 1, intBaseWork))
    End Function    
    
    ' ----------------------------------------------------------------------
    ' Convert a background to a foreground color
    '
    '
    ' This method converts a background color to a foreground color according
    ' to these default rules:
    '
    '
    '      *  Background is red: return white
    '
    '      *  Background is green: return black
    '
    '      *  Background is blue: return white
    '
    '      *  Background is yellow: return black
    '
    '      *  Otherwise, return black when the color's value is 75% of the
    '         maximum 24-bit value, white otherwise
    '
    ' 
    ' This method has the basic syntax bg2fgColor(bg) and it returns a Color object,
    ' set to the foreground color.
    '
    ' Alternatively, the rules can be specified using the overloaded syntax
    ' bg2fgColor(bg, s) where s is a string in the form
    '
    '
    '      bgColor1 fgColor1 bgColor2 fgColor2 ... bgColorn fgColorn [pct] [*]
    '
    '
    ' In this string:
    '
    '
    '      *  bgColorn is a background color name, and fgColorn is the foreground
    '         color identified by name to be returned by this foreground color.
    '
    '         If bgColorn occurs more than once, the last entry is the effective
    '         entry.
    '
    '      *  Optionally, [pct] may be a percent value expressed in one of the
    '         following (string) styles:
    '
    '         + As an integer between 0 and 100
    '
    '         + As an integer between 0 and 100 that is immediately followed by a percent
    '           sign 
    '
    '         + As a real number between 0 and 1, optionally, immediately followed by a
    '           percent sign
    ' 
    '         which will be used as the percentage of 2^24 that determines,
    '         for a color not specified, the value below which white is returned,
    '         and above which, black is returned.
    '
    '         This value may appear anywhere in the string: if it appears more than
    '         once its last occurence is the effective entry.
    '
    '      *  Anywhere an asterisk appears (surrounded by blanks) in s, it is replaced
    '         by the default value of s:
    '
    '         red white green black blue white yellow black .75
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 01 03 03     Nilges       Version 1
    '
    '
    Public Shared Overloads Function bg2fgColor(ByVal objBG As Color) As Color
        Return(bg2fgColor(objBG, BG2FGCOLOR_DEFAULT))
    End Function    
    Public Shared Overloads Function bg2fgColor(ByVal objBG As Color, _
                                                ByVal strTranslation As String) As Color
        Dim sglWhiteBlackDivide As Single                                         
        Dim colTable As Collection = bg2fgColor_mkTable_(strTranslation, _
                                                         sglWhiteBlackDivide)
        If (colTable Is Nothing) Then Return(Nothing)
        Try
            Return(CType(colTable(UCase(objBG.Name)), Color))
        Catch
            errorHandler("Can't translate background to foreground color", _
                         "utilities", "bg2fgColor", _
                         "Background color name is " & _
                         enquote(objBG.Name) & ": " & _
                         Err.Number & " " & Err.Description)
            Return(Nothing)                         
        End Try        
    End Function    
    
    ' ----------------------------------------------------------------------
    ' Translate string to color translation table on behalf of bg2fgColor
    '
    '
    Private Shared Function bg2fgColor_mkTable_(ByVal strTranslation As String, _
                                                ByRef sglWhiteBlackDivide As Single) As Collection
        Dim colTable As Collection
        Try
            colTable = New Collection
        Catch
            errorHandler("Can't create color table collection", _
                         "utilities", "bg2fgColor_mkTable_", _
                         Err.Number & " " & Err.Description)
            Return(Nothing)                         
        End Try  
        Dim strSplit() As String
        Try
            strSplit = split(Replace(strTranslation, "*", " " & BG2FGCOLOR_DEFAULT & " "), " ")
        Catch
            errorHandler("Can't create color table collection", _
                         "utilities", "bg2fgColor_mkTable_", _
                         Err.Number & " " & Err.Description)
            Return(Nothing)                         
        End Try 
        Dim objColor As Color             
        Dim intIndex1 As Integer = 0
        Dim strNext As String
        Dim sglNext As Single
        Do While intIndex1 <= UBound(strSplit)
            If strSplit(intIndex1) <> "" Then
                strNext = strSplit(intIndex1)
                sglNext = string2Percent(strNext)
                If sglNext >= 0 Then
                    sglWhiteBlackDivide = sglNext
                Else        
                    Try
                        objColor.FromName(UCase(strNext))
                    Catch
                        errorHandler("Error in translation string: " & _
                                     "background color name " & _
                                     enquote(strNext) & " " & _
                                     "doesn't identify a color", _
                                     "utilities", "bg2fgColor_mkTable_", _
                                     "Returning Nothing")
                        Return(Nothing)                                     
                    End Try          
                    If intIndex1 = UBound(strSplit)
                        errorHandler("Error in translation string: " & _
                                     "last background color name " & _
                                     enquote(strNext) & " " & _
                                     "is not followed by a foreground color", _
                                     "utilities", "bg2fgColor_mkTable_", _
                                     "Returning Nothing")
                        Return(Nothing)                                     
                    End If  
                    intIndex1 += 1
                    Try
                        objColor.FromName(UCase(strSplit(intIndex1)))
                    Catch
                        errorHandler("Error in translation string: " & _
                                     "foreground color name " & _
                                     enquote(strNext) & " " & _
                                     "doesn't identify a color", _
                                     "utilities", "bg2fgColor_mkTable_", _
                                     "Returning Nothing")
                        Return(Nothing)                                     
                    End Try                                                             
                    Try
                        colTable.Add(objColor, strNext)
                    Catch
                        errorHandler("Cannot extend translation table", _
                                     "utilities", "bg2fgColor_mkTable_", _
                                     Err.Number & " " & Err.Description)
                        Return(Nothing)                                     
                    End Try            
                End If            
            End If 
            intIndex1 += 1               
        Loop        
        Return(colTable)
    End Function    

    ' ----------------------------------------------------------------------
    ' Convert an object's value to a canonical type identified by name or
    ' enumeration
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 10 23 01   Nilges         Add Boolean and Char
    '
    '
    Private Const CT_UNKNOWN As String = "UNKNOWN"
    Private Const CT_BOOLEAN As String = "BOOLEAN"
    Private Const CT_BYTE As String = "BYTE"
    Private Const CT_CHAR As String = "CHAR"
    Private Const CT_INT16 As String = "INT16"
    Private Const CT_SHORT As String = "SHORT"
    Private Const CT_INT32 As String = "INT32"
    Private Const CT_INTEGER As String = "INTEGER"
    Private Const CT_INT64 As String = "INT64"
    Private Const CT_LONG As String = "LONG"
    Private Const CT_SINGLE As String = "SINGLE"
    Private Const CT_DOUBLE As String = "DOUBLE"
    Private Const CT_STRING As String = "STRING"
    Public Const CANONICAL_TYPES As String = CT_BOOLEAN & " " & CT_BYTE & " " & CT_CHAR & " " & _
                                             CT_SHORT & " " & CT_INTEGER & " " & CT_LONG & " " & _
                                             CT_SINGLE & " " & CT_DOUBLE & " " & CT_STRING
    Public Shared Overloads Function canonicalTypeCast(ByVal objObjectValue As Object, _
                                                        ByVal strTypeName As String) As Object
        Return(canonicalTypeCast(objObjectValue, canonicalTypeName2Enum(strTypeName, True)))
    End Function
    Public Shared Overloads Function canonicalTypeCast(ByVal objObjectValue As Object, _
                                                        ByVal enuType As ENUcanonicalType) As Object
        Try
            Select Case enuType
                Case ENUcanonicalType.ctBoolean: Return(CBool(objObjectValue))
                Case ENUcanonicalType.ctByte: Return(CByte(objObjectValue))
                Case ENUcanonicalType.ctChar: Return(CChar(objObjectValue))
                Case ENUcanonicalType.ctShort: Return(CShort(objObjectValue))
                Case ENUcanonicalType.ctInteger: Return(CInt(objObjectValue))
                Case ENUcanonicalType.ctLong: Return(CLng(objObjectValue))
                Case ENUcanonicalType.ctSingle: Return(CSng(objObjectValue))
                Case ENUcanonicalType.ctDouble: Return(CDbl(objObjectValue))
                Case ENUcanonicalType.ctString: Return(CStr(objObjectValue))
                Case ENUcanonicalType.ctUnknown: Return(Nothing)
                Case Else: Return(Nothing)
            End Select
        Catch
            Return(Nothing)
        End Try
    End Function

    ' ----------------------------------------------------------------------
    ' Canonical type enumeration to name 
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 10 23 01   Nilges         Add Boolean and Char
    '
    '
    Public Shared Function canonicalTypeEnum2Name(ByVal enuCanonicalTypeValue As ENUcanonicalType) As String
        Select Case enuCanonicalTypeValue
            Case ENUcanonicalType.ctBoolean: Return(CT_BOOLEAN)
            Case ENUcanonicalType.ctChar: Return(CT_CHAR)
            Case ENUcanonicalType.ctByte: Return(CT_BYTE)
            Case ENUcanonicalType.ctShort: Return(CT_SHORT)
            Case ENUcanonicalType.ctInteger: Return(CT_INTEGER)
            Case ENUcanonicalType.ctLong: Return(CT_LONG)
            Case ENUcanonicalType.ctSingle: Return(CT_SINGLE)
            Case ENUcanonicalType.ctDouble: Return(CT_DOUBLE)
            Case ENUcanonicalType.ctString: Return(CT_STRING)
            Case ENUcanonicalType.ctUnknown: Return(CT_UNKNOWN)
            Case Else: Return(CT_UNKNOWN)
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Canonical type name to enumeration 
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 10 23 01   Nilges         Add Boolean and Char
    '
    '
    Public Shared Overloads Function canonicalTypeName2Enum(ByVal strName As String, _
                                                            ByVal booErrorHandle As Boolean) As ENUcanonicalType
        Dim enuType As ENUcanonicalType
        enuType = canonicalTypeName2Enum(strName)
        If Not booErrorHandle Then Return(enuType)
        If enuType = ENUcanonicalType.ctUnknown Then
            errorHandler("Canonical type, identified by name as " & enquote(strName) & " " & _
                         "is not valid")
        End If
        Return(enuType)
    End Function
    Public Shared Overloads Function canonicalTypeName2Enum(ByVal strName As String) As ENUcanonicalType
        Dim strNameWork As String = UCase(Trim(strName))
        If Len(strNameWork) >=7 AndAlso Mid(strNameWork, 1, 7) = "SYSTEM." Then
            strNameWork = Mid(strNameWork, 8)
        End If        
        Select Case UCase(Trim(strNameWork))
            Case CT_BOOLEAN: Return(ENUcanonicalType.ctBoolean)
            Case CT_BYTE: Return(ENUcanonicalType.ctByte)
            Case CT_INT16: Return(ENUcanonicalType.ctShort)
            Case CT_SHORT: Return(ENUcanonicalType.ctShort)
            Case CT_INT32: Return(ENUcanonicalType.ctInteger)
            Case CT_INTEGER: Return(ENUcanonicalType.ctInteger)
            Case CT_LONG: Return(ENUcanonicalType.ctLong)
            Case CT_INT64: Return(ENUcanonicalType.ctLong)
            Case CT_SINGLE: Return(ENUcanonicalType.ctSingle)
            Case CT_DOUBLE: Return(ENUcanonicalType.ctDouble)
            Case CT_STRING: Return(ENUcanonicalType.ctString)
            Case else: Return(ENUcanonicalType.ctUnknown)
        End Select
    End Function
    
    ' ---------------------------------------------------------------------
    ' Collection to columnar report
    '
    '    
    ' This method produces a report in columns showing the information in
    ' any collection.
    '
    ' This method can display two types of collections:
    '
    '
    '      *  Any collection containing items, none of which is a collection,
    '         is converted to a single column report.
    '
    '         Items having value types (Boolean, Byte, Short, Integer, Long,
    '         Single, Double or String) are converted to their string value
    '         by default (see below for the use of the booDeco parameter to
    '         add type information.)
    '
    '         Items, other than collections, having reference types are
    '         converted to the string value returned by utilities.object2String.
    '
    '      *  Any collection that contains a collection in each entry
    '         is converted to a multiple column report, where each column
    '         contains the subcollection item.  
    '
    ' 
    ' In the report, fields of numeric type are right aligned while fields
    ' of string or Boolean type are left aligned.
    '
    ' By default the columns are displayed without headers.  However, the
    ' optional parameter strColHeaders may be present and it may specify
    ' column headers as a comma-delimited list of words or phrases.
    '
    ' The intGutter optional parameter may specify the size in spaces of 
    ' the "gutter" the area between columns: by default intGutter is 1.
    '
    ' By default, numbers and strings are converted to strings and
    ' reference objects are converted to their representation as returned
    ' by object2String.  Use the optional booDeco:=True parameter to
    ' display each field in the "decorated" format type(value).
    '
    ' Note: output strings built with this method are best viewed in a 
    ' monospaced font such as Courier New.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE   PROGRAMMER     D E S C R I P T I O N
    ' -------- ----------     -----------------------------------------
    ' 01 12 03 Nilges         1.  Version 1.0
    '
    '
    Public Shared Function collection2Report(ByVal colCollection As Collection, _
                                             Optional ByVal strColHeaders As String = "", _
                                             Optional ByVal intGutter As Integer = 1, _
                                             Optional ByVal booDeco As Boolean = False) _
           As String
        If (colCollection) Is Nothing Then
            errorHandler("Collection is null", _
                         "utilities", "collection2Report", _
                         "Returning a null string")
            Return("")                         
        End If        
        If intGutter < 0 Then
            errorHandler("Invalid intGutter value is " & intGutter, _
                         "utilities", "collection2Report", _
                         "Returning a null string")
            Return("")                         
        End If        
        Dim intIndex1 As Integer
        Dim strColHeaderArray() As String
        Try
            strColHeaderArray = split(strColHeaders, ",")
        Catch
            errorHandler("Cannot convert strColHeaders into an array", _
                         "utilities", "collection2Report", _
                         Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                         "Returning a null string")
            Return("")                         
        End Try        
        Dim intMaxWidth() As Integer
        Dim strCollection(,) As String
        Try
            Redim intMaxWidth(UBound(strColHeaderArray) + 1)
            Redim strCollection(colCollection.Count, UBound(intMaxWidth))
        Catch
            errorHandler("Cannot create the strCollection array and/or intMaxWidth array", _
                         "utilities", "collection2Report", _
                         Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                         "Returning a null string")
            Return("")                         
        End Try   
        For intIndex1 = 1 To UBound(strColHeaderArray)
            intMaxWidth(intIndex1) = Len(strColHeaderArray(intIndex1 - 1))
        Next intIndex1        
        Dim booReferenceType As Boolean  
        Dim colSubcollection As Collection
        Dim intIndex2 As Integer
        Dim strNext As String
        With colCollection
            ' --- Pass one: determine max column widths and verify we have a usable collection
            For intIndex1 = 1 To .Count
                booReferenceType = hasReferenceType(.Item(intIndex1))
                If booReferenceType Then
                    If (TypeOf(.Item(intIndex1)) Is Collection) Then
                        colSubcollection = CType(.Item(intIndex1), Collection)
                    Else
                        booReferenceType = False
                        strNext = object2String(.Item(intIndex1))
                    End If
                Else
                    If booDeco Then 
                        strNext = object2String(.Item(intIndex1), True)
                    Else                        
                        strNext = CStr(.Item(intIndex1))
                    End If                                        
                End If                                            
                If Not booReferenceType Then
                    Try
                        colSubcollection = New Collection
                        colSubcollection.Add(strNext)
                    Catch
                        errorHandler("Cannot create temporary sub-collection", _
                                     "utilities", "collection2Report", _
                                     Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                                     "Returning a null string")
                        Return("")                         
                    End Try     
                End If 
                With colSubcollection
                    If .Count > UBound(intMaxWidth) Then
                        Try
                            Redim Preserve strCollection(UBound(strCollection), .Count)
                            Redim Preserve intMaxWidth(.Count)
                        Catch
                            errorHandler("Cannot expand intMaxWidth collection", _
                                        Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                                        "Returning a null string")
                            Return("")                         
                        End Try     
                    End If     
                    For intIndex2 = 1 To UBound(strCollection, 2)
                        Try
                            strCollection(intIndex1, intIndex2) = CStr(.Item(intIndex2))
                        Catch
                            errorHandler("Cannot convert collection entry " & _
                                         object2String(.Item(intIndex2)) & " " & _
                                         "at row " & intIndex1 & " " & _
                                         "and column " & intIndex2 & " " & _
                                         "to a string", _
                                         "utilities", "collection2Report", _
                                         Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                                         "Returning a null string")
                            Return("")                                         
                        End Try       
                        intMaxWidth(intIndex2) = Math.Max(intMaxWidth(intIndex2), _
                                                          Len(CStr(.Item(intIndex2))))
                    Next intIndex2                                                       
                End With          
                If Not booReferenceType Then colSubcollection = Nothing                     
            Next intIndex1
            ' --- Pass two: create the report
            Dim objStringBuilder As System.Text.StringBuilder
            Try
                objStringBuilder = New System.Text.StringBuilder
            Catch
                errorHandler("Cannot create StringBuilder", _
                             "utilities", "collection2Report", _
                             Err.Number & Err.Description & vbNewline & vbNewline & _
                             "Returning a null string")                
                Return("")                             
            End Try    
            Dim strGutter As String = copies(" ", intGutter)   
            If Trim(strColHeaders) <> "" Then
                For intIndex1 = 0 To UBound(strColHeaderArray)
                    append(objStringBuilder, _
                           " ", _
                           alignCenter(strColHeaderArray(intIndex1), _
                                       intMaxWidth(intIndex1 + 1)))
                Next intIndex1     
                objStringBuilder.Append(vbNewline)           
                For intIndex1 = 0 To UBound(strColHeaderArray)
                    append(objStringBuilder, _
                           strGutter, _
                           copies("-", intMaxWidth(intIndex1 + 1)))
                Next intIndex1                
            End If                 
            Dim objStringBuilder2 As System.Text.StringBuilder
            Try
                objStringBuilder2 = New System.Text.StringBuilder
            Catch
                errorHandler("Cannot create string builder for row display", _
                             "utilities", "collection2Report", _
                             Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                             "Returning a null string")
                Return("")                             
            End Try            
            For intIndex1 = 1 To UBound(strCollection, 1)
                For intIndex2 = 1 To UBound(strCollection, 2)
                    If IsNumeric(strCollection(intIndex1, intIndex2)) Then
                        strCollection(intIndex1, intIndex2) = _
                            alignRight(strCollection(intIndex1, intIndex2), intMaxWidth(intIndex2))
                    Else                            
                        strCollection(intIndex1, intIndex2) = _
                            alignLeft(strCollection(intIndex1, intIndex2), intMaxWidth(intIndex2))
                    End If                                          
                    append(objStringBuilder2, _
                           strGutter, _
                           strCollection(intIndex1, intIndex2))
                Next intIndex2
                append(objStringBuilder, vbNewline, objStringBuilder2.ToString)                
            Next intIndex1            
            Return(objStringBuilder.ToString)
        End With          
    End Function                                                          
    
    ' ---------------------------------------------------------------------
    ' Convert a collection to a string
    '
    '
    ' This function accepts a Collection and returns the comma-delimited
    ' string of its elements:
    '
    '
    '      *  Numeric members are included without change
    '
    '      *  System.String members are included in quotes
    '
    '      *  Collections, embedded in the main collection, are
    '         included as the parenthesized output of this function,
    '         called recursively.  
    '
    '         The recursion depth is limited to the value in the optional
    '         parameter intRecursionDepth.  When recursion depth is exceeded,
    '         contained collections are returned as [Collection] only.  This
    '         avoids the possibility of a loop when a collection contains
    '         a reference to a container.  If intRecursionDepth is -1 then
    '         there is no limit to recursion.  The default value of this
    '         parameter is 8 levels.
    '
    '      *  All other object members are converted using object2String
    '
    '
    ' Note that the string returned by this method may be ambiguous in that
    ' it separates objects in the collection, represented as above, using
    ' commas, but System.Strings may contain commas.
    '
    ' To obtain a string which can be accurately converted back to a collection
    ' by the string2Collection method (see below), use the optional parameter
    ' strDelim, and assign a nongraphic character to it such as ChrW(0) (the
    ' Ascii Nul character.)
    '
    ' If this is done, then string2Collection will be able to convert all
    ' flat collections, that contain System.Strings and numbers only, back to 
    ' identical collections.  Note that string2Collection may still have
    ' trouble with more complex collections.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE   PROGRAMMER     D E S C R I P T I O N
    ' -------- ----------     -----------------------------------------
    ' 11 20 02 Nilges         1.  Version 1.0
    ' 02 17 03 Nilges         1.  Added strDelim parameter
    '                         2.  Decorate the string
    '
    '
    Public Shared Function collection2String(ByVal colCollection As Collection, _
                                             Optional ByVal intRecursionDepth As Integer = 8, _ 
                                             Optional ByVal strDelim As String = ", ") _ 
                           As String
        If intRecursionDepth < -1 Then
            errorHandler("Invalid recursion depth limit of " & intRecursionDepth, _
                         "utilities", "collection2String", _
                         "This parameter must be zero, positive or -1 for no limit")
            Return("")   
        End If        
        If (colCollection Is Nothing) Then Return("")
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler("Cannot create stringBuilder", _
                         "utilities", "collection2String", _ 
                         Err.Number & " " & Err.Description)
            Return("")                         
        End Try       
        Dim intIndex1 As Integer
        Dim intRecursion As Integer
        Dim objNext As Object
        For intIndex1 = 1 To colCollection.Count
            objNext = colCollection.Item(intIndex1)
            If (Typeof objNext Is Collection) Then
                If intRecursion <> -1 AndAlso intRecursion >= intRecursionDepth Then
                    objNext = "[collection]"
                Else
                    objNext = "(" & _
                              collection2String(CType(objNext, Collection), _
                                                intRecursionDepth:=intRecursion + 1) & _
                              ")"                                                
                End If                
            Else
                objNext = object2String(objNext, True)                
            End If            
            append(objStringBuilder, strDelim, CStr(objNext))
        Next intIndex1           
        Return objStringBuilder.ToString     
    End Function                                                          
    
    ' ----------------------------------------------------------------------
    ' Return several copies of a character or a string
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 11 23 01     Nilges       1.  Bug: incorrect Mid 
    '
    '
    ' --- Legacy (reversed) syntax        
    Public Overloads Shared Function copies(ByVal strInstring As String, ByVal intCopies As Integer) As String
        Return(copies(intCopies, strInstring))
    End Function    
    ' --- Normal syntax
    Public Overloads Shared Function copies(ByVal intCopies As Integer, ByVal strInstring As String) As String
        Dim intIndex1 As Integer
        If intCopies < 0 Then
            errorHandler("Invalid intCopies parameter " & intCopies)
            Return("")
        End If
        Dim objStringBuilder As New System.Text.StringBuilder("")
        If Len(strInstring) = 1 Then
            objStringBuilder.Append(CChar(strInstring), intCopies)
        Else
            For intIndex1 = 1 To intCopies
                objStringBuilder.Append(strInstring)
            Next intIndex1
        End If
        Return(objStringBuilder.ToString)
    End Function

    ' -----------------------------------------------------------------------
    ' Datatype: test data type of string
    '
    '
    ' This method can determine the data type of a string, or it can test
    ' for a specific data type.  It has the following overloaded syntax.
    '
    ' 
    '      datatype(string): determines the data type and returns
    '           it as one of the enumerations below.  However, note
    '           that only the enumerations flagged are returned by 
    '           this syntax. 
    ' 
    '      datatype(string, expected): tests string for the expected
    '           data type, and returns True (string has data type) or
    '           False.  The expected data type can be any one of
    '           the ENUdatatype enumerations except for unspecifiedDataype
    '           or unknownDatatype, or it may be a string identifying
    '           one of these enumerations (except for unspecified/unknown.)
    ' 
    '      datatype(string, converted): determines the data type and returns
    '           it as one of the enumerations below.  Only the enumerations
    '           flagged are returned by this syntax.  Also, this overload
    '           sets converted to the value object representing the data
    '           type when converted to the returned type.
    ' 
    '      datatype(string, expected, converted): tests string for the expected
    '           data type, and returns True (string has data type) or
    '           False.  The expected data type can be any one of
    '           the ENUdatatype enumerations except for unspecifiedDatayype
    '           or unknownDatatype, or it may be a string as above.  Also, this overload
    '           sets converted to the value object representing the data
    '           type when converted to the expected type.
    ' 
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 03 25 02   Nilges         Added ability to specify enumeration as a
    '                           string
    '
    ' 
    ' --- Determines the data type
    Public Shared Overloads Function datatype(ByVal strInstring As String) _
           As ENUdatatype
        Return(datatype_(strInstring, _
               ENUdatatype.UnspecifiedDatatype, _
               Nothing, _
               False))
    End Function
    ' --- Tests for the enumerated data type
    Public Shared Overloads Function datatype(ByVal strInstring As String, _
                                              ByVal enuExpectedDatatype As ENUdatatype) _
           As Boolean
        Return(datatype_(strInstring, enuExpectedDatatype, Nothing, False) _
               = _
               ENUdatatype.TrueDatatype)
    End Function
    ' --- Tests for the named data type  
    Public Shared Overloads Function datatype(ByVal strInstring As String, _
                                              ByVal strExpectedDatatype As String) _
           As Boolean
        Return(datatype(strInstring, strExpectedDatatype, Nothing))
    End Function
    ' --- Sets the converted value
    Public Shared Overloads Function datatype(ByVal strInstring As String, _
                                              ByRef objConverted As Object) _
                     As ENUdatatype
        Return(datatype_(strInstring, _
               ENUdatatype.UnspecifiedDatatype, _
               objConverted, True))
    End Function
    ' --- Tests for the enumerated data type and sets the converted value
    Public Shared Overloads Function datatype(ByVal strInstring As String, _
                                              ByVal enuExpectedDatatype As ENUdatatype, _
                                              ByRef objConverted As Object) _
                     As Boolean
        Return(datatype_(strInstring, enuExpectedDatatype, objConverted, True) _
               = _
               ENUdatatype.TrueDatatype)
    End Function
    ' --- Tests for the named data type and sets the converted value
    Public Shared Overloads Function datatype(ByVal strInstring As String, _
                                              ByVal strExpectedDatatype As String, _
                                              ByRef objConverted As Object) _
           As Boolean
        Dim enuExpectedDatatype As ENUdatatype = datatypeName2Enum(strExpectedDatatype)
        If enuExpectedDatatype =  ENUdatatype.UnspecifiedDatatype Then
            errorHandler("Invalid data type name " & enquote(strExpectedDatatype), _
                         "utilities", "datatype", _
                         "Returning False")
            Return(False)                         
        End If                                      
        Return(datatype(strInstring, enuExpectedDatatype, objConverted))
    End Function
    ' --- Common functionality
    Private Shared Overloads Function datatype_(ByVal strInstring As String, _
                                                ByVal enuExpectedDatatype As ENUdatatype, _
                                                ByRef objConverted As Object, _
                                                ByVal booReturnConverted As Boolean) _
                     As ENUdatatype
        '
        ' Test string for data type
        '
        '
        Dim booCheckResult As Boolean
        Dim booDatatypeValue As Boolean
        Dim dblMax As Double
        Dim dblMin As Double
        Dim enuDatatypeValue As ENUdatatype
        Dim objValue As Object
        dblMin = 0: dblMax = 0
        If enuExpectedDatatype <> ENUdatatype.UnspecifiedDatatype Then
            booDatatypeValue = False
            Select Case enuExpectedDatatype
                Case ENUdatatype.AlphabeticDatatype:
                    If verify(UCase(strInstring), _
                              "ABCDEFGHIJKLMNOPQRSTUVWXYZ") _
                    = _
                    0 Then
                        booDatatypeValue = True
                        objValue = CStr(strInstring)
                    End If
                Case ENUdatatype.ByteDatatype:
                    booDatatypeValue = True
                    Try
                        objValue = CByte(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.DateDatatype:
                    If IsDate(strInstring) Then
                        booDatatypeValue = True
                        objValue = CDate(strInstring)
                    End If
                Case ENUdatatype.DoubleDatatype:
                    booDatatypeValue = True
                    Try
                        objValue = CDbl(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                Case ENUdatatype.IntegerDatatype:
                    booDatatypeValue = True
                    Try
                        objValue = CInt(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.LabelDatatype:
                    If Instr(strInstring, " ") = 0 Then
                        booDatatypeValue = True: objValue = strInstring
                    End If
                Case ENUdatatype.LongDatatype:
                    booDatatypeValue = True
                    Try
                        objValue = CLng(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.LowerDatatype:
                    If verify(strInstring, "abcdefghijklmnopqrstuvwxyz") = 0 Then
                        booDatatypeValue = True: objValue = strInstring
                    End If
                Case ENUdatatype.PercentValueDatatype:
                    If datatype(strInstring, ENUdatatype.ByteDatatype, objValue) Then
                        If CInt(objValue) <= 100 Then
                            booDatatypeValue = True
                        End if    
                    End If
                Case ENUdatatype.SerialDatatype:
                    If verify(strInstring, "0123456789") = 0 Then 
                        booDatatypeValue = True: objValue = strInstring
                    End If
                Case ENUdatatype.ShortDatatype:
                    booDatatypeValue = True 
                    Try
                        objValue = CShort(strInstring)
                    Catch
                        Return(ENUdatatype.FalseDatatype)
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.SignumDatatype:
                    If datatype(strInstring, ENUdatatype.ShortDatatype, objValue) Then 
                        booDatatypeValue = True
                        Select Case objValue
                            Case 1:
                            Case 0:
                            Case -1:
                            Case Else: Return(ENUdatatype.FalseDatatype)
                        End Select
                    End If
                Case ENUdatatype.SingleDatatype:
                    booDatatypeValue = True 
                    Try
                        objValue = CSng(strInstring)
                    Catch
                        booDatatypeValue = False 
                    End Try
                Case ENUdatatype.StringDatatype:
                    booDatatypeValue = True: objValue = strInstring
                Case ENUdatatype.UnsignedIntegerDatatype:
                    If datatype(strInstring, _
                                ENUdatatype.IntegerDatatype, _
                                objValue) Then
                        booDatatypeValue = (CInt(objValue) >= 0)
                    End If
                Case ENUdatatype.UnsignedLongDatatype:
                    If datatype(strInstring, ENUdatatype.LongDatatype, objValue) Then
                        booDatatypeValue = (Cint(objValue) >= 0)
                    End If
                Case ENUdatatype.UnsignedShortDatatype:
                    If datatype(strInstring, ENUdatatype.ShortDatatype, objValue) Then
                        booDatatypeValue = (CInt(objValue) >= 0)
                    End If
                Case ENUdatatype.UpperDatatype:
                    If verify(strInstring, _
                              "ABCDEFGHIJKLMNOPQRSTUVWXYZ") _
                    = _
                    0 Then
                        booDatatypeValue = True: objValue = strInstring
                    End If
                Case ENUdatatype.VBIdentifierDatatype:
                    If verify(LCase(strInstring), _
                              "0123456789_abcdefghijklmnopqrstuvwxyz") _
                    = _
                    0 Then
                        Dim strChar1 As String
                        strChar1 = Mid(strInstring, 1, 1)
                        booDatatypeValue = datatype(strChar1, _
                                                    ENUdatatype.AlphabeticDatatype) _
                                           OrElse _
                                           strChar1 = "_"
                        If booDatatypeValue Then objValue = strInstring
                    End If
                Case Else:
                    MsgBox("Error in datatype: invalid datatype code" )
            End Select
            If booReturnConverted Then 
                objConverted = Iif(booDatatypeValue, objValue, Nothing)
            End if
            enuDatatypeValue = CType(iif(booDatatypeValue, _
                                         ENUdatatype.TrueDatatype, _
                                         ENUdatatype.FalseDatatype), _
                                     ENUdatatype)
        Else
            If datatype(strInstring, ENUdatatype.ByteDatatype) Then
                enuDatatypeValue = ENUdatatype.ByteDatatype
            ElseIf datatype(strInstring, ENUdatatype.ShortDatatype) Then
                enuDatatypeValue =  ENUdatatype.ShortDatatype
            ElseIf datatype(strInstring, ENUdatatype.IntegerDatatype) Then
                enuDatatypeValue =  ENUdatatype.IntegerDatatype
            ElseIf datatype(strInstring, ENUdatatype.LongDatatype) Then
                enuDatatypeValue =  ENUdatatype.LongDatatype
            ElseIf datatype(strInstring, ENUdatatype.SingleDatatype) Then
                enuDatatypeValue =  ENUdatatype.SingleDatatype
            ElseIf datatype(strInstring, ENUdatatype.DoubleDatatype) Then
                enuDatatypeValue =  ENUdatatype.DoubleDatatype
            ElseIf datatype(strInstring, ENUdatatype.LowerDatatype) Then
                enuDatatypeValue =  ENUdatatype.LowerDatatype
            ElseIf datatype(strInstring, ENUdatatype.UpperDatatype) Then
                enuDatatypeValue =  ENUdatatype.UpperDatatype
            ElseIf datatype(strInstring, ENUdatatype.AlphabeticDatatype) Then
                enuDatatypeValue =  ENUdatatype.AlphabeticDatatype
            ElseIf datatype(strInstring, ENUdatatype.DateDatatype) Then
                enuDatatypeValue =  ENUdatatype.DateDatatype
            ElseIf datatype(strInstring, ENUdatatype.LabelDatatype) Then
                enuDatatypeValue =  ENUdatatype.LabelDatatype
            ElseIf datatype(strInstring, ENUdatatype.StringDatatype) Then
                enuDatatypeValue =  ENUdatatype.StringDatatype
            Else
                enuDatatypeValue = ENUdatatype.UnknownDatatype
            End If
        End If
        Return(enuDatatypeValue)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of datatype, check for an integer
    '
    '
    Private Shared Function datatype_integerCheck_(ByVal strInstring As String, _
                                                   ByRef objValue As Object) As Boolean
        If CDbl(strInstring) = CLng(strInstring) Then Return(True)
        objValue = Nothing
        Return(False)
    End Function

    ' ----------------------------------------------------------------------
    ' Data type enumerator to name
    '
    '
    Public Shared Function datatypeEnum2Name(ByVal enuDatatypeEnum As ENUdatatype) As String
        Select Case enuDatatypeEnum
            Case ENUdatatype.AlphabeticDatatype: Return("AlphabeticDatatype")
            Case ENUdatatype.ByteDatatype: Return("ByteDatatype")
            Case ENUdatatype.DateDatatype: Return("DateDatatype")
            Case ENUdatatype.DoubleDatatype: Return("DoubleDatatype")
            Case ENUdatatype.FalseDatatype: Return("FalseDatatype")
            Case ENUdatatype.IntegerDatatype: Return("IntegerDatatype")
            Case ENUdatatype.LabelDatatype: Return("LabelDatatype")
            Case ENUdatatype.LongDatatype: Return("LongDatatype")
            Case ENUdatatype.LowerDatatype: Return("LowerDatatype")
            Case ENUdatatype.PercentValueDatatype: Return("PercentValueDatatype")
            Case ENUdatatype.SerialDatatype: Return("SerialDatatype")
            Case ENUdatatype.SignumDatatype: Return("SignumDatatype")
            Case ENUdatatype.ShortDatatype: Return("ShortDatatype")
            Case ENUdatatype.SingleDatatype: Return("SingleDatatype")
            Case ENUdatatype.StringDatatype: Return("StringDatatype")
            Case ENUdatatype.TrueDatatype: Return("TrueDatatype")
            Case ENUdatatype.UnknownDatatype: Return("UnknownDatatype")
            Case ENUdatatype.UnsignedIntegerDatatype: Return("UnsignedIntegerDatatype")
            Case ENUdatatype.UnsignedLongDatatype: Return("UnsignedLongDatatype")
            Case ENUdatatype.UnsignedShortDatatype: Return("UnsignedShortDatatype")
            Case ENUdatatype.UnspecifiedDatatype: Return("UnspecifiedDatatype")
            Case ENUdatatype.UpperDatatype: Return("UpperDatatype")
            Case ENUdatatype.VBIdentifierDatatype: Return("VBIdentifierDatatype")
            Case Else: Return("?")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Data type enumerator to name
    '
    '
    Public Shared Function datatypeName2Enum(ByVal strName As String) As ENUdatatype
        Select Case LCase(Trim(strName))
            Case "alphabeticdatatype": Return(ENUdatatype.AlphabeticDatatype)
            Case "bytedatatype": Return(ENUdatatype.ByteDatatype)
            Case "datedatatype": Return(ENUdatatype.DateDatatype)
            Case "doubledatatype": Return(ENUdatatype.DoubleDatatype)
            Case "falsedatatype": Return(ENUdatatype.FalseDatatype)
            Case "integerdatatype": Return(ENUdatatype.IntegerDatatype)
            Case "labeldatatype": Return(ENUdatatype.LabelDatatype)
            Case "longdatatype": Return(ENUdatatype.LongDatatype)
            Case "lowerdatatype": Return(ENUdatatype.LowerDatatype)
            Case "percentvaluedatatype": Return(ENUdatatype.PercentValueDatatype)
            Case "serialdatatype": Return(ENUdatatype.SerialDatatype)
            Case "shortdatatype": Return(ENUdatatype.ShortDatatype)
            Case "signumdatatype": Return(ENUdatatype.SignumDatatype)
            Case "singledatatype": Return(ENUdatatype.SingleDatatype)
            Case "stringdatatype": Return(ENUdatatype.StringDatatype)
            Case "truedatatype": Return(ENUdatatype.TrueDatatype)
            Case "unknowndatatype": Return(ENUdatatype.UnknownDatatype)
            Case "unsignedintegerdatatype": Return(ENUdatatype.UnsignedIntegerDatatype)
            Case "unsignedlongdatatype": Return(ENUdatatype.UnsignedLongDatatype)
            Case "unsignedshortdatatype": Return(ENUdatatype.UnsignedShortDatatype)
            Case "unspecifieddatatype": Return(ENUdatatype.UnspecifiedDatatype)
            Case "upperdatatype": Return(ENUdatatype.UpperDatatype)
            Case "vbidentifierdatatype": Return(ENUdatatype.VBIdentifierDatatype)
            Case Else: Return(ENUdatatype.UnspecifiedDatatype)
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Dequote a string when it is quoted
    '
    '
    Public Shared Function dequote(ByVal strInstring As String) As String
        If Not isQuoted(strInstring) Then Return(strInstring)
        Return(Replace(Mid(strInstring, 2, Len(strInstring) - 2), Chr(34) & Chr(34), Chr(34)))
    End Function
    
    ' ----------------------------------------------------------------------
    ' Return True if a Windows directory exists: otherwise return False
    '
    '
    Public Shared Function directoryExists(ByVal strDirectory As String) As Boolean
        Dim booDirectoryExists As Boolean = True
        Dim strSaveDirectory As String = CurDir 
        Try
            ChDir(strDirectory)
        Catch
            booDirectoryExists = False
        End Try
        ChDir(strSaveDirectory)
        Return(booDirectoryExists)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert displayable string back to original value
    '
    '
    ' This method converts its strInstring parameter to a string using its
    ' enuSyntax parameter:
    '
    '
    '      *  If enuSyntax is C, the string is assumed to represent 
    '         nondisplayable characters as C escape sequences.  The 
    '         newline may be present as \n and a hex character may be 
    '         present as \xdd where dd is a hex value of an ASCII
    '         character.   
    '
    '      *  If enuSyntax is XML the string is assumed to represent
    '         nondisplayable characters as XML sequences of the form
    '         &#value, where value is a character value.    
    '
    '      *  If enuSyntax is VBexpression the string is assumed to represent
    '         strings as Visual Basic expressions; see string2Display for
    '         the complete VBexpression format.
    '
    '      *  If enuSyntax is VBexpressionCondensed the string is assumed 
    '         to represent strings as condensed Visual Basic expressions; see 
    '         string2Display for the complete VBexpressionCondensed format.
    '             
    '      *  If enuSyntax is Determined or unspecified then the syntax
    '         is determined automatically:
    '
    '         + If the input string contains at least one backslash and
    '           no character pairs "&#" then it is assumed to be in C
    '           syntax.
    '
    '         + If the input string contains at least one character pair
    '           "&#" and no backslash then it is assumed to be in XML
    '           syntax.
    '
    '         + If the input string converts without error from an
    '           uncondensed or condensed Visual Basic expression then
    '           VBexpression or VBexpressionCondensed is assumed.
    '
    '         + If none of the above conditions are true, the input string 
    '           is returned unchanged.
    '
    '         + If more than one of the above conditions are true an error results.
    '
    '
    ' Note that enuSyntax can also be the STRINGS C, XML, VBEXPRESSION, or DETERMINED
    ' (case-insensitive.)
    '
    ' By default, this method expects that character values in C escapes and XML 
    ' sequences will be fixed-length; it will expect that character values in C will 
    ' be 3 digits with left zeroes, and it will expect that values in XML will be 5 
    ' digits.  However, the  optional parameter booVariableLenVals may be True to 
    ' allow variable digit lengths in both C escapes and XML sequences.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     ---------------------------------------
    ' 03 18 02   Nilges         Version 1.0
    '
    ' 01 09 03   Nilges         Bug: failed to translate variable length
    '                           XML
    '
    ' --- Determine the syntax
    Public Overloads Shared Function display2String(ByVal strInstring As String, _
                                                    Optional ByVal booVariableLenVals _
                                                             As Boolean = False) As String
        Return(display2String(strInstring, ENUdisplay2StringSyntax.Determine, booVariableLenVals:=booVariableLenVals))
    End Function
    ' --- Specifies the syntax as a string
    Public Overloads Shared Function display2String(ByVal strInstring As String, _
                                                    ByVal strSyntax As String, _
                                                    Optional ByVal booVariableLenVals _
                                                             As Boolean = False) As String
        Dim enuSyntax As ENUdisplay2StringSyntax = display2String_syntax2Enum_(strSyntax)
        If enuSyntax = ENUdisplay2StringSyntax.Invalid Then Return("")                         
        Return(display2String(strInstring, enuSyntax, booVariableLenVals:=booVariableLenVals))
    End Function                          
    ' --- Common logic                                   
    Public Overloads Shared Function display2String(ByVal strInstring As String, _
                                                    ByVal enuSyntax _
                                                          As ENUdisplay2StringSyntax, _
                                                    Optional ByVal booVariableLenVals _
                                                             As Boolean = False) As String
        Dim booCondensed As Boolean
        Dim booCsyntax As Boolean
        Dim booVBexpressionSyntax As Boolean
        Dim booVBexpressionCondensedSyntax As Boolean
        Dim booXMLsyntax As Boolean
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intIndex3 As Integer
        Dim intIndex4 As Integer
        Dim intLength As Integer
        Dim objCommonRegularExpressions As commonRegularExpressions
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim shAsc As Short
        Dim strEscape As String
        Dim strNextChar As String
        Dim strOutstring As String
        Select Case enuSyntax
            Case ENUdisplay2StringSyntax.C:
                booCsyntax = True
            Case ENUdisplay2StringSyntax.Determine:
                booCsyntax = (Instr(strInstring, "\") <> 0)
                booXMLsyntax = (Instr(strInstring, "&#") <> 0)
                strOutstring = _
                display2String_VBexpression2String_(strInstring, _
                                                    booVBexpressionSyntax, _
                                                    booCondensed) 
                If Not (booCsyntax _
                        AndAlso _
                        Not booXMLsyntax _
                        AndAlso _
                        Not booVBexpressionSyntax _
                        OrElse _
                        Not booCsyntax _
                        AndAlso _
                        booXMLsyntax _
                        AndAlso _
                        Not booVBexpressionSyntax _
                        OrElse _
                        Not booCsyntax _
                        AndAlso _
                        Not booXMLsyntax _
                        AndAlso _
                        booVBexpressionSyntax) Then
                    errorHandler("Cannot determine input string's display syntax", _
                                 "display2String")
                End If
                If booVBexpressionSyntax Then Return(strOutstring)
            Case ENUdisplay2StringSyntax.XML:
                booXMLsyntax = True
            Case ENUdisplay2StringSyntax.VBExpression:
                If strInstring = "" Then Return("")
                strOutstring = _
                display2String_VBexpression2String_(strInstring, _
                                                    booVBexpressionSyntax, _
                                                    booCondensed)
                If Not booVBexpressionSyntax OrElse booCondensed Then
                    errorHandler("Cannot convert string " & vbNewline & vbNewline & _
                                 strInstring & vbNewline & vbNewline & _
                                 "assuming uncondensed VB expression syntax", _
                                 "display2String")
                    Return("")
                End If
                Return(strOutstring)
            Case ENUdisplay2StringSyntax.VBExpressionCondensed:
                If strInstring = "" Then Return("")
                strOutstring = _
                display2String_VBexpression2String_(strInstring, _
                                                    booVBexpressionSyntax, _
                                                    booCondensed)
                If Not booVBexpressionSyntax Then
                    errorHandler("Cannot convert string " & vbNewline & vbNewline & _
                                 strInstring & vbNewline & vbNewline & _
                                 "assuming condensed VB expression syntax")
                    Return("")
                End If
                Return(strOutstring)
            Case Else:
                errorHandler("Unexpected enuSyntax " & enuSyntax)
                Return("")
        End Select
        intIndex1 = 1
        intLength = Len(strInstring)
        Do While intIndex1 <= intLength
            If booCsyntax Then
                strEscape = "\"
            Else
                strEscape = "&#"
            End If
            intIndex2 = InStr(intIndex1, strInstring & strEscape, strEscape)
            objStringBuilder.Append(Mid(strInstring, _
                                        intIndex1, _
                                        intIndex2 - intIndex1))
            If intIndex2 > intLength Then Exit Do
            If booCsyntax Then
                strNextChar = Mid(strInstring, intIndex2 + 1, 1)
                If strNextChar = "\" Then
                    strEscape = strEscape  
                    intIndex1 = intIndex2 + 2
                ElseIf strNextChar = "x" _
                       AndAlso _
                       intLength - intIndex2 >= 3 Then
                    intIndex2 += 2
                    intIndex3 = verify(strInstring, _
                                       asciiCharsetEnum2String(ENUasciiCharset. _
                                                               hexDigits), _
                                       intIndex2) 
                    If intIndex3 = 0 Then intIndex3 = Len(strInstring) + 1
                    intIndex3 = CInt(Math.min(2, intIndex3 - intIndex2))
                    If booVariableLenVals Or intIndex3 = 2 Then
                        strEscape = _
                        ChrW(CInt(baseN2Long(Mid(strInstring & " ", _
                                                 intIndex2, _
                                                 intIndex3), _
                                             asciiCharsetEnum2String _
                                             (ENUasciiCharset.hexDigits), _
                                             booIgnoreCase:=True)))
                    Else
                        strEscape = Mid(strInstring, intIndex2 - 2, intIndex3 + 2)
                    End If
                    intIndex1 = intIndex2 + intIndex3
                ElseIf strNextChar = "n" Then
                    strEscape = vbNewline
                    intIndex1 = intIndex2 + 2
                Else
                    strEscape = Mid(strInstring, intIndex2, 1)
                    intIndex1 += 1
                End If
            Else
                If intLength - intIndex2 >= 2 Then
                    Dim objRegEx As System.Text.RegularExpressions.Regex
                    Dim objRegExMatch As System.Text.RegularExpressions.Match
                    objRegEx = _
                        New System.Text.RegularExpressions.Regex(objCommonRegularExpressions.FindSignedInteger)
                    objRegExMatch = objRegEx.Match(strInstring,  intIndex2)
                    intIndex3 = objRegExMatch.Index + 1
                    If intIndex3 = intIndex2 + 2 Then
                        intIndex4 = Cint(Math.min(objRegExMatch.Length, 5))
                        If booVariableLenVals OrElse intIndex4 = 5 Then
                            strEscape = ChrW(CInt(Mid(strInstring, _
                                                      intIndex3, _
                                                      intIndex4)))
                        Else
                            strEscape = Mid(strInstring, _
                                            intIndex3 - 2, _
                                            objRegExMatch.Length + 2)
                        End If
                        intIndex2 = intIndex3 + intIndex4
                    Else
                        intIndex2 += 2
                    End If
                Else
                    intIndex2 += 2
                End If
                intIndex1 = intIndex2
            End If
            objStringBuilder.Append(strEscape)
        Loop
        Return(objStringBuilder.ToString)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Convert syntax name to enumerator
    '
    '
    Private Shared Function display2String_syntax2Enum_(ByVal strSyntax As String) As ENUdisplay2StringSyntax
        Select Case UCase(Trim(strSyntax))
            Case "C": Return ENUdisplay2StringSyntax.C
            Case "XML": Return ENUdisplay2StringSyntax.XML
            Case "VBEXPRESSION": Return ENUdisplay2StringSyntax.VBExpression
            Case "VBEXPRESSIONCONDENSED": Return ENUdisplay2StringSyntax.VBExpressionCondensed
            Case "DETERMINE": Return ENUdisplay2StringSyntax.Determine
            Case Else: 
                errorHandler("Invalid syntax " & enquote(strSyntax), "utilities", "display2String", _
                             "The valid values are C, XML, VBEXPRESSION, VBEXPRESSIONCONDENSED and DETERMINE: returning " & _
                             "a null string")
                Return ENUdisplay2StringSyntax.Invalid
        End Select        
    End Function    

    ' ----------------------------------------------------------------------
    ' Convert from VB expression format returned by string2Display, on
    ' behalf of display2String
    '
    '
    Private Shared Function display2String_VBexpression2String_(ByVal strInstring _
                                                                      As String, _
                                                                ByRef booValid As Boolean, _
                                                                ByRef booCondensed _
                                                                      As Boolean) As String
        Dim intIndex1 As Integer
        Dim intLength As Integer
        Dim objStringBuilder As New System.Text.StringBuilder  
        Dim objRegEx As New System.Text.RegularExpressions.Regex _
            (Replace("(<quote>[^<quote>]*<quote>)|(ChrW\(([0123456789]+)\))" & _
                     "|" & _
                     "(copies\(((<quote>[^<quote>]*<quote>)" & _
                     "|" & _
                     "(ChrW\(([0123456789]+)\))), ([0123456789]+)\))" & _
                     "|" & _
                     "(range2String\(([0123456789]+), ([0123456789]+)\))|( & )", _
                     "<quote>", Chr(34)))
        Dim strArray() As String
        Dim strNext As String
        Dim objMatchCollection As System.Text.RegularExpressions.MatchCollection
        Dim strToken As String
        booCondensed = False
        objMatchCollection = objRegEx.Matches(strInstring)
        For intIndex1 = 0 To objMatchCollection.Count - 1
            With objMatchCollection.Item(intIndex1)
                strToken = Mid(strInstring, .Index + 1, .Length)
                intLength += Len(strToken)
                If isQuoted(strToken) Then
                    strToken = dequote(strToken)
                ElseIf Mid(strToken, 1, 4) = "ChrW" Then
                    Try
                        strToken = ChrW(CInt(Mid(strInstring, _
                                                 .Groups.Item(3).Index + 1, _
                                                 .Groups.Item(3).Length)))
                    Catch: End Try
                ElseIf Mid(strToken, 1, 6) = "copies" Then
                    strToken = Mid(strInstring, _
                                   .Groups.Item(5).Index + 1, _
                                   .Groups.Item(5).Length)
                    If isQuoted(strToken) Then
                        Try
                            strToken = copies(dequote(strToken), _
                                              CInt(Mid(strInstring, _
                                                   .Groups.Item(9).Index + 1, _
                                                   .Groups.Item(9).Length)))
                        Catch: End Try
                    Else
                        Try
                            strToken = copies(ChrW(CInt(Mid(strInstring, _
                                                            .Groups.Item(8).Index _
                                                            + _
                                                            1, _
                                                            .Groups.Item(8). _
                                                            Length))), _
                                              CInt(Mid(strInstring, _
                                                       .Groups.Item(9).Index + 1, _
                                                       .Groups.Item(9).Length)))
                        Catch: End Try
                    End If
                ElseIf Mid(strToken, 1, 12) = "range2String" Then
                    Try
                        strToken = range2String(Cint(Mid(strInstring, _
                                                     .Groups.Item(11).Index + 1, _
                                                     .Groups.Item(11).Length)), _
                                                CInt(Mid(strInstring, _
                                                         .Groups.Item(12).Index + 1, _
                                                         .Groups.Item(12).Length)))
                    Catch: End Try
                ElseIf strToken = " & " Then
                    strToken = ""
                Else
                    Exit For
                End If
            End With 
            objStringBuilder.Append(strToken)
        Next intIndex1
        booValid = (intIndex1 = objMatchCollection.Count _
                    AndAlso _
                    objMatchCollection.Count <> 0 _ 
                    AndAlso _
                    intLength = Len(strInstring))
        Return objStringBuilder.ToString
    End Function

    ' ----------------------------------------------------------------------
    '
    ' Replaces string by ellipsis
    '
    '
    ' This function is passed a string in strInstring and a Math.max length: it returns the string
    ' contents in a string that does not exceed intMaxLength.  If any characters have to be
    ' hidden to do this, the string returned ends in the strEllipsis value.
    '
    ' The ellipsis can be placed in FRONT of the string.  To do this when the string is overlong,
    ' pass booLeftEllipsis:=True.
    '
    '
    Public Shared Function ellipsis(ByVal strInstring As String, _
                                    ByVal intMaxLength As Integer, _
                                    Optional strEllipsis As String = "...", _
                                    Optional booLeftEllipsis As Boolean = False) _
                    As String
        Dim intEllipsisLength As Integer
        Dim intStringLength As Integer
        intStringLength = Len(strInstring): intEllipsisLength = Len(strEllipsis)
        If intMaxLength < 0 Then
            errorHandler("Internal programming error in call to ellipsis: " & _
                         "intMaxLength " & intMaxLength & " is less than zero")
            Return("")
        End If
        If intEllipsisLength > intMaxLength Then
            errorHandler("Internal programming error in call to ellipsis: " & _
                         "intEllipsisLength " & intEllipsisLength & " is greater than " & _
                         "intMaxLength " & intMaxLength)
            Return("")
        End If
        If intStringLength < intMaxLength Then
            Return(strInstring)
        Else
            If booLeftEllipsis Then
                Return(strEllipsis & Mid$(strInstring, 1, Math.max(0, intMaxLength - intEllipsisLength)))
            Else
                Return(Mid$(strInstring, 1, Math.max(0, intMaxLength - intEllipsisLength)) & strEllipsis)
            End If
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Quote a string
    '
    '
    ' This method returns its input string in double quotes, and it 
    ' guarantees a quoted string with two properties:
    '
    '
    '      *  It is a valid Visual Basic string
    '
    '      *  It converts back to the original string in all cases using
    '         the dequote method:
    '
    '              dequote(enquote(string)) == string
    '
    '
    ' Internal double quotes are replaced by double quote pairs.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 12 04 01     Nilges       1.  Version 1
    ' 10 20 02     Nilges       1.  Double internal double quotes
    '
    '
    Public Shared Function enquote(ByVal strInstring As String, _
                                   Optional ByVal booXML As Boolean = False, _
                                   Optional ByVal strGraphic As String = _
                                            "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                            "abcdefghijklmnopqrstuvwxyz" & _
                                            "0123456789 ") As String
        Return(Chr(34) & Replace(strInstring, Chr(34), Chr(34) & Chr(34)) & Chr(34)) 
    End Function

    ' ----------------------------------------------------------------------
    ' Error handling
    '
    ' 
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 01 09 03   Nilges         Log the error to the Windows application log
    '                           under the name utilities_errorHandler
    
    ' 01 12 03   Nilges         Added the booInfo parameter: when present and
    '                           True, error is eventlogged as info and no
    '                           error is Thrown
    '
    ' --- Message only
    Public Shared Overloads Sub errorHandler(ByVal strMessage As String, _
                                             Optional ByVal booInfo As Boolean = False)
        errorHandler_(strMessage, "", "", "", booInfo)
    End Sub
    ' --- Message and object name only
    Public Shared Overloads Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             Optional ByVal booInfo As Boolean = False)
        errorHandler_(strMessage, strObject, "", "", booInfo)
    End Sub
    ' --- Message, object name and procedure
    Public Shared Overloads Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             ByVal strProcedure As String, _
                                             Optional ByVal booInfo As Boolean = False)
        errorHandler_(strMessage, strObject, strProcedure, "", booInfo)
    End Sub
    ' --- Message, object name, procedure and help info
    Public Shared Overloads Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             ByVal strProcedure As String, _
                                             ByVal strHelp As String, _
                                             Optional ByVal booInfo As Boolean = False)
        errorHandler_(strMessage, strObject, strProcedure, strHelp, booInfo)
    End Sub
    ' --- Common logic
    Private Shared Sub errorHandler_(ByVal strMessage As String, _
                                     ByVal strObject As String, _
                                     ByVal strProcedure As String, _
                                     ByVal strHelp As String, _
                                     ByVal booInfo As Boolean)
        Dim strLocation As String = CStr(Iif(strObject = "" AndAlso strProcedure = "", _
                                             "", _
                                             CStr(Iif(strObject = "", "?", strObject)) & "." & _
                                             CStr(Iif(strProcedure = "", "?", strProcedure))))                                          
        Dim strMessageWork As String = CStr(Iif(booInfo, "Informational message", "Error")) & " " & _
                                       "at " & Now & ": " & strMessage & _
                                       CStr(Iif(strLocation <> "", vbNewline & vbNewline & "Location: " & strLocation, "")) & _
                                       CStr(Iif(strHelp <> "", vbNewline & vbNewline & strHelp, ""))  
        #If WINDOWS_LOGGING Then                                       
            Try                                       
                System.Diagnostics.EventLog.WriteEntry("utilities_errorHandler", strMessageWork)
            Catch
                Dim objExceptionEvt As New Exception("Cannot log the following message:" & _
                                                     vbNewline & vbNewline & _
                                                     strMessageWork & _
                                                     vbNewline & vbNewline & _
                                                     Err.Number & " " & Err.Description)
                Throw objExceptionEvt
                Return
            End Try 
        #End If            
        If booInfo Then Return           
        Dim objException As New Exception(strMessageWork)
        Throw objException
    End Sub

    ' ----------------------------------------------------------------------
    ' File to string  
    '
    '
    ' This method reads all or part of a file.  It has six overloads:
    '
    '
    '      *  file2String(f) reads all of the file f.  The file is
    '         opened as a file stream, read and closed.
    '
    '      *  file2String(f,i) reads all of file f starting at character i
    '         where i is numbered from one.  The file is opened as a 
    '         file stream, read and closed.
    '
    '      *  file2String(f,i,L) reads L characters of file f starting at
    '         character i where i is numbered from 1.  The file is opened
    '         as a file stream, read and closed.
    '
    '      *  file2String(s) reads all of the open file stream s
    '
    '      *  file2String(s,i) reads all of the open file stream f starting 
    '         at character i where i is numbered from one
    '
    '      *  file2String(s,i,L) reads L characters of the open file stream
    '         s starting at character i where i is numbered from 1
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     --------------------------------------------
    ' 11 26 02   Nilges         1.  Added ability to read segments
    '                           2.  Added ability to read from an open file
    '                               stream
    '
    '
    Public Overloads Shared Function file2String(ByVal strFileid As String) As String
        Return(file2String_(strFileid, 1, True, 0))
    End Function
    Public Overloads Shared Function file2String(ByVal strFileid As String, _
                                                 ByVal intStartIndex As Integer) As String
        Return(file2String_(strFileid, intStartIndex, True, 0))                                            
    End Function
    Public Overloads Shared Function file2String(ByVal strFileid As String, _
                                                 ByVal intStartIndex As Integer, _
                                                 ByVal intLength As Integer) As String
        Return(file2String_(strFileid, intStartIndex, False, intLength))
    End Function                                                 
    Public Overloads Shared Function file2String(ByVal objFileStream As IO.FileStream) As String
        Return(file2String_(objFileStream, 1, True, 0))
    End Function
    Public Overloads Shared Function file2String(ByVal objFileStream As IO.FileStream, _
                                                 ByVal intStartIndex As Integer) As String
        Return(file2String_(objFileStream, intStartIndex, True, 0))                                            
    End Function
    Public Overloads Shared Function file2String(ByVal objFileStream As IO.FileStream, _
                                                 ByVal intStartIndex As Integer, _
                                                 ByVal intLength As Integer) As String
        Return(file2String_(objFileStream, intStartIndex, False, intLength))
    End Function                                                 
    Private Shared Function file2String_(ByVal objFile As Object, _
                                         ByVal intStartIndex As Integer, _
                                         ByVal booReadAll As Boolean, _ 
                                         ByVal intLength As Integer) As String
        If intStartIndex < 1 Then
            errorHandler("Start index " & intStartIndex & " is not valid", _
                         "utilities", "file2String", _
                         "Returning null string")
            Return("")                         
        End If                                                 
        Dim intFileLength As Integer  
        Dim strFileid As String
        Dim objStream As IO.FileStream
        If (Typeof objFile Is System.String) Then
            strFileid = CType(objFile, System.String)
        Else
            objStream = CType(objFile, IO.FileStream)
            strFileid = objStream.Name
        End If        
        Try
            intFileLength = CInt(FileLen(strFileid))
        Catch
            errorHandler("Length of file " & _
                        "(" & strFileid & ") " & _
                        intFileLength & " " & _
                        "cannot be handled by file2String utility")
            Return("")
        End Try
        Dim intReadLength As Integer = intFileLength - intStartIndex + 1
        If Not booReadAll Then                                         
            intReadLength = Math.Min(intLength, intReadLength)                                          
        End If          
        If intReadLength = 0 Then Return("")
        If objStream Is Nothing Then
            Try
                Dim objStreamNew As New IO.FileStream(strFileId, IO.FileMode.Open)
                objStream = objStreamNew
            Catch
                errorHandler("Cannot open file " & strFileid & ": " & Err.Number & " " & Err.Description)
                Return("")
            End Try
        End If
        Dim bytArray() As Byte
        Try
            Redim bytArray(intReadLength - 1)
        Catch
            errorHandler("file2String cannot allocate a buffer: " & Err.Number & " " & Err.Description) 
        End Try
        objStream.Read(bytArray, intStartIndex - 1, intReadLength)
        If (Typeof objFile Is System.String) Then objStream.Close
        Dim intIndex1 As Integer
        Dim objStringBuilder As New System.Text.StringBuilder("")
        For intIndex1 = 0 To UBound(bytArray)
            objStringBuilder.Append(ChrW(bytArray(intIndex1)))
        Next intIndex1
        Return objStringBuilder.ToString
    End Function

    ' ----------------------------------------------------------------------
    ' Return True when file exists, False otherwise
    '
    '
    Public Shared Function fileExists(ByVal strFileid As String) As Boolean
        Dim lngFileLen As Long
        lngFileLen = -1
        Try
            lngFileLen = FileLen(strFileid)
        Catch
        End Try
        Return (lngFileLen <> -1)
    End Function
    
    ' -----------------------------------------------------------------------
    ' Find potentially abbreviated string
    '
    '
    ' This method searches the list of blank-delimited words in strKeywords
    ' for the blank-delimited word in strTarget and returns either the
    ' word position (from one) of the found keyword, or zero on
    ' either failure to find strTarget or multiple "hits."
    '
    ' The search completely ignores case.  Leading and trailing spaces are
    ' removed from strTarget for the purpose of searching only.
    '
    ' strTarget may be any unique abbreviation of the keyword consisting
    ' of 1..n characters from the beginning of the keyword.  If all keywords
    ' differ in character one, strTarget can be a one character abbreviation.
    '
    ' For example, findAbbrev("s", "seconds minutes hours") will return
    ' 1.
    '
    ' See also abbrev.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------------
    ' 12 13 02   Nilges         Version 1
    '
    '
    Public Shared Function findAbbrev(ByVal strTarget As String, _
                                      ByVal strKeywords As String) As Integer
        Dim strKeywordSplit() As String
        Try
            strKeywordSplit = split(UCase(strKeywords)) 
        Catch
            errorHandler("Could not split keywords", _
                         "utilities", "findAbbrev", _
                         err.Number & " " & Err.Description)
        End Try        
        Dim strTargetWork As String = Trim(UCase(strTarget))
        If Instr(strTargetWork, " ") <> 0 Then
            errorHandler("Target " & enquote(strTarget) & " contains internal spaces", _
                         "utilities", "findAbbrev", _
                         "Returning zero")
            Return(0)                         
        End If       
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        For intIndex1 = 0 To UBound(strKeywordSplit)
            If abbrev(strTargetWork, strKeywordSplit(intIndex1)) Then 
                If intIndex2 <> 0 Then Return(0)
                intIndex2 = intIndex1 + 1
            End If            
        Next intIndex1         
        Return(intIndex2)
    End Function                               

    ' -----------------------------------------------------------------------
    ' Find item in string
    '
    '
    Public Shared Function findItem(ByVal strInstring As String, _
                                    ByVal strTarget As String, _
                                    ByVal strDelimiter As String, _
                                    ByVal booSetDelimiter As Boolean, _
                                    Optional ByVal booIgnoreCase As Boolean = False) _
           As Integer
        Dim intDelimiterLength As Integer = Len(strDelimiter)
        If intDelimiterLength = 0 Then
            errorHandler("The delimiter can not be a null string")
        End If
        Dim intIndex1 As Integer
        Dim objCompareMethod As CompareMethod = _
            CType(Iif(booIgnoreCase, CompareMethod.Text, CompareMethod.Binary), CompareMethod)                
        Dim strSplitArray() As String
        Dim strTargetWork As String = strTarget
        If booIgnoreCase Then
            strTargetWork = UCase(strTargetWork)
        End If
        ' --- Try two shortcuts to avoid a search
        If Not booSetDelimiter _
           OrElse _
           booSetDelimiter AndAlso Len(strDelimiter) = 1 Then
            ' Shortcut 1: check for common case where string starts with target
            Dim strWork As String = Mid(strInstring, 1, Len(strTarget))
            If booIgnoreCase Then strWork = UCase(strWork)
            If strWork = strTargetWork Then Return(1)
            ' Shortcut 2: use Instr and split.  
            intIndex1 = Instr(strDelimiter & strInstring & strDelimiter, _
                              strDelimiter & strTarget & strDelimiter, _
                              objCompareMethod)
            If intIndex1 = 0 Then Return(0)
            strSplitArray = split(Mid(strInstring, 1, intIndex1 - Len(strDelimiter)), strDelimiter, Compare:=objCompareMethod)
            Return(UBound(strSplitArray) + 1)
        End If
        ' --- Brute-force search  
        Dim strDelimiter1 As String = strDelimiter
        If booSetDelimiter Then strDelimiter1 = Mid(strDelimiter, 1, 1)
        Dim strInstringWork As String = strInstring
        If booIgnoreCase Then strInstringWork = UCase(strInstringWork)
        strSplitArray = split(translate(strInstringWork, _
                                        strDelimiter, _
                                        copies(intDelimiterLength, strDelimiter)), _
                              strDelimiter1, _
                              Compare:=objCompareMethod)
        Dim intNullEntries As Integer
        For intIndex1 = 0 To UBound(strSplitArray)
            If strSplitArray(intIndex1) = "" Then
                intNullEntries += 1
            Else
                If strSplitArray(intIndex1) = strTargetWork Then Return(intIndex1 + 1 - intNullEntries)
            End If
        Next intIndex1
    End Function

    ' -----------------------------------------------------------------------
    ' Find blank-delimited word in string
    '
    '
    Public Shared Function findWord(ByVal strInstring As String, _
                                    ByVal strTarget As String, _
                                    Optional ByVal booIgnoreCase As Boolean = False) _
           As Integer
        Return(findItem(strInstring, strTarget, " ", True, booIgnoreCase:=booIgnoreCase))
    End Function

    ' -----------------------------------------------------------------------
    ' Find the eXtensible Markup Language tag
    '
    '
    ' This method searches for the next XML tag (an expression of the form <[/]name [attributes][/]>)
    ' or the next tag with a specific name.
    '
    ' It has three overloads:
    '
    '
    '      *  findXMLtag(x,t,start,len): searches all of the XML tag x for a tag with
    '         the name t that is not an end tag in the form </t...>.  Returns True on
    '         success, and, on success, places the start index of the tag and its length
    '         in start and in len, which should be Integers, and are passed ByRef.
    '         Returns False on failure and, on failure, leaves start and len unchanged.
    '
    '         If t is a null string, this overload searches for any non-end tag.
    '
    '      *  findXMLtag(x,t,start,len, startIdx, searchLen): searches part of the XML tag
    '         for a tag with the name t that is not an end tag in the form </t...>.  
    '         The search is restricted to the zone, defined by a 1-origin start index startIdx
    '         and a length, searchLen.  Returns True on success, and, on success, places the 
    '         start index of the tag and its length in start and in len, which should be Integers, 
    '         and are passed ByRef.
    '
    '         Returns False on failure and, on failure, leaves start and len unchanged.
    '
    '         If t is a null string, this overload searches for any non-end tag.
    '
    '      *  findXMLtag(x,t,start,len, startIdx, searchLen, True): searches part of the XML tag
    '         for an END tag with the name t.  
    '
    '         The search is restricted to the zone, defined by a 1-origin start index startIdx
    '         and a length, searchLen.  Returns True on success, and, on success, places the 
    '         start index of the tag and its length in start and in len, which should be Integers, 
    '         and are passed ByRef.
    '
    '         Returns False on failure and, on failure, leaves start and len unchanged.
    '
    '         If t is a null string, this overload searches for any non-end tag.
    '
    '
    Public Overloads Shared Function findXMLtag(ByVal strXML As String, _
                                                ByVal strTargetTag As String, _
                                                ByRef intEndTagStartIndex As Integer, _
                                                ByRef intEndTagLength As Integer) As Boolean
        Return(findXMLtag(strXML, _
                          strTargetTag, _
                          intEndTagStartIndex, _
                          intEndTagLength, _
                          1, _
                          Len(strXML), _
                          False))
    End Function                                                
    Public Overloads Shared Function findXMLtag(ByVal strXML As String, _
                                                ByVal strTargetTag As String, _
                                                ByRef intEndTagStartIndex As Integer, _
                                                ByRef intEndTagLength As Integer, _
                                                ByVal intStartIndex As Integer, _
                                                ByVal intLength As Integer) As Boolean
                                                
        Return(findXMLtag(strXML, _
                          strTargetTag, _
                          intEndTagStartIndex, _
                          intEndTagLength, _
                          intStartIndex, _
                          intLength, _
                          False))
    End Function                                                
    Public Overloads Shared Function findXMLtag(ByVal strXML As String, _
                                                ByVal strTargetTag As String, _
                                                ByRef intEndTagStartIndex As Integer, _
                                                ByRef intEndTagLength As Integer, _
                                                ByVal intStartIndex As Integer, _
                                                ByVal intLength As Integer, _
                                                ByVal booFindEndTag As Boolean) As Boolean
        If intStartIndex < 1 Then
            errorHandler("Invalid start index " & intStartIndex, _
                         "utilities", "findXMLtag", _
                         "The start index must not be less than 1.  " & _
                         "False is being returned as a result.")
            Return(False)                              
        End If             
        If intLength < 0 Then
            errorHandler("Invalid length " & intLength, _
                         "utilities", "findXMLtag", _
                         "The length must not be less than 0.  " & _
                         "False is being returned as a result.")
            Return(False)                              
        End If             
        If strTargetTag <> "" AndAlso Not isXMLname(strTargetTag) Then
            errorHandler("Invalid target tag " & enquote(strTargetTag), _
                         "utilities", "findXMLtag", _
                         "This target is not null and it is not a valid XML name.  " & _
                         "False is being returned as a result.")
            Return(False)                              
        End If             
        Dim booOK As Boolean
        Dim intIndex1 As Integer = intStartIndex - 1
        Dim intEndIndex As Integer = intStartIndex + intLength
        Dim strEndTagMark As String = CStr(Iif(booFindEndTag, "/", ""))
        Dim strTargetTagWork As String = UCase(Trim(strTargetTag))
        Dim objRegExp As New System.Text.RegularExpressions.Regex("<" & _
                                                                  strEndTagMark & _
                                                                  strTargetTagWork & _
                                                                  "[\x00-\x20]") 
        Dim objMatch As System.Text.RegularExpressions.Match                                                                  
        Dim strSentinel As String = "<" & strEndTagMark & Cstr(Iif(strTargetTag = "", "a", strTargetTag)) & " "                                                                  
        Dim strTagName As String
        Do
            objMatch = objRegExp.Match(strXML & strSentinel, intIndex1)
            intIndex1 = objMatch.Index + 1
            If intIndex1 <= intEndIndex AndAlso intIndex1 + objMatch.Length <= intEndIndex Then
                If strTargetTagWork <> "" Then
                    booOK = False
                    If parseXMLtag(Mid(strXML, intIndex1, objMatch.Length), strTagName) Then
                        booOK = (UCase(strTagName) = strTargetTag)
                    End If                    
                Else
                    booOK = True                    
                End If           
                If booOK Then     
                    intStartIndex = intIndex1
                    intLength = objMatch.Length
                    Return(True)
                End If                    
            End If                                                                      
        Loop        
        Return(False)
    End Function

    ' ----------------------------------------------------------------------
    ' Return True if object has reference type, False when object is a
    ' value type
    '
    '
    ' We make the needed test by creating a dummy collection, and then
    ' comparing objTest, the object to be checked, to see if it "Is" the
    ' collection:
    '
    '
    '      *  If objTest is a value in the stack, the compare will 
    '         Throw an exception
    '
    '      *  If objTest happens to be a nice collection, the compare will
    '         not Throw up, and it will return Truth
    '
    '      *  If objTest is a reference object but not a collection, the
    '         compare will return False...but not Throw up
    '
    '
    ' This is all rather lame, but it works.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 12 27 01   Nilges         Version 1
    '
    '
    Public Shared Function hasReferenceType(ByVal objTest As Object) As Boolean
        Dim objCollection As Collection
        Try
            objCollection = New Collection
        Catch
            errorHandler("Cannot create tester collection", _
                         "utilities", "hasReferenceType", _
                         Err.Number & " " & Err.Description)
            Return(False)                         
        End Try
        Try
            Dim booDummy As Boolean = (objTest Is objCollection)
        Catch
            Return(False)
        End Try
        Return(True)
    End Function

    ' ----------------------------------------------------------------------
    ' Model data values onto a range of values
    '
    '
    ' This method is useful for plotting and any application where you need
    ' to map a range of numbers onto a larger or smaller set of values, for it
    ' changes a double-precision value to a double-precision number in the
    ' range dblRangeMin..dblRangeMax, where the value can range from 
    ' dblValueMin to dblValueMax.
    '
    ' For example, if you want to plot values with a mininum value of 0 and
    ' a Math.max value of 32767 and you have a label, lblLabel with mininum width
    ' 10 pixels and Math.max width dblMaxWidth you can get the plot size for X:
    '
    '
    '      dblPlotSize = histogram(X, _
    '                              dblRangeMin:=10, _
    '                              dblRangeMax:=32767, _
    '                              dblValueMax:=32767)
    '
    '
    ' This method solves the equation
    '
    '
    '      (value - minValue) / (valueMax - valueMin) = (x - rangeMin) / (rangeMax - rangeMin)
    '
    '
    ' Multiplying both sides by (rangeMax - rangeMin) and adding the rangeMin we obtain the solution
    ' for x, which is the needed range unknown.  I always have to do this bit of high-school algebra
    ' when I plot, so histogram encapsulates this bit of math anxiety.
    '
    '
    '      x = (rangeMax - rangeMin) * ((value - minValue) / (valueMax - valueMin)) + rangeMin
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 12 09 01   Nilges         Version 1
    '
    '
    Public Shared Function histogram(ByVal dblValue As Double, _
                                     Optional ByVal dblRangeMin As Double = 0, _
                                     Optional ByVal dblRangeMax As Double = 32767, _
                                     Optional ByVal dblValueMin As Double = 0, _
                                     Optional ByVal dblValueMax As Double = 32767) As Double
        If dblValueMin = dblValueMax Then Return(0)
        Return(Math.Abs(dblRangeMax - dblRangeMin) _
                    * _
                    ((dblValue - dblValueMin) / Math.Abs(dblValueMax - dblValueMin)) _
                    + _
                    dblRangeMin)
    End Function

    ' ----------------------------------------------------------------------
    ' Indent each line of a string
    '
    '
    ' This function indents every line of a string consisting of 0, 1 or
    ' more lines.
    '
    ' Its overload indent(string) simply indents each line of the string
    ' four spaces, and it assumes that vbNewLine separates lines.  Note
    ' that vbNewLine is the newline appropriate to the environment,
    ' normally carriage return and line feed on Windows, or line feed on
    ' the Web.
    '
    ' Its overload indent(string, indentString, newline) indents each line 
    ' using the indent string and uses newline as the newline character
    ' or string.
    '
    ' Its overload indent(string, count, newline) indents each line 
    ' using a string of blanks equal in length to count, and uses newline 
    ' as the newline character or string.
    '
    ' 
    Public Shared Overloads Function indent(ByVal strInstring As String) As String
        Return(indent(strInstring, "    ", vbNewLine))
    End Function
    Public Shared Overloads Function indent(ByVal strInstring As String, _
                                            ByVal intCount As Integer, _
                                            ByVal strNewLine As String) As String
        If intCount < 0 Then
            errorHandler("Invalid intCount parameter has value " & intCount, "indent")
            Return(strInstring)
        End If
        Return(indent(strInstring, copies(intCount, " "), vbNewLine))
    End Function
    Public Shared Overloads Function indent(ByVal strInstring As String, _
                                            ByVal strIndent As String, _
                                            ByVal strNewLine As String) As String
        Dim intIndex1 As Integer
        Dim strArray() As String
        Try
            strArray = split(strInstring, strNewLine)
        Catch
            With Err
                errorHandler("Not able to parse input string using Split: " & .Number & " " & .Description, _
                             "indent")
            End With
            Return(strInstring)
        End Try
        Dim objStringBuilder As String
        For intIndex1 = 0 To UBound(strArray)
            append(objStringBuilder, strNewLine, strIndent & strArray(intIndex1))
        Next intIndex1
        Return(objStringBuilder.ToString)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Inspection and test append
    '
    '
    ' This method supports various inspect and test methods, found in EGNSF
    ' objects, which implement the corresponding methods of the Iegnsf interface.
    '
    ' strReport should be the inspection or test report.
    '
    ' strRule should be the inspection rule or test description.
    '
    ' booRuleResult should be the effect of applying the rule or making the
    ' test.
    '
    ' booResult should be a flag indicating the global success of the 
    ' inspection or test: booResult will be Anded with booRuleResult, and
    ' its value after the And will be returned as the function value.
    ' booResult should be declared with an initial value of True.
    '
    ' The optional overload parameter strComments may contain additional 
    ' test comments.
    '
    ' Note that this method is not a true utility in the sense that it is
    ' for EGNSF praxis.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 01 20 03     Nilges       1.  Version 1
    '
    '
    Public Overloads Shared Function inspectionAppend(ByRef strReport As String, _
                                                      ByVal strRule As String, _
                                                      ByVal booRuleResult As Boolean, _
                                                      ByRef booResult As Boolean) As Boolean
        Return(inspectionAppend(strReport, strRule, booRuleResult, booResult, ""))
    End Function                                                      
    Public Overloads Shared Function inspectionAppend(ByRef strReport As String, _
                                                      ByVal strRule As String, _
                                                      ByVal booRuleResult As Boolean, _
                                                      ByRef booResult As Boolean, _
                                                      ByVal strComments As String) As Boolean
        strReport =  append(strReport, _
                            vbNewline, _
                            strRule & ": " & CStr(Iif(booRuleResult, "OK", "Failed")) & _
                            CStr(Iif(strComments <> "", vbNewline, "")) & _
                            strComments)
        booResult = booResult AndAlso booRuleResult
        Return(booResult)                            
    End Function 
    
    ' ----------------------------------------------------------------------
    ' Convert a (positive) integer to the number of digits it containeth
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 12 01 01     Nilges       1.  Version 1
    '
    '
    Public Shared Function int2Digits(ByVal intValue As Integer) As Integer
        If intValue = 0 Then Return(1)
        Return(CInt(Math.Floor(Math.Log10(intValue))) + 1)
    End Function 

    ' ----------------------------------------------------------------------
    ' Return True (string is quoted using double quotes) or False
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 12 04 01     Nilges       1.  Version 1
    '
    '
    Public Shared Function isQuoted(ByVal strInstring As String) As Boolean
        Dim intLength As Integer = Len(strInstring)
        If intLength < 2 Then Return(False)
        Return(Mid(strInstring, 1, 1) = Chr(34) AndAlso Mid(strInstring, intLength) = Chr(34))             
    End Function
    

    ' ----------------------------------------------------------------------
    ' Determine whether string is an XML comment (W3C 1998)
    '
    '
    ' This method has four overloads:
    '
    '
    '      isXMLcomment(string): returns True if string is in the comment
    '           format <!-- anything -->, False otherwise
    '
    '      isXMLcomment(string, complete): "complete" is a Boolean, passed
    '           by reference.  This overload returns True as long as string
    '           starts with <!--.  It places True in the complete Boolean
    '           if the string ends with -->, False otherwise.
    '
    '      isXMLcomment(string, comment): "comment" is a String, passed
    '           by reference.  This overload returns True as long as string
    '           is a full comment with starting and ending symbol.  It also
    '           places the comment text (without the comment characters or
    '           leading/trailing blanks) in the comment reference variable.
    '
    '      isXMLcomment(string, complete, comment): complete is set to
    '           a flag indicating complete comment, and comment is set to
    '           the actual comment.  Combines the above two overloads.
    '
    '
    ' Note that the input string is trimmed, removing starting and ending
    ' spaces, before further tests.
    '
    '
    Public Shared Overloads Function isXMLcomment(ByVal strInstring As String) As Boolean
        Dim booComplete As Boolean
        Dim strCommentText As String
        Return isXMLcomment(strInstring, booComplete, strCommentText) And booComplete
    End Function
    Public Shared Overloads Function isXMLcomment(ByVal strInstring As String, _
                                                  ByRef booComplete As Boolean) As Boolean
        Dim strCommentText As String
        Return isXMLcomment(strInstring, booComplete, strCommentText)
    End Function
    Public Shared Overloads Function isXMLcomment(ByVal strInstring As String, _
                                                  ByRef strCommentText As String) As Boolean
        Dim booComplete As Boolean
        Return isXMLcomment(strInstring, booComplete, strCommentText) And booComplete
    End Function
    Public Shared Overloads Function isXMLcomment(ByVal strInstring As String, _
                                                  ByRef booComplete As Boolean, _
                                                  ByRef strCommentText As String) As Boolean
        Dim strInstringWork As String = Trim(strInstring)
        Dim intLength As Integer = Len(strInstringWork)
        booComplete = False: strCommentText = ""
        If intLength < 4 Then Return(False)
        If Mid(strInstringWork, 1, 4) <> "<!--" Then Return(False)
        booComplete = (Mid(strInstringWork, intLength - 2) = "-->")
        strCommentText = Trim(Mid(strInstringWork, 5, intLength - 3))
        Return(True)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Return True when input string is an XML name
    '
    '
    ' Cf. W3C 1998: this method allows strInstring to start with a
    ' letter, an underscore or a colon.  It allows the rest of the
    ' input string to contain letters, digits, periods, dashes, 
    ' underscores and colons.  This method does NOT support the use
    ' of combinedChars or Extenders.
    '
    '
    Public Shared Function isXMLname(ByVal strInstring As String) As Boolean
        Dim strInstringUCase As String = UCase(strInstring)
        If Instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ_:", strInstringUCase) = 0 Then
            Return(False)
        End If            
        Return(verify(strInstringUCase, _
                      "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_:") _
               = _
               0)                      
    End Function    
    
    ' ----------------------------------------------------------------------
    ' Return the nth delimited item from strInstring
    '
    '
    Public Shared Overloads Function item(ByVal strInstring As String, _
                                          ByVal intIndex As Integer) _
           As String
        Return item(strInstring, intIndex, " ", True)
    End Function
    Public Shared Overloads Function item(ByVal strInstring As String, _
                                          ByVal intIndex As Integer, _
                                          ByVal strDelimiter As String, _
                                          ByVal booSetDelimiter As Boolean) _
           As String
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intUBound As Integer
        Dim strArray() As String
        Dim strItem As String
        If intIndex < 1 Then
            MsgBox("intIndex " & intIndex & " is not valid")
            Return("")
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
                Do While intIndex1 <= Len(strInstring)
                    intIndex1 = verify(strInstring, _
                                    strDelimiter, _
                                    intStartIndex:=intIndex1)
                    If intIndex1 = 0 Then Exit Do
                    strItem = Mid(strInstring & strDelimiter, _
                                  intIndex1, _
                                  verify(strInstring, _
                                         strDelimiter, _
                                         intIndex1 + 1, _
                                         True))
                    If intIndex = intCount Then Exit Do
                Loop
            End If
        Else
            ' Use Split to find item
            If item_string2Array_(strInstring, strDelimiter, strArray) Then
                If intIndex <= UBound(strArray) + 1 Then strItem = strArray(intIndex - 1)
            End If
        End If
        Return(strItem)
    End Function

    ' ----------------------------------------------------------------------
    ' Parse the input string of item using Split: return True on success 
    ' or False on failure
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 11 23 01     Nilges       1.  Bug: incorrect success test 
    '
    '
    Private Shared Function item_string2Array_(ByVal strInstring As String, _
                                                ByVal strDelimiter As String, _
                                                ByRef strArray() As String) As Boolean
        Dim intUBound As Integer
        strArray = Split(strInstring, strDelimiter)
        intUBound = 0
        Try
            intUBound = UBound(strArray)
        Catch
        End Try
        Return(intUBound <> -1)
    End Function

    ' ----------------------------------------------------------------------
    ' Return the normalized phrase of adjacent delimited items from 
    ' strInstring
    '
    '
    ' The phrase is normalized because exactly one delimiter appears
    ' between each pair of items, and no delimiters appear at the
    ' beginning or end of the string.  
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------------
    ' 01 23 03   Nilges         1.  Allow an itemCount of -1 to represent the 
    '                               end of the string
    '
    '                           2.  Create the string builder in an error
    '                               handler
    '
    '
    Public Shared Overloads Function itemPhrase(ByVal strInstring As String, _
                                                ByVal intStartIndex As Integer, _
                                                ByVal intItemCount As Integer, _
                                                ByVal strDelimiter As String, _
                                                ByVal booSetDelimiter As Boolean) _
                     As String
        Dim intIndex1 As Integer
        If intStartIndex < 1 Then
            MsgBox("Invalid intStartIndex parameter: " & intStartIndex)
        End If  
        Dim intEndIndex As Integer
        Dim intItems As Integer = items(strInstring, strDelimiter, booSetDelimiter)       
        If intItemCount < -1 Then
            MsgBox("Invalid intItemCount parameter: " & intStartIndex)
        ElseIf intItemCount = -1 Then
            intEndIndex = intItems
        Else
            intEndIndex = Math.Min(intItems, intStartIndex + intItemCount - 1)            
        End If        
        Dim objStringBuilder As System.Text.StringBuilder 
        Try
            objStringBuilder = New System.Text.StringBuilder("", Len(strInstring))
        Catch        
            errorHandler("Can't create string builder", _
                         "utilities", "itemPhrase", _
                         Err.Number & " " & Err.Description & vbNewline & vbNewline & _
                         "Returning a null item phrase")
            Return("")                               
        End Try        
        With objStringBuilder
            For intIndex1 = intStartIndex to intEndIndex
                append(objStringBuilder, _
                       strDelimiter, _
                       item(strInstring, intIndex1, strDelimiter, booSetDelimiter))
            Next intIndex1
            Return(objStringBuilder.ToString)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Return the count of delimited items from strInstring
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 12 13 01   Nilges         Improved efficiency: use split in all cases,
    '                           counting nulls as needed
    '
    '
    Public Shared Overloads Function items(ByVal strInstring As String) _
           As Integer
        Return(items(strInstring, " ", True))
    End Function
    Public Shared Overloads Function items(ByVal strInstring As String, _
                                           ByVal strDelimiter As String, _
                                           ByVal booSetDelimiter As Boolean) _
                     As Integer
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intItems As Integer
        Dim intLBound As Integer
        Dim strArray() As String
        ' Use Split to find raw item count
        strArray = Split(strInstring, strDelimiter)
        intCount = 0
        Try
            intCount = UBound(strArray) + 1
        Catch
        End Try
        If booSetDelimiter Then
            ' Don't count null items
            For intIndex1 = 0 To intCount - 1
                If strArray(intIndex1) = "" Then intCount -= 1
            Next intIndex1
        End If
        Return(intCount)
    End Function

    ' ----------------------------------------------------------------------
    ' Join two multiline lists: intersperse lines
    '
    '
    ' This method creates one string out of strInstring1 and strInstring2.
    ' Each line of strInstring1 is joined to the corresponding line of
    ' strInstring2 and separated by a divider, known as the "gutter", which
    ' can be specified in the optional strGutter parameter and which defaults
    ' to three characters.
    '
    ' Each line of strInstring1 (the string of left column values is padded to
    ' the maximum line width.
    '
    ' If the input strings are of unequal length missing lines are blank.
    '
    '
    Public Shared Function joinlines(ByVal strInstring1 As String, _
                                     ByVal strInstring2 As String, _
                                     Optional ByVal strGutter As String = "   ") As String
        Dim intIndex1 As Integer
        Dim intMaxLength As Integer
        Dim intNextLength As Integer
        Dim strCol1 As String
        Dim strCol2 As String
        Dim strSplit1() As String = split(strInstring1, vbNewline)
        Dim strSplit2() As String = split(strInstring2, vbNewline)
        ' --- Find maximum left column length
        For intIndex1 = 1 To UBound(strSplit1)
            intNextlength = Len(strSplit1(intIndex1))
            If intNextLength > intMaxLength Then intMaxLength = intNextLength
        Next intIndex1
        ' --- Create return value
        Dim objStringBuilder As New System.Text.StringBuilder
        For intIndex1 = 0 To Math.max(UBound(strSplit1), UBound(strSplit2))
            If intIndex1 <= UBound(strSplit1) Then
                strCol1 = strSplit1(intIndex1)
            Else
                strCol1 = copies(intMaxLength, " ")
            End If
            If intIndex1 <= UBound(strSplit2) Then
                strCol2 = strSplit2(intIndex1)
            Else
                strCol2 = copies(intMaxLength, " ")
            End If
            append(objStringBuilder, vbCrLf, strCol1 & strGutter & strCol2)
        Next intIndex1
        Return(objStringBuilder.ToString)
    End Function
    
    ' ---------------------------------------------------------------------
    ' Replace keywords by values
    '
    '
    Public Shared Function keywordChange(ByVal strInstring As String, _
                                         ByVal strControlString As String, _
                                         ParamArray strSubstitution() As String) As String
        Dim intIndex1 As Integer
        Dim strInstringWork As String
        Dim strNextKeyword As String
        strInstringWork = strInstring
        For intIndex1 = LBound(strSubstitution) To UBound(strSubstitution) Step 2
            If Mid$(strSubstitution(intIndex1), 1, Len(strControlString)) <> strControlString Then
                strNextKeyword = strControlString & strNextKeyword
            End If
            strInstringWork = Replace(strInstringWork, strNextKeyword, strSubstitution(intIndex1 + 1)) 
        Next intIndex1
        On Error GoTo 0
        keywordChange = strInstringWork
    End Function

    ' ----------------------------------------------------------------------
    '
    ' Convert Long number to base N
    '
    '
    ' This method converts a Long integer in lngBase10 to its representation 
    ' that uses the N distinct characters in strBaseN.  The first character 
    ' in strBaseN represents zero: the second represents one, and so on.
    '
    ' This method has the following overloaded syntax:
    '
    '
    '      long2BaseN(number, N): converts number to its string base N 
    '           representation using the specified base: N must be in the
    '           range 2..36, and, it is assumed that the symbols in use will be
    '           "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    '
    '      long2BaseN(number, N, wordsize): N must be in the range 2..36, and, 
    '           it is assumed that the symbols in use will be
    '           "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".  wordsize is the maximum length
    '           of the base N number.  If the result is too long to fit in 
    '           wordsize then a null string is returned.  
    '
    '      long2BaseN(number, digits): converts number to its string base N 
    '           representation using the specified set of digits to define
    '           the base.
    '
    '           The first digit represents 0, the second digit, 1, and so on,
    '           and the length of digits is the base N.
    '
    '      long2BaseN(number, digits, wordsize): wordsize is the maximum length
    '           of the base N number.  If the result is too long to fit in 
    '           wordsize then a null string is returned.
    '
    '
    ' C H A N G E   R E C O R D ----------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     ----------------------------------------
    ' 01 07 03   Nilges         Added overloads to allow the base to be
    '                           specified as a number, with conventional
    '                           digits
    '
    '
    Public Shared Overloads Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal intBase As Integer) As String
        Return long2BaseN(lngBase10, intBase, 0)
    End Function
    Public Shared Overloads Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal intBase As Integer, _
                                                ByVal bytWordSize As Byte) As String
        Return(long2BaseN(lngBase10, _
                          baseN2Long_base2Digits_(intBase), _
                          bytWordSize))                                   
    End Function                                                
    Public Shared Overloads Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal strDigits As String) As String
        Return long2BaseN(lngBase10, strDigits, 0)
    End Function
    Public Shared Overloads Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal strDigits As String, _
                                                ByVal bytWordSize As Byte) As String
        Dim intBaseValue As Integer
        Dim lngBase10Work As Long
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strBaseN As String
        If Len(strDigits) < 2 Then
            errorHandler("length of strDigits is less than 2")
            Return("")
        End If
        If lngBase10 = 0 Then
            objStringBuilder.Append(Mid(strDigits, 1, 1))
        Else
            intBaseValue = Len(strDigits)
            If bytWordSize <> 0 Then
                If intBaseValue <> 2 Then
                    errorHandler("Word size is specified but base is not 2")
                    Return(Mid(strDigits, 1, 1))
                End If
                If lngBase10 > 2^(bytWordSize - 1) - 1 _
                   OrElse _
                   lngBase10 < -2^(bytWordSize - 1) Then
                    errorHandler("Long value " & lngBase10 & " " & _
                                 "cannot be represented with word size " & _
                                 bytWordSize)
                    Return(Mid(strDigits, 1, 1))
                End If
            End If
            If lngBase10 < 0 Then
                If intBaseValue <> 2 OrElse bytWordSize = 0 Then
                    errorHandler("Input number " & lngBase10 & " is negative, " & _
                                 "but base is not 2 or word size is not specified")
                    Return("")
                End If
                lngBase10Work = Math.Abs(lngBase10) - 1 
            Else
                lngBase10Work = lngBase10 
            End If
            Do While lngBase10Work <> 0
                append(objStringBuilder, _
                       "", _
                       Mid(strDigits, _
                           CInt(lngBase10Work Mod intBaseValue + 1), _
                           1), _
                       booToStart:=True)
                lngBase10Work = lngBase10Work \ intBaseValue
            Loop
        End If
        strBaseN = objStringBuilder.ToString
        If bytWordSize <> 0 Then
            If objStringBuilder.Length <= bytWordSize Then
                strBaseN = align(strBaseN, _
                                 bytWordSize, _
                                 ENUalign.alignRight, _
                                 Mid(strDigits, 1, 1))
            Else
                errorHandler("Overflow: input number " & lngBase10 & " " & _
                             "cannot be represented in " & _
                             bytWordSize & " digits")
            End If
        End If
        If lngBase10 < 0 Then strBaseN = translate(strBaseN, _
                                                   strDigits, _
                                                   Mid(strDigits, 2, 1) & _
                                                   Mid(strDigits, 1, 1))
        Return(strBaseN)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Make a temporary file
    '
    '
    ' This method returns a filetitle of the form tmpn in the optional
    ' path, which defaults to c:\temp.  In tmpn, n is a sequence number.
    ' The file id returned is guaranteed not to exist.
    '
    ' If the path does not exist it is created.  
    '
    '
    Public Shared Overloads Function mkTempFile() As String
        Return(mkTempFile("C:\TEMP"))
    End Function    
    Public Shared Overloads Function mkTempFile(ByVal strPath As String) As String
        Try
            MkDir(strPath)
        Catch: End Try        
        Dim intIndex1 As Integer
        Dim strNext As String
        Do
            strNext = appendPath(strPath, "tmp" & intIndex1)
            If Not FileExists(strNext) Then
                Return(strNext)
            End If            
        Loop        
    End Function    
    
    
    ' ----------------------------------------------------------------------
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
    Public Shared Function mkXMLComment(ByVal strComment As String, _
                                        Optional ByVal booMultipleLineEdit As Boolean = False) As String
        If Instr(strComment, vbNewLine) <> 0 AndAlso booMultipleLineEdit Then
            Dim intIndex1 As Integer
            Dim objStringBuilder As New System.Text.StringBuilder("")
            Dim strArray() As String
            strArray = split(strComment, vbNewLine)
            For intIndex1 = LBound(strArray) To UBound(strArray)
                append(objStringBuilder, vbNewLine, mkXMLcomment(strArray(intIndex1))) 
            Next intIndex1
            Return(objStringBuilder.ToString)
        End If
        Return("<!-- " & strComment & " -->")
    End Function

    ' ----------------------------------------------------------------------
    ' Make an XML element
    '
    '
    ' This method, for a tag name and value creates the tagged XML
    ' element including a start tag, end tag and value.
    '
    ' This method has the following overloaded syntax:
    '
    ' 
    '      *  mkXMLelement(tag, value): a one-line XML element
    '         with a start tag, the value as passed, and an end
    '         tag.  Note that the value can consist of inner XML tags.  
    '         To force it to consist of data, use the string2Display method: 
    '         mkXMLelement(tag, string2Display(value)) will ensure that the value
    '         is treated as data.
    ' 
    '      *  mkXMLelement(tag, value, lineLen): a multiple-line
    '         XML element is returned.  A line break appears
    '         between the start tag, the data, and the end tag.
    '         In addition, an "actual" line break (not metacharacters)
    '         appears between lines such that each line's length
    '         does not exceed lineLen.
    '      
    '      
    ' Also, in either overload, an array of tag attributes for the start tag
    ' may be passed in the form
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
    '
    ' 
    Public Shared Overloads Function mkXMLElement(ByVal strTag As String, _
                                                  ByVal strValue As String, _
                                                  ParamArray strAttributes() As String) As String
        Return(mkXMLElement(strTag, strValue, 0))
    End Function
    Public Shared Overloads Function mkXMLElement(ByVal strTag As String, _
                                                  ByVal strValue As String, _
                                                  ByVal intLineLen As Integer, _
                                                  ParamArray strAttributes() As String) As String
        If intLineLen < 0 Then
            errorHandler("Internal programming error: intLineLen parameter " & intLineLen & " " & _
                            "is not valid", _
                            "mkXMLElement", "", _
                            "This parameter must be zero or positive.  A null string has been returned.")
            Return("")
        End If
        Dim intIndex1 As Integer
        Dim intLineLenWork As Integer
        Dim objStringBuilder As System.Text.StringBuilder
        Try 
            objStringBuilder = New System.Text.StringBuilder(mkXMLTag(strTag, strAttributes))
        Catch
            errorHandler("Can't create string builder", _
                         "utilities", "mkXMLElement", _
                         Err.Number & Err.Description)
        End Try            
        Dim strNewline As String = CStr(Iif(intLineLen = 0, "", vbNewline))
        With objStringBuilder
            intLineLenWork = CInt(Iif(intLineLen = 0, Len(strValue), intLineLen))
            For intIndex1 = 1 To Len(strValue) Step intLineLenWork
                objStringBuilder.Append(strNewLine & Mid(strValue, intIndex1, intLineLenWork))
            Next intIndex1 
            objStringBuilder.Append(strNewLine & mkXMLTag(strTag, booEndTag:=True))
            Return(objStringBuilder.ToString)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Make an XML tag
    '
    '
    ' This method, for a tag name, returns <tagName> or (when the optional
    ' parameter booEndTag is present and True) </tagName>.
    '
    ' When called using the overloaded syntax mkXMLTag(name, attr), the
    ' parameter attr should be in the format:
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
    '
    ' C H A N G E   R E C O R D --------------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------------
    ' 10 21 02     Nilges       Added attributes parameter array: made booEndTag
    '                           an overloaded parameter rather than an Optional
    '                           parameter
    '
    '
    Public Overloads Shared Function mkXMLTag(ByVal strTagName As String, _
                                              ByVal booEndTag As Boolean) As String

        Return("<" & CStr(Iif(booEndTag, "/", "")) & strTagName & ">")
    End Function                                              
    Public Overloads Shared Function mkXMLTag(ByVal strTagName As String, _
                                              ParamArray strAttributes() As String) As String
        Dim strTag As String = mkXMLTag(strTagName, False)     
        Dim strAttributesWork As String = trim(mkXMLTag_attributes2String_(strAttributes))
        If strAttributesWork <> "" Then strAttributesWork = " " & strAttributesWork
        Return Mid(strTag, 1, Len(strTag) - 1) & _
               strAttributesWork & _
               ">"
    End Function
    
    ' ----------------------------------------------------------------------
    ' Convert attribute parameter array to string on behalf of mkXMLtag
    '
    '
    Public Shared Function mkXMLTag_attributes2String_(ByRef strAttributes() As String) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler("Can't create string builder", _
                         "utilities", "mkXMLTag_attributes2String_", _
                         Err.Number & Err.Description)
        End Try        
        For intIndex1 = LBound(strAttributes) To UBound(strAttributes) Step 2
            If Not append(objStringBuilder, " ", _
                          strAttributes(intIndex1) & _
                          CStr(Iif(strAttributes(intIndex2) = "", "", "=" & strAttributes(intIndex2)))) Then
                errorHandler("Can't extend string builder", _
                             "utilities", "mkXMLTag_attributes2String_", _
                             Err.Number & Err.Description)
            End If                          
        Next intIndex1       
        Return(objStringBuilder.ToString) 
    End Function        

    ' ----------------------------------------------------------------------
    ' Object to string
    '
    '
    ' This method returns one of the following values, applying the following
    ' rules in the order listed.  If the optional overloaded parameter
    ' booDeco is present and True, the value returned is included in an
    ' expression of the form <type>(<value>) where <type> is the type of the
    ' object, as returned by GetType.ToString.
    '
    '
    '     *  If the object's value is numeric, this value is returned
    '
    '     *  If the object converts without error to a non-numeric string,
    '        then the quoted string value is returned.
    '
    '     *  If the object exposes a usable ToString method that returns
    '        a System.String, or a value that converts to a System.String,
    '        then this value is returned.
    '
    '     *  If the object is Nothing the string Nothing is returned.
    '
    '     *  In all other cases Object is returned (the string Object.)
    '
    '
    ' Note that the string2Object method is available to convert value-type objects
    ' from the string returned by object2String back to their original value.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 10 23 01     Nilges       1.  Use Name property
    '                           2.  Always decorate objects (when booDeco is
    '                               True.)
    '
    ' 11 13 01     Nilges       1.  Use single quotes around single characters
    '
    ' 12 04 01     Nilges       1.  Pasted in from utility.vb in main utility
    '                               folder: enquote simplified, ability to 
    '                               retrieve Name taken out
    '
    '
    Public Shared Overloads Function object2String(ByVal objObject As Object) As String
        Return(object2String(objObject, False))
    End Function
    Public Shared Overloads Function object2String(ByVal objObject As Object, _
                                                   ByVal booDeco As Boolean) As String
        Dim strObject As String
        strObject = ""
        Try
            strObject = CStr(objObject)
            If booDeco Then
                If Not IsNumeric(strObject) Then 
                    strObject = string2Display(strObject, _
                                               ENUdisplay2StringSyntax.VBExpressionCondensed, _
                                               booBlankConvert:=True, _
                                               booQuoteConvert:=True)
                End If                                                    
                strObject = objObject.GetType.ToString & "(" & strObject & ")"
            Else                                                    
                If Not isNumeric(strObject) Then strObject = enquote(strObject)
            End If                
        Catch
            Try
                strObject = CStr(objObject.ToString)
            Catch
                If (objObject Is Nothing) Then
                    strObject = "Nothing"
                Else
                    strObject = "Object"
                End If
            End Try
        End Try
        Return(strObject)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Parse the XML tag
    '
    '
    ' This method returns the constituents of an XML tag: name, attributes,
    ' whether the tag is an end tag like </a> and whether the tag is a
    ' self-contained tag like <a/>.
    '
    ' Although this method returns False when the tag cannot be parsed, a
    ' True indicator does NOT mean that the tag is valid.
    '
    ' This method has the following overloads.
    '
    '
    '      *  parseXMLtag(t,n) retrieves only the tagname n from the tag t
    '
    '      *  parseXMLtag(t,n,a) retrieves the tagname n and its atributes a
    '         from the tag t
    '
    '      *  parseXMLtag(t,n,a,e) retrieves the tagname n and its atributes a
    '         from the tag t, and sets e to True when t is an end tag like </a>
    '
    '      *  parseXMLtag(t,n,a,e,s) retrieves the tagname n and its atributes a
    '         from the tag t, sets e to True when t is an end tag like </a>, and
    '         sets s when t is a self-contained tag like <a/>
    '
    '
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String) As Boolean
        Dim strTagAttributes As String
        Return(parseXMLtag(strTag, strTagName, strTagAttributes))
    End Function                                
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String, _
                                                 ByRef strTagAttributes As String) As Boolean
        Dim booEndTag As Boolean
        Return(parseXMLtag(strTag, strTagName, strTagAttributes, booEndTag))
    End Function                                
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String, _
                                                 ByRef strTagAttributes As String, _
                                                 ByRef booEndTag As Boolean) As Boolean
        Dim booSelfContainedTag As Boolean
        Return(parseXMLtag(strTag, strTagName, strTagAttributes, booEndTag, booSelfContainedTag))
    End Function                                
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String, _
                                                 ByRef strTagAttributes As String, _
                                                 ByRef booEndTag As Boolean, _
                                                 ByRef booSelfContainedTag As Boolean) As Boolean
        If Len(strTag) < 3 Then Return(False)                                                 
        Dim intIndex1 As Integer = verify(strTag & " ", " >", 2, True)                                                 
        strTagName = Mid(strTag, 2, intIndex1 - 2)
        intIndex1 = verify(strTag & ">", " ", intIndex1)
        Dim intIndex2 As Integer = verify(strTag & ".", ">/", intIndex1 + 1, True)
        strTagAttributes = Mid(strTag, intIndex1, intIndex2 - intIndex1)
        booEndTag = (Mid(strTag, intIndex2, 1) = "/")
        Return(True)
    End Function                                
    
    ' ----------------------------------------------------------------------
    ' Return a phrase of blank-delimited words
    '
    '
    Public Shared Overloads Function phrase(ByVal strInstring As String, _
                                            ByVal intStartIndex As Integer) As String
        Return(phrase(strInstring, intStartIndex, words(strInstring)))
    End Function
    Public Shared Overloads Function phrase(ByVal strInstring As String, _
                                            ByVal intStartIndex As Integer, _
                                            ByVal intWordCount As Integer) As String
        Return(itemPhrase(strInstring, intStartIndex, intWordCount, " ", True))
    End Function
    
    ' ----------------------------------------------------------------------
    ' Indicate whether progress reporting is available
    '
    '
    ' This method returns True to indicate whether the utilities object
    ' was compiled with the UTILITIES_PROGRESS compile-time symbol set to
    ' True.
    '
    '
    Public Shared Function progress As Boolean
        #If UTILITIES_PROGRESS Then
            Return(True)
        #End If        
        Return(False)
    End Function    

    ' ----------------------------------------------------------------------
    ' Character range to string
    '
    '
    ' This method returns the string of characters, starting with intStart
    ' and ending with intEnd.  Note that the character values can lie
    ' outside the ASCII character range.
    '
    ' The three overloads of this method allow you to specify the characters
    ' as characters, one-character strings, or as integers, but unfortunately
    ' and for no good reason you cannot mix specification modes.
    '
    ' If intStart is less than intEnd, then an ascending range is returned.
    ' If intStart is greater than intEnd, then a descending range is returned.
    '
    '
    Public Shared Overloads Function range2String(ByVal strStart As String, _
                                                  ByVal strEnd As String) As String
        If Len(strStart) <> 1 OrElse Len(strEnd) <> 1 Then
            errorHandler("Multiple-character or null strings cannot be converted", "range2String")
        End If
        Return(range2String(AscW(strStart), AscW(strEnd)))
    End Function
    Public Shared Overloads Function range2String(ByVal chrStart As Char, _
                                                  ByVal chrEnd As Char) As String
        Return(range2String(AscW(chrStart), AscW(chrEnd)))
    End Function
    Public Shared Overloads Function range2String(ByVal intStart As Integer, _
                                                  ByVal intEnd As Integer) As String
        Dim intIndex1 As Integer
        Dim objStringBuilder As New System.Text.StringBuilder("")
        For intIndex1 = intStart To intEnd Step CInt(Iif(intStart <= intEnd, 1, -1))
            objStringBuilder.Append(ChrW(intIndex1)) 
        Next intIndex1   
        Return(objStringBuilder.ToString) 
    End Function

    ' ----------------------------------------------------------------------
    ' Replace the XML metacharacters with &#nnnnn
    '
    '
    Public Shared Overloads Function replaceXMLmetaChars(ByVal strValue As String) As String
        Dim strComments As String
        Return(replaceXMLmetaChars(strValue, strComments))
    End Function                                                
    Public Shared Overloads Function replaceXMLmetaChars(ByVal strValue As String, _
                                                  ByRef strComments As String) As String
        Return(replaceMetaChar_ _
                    (replaceMetaChar_ _
                    (replaceMetaChar_(strValue, "&", strComments), _
                        "<", _
                        strComments), _
                    ">", _
                    strComments))                         
    End Function    
    
    ' ----------------------------------------------------------------------
    ' Replace XML meta character...update comments
    '
    '
    Private Shared Function replaceMetaChar_(ByVal strValue As String, _
                                             ByVal strMetaChar As String, _
                                             ByRef strComments As String) As String
        If Instr(strValue, strMetaChar) = 0 Then Return(strValue)
        Dim strXML As String = "&#" & alignRight(CStr(AscW(strMetaChar)), 5, "0") & ";"
        Dim strComment As String = "Replaced meta character " & _
                                   enquote(strMetaChar) & " " & _
                                   "with " & _
                                   enquote(strXML)
        If Instr(strComments, strComment) = 0 Then
            If strComments <> "" Then strComments &= vbNewline
            strComments &= "*  " & strComment
        End If                                           
        Return(Replace(strValue, strMetaChar, strXML))                                              
    End Function                                                   

    
    ' -----------------------------------------------------------------
    ' Convert soft paragraph to hard paraggraph
    '
    '
    ' A "soft" paragraph is text terminated by one newline: a "hard" 
    ' paragraph is text broken into lines by newlines but terminated by 
    ' TWO newlines.  This function converts from the soft to the hard format.
    '
    ' strSoftParagraphs may consist of zero, one or more paragraphs in
    ' soft format, separated by single newlines.  Each soft paragraph is
    ' converted to the hardened format and the result is returned as the
    ' string value of this function.
    '
    ' If the optional parameter bytLineWidth is absent the hard paragraph
    ' consists of no more than 80 characters, and words are never broken up
    ' between hard paragraph lines.  The limit of 80 can be changed by
    ' specifying a new value for bytLineWidth.
    '
    ' Each line may be given a prefix (such as "' * " for comment boxes)
    ' using the optional strLinePrefix paramater.
    '
    ' Each line may be given a suffix (such as " *" for comment boxes)
    ' using the optional strSuffix parameter.
    '
    ' If booSameWidth is present and True then each line is padded to 80
    ' characters (or the value in bytLineWidth.)  When a suffix is defined
    ' and booSameWidth is True the padding is placed after the end of the
    ' text and before the suffix.
    '
    ' By default, the result is inspected for correctness:
    '
    '
    '      *  If booSameWidth is False the length of each line in the
    '         result must be less than or equal to bytLineWidth
    '
    '      *  If booSameWidth is True the length of each line in the
    '         result must be equal to bytLineWidth
    '
    '
    ' This inspection can be suppressed using booInspection:=False.
    '
    ' Whenever a hard LINE break is encountered in strSoftParagraphs,
    ' in the form of a single newline this is replaced, by default,
    ' with a hard PARAGRAPH break in the form of two newlines...which
    ' produces a blank line on the output in most cases.
    '
    ' This default can be overridden by passing a string in the
    ' strHardParaBreak parameter which is placed between paragraphs.
    '
    ' Between lines as formatted inside soft paragraphs a single newline
    ' character is inserted by default.  Override this choice using the
    ' strHardLineBreak parameter.
    '
    ' For example, you may need to output HTML where a hard line break
    ' needs to be <BR>, and a hard paragraph break needs to be <BR><BR>:
    ' handle this using strHardLineBreak:="<BR>" and strHardParaBreak:=
    ' "<BR><BR>".
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 12 06 98   Nilges         Version 1
    ' 01 21 00   Nilges         Renamed prefix and suffix parameters
    ' 11 18 02   Nilges         1.  Converted to vbNet
    '                           2.  Added strHardParaBreak parameter
    '
    '
    Public Shared Function soft2HardParagraph(ByVal strSoftParagraphs As String, _
                                              Optional ByVal bytLineWidth As Byte = 80, _
                                              Optional ByVal strLinePrefix As String = "", _
                                              Optional ByVal strLineSuffix As String = "", _
                                              Optional ByVal booSameWidth As Boolean = False, _
                                              Optional ByVal booInspection As Boolean = False, _
                                              Optional ByVal strHardParaBreak As String = vbNewline & vbNewline, _
                                              Optional ByVal strHardLineBreak As String = vbNewline) As String
        Dim booOkay As Boolean
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intSuffixLength As Integer
        Dim strOutput As String
        Dim strOutputLine As String
        Dim strSplitLines() As String
        Dim strSplitWords() As String
        If bytLineWidth < 1 Then
            errorHandler("Invalid line width parameter " & bytLineWidth, _
                         "utilities", "soft2HardParagraph")
            Return("")
        End If
        intSuffixLength = Len(strLineSuffix)
        Try
            strSplitLines = split(strSoftParagraphs, vbNewline)
        Catch
            errorHandler("Can't split input string", _
                         "utilities", "soft2HardParagraph", _
                         Err.Number & " " & Err.Description)
            Return("")
        End Try        
        For intIndex1 = 0 To UBound(strSplitLines)
            If strSplitLines(intIndex1) <> "" Then
                Try
                    strSplitWords = split(strSplitLines(intIndex1), " ")
                Catch
                    errorHandler("Can't split input line", _
                                "utilities", "soft2HardParagraph", _
                                Err.Number & " " & Err.Description)
                    Return("")
                End Try        
                strOutputLine = strLinePrefix
                For intIndex2 = 0 To UBound(strSplitWords)
                    If strSplitWords(intIndex2) <> "" Then
                        If Len(strOutputLine) _
                        + _
                        Len(strSplitWords(intIndex2)) _
                        + _
                        intSuffixLength _
                        + _
                        1 _
                        > _
                        bytLineWidth Then
                            soft2HardParagraph_attachOutputLine_(strOutputLine, _
                                                                 strOutput, _
                                                                 bytLineWidth, _
                                                                 strLinePrefix, _
                                                                 strLineSuffix, _
                                                                 booSameWidth, _
                                                                 strHardLineBreak)
                        End If
                        strOutputLine = strOutputLine & _
                                        CStr(IIf(strOutputLine = strLinePrefix, "", " ")) & _
                                        strSplitWords(intIndex2)
                    End If                                    
                Next intIndex2
                soft2HardParagraph_attachOutputLine_(strOutputLine, _
                                                     strOutput, _
                                                     bytLineWidth, _
                                                     strLinePrefix, _
                                                     strLineSuffix, _
                                                     booSameWidth, _
                                                     strHardLineBreak)
            Else
                strOutput &= strHardParaBreak                                                                
            End If                                                    
        Next intIndex1
        If strOutput = "" Then
            soft2HardParagraph_attachOutputLine_(strLinePrefix, _
                                                 strOutput, _
                                                 bytLineWidth, _
                                                 strLinePrefix, _
                                                 strLineSuffix, _
                                                 booSameWidth, _
                                                 strHardLineBreak)
        End If
        If booInspection Then
            ' Quality inspector
            Try
                strSplitLines = split(strOutput, vbNewline)
            Catch
                errorHandler("Can't split result string for inspection", _
                            "utilities", "soft2HardParagraph", _
                            Err.Number & " " & Err.Description)
                Return("")
            End Try        
            For intIndex1 = 0 To UBound(strSplitLines)
                If booSameWidth Then
                    booOkay = Len(strSplitLines(intIndex1)) _
                              = _
                              bytLineWidth
                Else
                    booOkay = Len(strSplitLines(intIndex1)) _
                              = _
                              bytLineWidth
                End If
                If Not booOkay Then Exit For
            Next intIndex1
            If intIndex1 <= UBound(strSplitLines) Then
                errorHandler _
                    ("Internal failure in soft2HardParagraph: " & _
                    "line " & intIndex1 & " " & _
                    "is of the unexpected width " & _
                    Len(strSplitLines(intIndex1)) & " " & _
                    "where each line width must be " & _
                    CStr(IIf(booSameWidth, "=", "<=")) & " " & _
                    bytLineWidth & ": " & _
                    enquote(string2Display(strSplitLines(intIndex1), _
                                           booBlankConvert:=True)))
            End If
        End If
        Return(strOutput)
    End Function
    
    ' -----------------------------------------------------------------
    ' Attach line to output on behalf of soft2HardParagraph
    '
    '
    Private Shared Sub soft2HardParagraph_attachOutputLine_(ByRef strOutputLine As String, _
                                                            ByRef strOutput As String, _
                                                            ByVal bytLineWidth As Byte, _
                                                            ByVal strPrefix As String, _
                                                            ByVal STRsuffix As String, _
                                                            ByVal booSameWidth As Boolean, _
                                                            ByVal strHardLineBreak As String)
        If strOutputLine = "" Then Exit Sub
        strOutput = strOutput & _
                    CStr(IIf(strOutput = "" _
                             OrElse _
                             Len(strOutput) >= Len(strHardLineBreak) _
                             AndAlso _
                             Mid(strOutput, Len(strOutput) - Len(strHardLineBreak) + 1) = strHardLineBreak, _
                             "", _
                             strHardLineBreak)) & _
                    Trim(strOutputLine)
        If booSameWidth Then
            strOutput = strOutput & _
                        copies(Math.max(0, bytLineWidth - Len(strOutputLine) - Len(STRsuffix)), " ")
        End If
        strOutput = strOutput & STRsuffix
        strOutputLine = strPrefix
    End Sub

    ' -----------------------------------------------------------------
    ' Spin on a locked resource
    '
    '
    ' This method waits until a locked resource is available or a
    ' timeout value is exceeded.  Optionally it can slow the wait
    ' by sleeping between its checks to the spinlock.
    '
    ' The basic overload is spinlock(intLock) where intLock is an
    ' integer passed by reference: the lock is set to 1 until its
    ' value before the set is zero indicating an available resource
    ' associated with the lock, or a timeout default value of 10
    ' seconds occurs.  
    '
    ' Between checks, spinlock(intLock) will sleep 1 second.
    '
    ' If the lock is never released spinlock(intLock) returns False:
    ' on success it returns True.
    '
    ' spinlock(intLock, intTimeout) can override the default timeout
    ' value: spinlock(intLock, intTimeout, intSleep) overrides the
    ' sleep value.
    '
    ' spinlock(intLock, -1) and spinlock(intLock, -1, intSleep) will
    ' spin indefinitely until the lock is released.
    '
    ' If a timeout or sleep value of -2 is specified the default timeout
    ' or default sleep value is used.
    '
    ' The overload spinlock(intLock, intTimeout, intSleep, intCount) will
    ' assign the number of queries made to the lock, to the intCount
    ' reference parameter.
    '
    ' Note that spinlock(intLock, 0, 0) will return True if there is no
    ' lock, False otherwise.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   ---------------------------------------
    ' 01 07 03     Nilges       Version 1.0
    '
    ' 
    Public Shared Overloads Function spinlock(ByRef intLock As Integer) As Boolean
        Return(spinlock(intLock, _
                        SPINLOCK_TIMEOUT_DEFAULT))
    End Function    
    Public Shared Overloads Function spinlock(ByRef intLock As Integer, _
                                              ByVal intTimeout As Integer) As Boolean
        Return(spinlock(intLock, _
                        intTimeout, _
                        SPINLOCK_SLEEPWAIT_DEFAULT))
    End Function    
    Public Shared Overloads Function spinlock(ByRef intLock As Integer, _
                                              ByVal intTimeout As Integer, _
                                              ByVal intSleep As Integer) As Boolean
        Dim intCount As Integer                                              
        Return(spinlock(intLock, _
                        intTimeout, _
                        SPINLOCK_SLEEPWAIT_DEFAULT, _
                        intCount))
    End Function                                              
    Public Shared Overloads Function spinlock(ByRef intLock As Integer, _
                                              ByVal intTimeout As Integer, _
                                              ByVal intSleep As Integer, _
                                              ByRef intCount As Integer) As Boolean
        Dim intSleepWork As Integer = intSleep                                              
        Dim intTimeoutWork As Integer = intTimeout                                              
        If intTimeoutWork < -2 Then
            errorHandler("Invalid intTimeout value " & intTimeoutWork, _
                         "utilities", "spinlock")
            Return(False)       
        ElseIf intTimeoutWork = -2 Then
            intTimeoutWork = SPINLOCK_TIMEOUT_DEFAULT                                          
        End If              
        If intSleepWork = -2 Then                                 
            intSleepWork = SPINLOCK_SLEEPWAIT_DEFAULT
        ElseIf intSleepWork < 0 Then
            errorHandler("Invalid intSleep value " & intSleepWork, _
                         "utilities", "spinlock")
            Return(False)                         
        End If  
        Dim datStart As Date
        If intTimeoutWork <> -1 Then datStart = Now                                             
        Dim objThreading As Thread
        Do 
            intCount += 1
            If Interlocked.Exchange(intLock, intLock) = 0 Then Return(True)
            If intSleep <> 0 Then objThreading.CurrentThread.Sleep(intSleepWork)
        Loop While intTimeoutWork = -1 _
                   OrElse _
                   intTimeoutWork > 0 AndAlso DateDiff(DateInterval.Second, datStart, Now) < intTimeoutWork
        Return(False)
    End Function                                      

    ' ----------------------------------------------------------------------
    ' Wrap some lines in a nice box that is suitable for viewing by nerds
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 11 23 01     Nilges       1.  Bug: handling of newlines 
    '
    '
    Public Shared Overloads Function string2Box(ByVal strInstring As String) As String
        Return string2Box(strInstring, "", "*")
    End Function
    Public Shared Overloads Function string2Box(ByVal strInstring As String, _
                                                ByVal strBoxLabel As String) As String
        Return string2Box(strInstring, strBoxLabel, "*")
    End Function
    Public Shared Overloads Function string2Box(ByVal strInstring As String, _
                                                ByVal strBoxLabel As String, _
                                                ByVal strBuildChar As String) As String
        Dim colLines As New Collection
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intMaxLength As Integer
        Dim intNextLength As Integer
        Dim strFence As String
        Dim strOutstring As String
        ' --- Build character must be a character
        If Len(strBuildChar) <> 1 Then
            errorHandler("build character is not one character")
            Return("")
        End If
        ' --- Find length of longest line, while parsing lines
        intMaxLength = 0: intIndex1 = 1
        Do While intIndex1 <= Len(strInstring)
            intIndex2 = Instr(intIndex1, strInstring & vbNewLine, vbNewLine)
            intNextLength = intIndex2 - intIndex1
            colLines.Add(Mid(strInstring, intIndex1, intNextLength))
            If intNextLength > intMaxLength Then intMaxLength = intNextLength
            intIndex1 = intIndex2 + Len(vbNewLine)
        Loop
        intMaxLength += 4
        intMaxLength = CInt(Math.max(intMaxLength, _
                                CInt(Iif(strBoxLabel = "", 10, Len(strBoxLabel) + 15))))
        ' --- Produce the text box
        strFence = copies(intMaxLength, strBuildChar)
        If strBoxLabel = "" Then
            strOutstring = strFence 
        Else
            strOutstring = copies(5, strBuildChar) & _
                           " " & strBoxLabel & " " & _
                           copies(Cint(Math.max(0, intMaxLength - 5 - Len(strBoxLabel) - 2)), _
                                  strBuildChar) 
        End If
        For intIndex1 = 1 To colLines.Count
            strOutstring = strOutstring & _
                           vbNewLine & _
                           strBuildChar & " " & _
                           CStr(colLines(intIndex1)) & _
                           copies(intMaxLength - Len(colLines(intIndex1)) - 4, " ") & _
                           " " & strBuildChar 
        Next intIndex1
        Return(strOutstring & vbNewLine & strFence)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Convert string returned by collection2String back to a collection
    '
    '
    ' This method converts a string that has been created by the
    ' collection2String method back into a new collection, where possible,
    ' and this method returns this new collection as a function value.
    '
    ' Where the string consists of a comma-delimited list of items, they
    ' are converted according to the following rules.
    '
    '
    '      *  Each value that is a number is converted to the narrowest
    '         numeric type possible
    '
    '      *  Each value that is a quoted string is converted to a string,
    '         containing the quoted characters
    '
    '      *  Each value that is a parenthesized list is converted 
    '         recursively into a collection using this method
    '
    '      *  The unquoted string Nothing results in a collection entry containing
    '         Nothing
    '
    '      *  All other values result in a collection entry containing
    '         Nothing
    '
    '
    ' The basic syntax of this method is string2Collection(s), returning a
    ' new collection c.  The overloaded syntax string2Collection(s, OK)
    ' will set OK to True or to False to indicate a fully valid conversion:
    '
    '
    '      *  If each entry in s was a number, a quoted string, or the
    '         unquoted string Nothing, the returned collection has been
    '         accurately rebuilt: OK is set to True.
    '
    '      *  Otherwise the returned collection is not accurate and OK is
    '         set to False
    '
    '
    ' By default this method assumes that the default separator of collection
    ' members, the comma, is in use.  Use the optional strDelim parameter
    ' to override this delimiter, and note that for best results when you need
    ' to be sure that the output of collection2String converts back accurately
    ' to the same collection, you can use a nongraphic strDelim for both methods.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 01 07 03   Nilges         Version 1.0
    ' 02 17 03   Nilges         Added strDelim and completed code
    '
    '
    ' --- Simple call
    Public Shared Overloads Function string2Collection(ByVal strInstring As String, _
                                                       Optional ByVal strDelim As String = ", ") As Collection
        Dim booOK As Boolean
        Return string2Collection(strInstring, booOK, strDelim:=strDelim)
    End Function    
    ' --- Call with OK set
    Public Shared Overloads Function string2Collection(ByVal strInstring As String, _
                                                       ByRef booOK As Boolean, _
                                                       Optional ByVal strDelim As String = ", ") As Collection
        booOK = False                                                       
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch
            errorHandler("Cannot create output collection object", _
                         "utilities", "string2Collection", _
                         Err.Number & " " & Err.Description & _
                         vbNewline & vbNewline & _
                         "Returning Nothing: booOK is False")
            Return(Nothing)                                     
        End Try        
        Dim strSplit() As String
        Try
            strSplit = split(strInstring, strDelim)
        Catch
            errorHandler("Cannot create split array", _
                         "utilities", "string2Collection", _
                         Err.Number & " " & Err.Description & _
                         vbNewline & vbNewline & _
                         "Returning empty collection: booOK is False")
            Return(colNew)                                     
        End Try        
        Dim intIndex1 As Integer
        Dim objNext As Object
        Dim booOKwork As Boolean = True
        Dim booIsNumeric As Boolean
        Dim dblNext As Double
        For intIndex1 = 0 To UBound(strSplit)
            objNext = string2Object(strSplit(intIndex1))
            booIsNumeric = True
            Try
                dblNext = CDbl(objNext)
            Catch
                booIsNumeric = False
            End Try            
            If Not (TypeOf objNext Is System.String) AndAlso Not booIsNumeric Then
                booOKwork = False
            Else
                Try
                    colNew.Add(objNext)
                Catch
                    errorHandler("Cannot add object to collection", _
                                 "utilities", "string2Collection", _
                                 Err.Number & " " & Err.Description & _
                                 vbNewline & vbNewline & _
                                 "Returning partial collection: booOK is False")
                    Return(colNew)                                     
                End Try                                
            End If            
        Next intIndex1        
        booOK = booOKWork
        Return(colNew)
    End Function    

    ' ----------------------------------------------------------------------
    ' String to display
    '
    '
    ' This method converts strings to representations that avoid "non-graphic"
    ' characters not commonly available on typical USA keyboards.  It replaces
    ' non-graphic characters with their representation in eXtensible Markup
    ' Language, C, or as Visual Basic expressions.
    '
    ' This method has the following overloaded syntaxes.
    '
    '
    '      *  string2Display(string) returns the string converted to XML syntax
    '
    '      *  string2Display(string, syntax) returns the string converted to the
    '         specified syntax:
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.C) converts to C
    '           syntax
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.XML) converts to XML
    '           syntax
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.VBExpression) converts  
    '           to Visual Basic expression syntax
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.VBExpressionCondensed) 
    '           converts to condensed Visual Basic expression syntax
    '
    '         The syntax can alternatively be specified as a string:
    '
    '         + string2Display(string, "C") converts to C syntax
    '
    '         + string2Display(string, "XML") converts to XML syntax
    '
    '         + string2Display(string, "VBExpression") converts to Visual Basic 
    '           expression syntax
    '
    '         + string2Display(string, "VBExpressionCondensed" converts to condensed 
    '           Visual Basic expression syntax
    '
    '
    ' In addition, the Optional parameter booBlankConvert can be used in all of the
    ' overloads.  This parameter can be passed as True to consider all blanks as
    ' nondisplayable characters:
    '
    '
    '      *  In C syntax, blanks will be converted to \x20.
    '      *  In XML syntax, blanks will be converted to &#00032.
    '      *  In VB expression syntax, blanks will be converted to ChrW(32)
    '
    '
    ' By default, blanks are considered displayable and not converted.
    '
    ' The Optional parameter booQuoteConvert will convert double quotes to C, XML, or
    ' VB syntax in a similar manner to booBlankConvert for the blank.
    ' 
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 03 25 02   Nilges         Documented ability to specify enumeration as  
    '                           a string
    '
    ' 01 07 03   Nilges         Added the strGraphic and the strNonGraphic
    '                           optional parameters: use charset object
    '
    ' 
    ' --- Convert to use XML notation
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    Optional ByVal booBlankConvert _
                                                             As Boolean = False, _
                                                    Optional ByVal booQuoteConvert _
                                                             As Boolean = False) As String
        Return(string2Display(strInstring, _
                              ENUdisplay2StringSyntax.XML, _
                              booBlankConvert:=booBlankConvert, _
                              booQuoteConvert:=booQuoteConvert))
    End Function
    ' --- Convert to use notation specified as a string
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    ByVal strSyntax As String, _
                                                    Optional ByVal booBlankConvert _
                                                             As Boolean = False, _
                                                    Optional ByVal booQuoteConvert _
                                                             As Boolean = False) As String
        Dim enuSyntax As ENUdisplay2StringSyntax = display2String_syntax2Enum_(strSyntax)
        If enuSyntax <> ENUdisplay2StringSyntax.Invalid Then                                                              
            Return(string2Display(strInstring, enuSyntax, _
                                  booBlankConvert:=booBlankConvert, _
                                  booQuoteConvert:=booQuoteConvert))
        End If                                  
    End Function
    ' --- Common logic
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    ByVal enuSyntax _
                                                          As ENUdisplay2StringSyntax, _
                                                    Optional ByVal booBlankConvert _
                                                             As Boolean = False, _
                                                    Optional ByVal booQuoteConvert _
                                                             As Boolean = False) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intLength As Integer
        Dim intTagEnd As Integer
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strNext As String
        Dim strGraphic As String = asciiCharsetEnum2String(ENUasciiCharset.graphic) 
        If enuSyntax <> ENUdisplay2StringSyntax.C _
           And _   
           enuSyntax <> ENUdisplay2StringSyntax.XML _
           And _   
           enuSyntax <> ENUdisplay2StringSyntax.VBExpression _
           And _   
           enuSyntax <> ENUdisplay2StringSyntax.VBExpressionCondensed Then
            errorHandler("Invalid enuSyntax parameter of " & enuSyntax)
            Return("")
        End if
        If booBlankConvert Then strGraphic = Replace(strGraphic, " ", "")
        If booQuoteConvert Then strGraphic = Replace(strGraphic, ChrW(34), "")
        If enuSyntax = ENUdisplay2StringSyntax.C Then 
            strGraphic = Replace(strGraphic, "\", "")
        End If
        If enuSyntax = ENUdisplay2StringSyntax.VBExpression _
           OrElse _
           enuSyntax = ENUdisplay2StringSyntax.VBExpressionCondensed Then 
            strGraphic = Replace(strGraphic, Chr(34), "")
        End If
        If enuSyntax = ENUdisplay2StringSyntax.C _
           AndAlso _
           verify(strInstring, asciiCharsetEnum2String(ENUasciiCharset.allAscii)) _
           <> _
           0 Then
            errorHandler("Syntax is C but input string contains non-ASCII characters", _
                         "string2Display", "utilities", _
                         "Unlike the other formats supported by string2Display " & _
                         "(XML and VB expression) the C format is not able to handle " & _
                         "international characters outside the extended ASCII range 0..255.")
            Return("")
        End If
        If strInstring = "" _
           AndAlso _
           (enuSyntax = ENUdisplay2StringSyntax.VBExpression _
            OrElse _
            enuSyntax = ENUdisplay2StringSyntax.VBExpressionCondensed) Then
            Return(Chr(34) & Chr(34))
        End If
        string2Display = ""
        intIndex1 = 1
        Do While intIndex1 <= Len(strInstring)
            intIndex2 = verify(strInstring & Chr(0), _
                               strGraphic, _
                               intIndex1)
            strNext = Mid(strInstring, intIndex1, intIndex2 - intIndex1)
            If strNext <> "" _
               AndAlso _
               (enuSyntax = ENUdisplay2StringSyntax.VBExpression _
                Or _
                enuSyntax = ENUdisplay2StringSyntax.VBExpressionCondensed) Then
                strNext = CStr(Iif(objStringBuilder.Length = 0, "", " & ")) & _
                string2Display_segment2Expression_(strNext, _
                                                    enuSyntax _
                                                    = _
                                                    ENUdisplay2StringSyntax. _
                                                    VBExpressionCondensed, _
                                                    True)
            End If
            objStringBuilder.Append(strNext)
            If intIndex2 > Len(strInstring) Then Exit Do
            intIndex1 = verify(strInstring & "A", _
                               strGraphic, _
                               intIndex2, _
                               True)
            strNext = Mid(strInstring, intIndex2, intIndex1 - intIndex2)
            If enuSyntax = ENUdisplay2StringSyntax.C Then
                strNext = string2Display_string2C_(strNext)
            ElseIf enuSyntax = ENUdisplay2StringSyntax.XML Then
                strNext = string2Display_string2XML_(strNext)
            Else
                strNext = CStr(Iif(objStringBuilder.Length = 0, "", " & ")) & _
                          string2Display_segment2expression_ _
                          (strNext, _
                           enuSyntax _
                           = _
                           ENUdisplay2StringSyntax.VBExpressionCondensed, _
                           False)
            End If 
            objStringBuilder.Append(strNext)
        Loop
        Return(objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert displayable characters to a VB subexpression
    '
    '
    Private Shared Function string2Display_segment2Expression_(ByVal strSegment As String, _
                                                               ByVal booCondense _
                                                                     As Boolean, _
                                                               ByVal booDisplayable _
                                                                     As Boolean) As String
        If Not booCondense Then
            If booDisplayable Then
                Return(enquote(strSegment))
            Else
                Return(string2Display_nondisplayable2Expression_(strSegment, False))
            End If
        Else
            Dim intIndex1 As Integer = 1
            Dim intIndex2 As Integer = 1
            Dim intIndex3 As Integer  
            Dim intIndex4 As Integer  
            Dim intSubstringLength As Integer
            Dim objStringBuilderChars As New System.Text.StringBuilder
            Dim objStringBuilderExp As New System.Text.StringBuilder
            Dim strStartChar As String 
            Do While intIndex1 <= Len(strSegment)
                strStartChar = Mid(strSegment, intIndex1, 1)
                If strStartChar <> Chr(34) Then
                    ' Find string of identical characters
                    For intIndex2 = intIndex1 + 1 To Len(strSegment)
                        If Mid(strSegment, intIndex2, 1) <> strStartChar Then Exit For 
                    Next intIndex2
                    intSubstringLength = intIndex2 - intIndex1
                    If intSubstringLength > 1 Then
                        If objStringBuilderChars.Length <> 0 Then
                            append(objStringBuilderExp, _
                                   " & ", _
                                   string2Display_segment2Expression_ _
                                   (objStringBuilderChars.ToString, _
                                    False, _
                                    booDisplayable))
                            objStringBuilderChars.Length = 0
                        End If
                        If booDisplayable Then
                            strStartChar = enquote(strStartChar)
                        Else
                            strStartChar = "ChrW(" & AscW(strStartChar) & ")"  
                        End If
                        append(objStringBuilderExp, _ 
                            " & ", _
                            "copies(" & _
                            strStartChar & _
                            ", " & _
                            intSubstringLength & _
                            ")")
                        intIndex1 = intIndex2
                    End If
                End If
                ' Find range of ascending or descending characters
                If intIndex1 <= Len(strSegment) Then
                    intIndex3 = AscW(Mid(strSegment, intIndex1, 1))
                    intIndex4 = intIndex3
                    For intIndex2 = intIndex1 To Len(strSegment)
                        If AscW(Mid(strSegment, intIndex2, 1)) <> intIndex3 Then 
                            Exit For 
                        End If
                        intIndex3 += 1
                    Next intIndex2
                    intSubstringLength = intIndex2 - intIndex1
                    If intSubstringLength > 1 Then
                        If objStringBuilderChars.Length <> 0 Then
                            append(objStringBuilderExp, _
                                   " & ", _
                                   string2Display_segment2Expression_ _
                                   (objStringBuilderChars.ToString, _
                                    False, _
                                    booDisplayable))
                            objStringBuilderChars.Length = 0
                        End If
                        append(objStringBuilderExp, _ 
                            " & ", _
                            "range2String(" & intIndex4 & ", " & intIndex3 - 1 & ")")
                        intIndex1 = intIndex2
                    Else
                        objStringBuilderChars.Append(Mid(strSegment, intIndex1, 1))
                        intIndex1 += 1
                    End If
                End If
            Loop
            If objStringBuilderChars.Length <> 0 Then
                append(objStringBuilderExp, _
                        " & ", _
                        string2Display_segment2Expression_ _
                        (objStringBuilderChars.ToString, _
                         False, _
                         booDisplayable))
            End If
            Return objStringBuilderExp.ToString
        End If
    End Function 

    ' ----------------------------------------------------------------------
    ' Convert displayable characters to a VB subexpression
    '
    '
    Private Shared Function string2Display_nondisplayable2Expression_ _
            (ByVal strNonDisplayable As String, _
             ByVal booCondense As Boolean) As String
        Dim intIndex1 As Integer
        Dim intLength As Integer
        Dim objStringBuilder As New System.Text.StringBuilder("")
        For intIndex1 = 1 To Len(strNondisplayable)
            append(objStringBuilder, _
                   " & ", _
                   "ChrW(" & AscW(Mid(strNonDisplayable, intIndex1, 1)) & ")")
        Next intIndex1
        Return objStringBuilder.ToString
    End Function 

    ' ----------------------------------------------------------------------
    ' On behalf of string2Display, convert all chars in string to C escapes
    '
    '
    Private Shared Function string2Display_string2C_(ByVal strInstring As String) As String
        Dim intIndex1 As Integer
        Dim intLength As Integer
        Dim lngAsc As Long
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strNextCharacter As String
        Dim strNextEscape As String
        intLength = Len(strInstring)
        intIndex1 = 1
        Do While intIndex1 <= intLength
            If intIndex1 <= intLength - 1 _
               AndAlso _
               Mid(strInstring, intIndex1, 2) = vbNewLine Then
                strNextEscape = "\n": intIndex1 += Len(vbNewline)
            Else
                strNextCharacter = Mid(strInstring, intIndex1, 1)
                If strNextCharacter = "\" Then
                    strNextEscape = "\\"
                Else
                    lngAsc = AscW(strNextCharacter)
                    If lngAsc < 0 OrElse lngAsc > 255 Then
                        errorHandler("International character values cannot be " & _
                                     "converted to C syntax")
                        Return("")
                    End If
                    strNextEscape = "\x" & _
                                    align(long2BaseN(lngAsc, _
                                          asciiCharsetEnum2String _
                                          (ENUasciiCharset.hexDigits)), _
                                          2, _
                                          ENUalign.alignRight, _
                                          "0") 
                End If
                intIndex1 += 1
            End If
            objStringBuilder.Append(strNextEscape)
        Loop
        Return(objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of string2Display, convert all chars in string to XML
    '
    '
    Private Shared Function string2Display_string2XML_(ByVal strInstring As String) As String
        Dim intIndex1 As Integer
        Dim objStringBuilder As New System.Text.StringBuilder
        For intIndex1 = 1 to Len(strInstring)
            objStringBuilder.Append("&#" & _
                                    align(CStr(AscW(Mid(strInstring, _
                                                        intIndex1, _
                                                        1))), _
                                          5, _
                                          ENUalign.alignRight, _
                                          "0"))
        Next intIndex1
        Return(objStringBuilder.ToString)
    End Function
    
    ' ----------------------------------------------------------------------
    ' String to file
    '
    '
    ' This method writes strOutstring to a file.  When it is called as a
    ' function this method returns True (file written OK) or False (error
    ' occured.)  Since this method writes the string a single byte at a
    ' time using synchronous stream output, this method is NOT suitable 
    ' for large volumes of data.
    '
    '
    Public Shared Function string2File(ByVal strOutstring As String, _
                                       ByVal strFileId As String) As Boolean
        Dim intIndex1 As Integer
        ' --- Erase old file when it exists
        If fileExists(strFileId) Then
            Try
                Kill(strFileId)
            Catch
                MsgBox("Cannot erase pre-existing file: " & Err.Number & Err.Description)
                Return(False)
            End Try
        End If
        ' --- Open the file for stream output
        Dim objStream As New IO.FileStream(strFileId, IO.FileMode.CreateNew)
        ' --- Write the file
        For intIndex1 = 1 To Len(strOutstring)
            Try
                objStream.WriteByte(CByte(Asc(Mid(strOutstring, intIndex1, 1))))
            Catch
                MsgBox("Could not write file: " & Err.Number & " " & Err.Description)
                objStream.Close 
                Return(False)
            End Try
        Next intIndex1
        ' --- Close the file
        objStream.Close
        Return(True)
    End Function

    ' ----------------------------------------------------------------------
    ' String to object
    '
    '
    ' This method converts a string containing the value of an object in the
    ' format returned by object2String back to the original value of the
    ' object, where possible.
    '
    ' The optional overload strign2Object(s, OK) will set OK to True if the
    ' string was converted back to an object successfully, False otherwise.
    ' The string was converted back to an object successfully when it was
    ' a number, a quoted string, or the unquoted string Nothing.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 11 13 01   Nilges         Version 1
    ' 12 04 01   Nilges         Pasted in and simplified from utility folder
    ' 01 07 03   Nilges         Added booOK parameter
    ' 02 17 03   Nilges         Examine and use decorated type 
    '
    '
    Public Shared Overloads Function string2Object(ByVal strInstring As String) As Object
        Dim booOK As Boolean
        Return(string2Object(strInstring, booOK))
    End Function    
    Public Shared Overloads Function string2Object(ByVal strInstring As String, _
                                                   ByRef booOK As Boolean) As Object
        Dim objValue As Object
        Dim strInstringWork As String = Trim(strInstring)
        Dim strType As String
        Dim strValue As String = strInstringWork
        string2Object_unDeco_(strInstring, strType, strValue)
        booOK = True
        If UCase(strValue) = "NOTHING" Then Return(Nothing)
        If Not string2Object_value2Type_(strValue, strType, objValue) Then 
            booOK = False: Return(Nothing)
        End If            
        Return canonicalTypeCast(objValue, strType)
    End Function  

    ' ----------------------------------------------------------------------
    ' On behalf of string2Object, undecorate string if possible
    '
    '
    Public Shared Function string2Object_unDeco_(ByVal strInstring As String, _
                                                    ByRef strType As String, _
                                                    ByRef strValue As String) As Boolean
        Dim intIndex1 As Integer = Instr(strInstring, "(")
        If intIndex1 < 2 _
           OrElse _
           intIndex1 > Instr(strInstring & Chr(34), Chr(34)) Then
            Return(False)
        End If
        If Mid(strInstring, Len(strInstring)) <> ")" Then Return(False)
        strType = Mid(strInstring, 1, intIndex1 - 1)
        strValue = Mid(strInstring, intIndex1 + 1, Len(strInstring) - intIndex1 -1)
        Return(True)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of string2Object, convert number, string or other to value
    ' and type
    '
    '
    ' Note that we set the type (to the narrowest canonical type available)
    ' only when it is not a null string, for the undecorator may have already
    ' found the type.
    '
    '
    Public Shared Function string2Object_value2Type_(ByVal strInstring As String, _
                                                     ByRef strType As String, _
                                                     ByRef objValue As Object) As Boolean
        If isQuoted(strInstring) Then
            objValue = dequote(strInstring)
            If Mid(strInstring, 1, 1) = "'" Then
                If Len(Cstr(objValue)) <> 1 Then
                    errorHandler("String " & strInstring & " uses single quotes but is not one char")
                    Return(False)
                End If
                objValue = CChar(objValue)
                strType = "SYSTEM.CHAR"
            Else
                strType = "SYSTEM.STRING"                
            End If
        ElseIf IsNumeric(strInstring) Then
            If strType = "" Then
                objValue = string2ValueObject(strInstring)
            Else
                Select Case UCase(strType)
                    Case "SYSTEM.BYTE": objValue = CByte(strInstring)
                    Case "SYSTEM.INT16": objValue = CShort(strInstring)
                    Case "SYSTEM.INT32": objValue = CInt(strInstring)
                    Case "SYSTEM.INT64": objValue = CLng(strInstring)
                    Case "SYSTEM.SINGLE": objValue = CSng(strInstring)
                    Case "SYSTEM.DOUBLE": objValue = CDbl(strInstring)
                    Case Else:
                        errorHandler("Unsupported string type " & _
                                     enquote(strType), _
                                     "utilities", _
                                     "string2Object_value2Type_", _
                                     "Not converting the value")
                End Select                
            End If                                            
            strType = objValue.GetType.ToString
        Else 
            objValue = Nothing
        End If
        Return(True)
    End Function
    
    ' ----------------------------------------------------------------------
    ' String to percent value
    '
    '
    ' This function converts a string to a percentage according to the 
    ' following rules.
    '
    '
    '      *  The string is trimmed, removing leading and trailing spaces
    '
    '      *  If it is purely an integer, equal to or greater than zero,
    '         it is divided by 100, converted to Single precision, and
    '         returned
    '
    '      *  If the string is a real number in Single precision range, 
    '         followed by a percent sign (perhaps, with one of more spaces 
    '         between the integer and the percent sign), then the string's
    '         Single value is returned.
    '
    '      *  If the string is a real number in Single precision range its
    '         Single value is returned.
    '
    '      *  Otherwise, -1 is returned to indicate the error
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------------
    ' 01 07 02   Nilges         Version 1.0
    '
    '
    Public Shared Function string2Percent(ByVal strInstring As String) As Single
        Dim strInstringWork As String = trim(strInstring)
        Dim intValue As Integer = -1
        Try
            intValue = CInt(strInstringWork)
        Catch
            Try
                Dim sngValue As Single = CSng(strInstringWork)
                Return(sngValue)
            Catch
                Dim objRegEx As System.Text.RegularExpressions.Regex
                Try
                    objRegEx = New System.Text.RegularExpressions.Regex("^[0123456789.+-]+[ ]*%$")
                Catch
                    errorHandler("Cannot create regular expression", _
                                 "utilities", "string2Percent", _
                                 Err.Number & " " & Err.Description & _
                                 vbNewline & vbNewline & _
                                 "Return -1 to indicate non-percent value")
                    Return(-1)                                 
                End Try                
                If objRegEx.IsMatch(strInstringWork) Then
                    Try
                        Return CSng(word(strInstringWork, 1))
                    Catch
                        Return(-1)
                    End Try                        
                End If                
            End Try            
        End Try        
        If intValue < 0 Then Return(-1)
        Return(CSng(intValue) / 100)
    End Function    
    
    ' ----------------------------------------------------------------------
    ' String to value object
    '
    '
    ' This method converts strInstring to number of the type that represents
    ' the best (narrowest) representation of the value of the input string.
    ' Otherwise it returns a string.
    '
    ' By default, this method tries conversion to the following types in the
    ' sequence listed: Boolean, Integer, Long, Double and String, returning
    ' the first type successfully converted.
    '
    ' However, an optional overloaded parameter can control what types are
    ' attempted and also in what order.  Use string2ValueObject(string, types)
    ' to do this listing the types to be attempted as a space-separated list
    ' of type names.
    '
    ' For example, string2ValueObject(string, "SHORT INTEGER STRING") will
    ' return a Short integer, a 32-bit integer, or a string depending on whether
    ' the string can be converted to 16 bits, 32 bits or a string.
    '
    ' The predefined constant CANONICAL_TYPES is a list of all the canonical
    ' types "BOOLEAN BYTE CHAR SHORT INTEGER LONG SINGLE DOUBLE STRING", and
    ' string2ValueObject(string, CANONICAL_TYPES) will convert string to the
    ' first canonical type found.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION
    ' --------     ----------   --------------------------------------------
    ' 10 23 01     Nilges       Added strTypePreference parameter and
    '                           Char and Boolean types.
    '
    '
    Public Shared Function string2ValueObject(ByVal strInstring As String) As Object
        Try
            Return(CByte(strInstring))
        Catch
            Try
                Return(CShort(strInstring))
            Catch
                Try
                    Return(CInt(strInstring))
                Catch
                    Try
                        Return(CLng(strInstring))
                    Catch
                        Try
                            Return(CSng(strInstring))
                        Catch
                            Try
                                Return(CDbl(strInstring))
                            Catch
                                Return(strInstring)
                            End Try
                        End Try
                    End Try
                End Try
            End Try
        End Try
    End Function
    
    ' -----------------------------------------------------------------------------
    ' Self-test
    '
    '
    Public Shared Function test(ByRef strReport As String) As Boolean
        strReport = "The test could not be completed"
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler("Cannot create the StringBuilder for the test report", _
                         "utilities", "test", _
                         Err.Number & " " & Err.Description)
            Return(False)                         
        End Try 
        Dim intErrCount As Integer       
        Dim strUtility As String
        Try
            ' --- Test abbrev
            strUtility = "abbrev"
            If Not test_appendResult_(objStringBuilder, _
                                        abbrev("  T", " test "), _
                                        strUtility, _
                                        "abbrev(check, master): " & _
                                        "tests string check for an abbreviation of string master", _
                                        intErrCount) Then
                Return(False)
            End If                                        
        Catch
            test_appendResult__(objStringBuilder, _
                                vbNewline & vbNewline & vbNewline, _
                                "Test terminated because of an error in " & strUtility & ": " & _
                                Err.Number & " " & Err.Description)
        End Try   
        If Not test_appendResult__(objStringBuilder, _
                                   vbNewline & vbNewline & vbNewline, _
                                   "Test completed: " & intErrCount & " error(s)") Then
            Return(False)
        End If                                   
        strReport = objStringBuilder.ToString
        Return(True)     
    End Function   
    
    ' -----------------------------------------------------------------------------
    ' Append test and its result to report on behalf of test method, and maintain
    ' an error count
    '
    '
    Private Shared Function test_appendResult_(ByRef objStringBuilder As System.Text.StringBuilder, _
                                                ByVal booOK As Boolean, _
                                                ByVal strUtility As String, _
                                                ByVal strDescription As String, _
                                                ByRef intErrCount As Integer) As Boolean
        If Not append(objStringBuilder, _
                      vbNewLine & vbNewLine, _
                      string2Box("Testing " & strUtility & " at " & Now & _
                                 vbNewLine & vbNewLine & _
                                 "Result is " & _
                                 Cstr(Iif(booOK, "OK", "ERROR")) & _
                                 vbNewline & vbNewline & _
                                 strDescription)) Then Return(False)
        If Not booOK Then intErrCount += 1                                 
    End Function                                   
    
    ' -----------------------------------------------------------------------------
    ' Append to report on behalf of test method
    '
    '
    Private Shared Overloads Function test_appendResult__(ByRef objStringBuilder As System.Text.StringBuilder, _
                                                            ByVal strAppend As String) As Boolean
        Return test_appendResult__(objStringBuilder, vbNewline, strAppend)
    End Function                                                   
    Private Shared Overloads Function test_appendResult__(ByRef objStringBuilder As System.Text.StringBuilder, _
                                                            ByVal strSeparator As String, _
                                                            ByVal strAppend As String) As Boolean
        Return append(objStringBuilder, strSeparator, strAppend)
    End Function                                   

    ' -----------------------------------------------------------------------------
    ' Translate source to target characters
    '
    '
    ' This method changes each occurence of a character from strSourceCharacterSet
    ' in strInstring to the corresponding character in strTargetCharacterSet, and
    ' returns the translated string as its value.
    '
    ' For each source character, the target character (to which it is translated) is
    ' taken from the corresponding position in strTargetCharacterSet.  If the target
    ' character set is shorter than the source character set then excess source
    ' characters are not translated.
    '
    '
    Public Shared Function translate(ByVal strInstring As String, _
                                     ByVal strSourceCharacterSet As String, _
                                     ByVal strTargetCharacterSet As String) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strOutstring As String
        For intIndex1 = 1 To Len(strInstring)
            intIndex2 = Instr(strSourceCharacterSet, Mid(strInstring, intIndex1, 1))
            If intIndex2 <> 0 AndAlso intIndex2 <= Len(strTargetCharacterSet) Then
                strOutString &= Mid(strTargetCharacterSet, intIndex2, 1)
            Else
                strOutString &= Mid(strInstring, intIndex1, 1)
            End If
        Next intIndex1
        Return(strOutstring)
    End Function
    
    ' ----------------------------------------------------------------------
    ' Call one of these methods from quickBasic or another environment
    '
    '
    ' Note these specific, restricted formats that have to be used.  Note
    ' also that you should probably read the documentation for the function
    ' to avoid pitfalls and gotyas.
    '
    '
    '      *  utility("abbrev", checkStr, masterStr)
    '
    '         Returns True when checkStr is an abbreviated form of
    '         masterStr, when case and leading or trailing spaces are
    '         ignored. 
    '
    '      *  utility("append", str1, sep, str2, True|False, -1|maxlen)
    '
    '         Appends str2 to str1, separating them by sep when str1 is
    '         not null.  Appends to the start of str1 when parm 4 is True,
    '         to the end otherwise.  Parm5 should be the maximum length of
    '         str1 or -1 for no limit.
    '
    '         Returns the appended string.
    '
    '      *  utility("appendPath", path, title)
    '
    '         Appends the file title to the path with needed separation,
    '         returns the appended path as a string.
    '
    '      *  utility("alignCenter", str, len, fill)
    '
    '         Centers str within len with the indicated fill characters.
    '
    '      *  utility("alignLeft", str, len, fill)
    '
    '         Left-aligns str within len with the indicated fill characters.
    '
    '      *  utility("alignRight", str, len, fill)
    '
    '         Right-aligns str within len with the indicated fill characters.
    '
    '      *  utility("baseN2Long", str, digits, wordsize, True|False)
    '
    '         The string str is interpreted as an integer in base N, and this
    '         utility returns its value as an integer.  The digits of the
    '         number system should be defined as a string in digits from the
    '         digit for zero to the highest place value.  wordsize can be
    '         zero if no word size is defined, or a word size (which enables
    '         negative numbers to be processed.)  The last parameter should be
    '         True to treat digits of the same case the same way.
    '
    '      *  utility("datatype", str, type)
    '
    '         Returns True when str has the data type identified in type, False
    '         otherwise.  See the datatype method for a list of the accepted
    '         types, which should be identified by the string form of their names.
    '
    '      *  utility("dequote", str)
    '
    '         If str is in quotes the quotes are removed and the result is
    '         returned.  If str is not in quotes it is returned without
    '         change.
    '
    '      *  utility("display2String", str, syntax, True|False)
    '
    '         str should be a string in the displayable syntax that is
    '         returned by string2Display and identified, as a string,
    '         in syntax.  syntax should be C, XML, vbExpression,
    '         vbExpressionCondensed or determine; note that determine
    '         will try to determine the syntax from clues.
    '
    '         str is converted back to its original form and returned.
    '
    '         The 4th parm should be True when C and XML nondisplayables
    '         such as \x00 and &#00000 can be variable length (such as the
    '         values \x0 and &#0.) 
    '
    '      *  utility("ellipsis", str, len, ellipsis, True|False)
    '
    '         Returns the string in str truncated, as needed, to len
    '         characters.  If this involves the removal of any characters
    '         then the last three characters returned are always the
    '         string in ellipsis.  Pass True in the fifth parm to obtain
    '         the ellipsis effect at the front of the string.
    '
    '      *  utility("file2String", fileid)
    '
    '         fileid should be a string, containing a file id.  This 
    '         interface returns the contents of the file as a string
    '         or a null string on any error.
    '
    '      *  utility("fileExists", fileid)
    '
    '         fileid should be a string, containing a file id.  This 
    '         interface True when the identified file exists, False
    '         when it does not exist.
    '
    '      *  utility("findAbbrev", target, masterList)
    '
    '         Returns the index (from one) of target in masterList when
    '         target is a unique abbreviation of a blank-delimited word
    '         in masterList.  Case is ignored, and leading and trailing
    '         spaces are removed from the target. 
    '
    '         Returns zero when the target is either not found or
    '         occurs more than once in abbreviated or complete form.
    '
    '      *  utility("findItem", str, tgt, delim, True|False, True|False)
    '
    '         Searches str for the item in tgt.
    '
    '         Parameters 4..6 define the syntax of str:
    '
    '              + delim is the character or string that separates
    '                items.
    '
    '              + Parameter 5 should be True or False:
    '
    '                - If tgt is a character or a set of alternative
    '                  characters, any one of which may be a delimiter,
    '                  then parameter 5 should be True.
    '
    '                - If tgt is a string which separates items then
    '                  parameter 5 should be False.
    '
    '              + Parm 6 should be True to ignore case
    '
    '      *  utility("findWord", str, tgtWord, True|False)
    '
    '         Searches str for the target word, performing a case-insensitive
    '         search when parameter 4 is True.
    '
    '      *  utility("histogram", val, minRange, maxRange, minVal, maxVal)
    '
    '         Maps the value in val onto a larger or smaller range of values,
    '         as in the case of visual progress reporting or charting.  The
    '         range is specified in minRange and in maxRange.  The possible
    '         range of values is specified in minVal and maxVal.  Returns the
    '         double precision form of the mapped value, which will be between
    '         minRange and maxRange. 
    '
    '      *  utility("indent", str, indent, newline)
    '
    '         Indents each line of str.  The third parm, indent, may be a 
    '         count of the number of blanks to be used or it may be the string
    '         that indents.  The fourth parm defines the character or string
    '         that separates lines in str. 
    '
    '      *  utility("int2Digits", val)
    '
    '         Returns the number of digits in the absolute value of the integer
    '         in val. 
    '
    '      *  utility("isQuoted", str)
    '
    '         Returns True when str is properly quoted in double quotes, False
    '         otherwise. 
    '
    '      *  utility("item", str, i, delim, True|False)
    '
    '         Returns the indexed item in str, where the index, from 1, is in i.
    '
    '         Parameters 4 and 5 define the syntax of str:
    '
    '              + delim is the character or string that separates
    '                items.
    '
    '              + Parameter 5 should be True or False:
    '
    '                - If tgt is a character or a set of alternative
    '                  characters, any one of which may be a delimiter,
    '                  then parameter 5 should be True.
    '
    '                - If tgt is a string which separates items then
    '                  parameter 5 should be False.
    '
    '      *  utility("itemPhrase", str, i, k, delim, True|False)
    '
    '         Returns the "phrase" of contiguous items in str, where the start
    '         index is in i, and the count of items in the phrase is in k.
    '
    '         Parameters 5 and 6 define the syntax of str:
    '
    '              + delim is the character or string that separates
    '                items.
    '
    '              + Parameter 6 should be True or False:
    '
    '                - If tgt is a character or a set of alternative
    '                  characters, any one of which may be a delimiter,
    '                  then parameter 6 should be True.
    '
    '                - If tgt is a string which separates items then
    '                  parameter 6 should be False.
    '
    '      *  utility("items", str, delim, True|False)
    '
    '         Returns the count of items in str.  
    '
    '         Parameters 3 and 4 define the syntax of str:
    '
    '              + delim is the character or string that separates
    '                items.
    '
    '              + Parameter 4 should be True or False:
    '
    '                - If tgt is a character or a set of alternative
    '                  characters, any one of which may be a delimiter,
    '                  then parameter 4 should be True.
    '
    '                - If tgt is a string which separates items then
    '                  parameter 4 should be False.
    '
    '      *  utility("long2BaseN", L, digits, wordsize)
    '
    '         The integer in L is interpreted as a base 10 integer, and this
    '         utility returns its value as a base N integer, inside a string.  
    '         The digits of the number system should be defined as a string 
    '         in the digits parameter from the digit for zero to the highest 
    '         place value.  wordsize can be zero if no word size is defined, 
    '         or a word size (which enables negative numbers to be processed.)  
    '
    '      *  utility("phrase", str, i, k)
    '
    '         Returns the phrase of contiguous blank-delimited words in str,
    '         commencing with word i for a length of k. 
    '
    '      *  utility("range2String", start, end)
    '
    '         Returns a range of sequentially ascending characters.  The
    '         start and the end parameters may specify either strings or
    '         integers.  If they specify strings, each must be one character
    '         long, and the range of characters from start to end (including
    '         start and end) is returned.  If start and end are integers, then
    '         the characters corresponding to the integer range is returned.
    '
    '         If start is greater than end, then a descending range is returned.  
    '
    '      *  utility("shell", command, appWinStyle, True|False, timeout)
    '
    '         Executes the system command in command and returns True on success
    '         or False on an error.  
    '
    '         Parm 3 should be the window style of the shell application as one of
    '         the strings Hide, MaximizedFocus, MinimizedFocus, MinimizedNoFocus,
    '         NormalFocus, NormalNoFocus. Parm 4 should be True to wait for
    '         termination.  Parm 5 should be the amount of time to wait for termination.
    '
    '         Note also that unlike most other utility interfaces  this method does 
    '         not correspond to a utility.  
    '
    '      *  utility("string2Box", str, lbl, char)
    '
    '         Creates a boxed display of str, suitable for viewing by geeks in
    '         a mono-spaced font.  lbl may be null or a label for the first line
    '         of the box.  char should be the character out of which the box
    '         should be constructed, such as an asterisk. 
    '
    '      *  utility("string2Display", str, syntax, True|False, True|False)
    '
    '         str (which may be a string containing Unicode and non-graphical
    '         characters) is converted to a string containing only characters
    '         available on USA keyboards and displayable in standard fonts.
    '
    '         syntax should be one of these string values:
    '
    '              + C: a string using C notation (\xnn) for nongraphics will be
    '                returned.  An error will result when the string contains
    '                Unicode values greater than 255. 
    '
    '              + XML: a string using eXtensible Markup Language (XML) notation 
    '                for nongraphics will be returned.  Nondisplay characters will
    '                be changed to &#ddddd where ddddd is their value.
    '
    '              + vbExpression: a string using Visual Basic expression notation 
    '                is returned.  This will be a series of quoted strings, and
    '                calls of the ChrW function for nondisplay values. 
    '
    '              + vbExpressionCondensed: a string using packed Visual Basic 
    '                expression notation is returned.  This will be a series of 
    '                quoted strings, calls of the ChrW function for nondisplay values, _
    '                and calls of the range2String and copies function as seen in
    '                this library.  range2String is used for sequences of ascending
    '                or descending values; copies is used for sequences of identical
    '                characters.  
    '
    '         Parms 4 and 5 govern the conversion of blanks and quotes.  If parm 4
    '         is True then blanks are considered nondisplay and are replaced as above.
    '         If parm 5 is True then double quotes are converted.
    '
    '      *  utility("string2File", str, file)
    '
    '         str is written to the file identified as a string, in file.  This 
    '         utility interface returns True on success and False on failure.
    '
    '      *  utility("translate", str, src, tgt)
    '
    '         str is translated, and the translation is returned.  To translate,
    '         each character in str that also appears in src is replaced by the
    '         corresponding character in tgt.
    '
    '      *  utility("verify", str, charset, i, m)
    '
    '         The index of the first character in str that IS NOT or IS in charset
    '         is returned.  The search starts at i.  If m is True, the index of the
    '         first character that IS in charset is returned, otherwise the index
    '         of the first non-matching character is returned.
    '
    '      *  utility("word", str, i)
    '
    '         Returns the ith blank-delimited word in str.
    '
    '      *  utility("words", str)
    '
    '         Returns the count of blank-delimited words.
    '
    '
    Public Shared Function utility(ByVal strUtilityName As String, _
                                   ParamArray objParameters() As Object) As Object
        Try
            Select Case UCase(Trim(strUtilityName))
                Case "ABBREV":
                    Return(abbrev(CStr(objParameters(0)), CStr(objParameters(1))))
                Case "APPEND":
                    Return(append(CStr(objParameters(0)), _
                                  CStr(objParameters(1)), _
                                  CStr(objParameters(2)), _
                                  booToStart:=CBool(objParameters(3)), _
                                  intMaxLength:=CInt(objParameters(4))))        
                Case "APPENDPATH":
                    Return(appendPath(CStr(objParameters(0)), CStr(objParameters(1))))         
                Case "ALIGNCENTER":
                    Return(align( CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  ENUalign.alignCenter, _ 
                                  CStr(objParameters(2))))        
                Case "ALIGNLEFT":
                    Return(align( CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  ENUalign.alignLeft, _ 
                                  CStr(objParameters(2))))        
                Case "ALIGNRIGHT":
                    Return(align( CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  ENUalign.alignRight, _ 
                                  CStr(objParameters(2))))        
                Case "BASEN2LONG":
                    Return(baseN2Long(CStr(objParameters(0)), _
                                      CStr(objParameters(1)), _
                                      CByte(objParameters(2)), _
                                      booIgnoreCase:=CBool(objParameters(3))))        
                Case "DATATYPE":
                    Return(datatype(CStr(objParameters(0)), CStr(objParameters(1))))        
                Case "DEQUOTE":
                    Return(dequote(CStr(objParameters(0))))        
                Case "DISPLAY2STRING":
                    Return(display2String(CStr(objParameters(0)), _
                                          CStr(objParameters(1)), _
                                          booVariableLenVals:=CBool(objParameters(2))))        
                Case "ELLIPSIS":
                    Return(ellipsis(CStr(objParameters(0)), _
                                    CInt(objParameters(1)), _
                                    strEllipsis:=CStr(objParameters(2)), _
                                    booLeftEllipsis:=CBool(objParameters(3))))        
                Case "FILE2STRING":
                    Return(file2String(CStr(objParameters(0))))        
                Case "FILEEXISTS":
                    Return(fileExists(CStr(objParameters(0))))        
                Case "FINDABBREV":
                    Return(findAbbrev(CStr(objParameters(0)), CStr(objParameters(1))))
                Case "FINDITEM":
                    Return(findItem(CStr(objParameters(0)), _
                                    CStr(objParameters(1)), _
                                    CStr(objParameters(2)), _
                                    CBool(objParameters(3)), _
                                    booIgnoreCase:=CBool(objParameters(4))))        
                Case "FINDWORD":
                    Return(findWord(CStr(objParameters(0)), _
                                    CStr(objParameters(1)), _
                                    booIgnoreCase:=CBool(objParameters(4))))        
                Case "HISTOGRAM":
                    Return(histogram(CDbl(objParameters(0)), _
                                     CDbl(objParameters(1)), _
                                     CDbl(objParameters(2)), _        
                                     CDbl(objParameters(2)), _        
                                     CDbl(objParameters(2))))        
                Case "INDENT":
                    If IsNumeric(objParameters(1)) Then
                        Return(indent(CStr(objParameters(0)), _
                                      CInt(objParameters(1)), _
                                      CStr(objParameters(2)))) 
                    Else                                         
                        Return(indent(CStr(objParameters(0)), _
                                      CStr(objParameters(1)), _
                                      CStr(objParameters(2))))    
                    End If                                      
                Case "INT2DIGITS":
                    Return(int2Digits(CInt(objParameters(0))))        
                Case "ISQUOTED":
                    Return(isQuoted(CStr(objParameters(0))))        
                Case "ITEM":
                    Return(item(CStr(objParameters(0)), _
                                CInt(objParameters(1)), _
                                CStr(objParameters(2)), _
                                CBool(objParameters(3))))        
                Case "ITEMPHRASE":
                    Return(itemPhrase(CStr(objParameters(0)), _
                                      CInt(objParameters(1)), _
                                      CInt(objParameters(1)), _
                                      CStr(objParameters(2)), _
                                      CBool(objParameters(3))))        
                Case "ITEMS":
                    Return(items(CStr(objParameters(0)), CStr(objParameters(1)), CBool(objParameters(2))))        
                Case "LONG2BASEN":
                    Return(long2BaseN(CLng(objParameters(0)), _
                                      CStr(objParameters(1)), _
                                      CByte(objParameters(2))))        
                Case "PHRASE":
                    Return(phrase(CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  CInt(objParameters(2))))        
                Case "RANGE2STRING":
                    If UCase(objParameters(0).GetType.ToString) = "SYSTEM.STRING" Then
                        Return(range2String(CStr(objParameters(0)), _
                                            CStr(objParameters(1))))   
                    Else                                            
                        Return(range2String(CInt(objParameters(0)), _
                                            CInt(objParameters(1))))   
                    End If                                                  
                Case "SHELL":
                    Dim objAppWinStyle As AppWinStyle
                    Dim strFocus As String = CStr(objParameters(1))
                    Select Case UCase(Trim(strFocus))
                        Case "HIDE": objAppWinStyle = AppWinStyle.Hide
                        Case "MAXIMIZEDFOCUS": objAppWinStyle = AppWinStyle.MaximizedFocus
                        Case "MINIMIZEDFOCUS": objAppWinStyle = AppWinStyle.MinimizedFocus
                        Case "MINIMIZEDNOFOCUS": objAppWinStyle = AppWinStyle.MinimizedNoFocus
                        Case "NORMALFOCUS": objAppWinStyle = AppWinStyle.NormalFocus
                        Case "NORMALNOFOCUS": objAppWinStyle = AppWinStyle.NormalNoFocus
                        Case Else:
                            errorHandler("Invalid application focus name " & enquote(strFocus), _
                                         "utilities", "utility", _
                                         "Focus has been set to Hide")
                            objAppWinStyle = AppWinStyle.Hide                                        
                    End Select                    
                    Return(shell(CStr(objParameters(0)), _
                                 Style:=objAppWinStyle, _    
                                 Wait:=CBool(objParameters(1))))        
                Case "STRING2BOX"
                    Return(string2Box(CStr(objParameters(0)), CStr(objParameters(1)), CStr(objParameters(2))))                                 
                Case "STRING2DISPLAY"
                    Return(string2Display(CStr(objParameters(0)), _
                                          CStr(objParameters(1)), _
                                          CBool(objParameters(2)), _
                                          CBool(objParameters(3))))                                 
                Case "STRING2FILE"
                    Return(string2Display(CStr(objParameters(0)), _
                                          CStr(objParameters(1))))                                 
                Case "TRANSLATE"
                    Return(translate(CStr(objParameters(0)), _
                                     CStr(objParameters(1)), _
                                     CStr(objParameters(2))))                                 
                Case "VERIFY"
                    Return(verify(CStr(objParameters(0)), _
                                  CStr(objParameters(1)), _
                                  CInt(objParameters(2)), _                                 
                                  CBool(objParameters(2))))                                 
                Case "WORD"
                    Return(word(CStr(objParameters(0)), _
                                CInt(objParameters(1))))                                 
                Case "WORDS"
                    Return(words(CStr(objParameters(0))))                                 
                Case Else
                    errorHandler("Unsupported utility name " & enquote(strUtilityName), _
                                 "utilities", "utility", _
                                 "Returning Nothing")                    
            End Select            
        Catch
            errorHandler("Error in calling utility " & enquote(strUtilityName) & ": " & Err.Number & Err.Description, _
                         "utilities", "utility", _
                         "Returning Nothing")                    
        End Try        
        Return(Nothing)
    End Function                            

    ' ----------------------------------------------------------------------
    ' Scan string for character sets
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 11 23 01     Nilges       1.  Bug: incorrect Mid 
    '
    '
    Public Shared Overloads Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String) _
           As Integer
        verify = verify(strInstring, _
                        strCharacterSet, _
                        1, _
                        False)
    End Function
    Public Shared Overloads Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String, _
                                            ByVal intStartIndex As Integer) _
           As Integer
        verify = verify(strInstring, _
                        strCharacterSet, _
                        intStartIndex, _
                        False)
    End Function
    Public Shared Overloads Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String, _
                                            ByVal booMatch As Boolean) _
           As Integer
        verify = verify(strInstring, _
                        strCharacterSet, _
                        1, _
                        booMatch)
    End Function
    ' --- Common overload
    Public Shared Overloads Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String, _
                                            ByVal intStartIndex As Integer, _
                                            ByVal booMatch As Boolean) _
                As Integer
        '
        ' Verify string
        '
        '
        Dim bytIndex1 As Byte
        Dim bytIndex2 As Byte
        Dim intCharacterSetLength As Integer = Len(strCharacterSet)
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intVerify As Integer
        ' --- Error checking
        If intStartIndex <= 0 Then
            MsgBox("Error in verify function: intStartIndex " & intStartIndex & " is not valid")
            Exit Function
        End If
        ' --- Degenerate cases: null input string, unity  set
        If strInstring = "" Then Exit Function
        If booMatch AndAlso intCharacterSetLength = 1 Then
            Return(InStr(intStartIndex, strInstring, strCharacterSet))
            Exit Function
        End If
        ' --- Scan the string
        intVerify = 0
        If booMatch Then
            If len(strCharacterSet) <= 26 And Len(strInstring) > 1024 Then
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
        Return(intVerify)
    End Function
    ' ----------------------------------------------------------------------
    ' Parse a blank-delimited word
    '
    '
    Public Shared Function word(ByVal strInstring As String, ByVal intIndex As Integer) As String
        Return(item(strInstring, intIndex)) 
    End Function

    ' ----------------------------------------------------------------------
    ' Count blank-delimited words
    '
    '
    Public Shared Function words(ByVal strInstring As String) As Integer
        Return(items(strInstring)) 
    End Function
    
    ' ----------------------------------------------------------------------
    ' Create a Point object from its X and its Y coordinates
    '
    '
    Public Shared Function xy2Point(ByVal intX As Integer, ByVal intY As Integer) As Point
        If intX < 0 OrElse intY < 0 Then
            errorHandler("Invalid X and/or y coordinates " & _
                         intX & "," & intY, _
                         "utilities", "xy2Point", _
                         "These values must be zero or positive")
            Return(Nothing)                         
        End If        
        Dim objPoint As Point 
        Try
            objPoint = New Point(intX, intY)
        Catch
            errorHandler("Cannot create Point", _
                         "utilities", "xy2Point", _
                         Err.Number & " " & Err.Description)
            Return(Nothing)                         
        End Try        
        Return(objPoint)
    End Function    
    
    ' ----------------------------------------------------------------------
    ' Create a Size object from its X and its Y values
    '
    '
    Public Shared Function xy2Size(ByVal intX As Integer, ByVal intY As Integer) As Size
        If intX < 0 OrElse intY < 0 Then
            errorHandler("Invalid X and/or y values " & _
                         intX & "," & intY, _
                         "utilities", "xy2Size", _
                         "These values must be zero or positive")
            Return(Nothing)                         
        End If        
        Dim objSize As Size 
        Try
            objSize = New Size(intX, intY)
        Catch
            errorHandler("Cannot create Size", _
                         "utilities", "xy2Size", _
                         Err.Number & " " & Err.Description)
            Return(Nothing)                         
        End Try        
        Return(objSize)
    End Function    

End Class

