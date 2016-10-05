Option Strict

Imports System
Imports Microsoft.VisualBasic
Imports DotNetEnum = System.Enum
Imports System.Threading
Imports System.Drawing

' **************************************************************************
' *                                                                        *
' * A collection of utility methods                                        *
' *                                                                        *
' *                                                                        *
' * This Shared and stateless object exposes a number of utility methods   *
' * for development tasks common to Windows and Web applications.          *
' *                                                                        *
' * The rest of this header block addresses the following topics.          *
' *                                                                        *
' *                                                                        *
' *      *  Methods (and some properties) exposed by this object           *
' *      *  Compile-time symbols                                           *
' *      *  Multithreading considerations                                  *
' *      *  References                                                     *
' *      *  Change record                                                  *
' *      *  Issues                                                         *
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
' *      *  breakLongWords: ensure words are of manageable size in text    *
' *      *  canonicalTypeCast: convert an object to a standard type        *
' *      *  canonicalTypeEnum2Name: convert the type enum to name          *
' *      *  canonicalTypeName2Enum: convert the type name to enum          *
' *      *  changeRecord: return change record in EGNSF format             *
' *      *  char2Name: convert the special character to its name           *
' *      *  charset: return international character sets                   *
' *      *  commonRegularExpressions: return useful regular expressions    *
' *      *  copies: make many copies of a string                           *
' *      *  datatype: test a string for a data type                        *
' *      *  datatypeEnum2Name: data type enumerator to name                *
' *      *  datatypeName2Enum: data type name to enumerator                *
' *      *  dequote: remove quotes from a string                           *
' *      *  determineNewline: determines the newline character             *
' *      *  deweyParent: returns the parent Dewey decimal number           *
' *      *  directoryExists: check for Windows directory                   *
' *      *  display2String: convert displayable string to value            *
' *      *  ellipsis: abbreviates string to ellipsis...                    *
' *      *  enquote: intelligently enquote a string                        *
' *      *  errorHandler: trivial error handler                            *
' *      *  extendTextBox: extend a text box (Windows)                     *
' *      *  file2String: read a file, return its contents as a string      *
' *      *  fileExists: return True when file exists                       *
' *      *  fileidParse: lightweight parser of a broad range of file ids   *
' *      *  findAbbrev: find potentially abbreviated string                *
' *      *  findBalParenthesis: find balancing parenthesis                 *
' *      *  findItem: locate item in string                                *
' *      *  formatOutline: format an outline for a monospace font          *
' *      *  hasReferenceType: determines if object is a reference object   *
' *      *  histogram: models data on a value range                        *
' *      *  incrementDewey: add one to a Dewey decimal number              *
' *      *  inspectionAppend: append inspection and test reports           *
' *      *  int2Digits: convert positive integer to width                  *
' *      *  isQuoted: return True when string is quoted                    *
' *      *  isXMLcomment: return True when a string is an XML comment      *
' *      *  isXMLname: return True when a string is an XML name            *
' *      *  item: return the nth delimited item from a string              *
' *      *  itemPhrase: return the nth through mth adjacent items          *
' *      *  items: return the count of items                               *
' *      *  itemTest: test the item method                                 *
' *      *  joinLines: join two multiple-line lists                        *
' *      *  line: return the newline-delimited item                        *
' *      *  lines: return the count of newline-delimited items             *
' *      *  listBox2Registry: save the list box in the Registry (Windows)  *
' *      *  listItem: return the comma-delimited list item                 *
' *      *  listItems: return the count of comma-delimited list items      *
' *      *  long2BaseN: convert Long integers to nondecimal bases          *
' *      *  mkXMLcomment: make an XML comment                              *
' *      *  mkXMLelement: make an XML element                              *
' *      *  mkXMLtag: make an XML tag                                      *
' *      *  Name: merely returns the name of this class                    *
' *      *  name2Char: converts the character name back to the character   *
' *      *  numbers2Variables: convert each number in a string to variable *
' *      *  object2Scalar: convert an object to a scalar                   *
' *      *  object2String: convert an object to a string                   *
' *      *  objectInfo2XML: convert object about info and state to XML tag *
' *      *  parseXMLtag: parse an XML tag                                  *
' *      *  phrase: return the nth through mth adjacent words              *
' *      *  properCase: capitalizes first letter, lowercases rest of string*
' *      *  randomSentence: create a random sentence                       *
' *      *  range2String: create a single string from a start/end range    *
' *      *  replaceXMLmetaChars: replace XML special chars <, > and &      *
' *      *  soft2HardParagraph: format paragraphs                          *
' *      *  smallBusinessList: comma-delimited list of small business names*
' *      *  spinlock: wait for a locked resource, then lock it             *
' *      *  string2Box: place a string in a box of asterisks               *
' *      *  string2Display: convert a string to a displayable form         *
' *      *  string2File: write a string to a file                          *
' *      *  string2Object: convert many strings to an object               *
' *      *  string2Percent: convert string to percent                      *
' *      *  string2Range: convert a string to a range of strings           *
' *      *  string2Sentinel: find string, that doesn't occur in input str  *
' *      *  string2ValueObject: convert a string to a value object         *
' *      *  tempFileid: return an available temporary file identifier      *
' *      *  test: run stress and smoke tests on the utilities library      *
' *      *  testAvailable: returns True (stress tests are doable) or False *
' *      *  translate: translate source to target characters               *
' *      *  trimContainer: adjust form/container width and height (Windows)*
' *      *  utility: runs one of these methods from the quickBasic engine  *
' *      *  verify: scan a string                                          *
' *      *  word: return the nth blank-delimited word from a string        *
' *      *  xmlMeta2Name: convert XML meta-characters to their names       *
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
' *      *  WINDOWS_LOGGING to control writing of errorHandler messages to *
' *         the Windows 2000 log                                           *
' *                                                                        *
' *      *  INCLUDE_TEST to include the test method                        *
' *                                                                        *
' *                                                                        *
' * ----- WINDOWS_LOGGING                                                  *
' *                                                                        *
' * The WINDOWS_LOGGING compile-time symbol should be True, to write all   *
' * messages sent (inside runtime usage of the utilities) to utilities.    *
' * errorHandler to the Windows log named utilities_errorHandler.          *
' *                                                                        *
' * When running on Windows 2000 when WINDOWS_LOGGING is in effect, you can*
' * examine errorHandler messages by accessing the MMC display for Computer*
' * Management.  In Event Viewer within System Tools, view the log named   *
' * Application.                                                           *
' *                                                                        *
' * Errors containing utilities_errorHandler in the column labeled Source  *
' * will be from the utilities error handler.                              *
' *                                                                        *
' *                                                                        *
' * ----- INCLUDE_TEST                                                     *
' *                                                                        *
' * The INCLUDE_TEST compile-time symbol should be True to include the     *
' * test method. The test method conducts a comprehensive test of the      *
' * utilities library.                                                     *
' *                                                                        *
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
' * 05 01 03   Nilges         1.  Added collectionClear                    *
' * 05 02 03   Nilges         1.  Added objectInfo2XML                     *
' * 05 05 03   Nilges         1.  Added numbers2Variables                  *
' * 05 06 03   Nilges         1.  Converted to Visual Studio Pro 2003      *
' * 05 09 03   Nilges         1.  Changed errorHandler                     *
' * 05 12 03   Nilges         1.  Added string2Sentinel                    *
' *                           2.  Added properCase                         *
' * 05 13 03   Nilges         1.  Changed string2File                      *
' *                           2.  Changed item                             *
' *                           3.  Changed datatype                         *
' *                           4.  Changed itemPhrase                       *
' *                           5.  Changed verify                           *
' * 05 15 03   Nilges         1.  Added changeRecord                       *
' * 05 18 03   Nilges         1.  Added commonRegularExpressions           *
' * 05 22 03   Nilges         1.  Changed display2String                   *
' * 05 23 03   Nilges         1.  Changed commonRegularExpressions         *
' * 05 24 03   Nilges         1.  Added commonRegularExpressionsTester     *
' * 05 24 03   Nilges         1.  Changed object2String                    *
' * 05 24 03   Nilges         1.  Added determineNewline                   *
' * 05 25 03   Nilges         1.  Changed commonRegularExpressions         *
' * 05 25 03   Nilges         1.  Found three Public procedures that should*
' *                               be Private                               *
' *                           2.  Changed item                             *
' * 05 26 03   Nilges         1.  Changed utility                          *
' *                           2.  Added vbCodeType                         *
' * 05 27 03   Nilges         1.  Changed object2XML                       *
' * 05 29 03   Nilges         1.  Changed commonRegularExpressions         *
' * 05 31 03   Nilges         1.  Added test                               *
' *                           2.  Added string2RLE and RLE2string          *
' *                           3.  Removed progress                         *
' *                           4.  Changed utility                          *
' * 06 01 03   Nilges         1.  Added tempFileid                         *
' *                           2.  Changed findItem                         *
' * 06 01 03   Nilges         1.  Changed string2Display                   *
' *                           2.  Changed appendPath                       *
' * 06 03 03   Nilges         1.  Changed test                             *
' *                           2.  Removed findXMLtag                       *
' *                           3.  Changed range2String                     *
' * 06 03 03   Nilges         1.  Changed string2ValueObject               *
' * 06 05 03   Nilges         1.  Added About property                     *
' * 06 07 03   Nilges         1.  Added Name property                      *
' *                           2.  Added collectionCopy/Compare             *
' *                           3.  Changed test method                      *
' * 06 10 03   Nilges         1.  Changed commonRegularExpressions         *
' * 06 11 03   Nilges         1.  Added sortItems                          *
' * 06 22 03   Nilges         1.  Note: changes on or after 6-22 are NOT   *
' *                               documented in the Addison-Wesley draft   *
' *                           2.  Changed inspectionAppend                 *
' *                           3.  Added itemMaxWidth                       *
' * 06 24 03   Nilges         1.  Added collectionDepth                    *
' *                           2.  Changed test                             *
' * 06 25 03   Nilges         1.  Added issues section                     *
' *                           2.  Added collectionTypes                    *
' * 06 26 03   Nilges         1.  Changed string2Collection                *
' * 06 26 03   Nilges         1.  Changed test                             *
' * 07 02 03   Nilges         1.  Added findBalParenthesis                 *
' *                           2.  Changed collection2String                *
' *                           3.  Changed string2Collection                *
' *                           4.  Changed collectionStringsCompare         *
' * 07 02 03   Nilges         1.  Changed collection2String                *
' * 07 05 03   Nilges         1.  Changed string2Object                    *
' *                           2.  Transferred collection utilities to own  *
' *                               class library and project                *
' * 07 14 03   Nilges         1.  Added object2Scalar                      *
' * 10 03 03   Nilges         1.  Changed errorHandler                     *
' * 10 05 03   Nilges         1.  Changed string2Object                    *
' * 10 17 03   Nilges         1.  Changed mkXMLelement                     *
' * 10 18 03   Nilges         1.  Added listItem, listItems, line and lines*
' *                           2.  Changed string2Display                   *
' * 11 25 03   Nilges         1.  Changed properCase                       *
' * 11 30 03   Nilges         1.  Changed canonicalTypeCast                *
' *                           2.  Changed items                            *
' * 12 06 03   Nilges         1.  Changed test                             *
' *                           2.  Added smallBusinessList                  *
' * 12 08 03   Nilges         1.  Changed item                             *
' * 12 09 03   Nilges         1.  Added itemTest                           *
' * 01 24 04   Nilges         1.  Added char2Name/name2Char                *
' *                           2.  Changed mkXMLtag                         * 
' * 01 26 04   Nilges         1.  Changed mkXMLtag                         *
' *                           2.  Changed mkXMLelement                     *
' * 01 28 04   Nilges         1.  Changed mkXMLtag                         *
' * 01 30 04   Nilges         1.  Added formatOutline                      *
' *                           2.  Added INCLUDE_TEST                       *
' *                           3.  Added testAvailable                      *
' *                           4.  Added incrementDewey                     *
' *                           5.  Added deweyParent                        *
' * 02 10 04   Nilges         1.  Changed display2String                   *
' * 02 21 04   Nilges         1.  Added breakLongWords                     *
' *                           2.  Changed utility                          *
' * 02 21 04   Nilges         1.  Added breakLongWords                     *
' *                           2.  Changed utility                          *
' * 04 09 04   Nilges         1.  Changed dequote                          *
' *                                                                        *
' *                                                                        *
' * I S S U E S ---------------------------------------------------------- *
' *   DATE         MODULE        POSTER     DESCRIPTION                    *
' * --------   -------------- ------------  ------------------------------ *
' * 06 25 03   string2Display Nilges        Overload may be problematic:   *
' *                                         when used enumerator overload, *
' *                                         enum value was changed to 3    *
' *                                                                        *
' * 10 17 03   string2Display Nilges        May not generate true XML      *
' *                                         syntax because the &# sequences*
' *                                         are not followed by semicolons.*
' *                                                                        *
' * 12 11 03   append         Nilges        It appears that appending a    *
' *                                         separator containing a null    *
' *                                         character (0) isn't supported  *
' *                                         because, perhaps, the string   *
' *                                         builder can't handle it.       *
' *                                                                        *
' * 01 25 04   mkXMLtag       Nilges        Add support for self-contained *
' *                                         tags.                          *
' *                                                                        *
' * 01 30 04   mkXMLtag       Nilges        Add outline2XML: use in        *
' *                                         regex2Explanation              *
' *                                                                        *
' * 02 01 04   mkXMLtag       Nilges        Although regex2Explanation is  *
' *                                         in the project folder it is at *
' *                                         this time incomplete, does not *
' *                                         work, and not mentioned in the *
' *                                         project readme. This class     *
' *                                         needs work to provide its      *
' *                                         promised ability to "explain   *
' *                                         regular expressions".          *
' *                                                                        *
' * 02 22 04   mkXMLtag       Nilges        breakLongWords merely replaces *
' *                                         a consistent slash or consist- *
' *                                         ent backslash in an apparent   *
' *                                         fileid with that character     *
' *                                         followed by a space.           *
' *                                                                        *
' *                                         Use COM file id parse and the  *
' *                                         recursive solution for breaking*
' *                                         the file id in archive\        *
' *                                         02222004 for better results    *
' *                                                                        *
' * 03 10 04   mkXMLtag       Nilges        Get error messages for undefin-*
' *                                         ed object/procedure because    *
' *                                         internal error message calls   *
' *                                         don't identify this info       * 
' **************************************************************************

Public Class utilities

    ' ***** Easter Egg *****
    Private Const ABOUTINFO As String = _
    "This Shared and stateless object exposes a number of utility methods " & _
    "for development tasks common to Windows and Web applications." & _
    vbNewline & vbNewline & _
    "Edward G. Nilges" & vbNewLine & _
    "spinoza1111@yahoo.COM" & vbNewLine & _
    "http://members.screenz.com/edNilges"

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
        XMLcondensed
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
    
    ' ***** The standard or canonical types *****
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
        Dim intLength As Integer = Len(strCheck) 'boop
        If intLength > Len(strMaster) Then Return(False)                            
        Return(Mid(Ucase(Trim(strMaster)), 1, Len(strCheck)) _
               = _
               UCase(Trim(strCheck))) 
    End Function                           
            
    ' -----------------------------------------------------------------------
    ' Return "about" information
    '
    '
    Public Shared ReadOnly Property About() As String
        Get
            Return(ABOUTINFO)
        End Get         
    End Property                           

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
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 05 13 03   Nilges         Replaced msgbox with errorHandler
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
            errorHandler("Invalid intMaxLength = " & intMaxLength, _
                         Name, "append", _
                         "Returning error indicator")
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
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 06 01 03   Nilges         Bug: the combine method previously
    '                           used did not meet the contract in the
    '                           documentation.
    '
    '
    Public Shared Function appendPath(ByVal strPath As String, _
                                      ByVal strFileTitle As String) As String
        If Len(strPath) = 0 Then Return(strFileTitle)                                      
        If Mid(strPath, Len(strPath), 1) <> "\" Then
            Return(strPath & "\" & strFileTitle)
        End If        
        Return(strPath & strFileTitle)
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
                         Name, "baseN2Long", _
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
                         Name, "bg2fgColor", _
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
                         Name, "bg2fgColor_mkTable_", _
                         Err.Number & " " & Err.Description)
            Return(Nothing)                         
        End Try  
        Dim strSplit() As String
        Try
            strSplit = split(Replace(strTranslation, "*", " " & BG2FGCOLOR_DEFAULT & " "), " ")
        Catch
            errorHandler("Can't create color table collection", _
                         Name, "bg2fgColor_mkTable_", _
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
                                     Name, "bg2fgColor_mkTable_", _
                                     "Returning Nothing")
                        Return(Nothing)                                     
                    End Try          
                    If intIndex1 = UBound(strSplit)
                        errorHandler("Error in translation string: " & _
                                     "last background color name " & _
                                     enquote(strNext) & " " & _
                                     "is not followed by a foreground color", _
                                     Name, "bg2fgColor_mkTable_", _
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
                                     Name, "bg2fgColor_mkTable_", _
                                     "Returning Nothing")
                        Return(Nothing)                                     
                    End Try                                                             
                    Try
                        colTable.Add(objColor, strNext)
                    Catch
                        errorHandler("Cannot extend translation table", _
                                     Name, "bg2fgColor_mkTable_", _
                                     Err.Number & " " & Err.Description)
                        Return(Nothing)                                     
                    End Try            
                End If            
            End If 
            intIndex1 += 1               
        Loop        
        Return(colTable)
    End Function    

    ' ----------------------------------------------------------------
    ' Break up overlong words
    '
    '
    ' This method tries to ensure that all the words of strInstring
    ' are of manageable size, so they can be edited by soft2hardParagraph
    ' and other tools that edit blank-delimited words, without "loss of
    ' signal" in the form of unnecessarily truncated words.
    '
    ' For each word that exceeds intMax length this method tries two strategies
    ' to shorten the word.
    '
    '
    '      1.  If the word is a file id it is broken into words around
    '          its slash or backslash separators. It is broken in the
    '          middle, then at 1/4 and 3/4, then 1/8, 3/8, 5/8 and 7/8,
    '          and so on.
    '
    '      2.  If strategy (1) fails, this method breaks up long
    '          words into segments by inserting a blank at 75% of
    '          the max length, and multiples of the max length.
    '
    '
    Public Shared Function breakLongWords(ByVal strInstring As String, _
                                          ByVal intMax As Integer) As String
        Dim intIndex1 As Integer
        Dim strEdited(1) As String
        Dim strOutstring As String = ""
        For intIndex1 = 1 To words(strInstring)
            strEdited(0) = word(strInstring, intIndex1)
            If Len(strEdited(0)) > intMax Then
                strEdited(1) = breakLongWords_fileidEdit_(strEdited(0), intMax)
                If strEdited(1) = strEdited(0) Then
                    strEdited(1) = breakLongWords_addBlanks_(strEdited(1), intMax)
                End If
                strEdited(0) = strEdited(1)
            End If
            append(strOutstring, " ", strEdited(0))
        Next intIndex1
        Return strOutstring
    End Function

    ' ----------------------------------------------------------------
    ' Segment with blanks on behalf of breakLongWords
    '
    '
    Private Shared Function breakLongWords_addBlanks_(ByVal strInstring As String, _
                                                      ByVal intMax As Integer) As String
        Dim intIndex1 As Integer
        Dim intStep As Integer = CInt(intMax * 0.75)
        Dim strOutstring As String
        For intIndex1 = 1 To Len(strInstring) Step intStep
            append(strOutstring, _
                   " ", _
                   Mid(strInstring, intIndex1, intStep))
        Next intIndex1
        Return (strOutstring)
    End Function

    ' ----------------------------------------------------------------
    ' Segment fileid on behalf of breakLongWords
    '
    '
    Private Shared Function breakLongWords_fileidEdit_(ByVal strInstring As String, _
                                                       ByVal intMax As Integer) As String
        Dim strSlashSep As String = breakLongWords_slashType(strInstring)
        If strSlashSep = "" Then Return strInstring
        Return (Replace(strInstring, strSlashSep, strSlashSep & " "))
    End Function

    ' ----------------------------------------------------------------
    ' Determine the file id separator (forward slash or backward
    ' slash) on behalf of breakLongWords
    '
    '
    Private Shared Function breakLongWords_slashType(ByVal strInstring As String) As String
        Dim intIndex1 As Integer = InStr(strInstring, "/")
        Dim intIndex2 As Integer = InStr(strInstring, "\")
        If intIndex1 <> 0 And intIndex2 = 0 Then Return "/"
        If intIndex2 <> 0 Then Return "\"
        Return ""
    End Function

    ' ----------------------------------------------------------------------
    ' Return True (word exists with length > max) or False on behalf of
    ' breakLongWords
    '
    '
    Private Shared Function breakLongWords_wordTooLong_(ByVal strInstring As String, _
                                                        ByVal intMax As Integer) As Boolean
        Dim intIndex1 As Integer
        For intIndex1 = 1 To words(strInstring)
            If Len(word(strInstring, intIndex1)) > intMax Then Return True
        Next intIndex1
        Return False
    End Function

    ' ----------------------------------------------------------------------
    ' Convert to one of the standard types
    '
    ' Converts an object's value to a canonical type identified by name or
    ' enumeration
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 10 23 01   Nilges         Add Boolean and Char
    '
    ' 11 30 03   Nilges         Bug: for certain conversions including real
    '                           to Byte, significance is lost with no
    '                           error indication.
    '
    '                           Return Nothing after successful conversion
    '                           if string representations of converted data
    '                           don't match.
    '
    '
    Public Overloads Shared Function canonicalTypeCast(ByVal objObjectValue As Object, _
                                                        ByVal strTypeName As String) As Object
        Return (canonicalTypeCast(objObjectValue, canonicalTypeName2Enum(strTypeName, True)))
    End Function
    Public Overloads Shared Function canonicalTypeCast(ByVal objObjectValue As Object, _
                                                        ByVal enuType As ENUcanonicalType) As Object
        Dim objConverted As Object
        Try
            Select Case enuType
                Case ENUcanonicalType.ctBoolean : objConverted = (CBool(objObjectValue))
                Case ENUcanonicalType.ctByte : objConverted = (CByte(objObjectValue))
                Case ENUcanonicalType.ctChar : objConverted = (CChar(objObjectValue))
                Case ENUcanonicalType.ctShort : objConverted = (CShort(objObjectValue))
                Case ENUcanonicalType.ctInteger : objConverted = (CInt(objObjectValue))
                Case ENUcanonicalType.ctLong : objConverted = (CLng(objObjectValue))
                Case ENUcanonicalType.ctSingle : objConverted = (CSng(objObjectValue))
                Case ENUcanonicalType.ctDouble : objConverted = (CDbl(objObjectValue))
                Case ENUcanonicalType.ctString : objConverted = (CStr(objObjectValue))
                Case ENUcanonicalType.ctUnknown : Return (Nothing)
                Case Else : Return (Nothing)
            End Select
            If objConverted.ToString <> objObjectValue.ToString Then Return (Nothing)
        Catch
            Return (Nothing)
        End Try
        Return (objConverted)
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
            Case ENUcanonicalType.ctBoolean : Return (CT_BOOLEAN)
            Case ENUcanonicalType.ctChar : Return (CT_CHAR)
            Case ENUcanonicalType.ctByte : Return (CT_BYTE)
            Case ENUcanonicalType.ctShort : Return (CT_SHORT)
            Case ENUcanonicalType.ctInteger : Return (CT_INTEGER)
            Case ENUcanonicalType.ctLong : Return (CT_LONG)
            Case ENUcanonicalType.ctSingle : Return (CT_SINGLE)
            Case ENUcanonicalType.ctDouble : Return (CT_DOUBLE)
            Case ENUcanonicalType.ctString : Return (CT_STRING)
            Case ENUcanonicalType.ctUnknown : Return (CT_UNKNOWN)
            Case Else : Return (CT_UNKNOWN)
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
    Public Overloads Shared Function canonicalTypeName2Enum(ByVal strName As String, _
                                                            ByVal booErrorHandle As Boolean) As ENUcanonicalType
        Dim enuType As ENUcanonicalType
        enuType = canonicalTypeName2Enum(strName)
        If Not booErrorHandle Then Return (enuType)
        If enuType = ENUcanonicalType.ctUnknown Then
            errorHandler("Canonical type, identified by name as " & enquote(strName) & " " & _
                         "is not valid")
        End If
        Return (enuType)
    End Function
    Public Overloads Shared Function canonicalTypeName2Enum(ByVal strName As String) As ENUcanonicalType
        Dim strNameWork As String = UCase(Trim(strName))
        If Len(strNameWork) >= 7 AndAlso Mid(strNameWork, 1, 7) = "SYSTEM." Then
            strNameWork = Mid(strNameWork, 8)
        End If
        Select Case UCase(Trim(strNameWork))
            Case CT_BOOLEAN : Return (ENUcanonicalType.ctBoolean)
            Case CT_BYTE : Return (ENUcanonicalType.ctByte)
            Case CT_INT16 : Return (ENUcanonicalType.ctShort)
            Case CT_SHORT : Return (ENUcanonicalType.ctShort)
            Case CT_INT32 : Return (ENUcanonicalType.ctInteger)
            Case CT_INTEGER : Return (ENUcanonicalType.ctInteger)
            Case CT_LONG : Return (ENUcanonicalType.ctLong)
            Case CT_INT64 : Return (ENUcanonicalType.ctLong)
            Case CT_SINGLE : Return (ENUcanonicalType.ctSingle)
            Case CT_DOUBLE : Return (ENUcanonicalType.ctDouble)
            Case CT_STRING : Return (ENUcanonicalType.ctString)
            Case Else : Return (ENUcanonicalType.ctUnknown)
        End Select
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
    '
    Public Shared Function char2Name(ByVal strChar As String) As String
        If Len(strChar) <> 1 Then
            errorHandler("strChar is not a single character", _
                         "utilities", _
                         "char2Name", _
                         "Returning a null string")
            Return ""
        End If
        If strChar >= "0" AndAlso strChar <= "9" Then Return (strChar)
        If strChar >= "A" AndAlso strChar <= "Z" Then Return (strChar)
        If strChar >= "a" AndAlso strChar <= "z" Then Return (strChar)
        Select Case strChar
            Case "~" : Return "tilde"
            Case "`" : Return "acuteAccent"
            Case "!" : Return "exclamationPoint"
            Case "@" : Return "atSign"
            Case "#" : Return "poundSign"
            Case "$" : Return "dollarSign"
            Case "%" : Return "percent"
            Case "^" : Return "carat"
            Case "&" : Return "ampersand"
            Case "*" : Return "asterisk"
            Case "(" : Return "leftParenthesis"
            Case ")" : Return "rightParenthesis"
            Case "_" : Return "underscore"
            Case "-" : Return "dash"
            Case "+" : Return "plus"
            Case "=" : Return "equals"
            Case "{" : Return "leftBrace"
            Case "[" : Return "leftBracket"
            Case "}" : Return "rightBrace"
            Case "]" : Return "rightBracket"
            Case "|" : Return "verticalStroke"
            Case "\" : Return "backslash"
            Case ":" : Return "colon"
            Case ";" : Return "semicolon"
            Case "~" : Return "tilde"
            Case """" : Return "doubleQuote"
            Case "'" : Return "singleQuote"
            Case "<" : Return "lessThan"
            Case "," : Return "comma"
            Case ">" : Return "greaterThan"
            Case "." : Return "period"
            Case "?" : Return "questionMark"
            Case "/" : Return "slash"
            Case Chr(0) : Return "NUL"
            Case vbTab : Return "Tab"
            Case Chr(10) : Return "LineFeed"
            Case Chr(13) : Return "carriageReturn"
            Case " " : Return "Blank"
            Case Else
                Return "Chr" & alignRight(CStr(AscW(strChar)), 5, "0")
        End Select
    End Function

    ' ---------------------------------------------------------------------
    ' Generate change record in my format
    '
    '
    ' This method generates a change record, in EGNSF format, that shows
    ' date, programmer and description of change in a columnar format
    ' suitable for inclusion in a comment.
    '
    ' If its parameter array strChanges is not present column headers only
    ' are generated in the format shown in the change record of this method.
    '
    ' The parameter array strChanges may be a series consisting, for each
    ' change, of date, programmer name and description, and, when it is
    ' present, the change record includes this information, and the column
    ' widths are adjusted for the widest value in each column.
    '
    '
    ' C H A N G E   R E C O R D -------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -------------------------------------------
    ' 05 15 03   Nilges         Version 1
    '
    '
    Public Shared Function changeRecord(ByVal ParamArray strChanges() As String) As String
        ' --- Build the change matrix while determining column widths
        Dim strChangeRecords(,) As String
        Try
            ReDim strChangeRecords(UBound(strChanges) \ 3, 2)
        Catch ex As Exception
            errorHandler("Cannot allocate change record as an array: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "changeRecord", _
                         "Returning a null string" & vbNewLine & vbNewLine & _
                         ex.ToString)
            Return ("")
        End Try
        Dim datChange As Date
        Dim strDate As String
        Dim intIndex0 As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intMaxDate As Integer
        Dim intMaxProgrammer As Integer
        Dim intMaxDescription As Integer
        intIndex0 = 0
        For intIndex1 = 0 To UBound(strChanges) Step 3
            strDate = strChanges(intIndex1)
            Try
                datChange = CDate(strChanges(intIndex1))
                strDate = CStr(datChange)
            Catch : End Try
            intMaxDate = Math.Max(intMaxDate, Len(strDate))
            strChangeRecords(intIndex0, 0) = strDate
            If intIndex1 < UBound(strChanges) Then
                intIndex2 = intIndex1 + 1
                strChangeRecords(intIndex0, 1) = strChanges(intIndex2)
                intMaxProgrammer = Math.Max(intMaxProgrammer, Len(strChanges(intIndex2)))
                If intIndex2 < UBound(strChanges) Then
                    strChangeRecords(intIndex0, 2) = strChanges(intIndex2 + 1)
                    intMaxDescription = Math.Max(intMaxDescription, Len(strChanges(intIndex2 + 1)))
                End If
            End If
            intIndex0 += 1
        Next intIndex1
        ' --- Build the final report
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch ex As Exception
            errorHandler("Cannot create string builder: " & Err.Number & " " & Err.Description, _
                         Name, "changeRecord", _
                         "Returning a null string" & vbNewLine & vbNewLine & _
                         ex.ToString)
            Return ("")
        End Try
        intMaxDate = Math.Max(intMaxDate, 4)
        intMaxProgrammer = Math.Max(intMaxProgrammer, 20)
        intMaxDescription = Math.Max(intMaxDescription, 21)
        Dim strColheader As String = copies("-", intMaxDate) & " " & _
                                     copies("-", intMaxProgrammer) & "     " & _
                                     copies("-", intMaxDescription)
        objStringBuilder.Append(alignLeft("C H A N G E   R E C O R D ", Len(strColheader), "-") & _
                                vbNewLine & _
                                alignCenter("DATE", intMaxDate) & " " & _
                                alignCenter("PROGRAMMER/GENERATOR", intMaxProgrammer) & "     " & _
                                alignCenter("D E S C R I P T I O N", intMaxDescription) & _
                                vbNewLine & _
                                strColheader)
        For intIndex1 = 0 To UBound(strChangeRecords, 1)
            append(objStringBuilder, _
                   vbNewLine, _
                   alignLeft(strChangeRecords(intIndex1, 0), intMaxDate) & " " & _
                   alignLeft(strChangeRecords(intIndex1, 1), intMaxProgrammer) & "     " & _
                   alignLeft(strChangeRecords(intIndex1, 2), intMaxDescription))
        Next intIndex1
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' Return common regular expressions
    '
    '
    ' This method returns the following regular expressions (as strings), 
    ' where the name is identified in the strREname parameter:
    '
    '
    '      *  identifierCOM: captures identifiers (Visual Basic 6 syntax)
    '
    '      *  identifierNET: captures identifiers (Visual Basic .Net syntax)
    '
    '      *  graphicalCharacterSet: returns the set of characters that are available
    '         on standard USA keyboards as a regular expression of the form
    '         ([charset])
    '
    '      *  newlineWebWindows: returns a newline when it is not known 
    '         whether Web newlines (linefeed) or Windows newlines (carriage 
    '         return and linefeed) are in use
    '
    '      *  VBcomment: captures comments (Visual Basic syntax).  The comment
    '         cannot contain white space characters other than blanks (characters,
    '         that is, in the range hex 0 through hex 19.)
    '
    '      *  VBcommentLine: captures comment lines (Visual Basic syntax).
    '
    '      *  VBcommentBlock: captures comment blocks, where this is a series
    '         of comment lines (Visual Basic syntax)
    '
    '      *  VBparameterDefCOM: captures Visual Basic 6 parameter declarations
    '
    '      *  VBparameterDefNET: captures Visual Basic .Net parameter declarations
    '
    '      *  VBprocedureHeaderCOM: captures procedure headers (Visual Basic
    '         6 syntax). 
    '
    '         The procedure header cannot be on the first line of its input 
    '         file or string. The header text up through Sub, Function,
    '         Property Get, Property Let or Property Set, but before the
    '         name, is captured. The text must be formatted as found
    '         in the code window: there cannot be extra spaces between 
    '         elements removed: all keywords must be in the Proper case.
    '
    '      *  VBprocedureHeaderNET: captures procedure headers (Visual Basic
    '         .Net syntax). 
    '
    '         The procedure header cannot be on the first line of its input 
    '         file or string. The header text up through Sub, Function,
    '         Property Get, Property Let or Property Set, but before the
    '         name, is captured. The text must be formatted as found
    '         in the code window: there cannot be extra spaces between 
    '         elements removed: all keywords must be in the Proper case.
    '
    '         In addition, and unlike procedureHeaderCOM, the procedure header
    '         must start with a scope keyword: Public, Private, and Friend.
    '
    '
    ' In order to avoid problems in reusing these regular expressions inside larger
    ' constructs, each regular expression returned is parenthesized.
    '
    ' Alternatively this method may return a new collection of all the regular
    ' expressions it is able to support when it is called with the overloaded
    ' syntax commonRegularExpressions (with no nonoptional parameters). 
    '
    ' When commonRegularExpressions is called with no parameters at all (optional
    ' or nonoptional), this method creates and returns a collection. Each item is a 
    ' subcollection containing three items:
    '
    '
    '      *  The first element of each subcollection is the name of the regular
    '         expression (identifierCOM, identifierNET, and so forth, as above).
    '
    '      *  The second element of each subcollection is the regular expression
    '         itself.
    '
    '      *  The third element of each subcollection explains what the regular
    '         expression accomplishes.
    '
    '
    ' The key of each element in the returned collection is the name of the
    ' regular expression in the subcollection.
    '
    ' An optional booExtension parameter returns test cases; this is primarily
    ' for internal use. When commonRegularExpressions is called with 
    ' booExtension:=True, this method creates and returns a five-element collection 
    ' in this format.
    '
    '
    '      *  The first element of each subcollection is the name of the regular
    '         expression (identifierCOM, identifierNET, and so forth, as above).
    '
    '      *  The second element of each subcollection is the regular expression
    '         itself.
    '
    '      *  The third element of each subcollection explains what the regular
    '         expression accomplishes.
    '
    '      *  The fourth element of each subcollection contains a test string.
    '
    '      *  The fifth element of the subcollection contains expected test
    '         results, from applying the regular expression repeatedly to the
    '         test string until no match occurs.
    '
    '         The expected test results are an unkeyed sub-sub-collection containing
    '         an even number of entries. The first entry of each pair is the 
    '         expected index of a match (based on 0 and not one) and the second is
    '         the expected match length. 
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 05 23 03   Nilges         1.  Removed newline from the procedure
    '                               header regular expressions
    '
    '                           2.  Added comment/comment line/comment 
    '                               block/newlineOrStart/newlineOrEnd
    '
    '                           3.  Prefixed VB language parsers with
    '                               VB
    '
    '                           4.  Surround each regular expression
    '                               with parentheses
    '
    ' 05 24 03   Nilges         1.  Added return of collection that
    '                               includes test cases
    '
    ' 05 25 03   Nilges         1.  Added vbParameterDefCOM/NET
    '
    ' 06 10 03   Nilges         1.  Added graphicalCharacters
    '
    '
    ' --- Return collection
    Public Overloads Shared Function commonRegularExpressions _
                                     (Optional ByVal booExtension As Boolean = False) As Collection
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch ex As Exception
            errorHandler("Cannot create collection: " & Err.Number & " " & Err.Description, _
                         Name, "commonRegularExpressions", _
                         "Returning Nothing")
            Return (Nothing)
        End Try
        Dim strSplit() As String
        Try
            strSplit = Split(commonRegularExpressions_substitute_(COMMON_REGULAR_EXPRESSIONS), _
                             COMMON_REGULAR_EXPRESSIONS_SEP)
        Catch ex As Exception
            errorHandler("Cannot split string: " & Err.Number & " " & Err.Description, _
                         Name, "commonRegularExpressions", _
                         "Returning empty collection")
            Return (colNew)
        End Try
        Dim colSubEntry(1) As Collection
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strSplit2() As String
        For intIndex1 = 0 To UBound(strSplit) Step 4
            Try
                colSubEntry(0) = New Collection
                With colSubEntry(0)
                    .Add(word(strSplit(intIndex1), 1))      ' re name
                    .Add(phrase(strSplit(intIndex1), 2))    ' re
                    .Add(strSplit(intIndex1 + 1))           ' re explanation
                    If booExtension Then
                        .Add(strSplit(intIndex1 + 2))       ' re test string
                        colSubEntry(1) = New Collection
                        With colSubEntry(1)
                            strSplit2 = Split(strSplit(intIndex1 + 3))
                            For intIndex2 = 0 To UBound(strSplit2) Step 2
                                colSubEntry(1).Add(CInt(strSplit2(intIndex2)))
                                colSubEntry(1).Add(CInt(strSplit2(intIndex2 + 1)))
                            Next intIndex2
                        End With
                        .Add(colSubEntry(1))       ' re test results expected
                    End If
                    colNew.Add(colSubEntry(0), CStr(.Item(1)))
                End With
            Catch ex As Exception
                colSubEntry = Nothing
                errorHandler("Cannot create subcollection (or extend main collection): " & _
                             Err.Number & " " & Err.Description, _
                             Name, "commonRegularExpressions", _
                             "Returning available regular expressions")
                Return (colNew)
            End Try
        Next intIndex1
        Return (colNew)
    End Function
    ' --- Return named regular expression
    Public Overloads Shared Function commonRegularExpressions(ByVal strREname As String) _
           As String
        Dim colRE As Collection = commonRegularExpressions()
        If (colRE Is Nothing) Then Return ("")
        Dim strRE As String
        Try
            strRE = CStr(CType(colRE.Item(strREname), Collection).Item(2))
        Catch ex As Exception
            errorHandler("Undefined regular expression" & enquote(strREname), _
                         Name, "commonRegularExpressions", _
                         "Returning null string")
            Return ("")
        End Try
        Return ("(" & strRE & ")")
    End Function

    ' -----------------------------------------------------------------
    ' Substitute keywords on behalf of commonRegularExpressions
    '
    '
    Private Shared Function commonRegularExpressions_substitute_ _
                            (ByVal strCRE As String) As String
        Dim strTests As String
        Dim strExpectedResults As String
        commonRegularExpressions_substitute__extendCommon_(strTests, strExpectedResults)
        commonRegularExpressions_substitute__extendProperty_(strTests, strExpectedResults, "Get")
        commonRegularExpressions_substitute__extendProperty_(strTests, strExpectedResults, "Let")
        commonRegularExpressions_substitute__extendProperty_(strTests, strExpectedResults, "Set")
        Dim strSub As String = Replace(strCRE, _
                                       COMMON_REGULAR_EXPRESSIONS_LD & _
                                       "vbProcedureHeaderCOMtests" & _
                                       COMMON_REGULAR_EXPRESSIONS_RD, _
                                       strTests & _
                                       COMMON_REGULAR_EXPRESSIONS_SEP & _
                                       strExpectedResults)
        strTests = "" : strExpectedResults = ""
        commonRegularExpressions_substitute__extendCommon_(strTests, strExpectedResults)
        commonRegularExpressions_substitute__extendProperty_(strTests, strExpectedResults, "")
        strSub = Replace(strSub, _
                         COMMON_REGULAR_EXPRESSIONS_LD & _
                         "vbProcedureHeaderNettests" & _
                         COMMON_REGULAR_EXPRESSIONS_RD, _
                         strTests & _
                         COMMON_REGULAR_EXPRESSIONS_SEP & _
                         strExpectedResults)
        Return (strSub)
    End Function

    ' ----------------------------------------------------------------------
    ' Extend the cre tests on behalf of commonRegularExpressions_substitute
    '
    '
    Private Shared Sub commonRegularExpressions_substitute__extend_ _
                       (ByRef strTests As String, _
                        ByRef strExpectedResults As String, _
                        ByVal strNextTest As String)
        Dim strNewline As String = CStr(IIf(Rnd() > 0.25, vbNewLine, vbLf))
        Dim strNextTestWork As String = copies(" ", CInt(Rnd() * 7)) & _
                                        strNextTest & _
                                        " Procedure" & CStr(CInt(Rnd() * 1024))
        strExpectedResults = append(strExpectedResults, _
                                    " ", _
                                    Len(strTests) & " " & Len(strNextTestWork) + Len(strNewline))
        strTests &= strNewline & strNextTestWork
    End Sub

    ' ----------------------------------------------------------------------
    ' Extend the cre tests with syntax common to COM and Net, on behalf of 
    ' commonRegularExpressions_substitute
    '
    '
    Private Shared Sub commonRegularExpressions_substitute__extendCommon_ _
                       (ByRef strTests As String, _
                        ByRef strExpectedResults As String)
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Sub")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Public Sub")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Private Sub")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Friend Sub")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Function")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Public Function")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Private Function")
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Friend Function")
    End Sub

    ' ----------------------------------------------------------------------
    ' Extend the cre tests with COM property syntax, on behalf of 
    ' commonRegularExpressions_substitute
    '
    '
    Private Shared Sub commonRegularExpressions_substitute__extendProperty_ _
                       (ByRef strTests As String, _
                        ByRef strExpectedResults As String, _
                        ByVal strGetLetSet As String)
        Dim strGetLetSetWork As String = CStr(IIf(strGetLetSet = "", "", " ")) & _
                                         strGetLetSet
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Property" & strGetLetSetWork)
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Public Property" & strGetLetSetWork)
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Private Property" & strGetLetSetWork)
        commonRegularExpressions_substitute__extend_(strTests, _
                                                     strExpectedResults, _
                                                     "Friend Property" & strGetLetSetWork)
    End Sub

    ' ----------------------------------------------------------------------
    ' Test common regular expressions
    '
    '
    ' This method tests the common regular expressions, and it places a
    ' test report in its reference string parameter strReport. It returns
    ' True on a successful test and False on failure.
    '
    '
    Public Shared Function commonRegularExpressionTester(ByRef strReport As String) _
           As Boolean
        strReport = "Unknown test failure"
        Dim colRE As Collection = commonRegularExpressions(booExtension:=True)
        If (colRE Is Nothing) Then Return (False)
        strReport = "COMMON REGULAR EXPRESSION TEST     " & Now & _
                    vbNewLine & vbNewLine & vbNewLine
        Dim colREentry As Collection
        Dim colREsubEntry As Collection
        Dim intErrors As Integer
        With colRE
            Dim intExpectedIndex As Integer
            Dim intExpectedLength As Integer
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            Dim intStartAt As Integer
            Dim objRE As System.Text.RegularExpressions.Regex
            Dim objREmatch As System.Text.RegularExpressions.Match
            Dim strRE As String
            Dim strTest As String
            Dim strTestDisplay As String
            For intIndex1 = 1 To .Count
                colREentry = CType(.Item(intIndex1), Collection)
                With colREentry
                    strRE = CStr(.Item(2))
                    strTest = CStr(.Item(4))
                    strReport = append(strReport, _
                                       vbNewLine & vbNewLine & vbNewLine, _
                                       "Testing the " & _
                                       CStr(.Item(1)) & " " & _
                                       "regular expression " & _
                                       enquote(strRE) & ": " & _
                                       CStr(.Item(3)) & vbNewLine & _
                                       "Test string: " & string2Display(strTest))
                    Try
                        objRE = New System.Text.RegularExpressions.Regex(strRE)
                    Catch ex As Exception
                        errorHandler("Can't create or can't apply regular expression: " & _
                                     Err.Number & " " & Err.Description, _
                                     Name, "commonRegularExpressionTester", _
                                     "Returning failed test" & vbNewLine & vbNewLine & _
                                     ex.ToString)
                    End Try
                    colREsubEntry = CType(.Item(5), Collection)
                    With colREsubEntry
                        intStartAt = 0
                        For intIndex2 = 1 To .Count Step 2
                            strReport = append(strReport, _
                                               vbNewLine & vbNewLine, _
                                               "Test case " & intIndex2 \ 2 + 1 & " of " & .Count \ 2)
                            intExpectedIndex = CInt(.Item(intIndex2))
                            intExpectedLength = CInt(.Item(intIndex2 + 1))
                            objREmatch = objRE.Match(strTest, intStartAt)
                            If (objREmatch Is Nothing) Then Exit For
                            With objREmatch
                                If .Length = 0 Then Exit For
                                intStartAt = .Index + .Length
                                strReport = append(strReport, _
                                                    "    " & vbNewLine, _
                                                    "Expected index (zero based) is " & _
                                                    intExpectedIndex & ": " & _
                                                    "expected length is " & _
                                                    intExpectedLength)
                                strReport = append(strReport, _
                                                    "    " & vbNewLine, _
                                                    "Actual index (zero based) is " & _
                                                    .Index & ": " & _
                                                    "actual length is " & _
                                                    .Length)
                                If .Index <> intExpectedIndex _
                                   OrElse _
                                   .Length <> intExpectedLength Then
                                    strReport = append(strReport, _
                                                       vbNewLine, _
                                                       "Expected and actual values do NOT match")
                                    intErrors += 1
                                End If
                            End With
                        Next intIndex2
                        If intIndex2 <> .Count + 1 Then
                            strReport = append(strReport, _
                                               vbNewLine, _
                                               "Expected and actual matches are not the same")
                            intErrors += 1
                        End If
                    End With
                End With
            Next intIndex1
            strReport = append(strReport, _
                               vbNewLine & vbNewLine & vbNewLine, _
                               CStr(IIf(intErrors = 0, "Test passed", "Test failed")))
            Return (intErrors = 0)
        End With
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
        Return (copies(intCopies, strInstring))
    End Function
    ' --- Normal syntax
    Public Overloads Shared Function copies(ByVal intCopies As Integer, ByVal strInstring As String) As String
        Dim intIndex1 As Integer
        If intCopies < 0 Then
            errorHandler("Invalid intCopies parameter " & intCopies)
            Return ("")
        End If
        Dim objStringBuilder As New System.Text.StringBuilder("")
        If Len(strInstring) = 1 Then
            objStringBuilder.Append(CChar(strInstring), intCopies)
        Else
            For intIndex1 = 1 To intCopies
                objStringBuilder.Append(strInstring)
            Next intIndex1
        End If
        Return (objStringBuilder.ToString)
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
    ' 05 13 03   Nilges         1.  Replaced msgbox by errorHandler
    '                           2.  Return unknown data type when data type
    '                               is not valid. 
    '
    ' 
    ' --- Determines the data type
    Public Overloads Shared Function datatype(ByVal strInstring As String) _
           As ENUdatatype
        Return (datatype_(strInstring, _
               ENUdatatype.UnspecifiedDatatype, _
               Nothing, _
               False))
    End Function
    ' --- Tests for the enumerated data type
    Public Overloads Shared Function datatype(ByVal strInstring As String, _
                                              ByVal enuExpectedDatatype As ENUdatatype) _
           As Boolean
        Return (datatype_(strInstring, enuExpectedDatatype, Nothing, False) _
               = _
               ENUdatatype.TrueDatatype)
    End Function
    ' --- Tests for the named data type  
    Public Overloads Shared Function datatype(ByVal strInstring As String, _
                                              ByVal strExpectedDatatype As String) _
           As Boolean
        Return (datatype(strInstring, strExpectedDatatype, Nothing))
    End Function
    ' --- Sets the converted value
    Public Overloads Shared Function datatype(ByVal strInstring As String, _
                                              ByRef objConverted As Object) _
                     As ENUdatatype
        Return (datatype_(strInstring, _
               ENUdatatype.UnspecifiedDatatype, _
               objConverted, True))
    End Function
    ' --- Tests for the enumerated data type and sets the converted value
    Public Overloads Shared Function datatype(ByVal strInstring As String, _
                                              ByVal enuExpectedDatatype As ENUdatatype, _
                                              ByRef objConverted As Object) _
                     As Boolean
        Return (datatype_(strInstring, enuExpectedDatatype, objConverted, True) _
               = _
               ENUdatatype.TrueDatatype)
    End Function
    ' --- Tests for the named data type and sets the converted value
    Public Overloads Shared Function datatype(ByVal strInstring As String, _
                                              ByVal strExpectedDatatype As String, _
                                              ByRef objConverted As Object) _
           As Boolean
        Dim enuExpectedDatatype As ENUdatatype = datatypeName2Enum(strExpectedDatatype)
        If enuExpectedDatatype = ENUdatatype.UnspecifiedDatatype Then
            errorHandler("Invalid data type name " & enquote(strExpectedDatatype), _
                         Name, "datatype", _
                         "Returning False")
            Return (False)
        End If
        Return (datatype(strInstring, enuExpectedDatatype, objConverted))
    End Function
    ' --- Common functionality
    Private Overloads Shared Function datatype_(ByVal strInstring As String, _
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
        dblMin = 0 : dblMax = 0
        If enuExpectedDatatype <> ENUdatatype.UnspecifiedDatatype Then
            booDatatypeValue = False
            Select Case enuExpectedDatatype
                Case ENUdatatype.AlphabeticDatatype
                    If verify(UCase(strInstring), _
                              "ABCDEFGHIJKLMNOPQRSTUVWXYZ") _
                    = _
                    0 Then
                        booDatatypeValue = True
                        objValue = CStr(strInstring)
                    End If
                Case ENUdatatype.ByteDatatype
                    booDatatypeValue = True
                    Try
                        objValue = CByte(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.DateDatatype
                    If IsDate(strInstring) Then
                        booDatatypeValue = True
                        objValue = CDate(strInstring)
                    End If
                Case ENUdatatype.DoubleDatatype
                    booDatatypeValue = True
                    Try
                        objValue = CDbl(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                Case ENUdatatype.IntegerDatatype
                    booDatatypeValue = True
                    Try
                        objValue = CInt(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.LabelDatatype
                    If InStr(strInstring, " ") = 0 Then
                        booDatatypeValue = True : objValue = strInstring
                    End If
                Case ENUdatatype.LongDatatype
                    booDatatypeValue = True
                    Try
                        objValue = CLng(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.LowerDatatype
                    If verify(strInstring, "abcdefghijklmnopqrstuvwxyz") = 0 Then
                        booDatatypeValue = True : objValue = strInstring
                    End If
                Case ENUdatatype.PercentValueDatatype
                    If datatype(strInstring, ENUdatatype.ByteDatatype, objValue) Then
                        If CInt(objValue) <= 100 Then
                            booDatatypeValue = True
                        End If
                    End If
                Case ENUdatatype.SerialDatatype
                    If verify(strInstring, "0123456789") = 0 Then
                        booDatatypeValue = True : objValue = strInstring
                    End If
                Case ENUdatatype.ShortDatatype
                    booDatatypeValue = True
                    Try
                        objValue = CShort(strInstring)
                    Catch
                        Return (ENUdatatype.FalseDatatype)
                    End Try
                    booDatatypeValue = booDatatypeValue _
                                       AndAlso _
                                       datatype_integerCheck_(strInstring, objValue)
                Case ENUdatatype.SignumDatatype
                    If datatype(strInstring, ENUdatatype.ShortDatatype, objValue) Then
                        booDatatypeValue = True
                        Select Case objValue
                            Case 1
                            Case 0
                            Case -1
                            Case Else : Return (ENUdatatype.FalseDatatype)
                        End Select
                    End If
                Case ENUdatatype.SingleDatatype
                    booDatatypeValue = True
                    Try
                        objValue = CSng(strInstring)
                    Catch
                        booDatatypeValue = False
                    End Try
                Case ENUdatatype.StringDatatype
                    booDatatypeValue = True : objValue = strInstring
                Case ENUdatatype.UnsignedIntegerDatatype
                    If datatype(strInstring, _
                                ENUdatatype.IntegerDatatype, _
                                objValue) Then
                        booDatatypeValue = (CInt(objValue) >= 0)
                    End If
                Case ENUdatatype.UnsignedLongDatatype
                    If datatype(strInstring, ENUdatatype.LongDatatype, objValue) Then
                        booDatatypeValue = (CInt(objValue) >= 0)
                    End If
                Case ENUdatatype.UnsignedShortDatatype
                    If datatype(strInstring, ENUdatatype.ShortDatatype, objValue) Then
                        booDatatypeValue = (CInt(objValue) >= 0)
                    End If
                Case ENUdatatype.UpperDatatype
                    If verify(strInstring, _
                              "ABCDEFGHIJKLMNOPQRSTUVWXYZ") _
                    = _
                    0 Then
                        booDatatypeValue = True : objValue = strInstring
                    End If
                Case ENUdatatype.VBIdentifierDatatype
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
                Case Else
                    errorHandler("Error in datatype: invalid datatype code", _
                                 Name, "datatype", _
                                 "Returning the unknown data type")
                    Return (ENUdatatype.UnknownDatatype)
            End Select
            If booReturnConverted Then
                objConverted = IIf(booDatatypeValue, objValue, Nothing)
            End If
            enuDatatypeValue = CType(IIf(booDatatypeValue, _
                                         ENUdatatype.TrueDatatype, _
                                         ENUdatatype.FalseDatatype), _
                                     ENUdatatype)
        Else
            If datatype(strInstring, ENUdatatype.ByteDatatype) Then
                enuDatatypeValue = ENUdatatype.ByteDatatype
            ElseIf datatype(strInstring, ENUdatatype.ShortDatatype) Then
                enuDatatypeValue = ENUdatatype.ShortDatatype
            ElseIf datatype(strInstring, ENUdatatype.IntegerDatatype) Then
                enuDatatypeValue = ENUdatatype.IntegerDatatype
            ElseIf datatype(strInstring, ENUdatatype.LongDatatype) Then
                enuDatatypeValue = ENUdatatype.LongDatatype
            ElseIf datatype(strInstring, ENUdatatype.SingleDatatype) Then
                enuDatatypeValue = ENUdatatype.SingleDatatype
            ElseIf datatype(strInstring, ENUdatatype.DoubleDatatype) Then
                enuDatatypeValue = ENUdatatype.DoubleDatatype
            ElseIf datatype(strInstring, ENUdatatype.LowerDatatype) Then
                enuDatatypeValue = ENUdatatype.LowerDatatype
            ElseIf datatype(strInstring, ENUdatatype.UpperDatatype) Then
                enuDatatypeValue = ENUdatatype.UpperDatatype
            ElseIf datatype(strInstring, ENUdatatype.AlphabeticDatatype) Then
                enuDatatypeValue = ENUdatatype.AlphabeticDatatype
            ElseIf datatype(strInstring, ENUdatatype.DateDatatype) Then
                enuDatatypeValue = ENUdatatype.DateDatatype
            ElseIf datatype(strInstring, ENUdatatype.LabelDatatype) Then
                enuDatatypeValue = ENUdatatype.LabelDatatype
            ElseIf datatype(strInstring, ENUdatatype.StringDatatype) Then
                enuDatatypeValue = ENUdatatype.StringDatatype
            Else
                enuDatatypeValue = ENUdatatype.UnknownDatatype
            End If
        End If
        Return (enuDatatypeValue)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of datatype, check for an integer
    '
    '
    Private Shared Function datatype_integerCheck_(ByVal strInstring As String, _
                                                   ByRef objValue As Object) As Boolean
        If CDbl(strInstring) = CLng(strInstring) Then Return (True)
        objValue = Nothing
        Return (False)
    End Function

    ' ----------------------------------------------------------------------
    ' Data type enumerator to name
    '
    '
    Public Shared Function datatypeEnum2Name(ByVal enuDatatypeEnum As ENUdatatype) As String
        Select Case enuDatatypeEnum
            Case ENUdatatype.AlphabeticDatatype : Return ("AlphabeticDatatype")
            Case ENUdatatype.ByteDatatype : Return ("ByteDatatype")
            Case ENUdatatype.DateDatatype : Return ("DateDatatype")
            Case ENUdatatype.DoubleDatatype : Return ("DoubleDatatype")
            Case ENUdatatype.FalseDatatype : Return ("FalseDatatype")
            Case ENUdatatype.IntegerDatatype : Return ("IntegerDatatype")
            Case ENUdatatype.LabelDatatype : Return ("LabelDatatype")
            Case ENUdatatype.LongDatatype : Return ("LongDatatype")
            Case ENUdatatype.LowerDatatype : Return ("LowerDatatype")
            Case ENUdatatype.PercentValueDatatype : Return ("PercentValueDatatype")
            Case ENUdatatype.SerialDatatype : Return ("SerialDatatype")
            Case ENUdatatype.SignumDatatype : Return ("SignumDatatype")
            Case ENUdatatype.ShortDatatype : Return ("ShortDatatype")
            Case ENUdatatype.SingleDatatype : Return ("SingleDatatype")
            Case ENUdatatype.StringDatatype : Return ("StringDatatype")
            Case ENUdatatype.TrueDatatype : Return ("TrueDatatype")
            Case ENUdatatype.UnknownDatatype : Return ("UnknownDatatype")
            Case ENUdatatype.UnsignedIntegerDatatype : Return ("UnsignedIntegerDatatype")
            Case ENUdatatype.UnsignedLongDatatype : Return ("UnsignedLongDatatype")
            Case ENUdatatype.UnsignedShortDatatype : Return ("UnsignedShortDatatype")
            Case ENUdatatype.UnspecifiedDatatype : Return ("UnspecifiedDatatype")
            Case ENUdatatype.UpperDatatype : Return ("UpperDatatype")
            Case ENUdatatype.VBIdentifierDatatype : Return ("VBIdentifierDatatype")
            Case Else : Return ("?")
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Data type enumerator to name
    '
    '
    Public Shared Function datatypeName2Enum(ByVal strName As String) As ENUdatatype
        Select Case LCase(Trim(strName))
            Case "alphabeticdatatype" : Return (ENUdatatype.AlphabeticDatatype)
            Case "bytedatatype" : Return (ENUdatatype.ByteDatatype)
            Case "datedatatype" : Return (ENUdatatype.DateDatatype)
            Case "doubledatatype" : Return (ENUdatatype.DoubleDatatype)
            Case "falsedatatype" : Return (ENUdatatype.FalseDatatype)
            Case "integerdatatype" : Return (ENUdatatype.IntegerDatatype)
            Case "labeldatatype" : Return (ENUdatatype.LabelDatatype)
            Case "longdatatype" : Return (ENUdatatype.LongDatatype)
            Case "lowerdatatype" : Return (ENUdatatype.LowerDatatype)
            Case "percentvaluedatatype" : Return (ENUdatatype.PercentValueDatatype)
            Case "serialdatatype" : Return (ENUdatatype.SerialDatatype)
            Case "shortdatatype" : Return (ENUdatatype.ShortDatatype)
            Case "signumdatatype" : Return (ENUdatatype.SignumDatatype)
            Case "singledatatype" : Return (ENUdatatype.SingleDatatype)
            Case "stringdatatype" : Return (ENUdatatype.StringDatatype)
            Case "truedatatype" : Return (ENUdatatype.TrueDatatype)
            Case "unknowndatatype" : Return (ENUdatatype.UnknownDatatype)
            Case "unsignedintegerdatatype" : Return (ENUdatatype.UnsignedIntegerDatatype)
            Case "unsignedlongdatatype" : Return (ENUdatatype.UnsignedLongDatatype)
            Case "unsignedshortdatatype" : Return (ENUdatatype.UnsignedShortDatatype)
            Case "unspecifieddatatype" : Return (ENUdatatype.UnspecifiedDatatype)
            Case "upperdatatype" : Return (ENUdatatype.UpperDatatype)
            Case "vbidentifierdatatype" : Return (ENUdatatype.VBIdentifierDatatype)
            Case Else : Return (ENUdatatype.UnspecifiedDatatype)
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Dequote a string when it is quoted
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 04 09 04   Nilges         1. Bug: returned Nothing and not a null
    '                              string for two adjacent quotes
    '                              1.1 Added special test
    '
    Public Shared Function dequote(ByVal strInstring As String) As String
        If Not isQuoted(strInstring) Then Return (strInstring)
        If strInstring = """""" Then Return ""
        Return (Replace(Mid(strInstring, 2, Len(strInstring) - 2), Chr(34) & Chr(34), Chr(34)))
    End Function

    ' ----------------------------------------------------------------------
    ' Determine the newline string
    '
    '
    ' This method tests strInstring to see if all strings consisting of
    ' carriage return and newline are either repetitions of the carriage
    ' return and the newline, or repetitions of the newline. This method
    ' returns carriage return and newline in the former case, or linefeed
    ' in the latter case. In any other situation, this method returns
    ' a null string to indicate failure to determine the newline string.
    '
    ' Optionally, to save time for large files, use the intMaxLength parameter
    ' to scan only the characters at the beginning of the file, up to
    ' intMaxLength.
    '
    '
    Public Shared Function determineNewline(ByVal strInstring As String, _
                                            Optional ByVal intMaxLength As Integer = -1) _
                           As String
        If intMaxLength < -1 Then
            errorHandler("intMaxLength is not valid", _
                         Name, _
                         "determineNewline", _
                         "Returning a null string")
            Return ("")
        End If
        Dim intLength As Integer = CInt(IIf(intMaxLength = -1, Len(strInstring), intMaxLength))
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim strNext As String
        Dim strNewline As String
        Dim strNextNewline As String
        Do While intIndex1 <= intLength
            intIndex1 = verify(strInstring, vbCr & vbLf, intIndex1, True)
            If intIndex1 = 0 Then Exit Do
            intIndex2 = verify(strInstring & "A", vbCr & vbLf, intIndex1, False)
            strNext = Mid(strInstring, intIndex1, intIndex2 - intIndex1)
            If strNext = copies(vbCrLf, Len(strNext) \ 2) Then
                strNextNewline = vbCrLf
            ElseIf strNext = copies(vbLf, Len(strNext)) Then
                strNextNewline = vbLf
            End If
            If strNextNewline <> strNewline Then
                If strNewline = "" Then
                    strNewline = strNextNewline
                Else
                    Return ("")
                End If
            End If
            intIndex1 = intIndex2
        Loop
        Return (strNewline)
    End Function

    ' ----------------------------------------------------------------------
    ' Return the parent Dewey decimal number
    '
    '
    Public Shared Function deweyParent(ByVal strDewey As String) As String
        Dim intItems As Integer = items(strDewey, ".", False)
        If intItems = 0 Then Return ""
        Return itemPhrase(strDewey, 1, intItems - 1, ".", False)
    End Function

    ' ----------------------------------------------------------------------
    ' Return True if a Windows directory exists: otherwise return False
    '
    '
    Public Shared Function directoryExists(ByVal strDirectory As String) As Boolean
        Dim booDirectoryExists As Boolean = True
        Dim strSaveDirectory As String = CurDir()
        Try
            ChDir(strDirectory)
        Catch
            booDirectoryExists = False
        End Try
        ChDir(strSaveDirectory)
        Return (booDirectoryExists)
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
    '         Note that over and above the VBexpression format this method
    '         accepts True and False as the string form of the display.
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
    ' 05 22 03   Nilges         Bug: error in application of regular
    '                           expression to find XML character value
    '
    ' 02 10 04   Nilges         Allow the vbExpression to contain True or
    '                           False, and convert to Boolean
    '
    ' --- Determine the syntax
    Public Overloads Shared Function display2String(ByVal strInstring As String, _
                                                    Optional ByVal booVariableLenVals _
                                                             As Boolean = False) As String
        Return (display2String(strInstring, ENUdisplay2StringSyntax.Determine, booVariableLenVals:=booVariableLenVals))
    End Function
    ' --- Specifies the syntax as a string
    Public Overloads Shared Function display2String(ByVal strInstring As String, _
                                                    ByVal strSyntax As String, _
                                                    Optional ByVal booVariableLenVals _
                                                             As Boolean = False) As String
        Dim enuSyntax As ENUdisplay2StringSyntax = display2String_syntax2Enum_(strSyntax)
        If enuSyntax = ENUdisplay2StringSyntax.Invalid Then Return ("")
        Return (display2String(strInstring, enuSyntax, booVariableLenVals:=booVariableLenVals))
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
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim shAsc As Short
        Dim strEscape As String
        Dim strNextChar As String
        Dim strOutstring As String
        Select Case enuSyntax
            Case ENUdisplay2StringSyntax.C
                booCsyntax = True
            Case ENUdisplay2StringSyntax.Determine
                booCsyntax = (InStr(strInstring, "\") <> 0)
                booXMLsyntax = (InStr(strInstring, "&#") <> 0)
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
                If booVBexpressionSyntax Then Return (strOutstring)
            Case ENUdisplay2StringSyntax.XML
                booXMLsyntax = True
            Case ENUdisplay2StringSyntax.VBExpression
                If strInstring = "" Then Return ("")
                strOutstring = _
                display2String_VBexpression2String_(strInstring, _
                                                    booVBexpressionSyntax, _
                                                    booCondensed)
                If Not booVBexpressionSyntax OrElse booCondensed Then
                    errorHandler("Cannot convert string " & vbNewLine & vbNewLine & _
                                 strInstring & vbNewLine & vbNewLine & _
                                 "assuming uncondensed VB expression syntax", _
                                 "display2String")
                    Return ("")
                End If
                Return (strOutstring)
            Case ENUdisplay2StringSyntax.VBExpressionCondensed
                If strInstring = "" Then Return ("")
                strOutstring = _
                display2String_VBexpression2String_(strInstring, _
                                                    booVBexpressionSyntax, _
                                                    booCondensed)
                If Not booVBexpressionSyntax Then
                    errorHandler("Cannot convert string " & vbNewLine & vbNewLine & _
                                 strInstring & vbNewLine & vbNewLine & _
                                 "assuming condensed VB expression syntax")
                    Return ("")
                End If
                Return (strOutstring)
            Case Else
                errorHandler("Unexpected enuSyntax " & enuSyntax)
                Return ("")
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
                    intIndex3 = CInt(Math.Min(2, intIndex3 - intIndex2))
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
                    strEscape = vbNewLine
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
                        New System.Text.RegularExpressions.Regex("[\+\-]{0,1}[0123456789]+")
                    objRegExMatch = objRegEx.Match(strInstring, intIndex2 - 1)
                    intIndex3 = objRegExMatch.Index + 1
                    If intIndex3 = intIndex2 + 2 Then
                        intIndex4 = CInt(Math.Min(objRegExMatch.Length, 5))
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
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert syntax name to enumerator
    '
    '
    Private Shared Function display2String_syntax2Enum_(ByVal strSyntax As String) As ENUdisplay2StringSyntax
        Select Case UCase(Trim(strSyntax))
            Case "C" : Return ENUdisplay2StringSyntax.C
            Case "XML" : Return ENUdisplay2StringSyntax.XML
            Case "VBEXPRESSION" : Return ENUdisplay2StringSyntax.VBExpression
            Case "VBEXPRESSIONCONDENSED" : Return ENUdisplay2StringSyntax.VBExpressionCondensed
            Case "DETERMINE" : Return ENUdisplay2StringSyntax.Determine
            Case Else
                errorHandler("Invalid syntax " & enquote(strSyntax), Name, "display2String", _
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
        Select Case Trim(UCase(strInstring))
            Case "TRUE"
                booValid = True : booCondensed = False : Return strInstring
            Case "FALSE"
                booValid = True : booCondensed = False : Return strInstring
        End Select
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
                    Catch : End Try
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
                        Catch : End Try
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
                        Catch : End Try
                    End If
                ElseIf Mid(strToken, 1, 12) = "range2String" Then
                    Try
                        strToken = range2String(CInt(Mid(strInstring, _
                                                     .Groups.Item(11).Index + 1, _
                                                     .Groups.Item(11).Length)), _
                                                CInt(Mid(strInstring, _
                                                         .Groups.Item(12).Index + 1, _
                                                         .Groups.Item(12).Length)))
                    Catch : End Try
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
    Optional ByVal strEllipsis As String = "...", _
    Optional ByVal booLeftEllipsis As Boolean = False) _
                    As String
        Dim intEllipsisLength As Integer
        Dim intStringLength As Integer
        intStringLength = Len(strInstring) : intEllipsisLength = Len(strEllipsis)
        If intMaxLength < 0 Then
            errorHandler("Internal programming error in call to ellipsis: " & _
                         "intMaxLength " & intMaxLength & " is less than zero")
            Return ("")
        End If
        If intEllipsisLength > intMaxLength Then
            errorHandler("Internal programming error in call to ellipsis: " & _
                         "intEllipsisLength " & intEllipsisLength & " is greater than " & _
                         "intMaxLength " & intMaxLength)
            Return ("")
        End If
        If intStringLength < intMaxLength Then
            Return (strInstring)
        Else
            If booLeftEllipsis Then
                Return (strEllipsis & Mid$(strInstring, 1, Math.Max(0, intMaxLength - intEllipsisLength)))
            Else
                Return (Mid$(strInstring, 1, Math.Max(0, intMaxLength - intEllipsisLength)) & strEllipsis)
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
        Return (Chr(34) & Replace(strInstring, Chr(34), Chr(34) & Chr(34)) & Chr(34))
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
    '
    ' 01 12 03   Nilges         Added the booInfo parameter: when present and
    '                           True, error is eventlogged as info and no
    '                           error is Thrown
    '
    ' 05 08 03   Nilges         Added ability to get full message text
    '
    ' 05 08 03   Nilges         Added strEventLog and booDecoration
    '
    ' 10 03 03   Nilges         Added booMsgBox
    '
    ' --- Message only
    Public Overloads Shared Sub errorHandler(ByVal strMessage As String, _
                                             Optional ByVal booInfo As Boolean = False, _
                                             Optional ByVal strEventLog As String = "", _
                                             Optional ByVal booDecoration As Boolean = True, _
                                             Optional ByVal booMsgBox As Boolean = False)
        Dim strFullMessage As String
        errorHandler_(strMessage, _
                      "unidentified object", _
                      "unidentified procedure", _
                      "", _
                      strFullMessage, _
                      booInfo, _
                      strEventLog, _
                      booDecoration, _
                      booMsgBox)
    End Sub
    ' --- Message and object name only
    Public Overloads Shared Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             Optional ByVal booInfo As Boolean = False, _
                                             Optional ByVal strEventLog As String = "", _
                                             Optional ByVal booDecoration As Boolean = True, _
                                             Optional ByVal booMsgBox As Boolean = False)
        Dim strFullMessage As String
        errorHandler_(strMessage, _
                      strObject, _
                      "unidentified procedure", _
                      "", _
                      strFullMessage, _
                      booInfo, _
                      strEventLog, _
                      booDecoration, _
                      booMsgBox)
    End Sub
    ' --- Message, object name and procedure
    Public Overloads Shared Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             ByVal strProcedure As String, _
                                             Optional ByVal booInfo As Boolean = False, _
                                             Optional ByVal strEventLog As String = "", _
                                             Optional ByVal booDecoration As Boolean = True, _
                                             Optional ByVal booMsgBox As Boolean = False)
        Dim strFullMessage As String
        errorHandler_(strMessage, _
                      strObject, _
                      strProcedure, _
                      "", _
                      strFullMessage, _
                      booInfo, _
                      strEventLog, _
                      booDecoration, _
                      booMsgBox)
    End Sub
    ' --- Message, object name, procedure and help info
    Public Overloads Shared Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             ByVal strProcedure As String, _
                                             ByVal strHelp As String, _
                                             Optional ByVal booInfo As Boolean = False, _
                                             Optional ByVal strEventLog As String = "", _
                                             Optional ByVal booDecoration As Boolean = True, _
                                             Optional ByVal booMsgBox As Boolean = False)
        Dim strFullMessage As String
        errorHandler_(strMessage, _
                      strObject, _
                      strProcedure, _
                      strHelp, _
                      strFullMessage, _
                      booInfo, _
                      strEventLog, _
                      booDecoration, _
                      booMsgBox)
    End Sub
    ' --- Message, object name, procedure and help info: return complete message
    Public Overloads Shared Sub errorHandler(ByVal strMessage As String, _
                                             ByVal strObject As String, _
                                             ByVal strProcedure As String, _
                                             ByVal strHelp As String, _
                                             ByRef strFullMessage As String, _
                                             Optional ByVal booInfo As Boolean = False, _
                                             Optional ByVal strEventLog As String = "", _
                                             Optional ByVal booDecoration As Boolean = True, _
                                             Optional ByVal booMsgBox As Boolean = False)
        errorHandler_(strMessage, _
                      strObject, _
                      strProcedure, _
                      strHelp, _
                      strFullMessage, _
                      booInfo, _
                      strEventLog, _
                      booDecoration, _
                      booMsgBox)
    End Sub
    ' --- Common logic
    Private Shared Sub errorHandler_(ByVal strMessage As String, _
                                     ByVal strObject As String, _
                                     ByVal strProcedure As String, _
                                     ByVal strHelp As String, _
                                     ByRef strFullMessage As String, _
                                     ByVal booInfo As Boolean, _
                                     ByVal strEventLog As String, _
                                     ByVal booDecoration As Boolean, _
                                     ByVal booMsgBox As Boolean)
        Dim strLocation As String = CStr(IIf(strObject = "" AndAlso strProcedure = "", _
                                             "", _
                                             CStr(IIf(strObject = "", "?", strObject)) & "." & _
                                             CStr(IIf(strProcedure = "", "?", strProcedure))))
        strFullMessage = strMessage
        If booDecoration Then
            strFullMessage = CStr(IIf(booInfo, "Informational message", "Error")) & " " & _
                             "at " & Now & ": " & strFullMessage & _
                             CStr(IIf(strLocation <> "", vbNewLine & vbNewLine & "Location: " & strLocation, "")) & _
                             CStr(IIf(strHelp <> "", vbNewLine & vbNewLine & strHelp, ""))
        End If
        If Trim(strEventLog) <> "" Then
            Try
                System.Diagnostics.EventLog.WriteEntry(strEventLog, strFullMessage)
            Catch
                Dim objExceptionEvt As New Exception("Cannot log the following message:" & _
                                                     vbNewLine & vbNewLine & _
                                                     strFullMessage & _
                                                     vbNewLine & vbNewLine & _
                                                     Err.Number & " " & Err.Description)
                Throw objExceptionEvt
                Return
            End Try
        End If
        If booMsgBox Then MsgBox(strFullMessage)
        If booInfo Then Return
        Dim objException As New Exception(strFullMessage)
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
        Return (file2String_(strFileid, 1, True, 0))
    End Function
    Public Overloads Shared Function file2String(ByVal strFileid As String, _
                                                 ByVal intStartIndex As Integer) As String
        Return (file2String_(strFileid, intStartIndex, True, 0))
    End Function
    Public Overloads Shared Function file2String(ByVal strFileid As String, _
                                                 ByVal intStartIndex As Integer, _
                                                 ByVal intLength As Integer) As String
        Return (file2String_(strFileid, intStartIndex, False, intLength))
    End Function
    Public Overloads Shared Function file2String(ByVal objFileStream As IO.FileStream) As String
        Return (file2String_(objFileStream, 1, True, 0))
    End Function
    Public Overloads Shared Function file2String(ByVal objFileStream As IO.FileStream, _
                                                 ByVal intStartIndex As Integer) As String
        Return (file2String_(objFileStream, intStartIndex, True, 0))
    End Function
    Public Overloads Shared Function file2String(ByVal objFileStream As IO.FileStream, _
                                                 ByVal intStartIndex As Integer, _
                                                 ByVal intLength As Integer) As String
        Return (file2String_(objFileStream, intStartIndex, False, intLength))
    End Function
    Private Shared Function file2String_(ByVal objFile As Object, _
                                         ByVal intStartIndex As Integer, _
                                         ByVal booReadAll As Boolean, _
                                         ByVal intLength As Integer) As String
        If intStartIndex < 1 Then
            errorHandler("Start index " & intStartIndex & " is not valid", _
                         Name, "file2String", _
                         "Returning null string")
            Return ("")
        End If
        Dim intFileLength As Integer
        Dim strFileid As String
        Dim objStream As IO.FileStream
        If (TypeOf objFile Is System.String) Then
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
            Return ("")
        End Try
        Dim intReadLength As Integer = intFileLength - intStartIndex + 1
        If Not booReadAll Then
            intReadLength = Math.Min(intLength, intReadLength)
        End If
        If intReadLength = 0 Then Return ("")
        If objStream Is Nothing Then
            Try
                Dim objStreamNew As New IO.FileStream(strFileid, IO.FileMode.Open)
                objStream = objStreamNew
            Catch
                errorHandler("Cannot open file " & strFileid & ": " & Err.Number & " " & Err.Description)
                Return ("")
            End Try
        End If
        Dim bytArray() As Byte
        Try
            ReDim bytArray(intReadLength - 1)
        Catch
            errorHandler("file2String cannot allocate a buffer: " & Err.Number & " " & Err.Description)
        End Try
        objStream.Read(bytArray, intStartIndex - 1, intReadLength)
        If (TypeOf objFile Is System.String) Then objStream.Close()
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
            strKeywordSplit = Split(UCase(strKeywords))
        Catch
            errorHandler("Could not split keywords", _
                         Name, "findAbbrev", _
                         Err.Number & " " & Err.Description)
        End Try
        Dim strTargetWork As String = Trim(UCase(strTarget))
        If InStr(strTargetWork, " ") <> 0 Then
            errorHandler("Target " & enquote(strTarget) & " contains internal spaces", _
                         Name, "findAbbrev", _
                         "Returning zero")
            Return (0)
        End If
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        For intIndex1 = 0 To UBound(strKeywordSplit)
            If abbrev(strTargetWork, strKeywordSplit(intIndex1)) Then
                If intIndex2 <> 0 Then Return (0)
                intIndex2 = intIndex1 + 1
            End If
        Next intIndex1
        Return (intIndex2)
    End Function

    ' -----------------------------------------------------------------------
    ' Find balancing parenthesis
    '
    '
    ' This method returns, in a string, the position from 1 of the parenthesis 
    ' (or character to be treated as a parenthesis) that balances the parenthesis
    ' or character at a given position. If no balancing parenthesis is found then
    ' this method returns one beyond the end of the string with no other error
    ' indication.
    '
    ' strInstring is the string to be scanned. By default, the scan starts
    ' at location 2 and balances the parenthesis or character at location 1.
    ' 
    ' The optional overload syntax findBalParenthesis(strInstring, intStartIndex)
    ' may start the search at the character or parenthesis at a different
    ' location. 
    '
    ' If the character or parenthesis at index 1 or intStartIndex is a round
    ' left parenthesis, this method searches for the balancing round right
    ' parenthesis; if the character is a left square bracket, this method searches
    ' for a right square bracket; if the character is a left curly brace this
    ' method method searches for a right curly brace.
    '
    ' Otherwise this method will raise an error. To avoid this problem use the
    ' optional overloaded syntax findBalParenthesis(strInstring, intStartIndex,
    ' strRightParenthesis) which specifies the right parenthesis character.
    '
    '
    ' --- Start at one and look for default right parenthesis
    Public Overloads Shared Function findBalParenthesis(ByVal strInstring As String) _
           As Integer
        Return (findBalParenthesis(strInstring, 1))
    End Function
    ' --- Start at one and look for default right parenthesis
    Public Overloads Shared Function findBalParenthesis(ByVal strInstring As String, _
                                                        ByVal intStartIndex As Integer) _
           As Integer
        Dim strLeftParenthesis As String = Mid(strInstring, intStartIndex, 1)
        Dim strRightParenthesis As String
        Select Case strLeftParenthesis
            Case "(" : strRightParenthesis = ")"
            Case "[" : strRightParenthesis = "]"
            Case "{" : strRightParenthesis = "}"
            Case Else
                errorHandler("Cannot determine the balancing right parenthesis for the " & _
                             "character " & enquote(strLeftParenthesis), _
                             Name, "findBalParenthesis", _
                             "Returning 0")
                Return (0)
        End Select
        Return (findBalParenthesis(strInstring, intStartIndex, strRightParenthesis))
    End Function
    ' --- Start at one and look for default right parenthesis
    Public Overloads Shared Function findBalParenthesis(ByVal strInstring As String, _
                                                        ByVal intStartIndex As Integer, _
                                                        ByVal strRightParenthesis As String) _
           As Integer
        If Len(strRightParenthesis) <> 1 Then
            errorHandler("Right parenthesis " & enquote(strRightParenthesis) & " " & _
                         "is not a character", _
                         Name, "findBalParenthesis", _
                         "Returning 0")
            Return (0)
        End If
        Dim intIndex1 As Integer
        Dim intLevel As Integer = 1
        Dim strLeftParenthesis As String = Mid(strInstring, intStartIndex, 1)
        For intIndex1 = intStartIndex + 1 To Len(strInstring)
            Select Case Mid(strInstring, intIndex1, 1)
                Case strLeftParenthesis : intLevel += 1
                Case strRightParenthesis
                    intLevel -= 1
                    If intLevel < 1 Then Exit For
            End Select
        Next intIndex1
        Return (intIndex1)
    End Function

    ' -----------------------------------------------------------------------
    ' Find item in string
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION
    ' --------   ----------     ---------------------------------------------
    ' 06 01 03   Nilges         Bug: incorrect offsets
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
            CType(IIf(booIgnoreCase, CompareMethod.Text, CompareMethod.Binary), CompareMethod)
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
            If strWork = strTargetWork Then Return (1)
            ' Shortcut 2: use Instr and split.  
            intIndex1 = InStr(strDelimiter & strInstring & strDelimiter, _
                              strDelimiter & strTarget & strDelimiter, _
                              objCompareMethod)
            If intIndex1 = 0 Then Return (0)
            strSplitArray = Split(Mid(strInstring, 1, intIndex1 - Len(strDelimiter)), strDelimiter, Compare:=objCompareMethod)
            Return (UBound(strSplitArray) + 2)
        End If
        ' --- Brute-force search  
        Dim strDelimiter1 As String = strDelimiter
        If booSetDelimiter Then strDelimiter1 = Mid(strDelimiter, 1, 1)
        Dim strInstringWork As String = strInstring
        If booIgnoreCase Then strInstringWork = UCase(strInstringWork)
        strSplitArray = Split(translate(strInstringWork, _
                                        strDelimiter, _
                                        copies(intDelimiterLength, strDelimiter)), _
                              strDelimiter1, _
                              Compare:=objCompareMethod)
        Dim intNullEntries As Integer
        For intIndex1 = 0 To UBound(strSplitArray)
            If strSplitArray(intIndex1) = "" Then
                intNullEntries += 1
            Else
                If strSplitArray(intIndex1) = strTargetWork Then Return (intIndex1 + 1 - intNullEntries)
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
        Return (findItem(strInstring, strTarget, " ", True, booIgnoreCase:=booIgnoreCase))
    End Function

    ' ----------------------------------------------------------------------
    ' Format an outline
    '
    '
    ' This method formats an outline which already has Dewey Decimal numbers
    ' into a string suitable for display in a monospace font.
    '
    ' Its input string should be a series of newline-separated lines. Each
    ' line should start with a number consisting of unsigned integers,
    ' separated by periods.
    '
    ' This method returns an indented outline which (because it uses spaces
    ' to format the outline) is best displayed in a monospace font such as
    ' Courier New.
    '
    ' By default, the syntax formatOutline(s) will format the input string
    ' s into a series of lines, none of which exceeds 72 characters: however,
    ' the syntax formatOutline(s,L) will use a line length specified in L.
    '
    ' By default, a period and one extra space will be inserted after toplevel
    ' outline numbers containing one number and no period: to suppress this
    ' option use formatOutline(s, False) or formatOutline(s, L, False).
    '
    ' By default the outline as formatted will be single-spaced. Specify spacing
    ' requirements in the optional parameters intSpacing and booSpaceOnMultiline:
    '
    '
    '      *  intSpacing may specify the number of blank lines between outline
    '         entries
    '
    '      *  booSpaceOnMultiline: this controls the application of intSpacing:
    '
    '         + If this parameter is absent or False intSpacing is applied
    '           consistently
    '
    '         + If this parameter is present and True intSpacing is applied
    '           only when at least one line of the outline has been edited to
    '           a multiple-line paragraph
    '
    '
    ' Note that these optional parameter are only available when the input string,
    ' length and the option to edit toplevel numbers are all explicitly specified.
    '
    '
    ' --- 72-character lines with toplevel edit
    Public Overloads Shared Function formatOutline(ByVal strInstring As String) As String
        Return formatOutline(strInstring, 72, True)
    End Function
    ' --- Line width is specified
    Public Overloads Shared Function formatOutline(ByVal strInstring As String, _
                                                   ByVal intWidth As Integer) As String
        Return formatOutline(strInstring, 72, True)
    End Function
    ' --- Toplevel edit is specified
    Public Overloads Shared Function formatOutline(ByVal strInstring As String, _
                                                   ByVal booToplevelEdit As Boolean) As String
        Return formatOutline(strInstring, 72, True)
    End Function
    ' --- Line width and toplevel editing behavior is specified
    Public Structure TYPindentStack
        Dim intRank As Integer
        Dim intIndentation As Integer
        Dim strDewey As String
    End Structure
    Public Overloads Shared Function formatOutline(ByVal strInstring As String, _
                                                   ByVal intWidth As Integer, _
                                                   ByVal booToplevelEdit As Boolean, _
                                                   Optional ByVal intSpacing As Integer = 0, _
                                                   Optional ByVal booSpaceOnMultiline As Boolean = False) As String
        Dim strSplit() As String
        Try
            strSplit = Split(strInstring, vbNewLine)
        Catch
            errorHandler("Cannot split: " & Err.Number & " " & Err.Description, _
                         "utilities", "formatOutline", _
                         "Returning outline with no change")
            Return strInstring
        End Try
        Dim usrIndentStack() As TYPindentStack
        Try
            ReDim usrIndentStack(0)
        Catch
            errorHandler("Cannot create stack: " & Err.Number & " " & Err.Description, _
                         "utilities", "formatOutline", _
                         "Returning outline with no change")
            Return strInstring
        End Try
        Dim booMultiline As Boolean
        Dim colFirstLines As Collection                 ' Key: _lineNumber: data: True
        If intSpacing <> 0 Then
            Try
                colFirstLines = New Collection
            Catch
                errorHandler("Cannot create collection: " & Err.Number & " " & Err.Description, _
                            "utilities", "formatOutline", _
                            "Returning outline with no change")
                Return strInstring
            End Try
        End If
        Dim intIndex1 As Integer
        Dim intLineCount As Integer
        Dim intRank As Integer
        Dim strNext As String
        Dim strNextDewey As String
        Dim strOutstring As String
        For intIndex1 = 0 To UBound(strSplit)
            strNextDewey = word(strSplit(intIndex1), 1)
            intRank = items(strNextDewey, ".", False)
            If intRank = 1 AndAlso booToplevelEdit AndAlso strNextDewey <> "" Then
                strNextDewey &= ".  "
            Else
                strNextDewey &= " "
            End If
            With usrIndentStack(UBound(usrIndentStack))
                If .intRank < intRank Then
                    ' Deeper rank: push new rank
                    Try
                        ReDim Preserve usrIndentStack(UBound(usrIndentStack) + 1)
                    Catch
                        errorHandler("Cannot push indentation stack: " & _
                                     Err.Number & " " & Err.Description, _
                                     "utilities", "formatOutline", _
                                     "Returning outline with no change")
                        Return strInstring
                    End Try
                    With usrIndentStack(UBound(usrIndentStack))
                        .intRank = intRank
                        .intIndentation = usrIndentStack(UBound(usrIndentStack) - 1).intIndentation _
                                          + _
                                          Len(strNextDewey)
                        .strDewey = strNextDewey
                    End With
                Else
                    ' Pop until rank at stack top is less or equal to than incoming rank
                    Do Until usrIndentStack(UBound(usrIndentStack)).intRank <= intRank
                        Try
                            ReDim Preserve usrIndentStack(UBound(usrIndentStack) - 1)
                        Catch
                            errorHandler("Cannot pop indentation stack: " & _
                                        Err.Number & " " & Err.Description, _
                                        "utilities", "formatOutline", _
                                        "Returning outline with no change")
                            Return strInstring
                        End Try
                    Loop
                    If UBound(usrIndentStack) = 0 Then
                        errorHandler("Cannot pop indentation stack: " & _
                                     Err.Number & " " & Err.Description, _
                                     "utilities", "formatOutline", _
                                     "Returning outline with no change")
                        Return strInstring
                    End If
                    With usrIndentStack(UBound(usrIndentStack))
                        .intIndentation -= Len(.strDewey)
                        .intIndentation = usrIndentStack(UBound(usrIndentStack) - 1).intIndentation _
                                        + _
                                        Len(strNextDewey)
                        .strDewey = strNextDewey
                        .intRank = intRank
                    End With
                End If
                strNext = strSplit(intIndex1)
                strNext = phrase(strNext, 2)
                strNext = soft2HardParagraph(strNext, _
                                             bytLineWidth:=CByte(intWidth _
                                                                 - _
                                                                 usrIndentStack(UBound(usrIndentStack)).intIndentation))
                booMultiline = booMultiline OrElse InStr(strNext, vbNewLine) <> 0
                strNext = Replace(strNext, _
                                  vbNewLine, _
                                  vbNewLine & _
                                  copies(usrIndentStack(UBound(usrIndentStack)).intIndentation, " "))
                strOutstring &= CStr(IIf(strOutstring = "", "", vbNewLine)) & _
                                copies(usrIndentStack(UBound(usrIndentStack) - 1).intIndentation, " ") & _
                                strNextDewey & _
                                strNext
                intLineCount += items(strNext, vbNewLine, False)
                If intSpacing <> 0 AndAlso strOutstring <> "" Then
                    Try
                        colFirstLines.Add(True, "_" & intLineCount - 1)
                    Catch
                        errorHandler("Cannot extend collection: " & _
                                     Err.Number & " " & Err.Description, _
                                     "utilities", "formatOutline", _
                                     "Returning outline with no change")
                        Return strInstring
                    End Try
                End If
            End With
        Next intIndex1
        If intSpacing <> 0 AndAlso (Not booSpaceOnMultiline OrElse booMultiline) Then
            Try
                strSplit = Split(strOutstring, vbNewLine)
            Catch
                errorHandler("Cannot split in preparation for postedit: " & _
                             Err.Number & " " & Err.Description, _
                             "utilities", "formatOutline", _
                             "Returning outline without extra formatting requested")
                Return strOutstring
            End Try
            Dim booSpacing As Boolean
            strOutstring = strSplit(0)
            For intIndex1 = 1 To UBound(strSplit)
                strOutstring &= vbNewLine
                booSpacing = False
                Try
                    booSpacing = CBool(colFirstLines.Item("_" & CStr(intIndex1 - 1)))
                Catch : End Try
                If booSpacing Then
                    strOutstring &= copies(vbNewLine, intSpacing)
                End If
                strOutstring &= strSplit(intIndex1)
            Next intIndex1
        End If
        Return strOutstring
    End Function

    ' ----------------------------------------------------------------------
    ' Return True if object has reference type, False when object is a
    ' value type
    '
    '
    ' This function returns True when objTest has a reference type other
    ' than System.String. It returns False when objTest has one of the
    ' scalar types: Boolean, Byte, Short, Integer, Single, Double or String.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 12 27 01   Nilges         Version 1
    ' 07 15 03   Nilges         Didn't work under VB 2003
    '
    '
    Public Shared Function hasReferenceType(ByVal objTest As Object) As Boolean
        Dim strType As String = UCase(objTest.GetType.ToString)
        If Mid(strType, 1, 7) = "SYSTEM." Then strType = Mid(strType, 8)
        Select Case strType
            Case "BOOLEAN" : Return (False)
            Case "BYTE" : Return (False)
            Case "SHORT" : Return (False)
            Case "INTEGER" : Return (False)
            Case "INT16" : Return (False)
            Case "INT32" : Return (False)
            Case "INT64" : Return (False)
            Case "SINGLE" : Return (False)
            Case "DOUBLE" : Return (False)
            Case "STRING" : Return (False)
        End Select
        Return (True)
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
        If dblValueMin = dblValueMax Then Return (0)
        Return (Math.Abs(dblRangeMax - dblRangeMin) _
                    * _
                    ((dblValue - dblValueMin) / Math.Abs(dblValueMax - dblValueMin)) _
                    + _
                    dblRangeMin)
    End Function

    ' ----------------------------------------------------------------------
    ' Increment the Dewey number
    '
    '
    ' --- Increment last period-separated number
    Public Shared Function incrementDewey(ByVal strDewey As String) As String
        Return incrementDewey(strDewey, items(strDewey, ".", False))
    End Function
    ' --- Increment indexed period-separated number
    Public Shared Function incrementDewey(ByVal strDewey As String, _
                                          ByVal intIndex As Integer) As String
        If strDewey = "" Then Return CStr(intIndex)
        If intIndex < 1 Then
            errorHandler("Invalid index " & intIndex, _
                         Name, "incrementDewey", _
                         "Returning original string")
            Return strDewey
        End If
        Dim strOutstring As String
        If InStr(strDewey, ".") <> 0 Then
            strOutstring = itemPhrase(strDewey, 1, intIndex - 1, ".", False)
            If strOutstring <> "" Then strOutstring &= "."
        End If
        Dim strNext As String = "1"
        Dim intItems As Integer = items(strDewey, ".", False)
        If intIndex <= intItems Then
            strNext = CStr(CInt(item(strDewey, intIndex, ".", False)) + 1)
        End If
        strOutstring = itemPhrase(strDewey, 1, intIndex - 1, ".", False) & _
                       CStr(IIf(intIndex > 1, ".", "")) & _
                       strNext & _
                       CStr(IIf(intIndex < intItems, ".", "")) & _
                       itemPhrase(strDewey, intIndex + 1, -1, ".", False)
        Return strOutstring
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
    Public Overloads Shared Function indent(ByVal strInstring As String) As String
        Return (indent(strInstring, "    ", vbNewLine))
    End Function
    Public Overloads Shared Function indent(ByVal strInstring As String, _
                                            ByVal intCount As Integer, _
                                            ByVal strNewLine As String) As String
        If intCount < 0 Then
            errorHandler("Invalid intCount parameter has value " & intCount, "indent")
            Return (strInstring)
        End If
        Return (indent(strInstring, copies(intCount, " "), vbNewLine))
    End Function
    Public Overloads Shared Function indent(ByVal strInstring As String, _
                                            ByVal strIndent As String, _
                                            ByVal strNewLine As String) As String
        Dim intIndex1 As Integer
        Dim strArray() As String
        Try
            strArray = Split(strInstring, strNewLine)
        Catch
            With Err()
                errorHandler("Not able to parse input string using Split: " & .Number & " " & .Description, _
                             "indent")
            End With
            Return (strInstring)
        End Try
        Dim objStringBuilder As String
        For intIndex1 = 0 To UBound(strArray)
            append(objStringBuilder, strNewLine, strIndent & strArray(intIndex1))
        Next intIndex1
        Return (objStringBuilder.ToString)
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
    ' The optional parameter booString2Box may be used to format the rule
    ' display into a box constructed from asterisks.
    '
    ' Note that this method is not a true utility in the sense that it is
    ' for EGNSF praxis.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 01 20 03     Nilges       1.  Version 1
    ' 06 22 03     Nilges       1.  Added booString2Box 
    '
    '
    Public Overloads Shared Function inspectionAppend(ByRef strReport As String, _
                                                      ByVal strRule As String, _
                                                      ByVal booRuleResult As Boolean, _
                                                      ByRef booResult As Boolean, _
                                                      Optional ByVal booString2Box As Boolean = False) As Boolean
        Return (inspectionAppend(strReport, strRule, booRuleResult, booResult, ""))
    End Function
    Public Overloads Shared Function inspectionAppend(ByRef strReport As String, _
                                                      ByVal strRule As String, _
                                                      ByVal booRuleResult As Boolean, _
                                                      ByRef booResult As Boolean, _
                                                      ByVal strComments As String, _
                                                      Optional ByVal booString2Box As Boolean = False) As Boolean
        If booString2Box Then
            Dim strResult As String = "This rule is " & _
                                      CStr(IIf(booRuleResult, "OK", "NOT MET")) & _
                                      CStr(IIf(strComments = "", _
                                               "", _
                                               vbNewLine & vbNewLine & _
                                               strComments))
            strReport = append(strReport, _
                                vbNewLine & vbNewLine, _
                                string2Box(vbNewLine & _
                                            strRule & _
                                            vbNewLine & vbNewLine & _
                                            copies("-", _
                                                   Math.Max(Math.Max(itemMaxWidth(strRule, _
                                                                                  vbNewLine, _
                                                                                  False), _
                                                                     itemMaxWidth(strResult, _
                                                                                  vbNewLine, _
                                                                                  False)), _
                                                            80)) & _
                                            vbNewLine & _
                                            strResult, _
                                            "I N S P E C T I O N   R U L E"))
        Else
            strReport = append(strReport, _
                                vbNewLine, _
                                strRule & ": " & CStr(IIf(booRuleResult, "OK", "Failed")) & _
                                CStr(IIf(strComments <> "", vbNewLine, "")) & _
                                strComments)
        End If
        booResult = booResult AndAlso booRuleResult
        Return (booResult)
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
        If intValue = 0 Then Return (1)
        Return (CInt(Math.Floor(Math.Log10(intValue))) + 1)
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
        If intLength < 2 Then Return (False)
        Return (Mid(strInstring, 1, 1) = Chr(34) AndAlso Mid(strInstring, intLength) = Chr(34))
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
    Public Overloads Shared Function isXMLcomment(ByVal strInstring As String) As Boolean
        Dim booComplete As Boolean
        Dim strCommentText As String
        Return isXMLcomment(strInstring, booComplete, strCommentText) And booComplete
    End Function
    Public Overloads Shared Function isXMLcomment(ByVal strInstring As String, _
                                                  ByRef booComplete As Boolean) As Boolean
        Dim strCommentText As String
        Return isXMLcomment(strInstring, booComplete, strCommentText)
    End Function
    Public Overloads Shared Function isXMLcomment(ByVal strInstring As String, _
                                                  ByRef strCommentText As String) As Boolean
        Dim booComplete As Boolean
        Return isXMLcomment(strInstring, booComplete, strCommentText) And booComplete
    End Function
    Public Overloads Shared Function isXMLcomment(ByVal strInstring As String, _
                                                  ByRef booComplete As Boolean, _
                                                  ByRef strCommentText As String) As Boolean
        Dim strInstringWork As String = Trim(strInstring)
        Dim intLength As Integer = Len(strInstringWork)
        booComplete = False : strCommentText = ""
        If intLength < 4 Then Return (False)
        If Mid(strInstringWork, 1, 4) <> "<!--" Then Return (False)
        booComplete = (Mid(strInstringWork, intLength - 2) = "-->")
        strCommentText = Trim(Mid(strInstringWork, 5, intLength - 3))
        Return (True)
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
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ_:", strInstringUCase) = 0 Then
            Return (False)
        End If
        Return (verify(strInstringUCase, _
                      "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_:") _
               = _
               0)
    End Function

    ' ----------------------------------------------------------------------
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
    ' This method has the following overloads.
    '
    '
    '      *  item(string, number): returns the numbered blank-delimited
    '         word (where number starts at one for the leftmost word) because
    '         the default delimiter is blank and the default style is set
    '         delimiter.
    '
    '      *  item(string, number, delimiter, style): returns the numbered
    '         item where the delimiter is defined and style is True or False:
    '
    '         + When style is True, the item is assumed to be delimited by
    '           "set" delimiters such that no item is null and delimiter is
    '           a set of alternative characters
    '
    '         + When style is False, the item is assumed to be delimited by
    '           "string" delimiters such that no item is null and delimiter is
    '           a fixed string or character.
    '
    '      *  item(string, number, delimiter, style, start): returns the numbered
    '         item where the delimiter is defined and style is True or False as
    '         above. The starting index of the delimiter is placed in the start
    '         parameter which must be an Integer.
    '
    '
    ' C H A N G E   R E C O R D -------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -------------------------------------------
    ' 05 13 03   Nilges         1. Added documentation
    '                           2. Replaced message box by error handling
    ' 05 25 03   Nilges         1. Bug in parsing for set mode delimiter
    ' 11 30 03   Nilges         1. Add intStartIndex parameter
    '                           2. Bug: did not correctly parse using
    '                              a multicharacter delimiter
    ' 12 08 03   Nilges         1. Added check for a null delimiter
    '
    '
    ' --- Blank-delimited item
    Public Overloads Shared Function item(ByVal strInstring As String, _
                                          ByVal intIndex As Integer) _
           As String
        Return item(strInstring, intIndex, " ", True)
    End Function
    ' --- Various other items
    Public Overloads Shared Function item(ByVal strInstring As String, _
                                          ByVal intIndex As Integer, _
                                          ByVal strDelimiter As String, _
                                          ByVal booSetDelimiter As Boolean) _
           As String
        Dim intStartIndex As Integer = 1
        Return (item_(strInstring, _
                      intIndex, _
                      strDelimiter, _
                      booSetDelimiter, _
                      intStartIndex, _
                      False))
    End Function
    ' --- Various other items
    Public Overloads Shared Function item(ByVal strInstring As String, _
                                          ByVal intIndex As Integer, _
                                          ByVal strDelimiter As String, _
                                          ByVal booSetDelimiter As Boolean, _
                                          ByRef intStartIndex As Integer) _
           As String
        Return (item_(strInstring, _
                      intIndex, _
                      strDelimiter, _
                      booSetDelimiter, _
                      intStartIndex, _
                      True))
    End Function
    ' --- Common logic
    Private Shared Function item_(ByVal strInstring As String, _
                                  ByVal intIndex As Integer, _
                                  ByVal strDelimiter As String, _
                                  ByVal booSetDelimiter As Boolean, _
                                  ByRef intStartIndex As Integer, _
                                  ByVal booFindStartIndex As Boolean) _
           As String
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intUBound As Integer
        Dim strArray() As String
        Dim strItem As String
        If strDelimiter = "" Then
            errorHandler("Delimiter cannot be a null string", _
                         Name, _
                         "item", _
                         "Returning a null item")
            Return ("")
        End If
        If intIndex < 1 Then
            errorHandler("intIndex " & intIndex & " is not valid", _
                         Name, _
                         "item", _
                         "Returning a null item")
            Return ("")
        End If
        If booFindStartIndex _
           OrElse _
           booSetDelimiter AndAlso Len(strDelimiter) <> 1 Then
            ' Scan the string
            intCount = 1
            intIndex1 = 1
            Dim intIndex2 As Integer
            Do While intIndex1 <= Len(strInstring)
                If booSetDelimiter Then
                    intIndex1 = verify(strInstring, _
                                        strDelimiter, _
                                        intIndex1, _
                                        False)
                End If
                If intIndex1 = 0 Then
                    intIndex1 = Len(strInstring) + 1
                    Exit Do
                End If
                If booSetDelimiter Then
                    intIndex2 = verify(strInstring & Mid(strDelimiter, 1, 1), _
                                        strDelimiter, _
                                        intIndex1 + 1, _
                                        True)
                Else
                    intIndex2 = InStr(intIndex1, _
                                      strInstring & strDelimiter, _
                                      strDelimiter)
                End If
                strItem = Mid(strInstring, _
                              intIndex1, _
                              intIndex2 - intIndex1)
                If intIndex = intCount Then
                    intStartIndex = intIndex1
                    Exit Do
                End If
                intCount += 1
                intIndex1 += Len(strItem)
                If Not booSetDelimiter Then intIndex1 += Len(strDelimiter)
            Loop
            If intIndex1 > Len(strInstring) Then
                strItem = "" : intStartIndex = 0
            End If
        Else
            If booSetDelimiter Then
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
                ' Use Split to find item
                If item_string2Array_(strInstring, strDelimiter, strArray) Then
                    If intIndex <= UBound(strArray) + 1 Then strItem = strArray(intIndex - 1)
                End If
            End If
        End If
        Return (strItem)
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
        Return (intUBound <> -1)
    End Function

    ' ----------------------------------------------------------------------
    ' Return the length, and, as an optional overload, the index of the 
    ' longest item
    '
    '
    ' strInstring is parsed into string-delimited items (as described for 
    ' the item method) or character set-delimited items (as described), and
    ' the length of the longest item is returned as the function value.
    '
    ' When the optional overload parameter, intIndex, is present, it is
    ' set to the index of the longest item.
    '
    '
    Public Overloads Shared Function itemMaxWidth(ByVal strInstring As String, _
                                                  ByVal strDelimiter As String, _
                                                  ByVal booSetDelimiter As Boolean) As Integer
        Dim intIndex1 As Integer
        Return (itemMaxWidth(strInstring, strDelimiter, booSetDelimiter, intIndex1))
    End Function
    Public Overloads Shared Function itemMaxWidth(ByVal strInstring As String, _
                                            ByVal strDelimiter As String, _
                                            ByVal booSetDelimiter As Boolean, _
                                            ByRef intIndex As Integer) As Integer
        Dim intMaxWidth As Integer
        Dim intIndex1 As Integer
        Dim intLength As Integer
        intIndex = 0
        For intIndex1 = 1 To items(strInstring, strDelimiter, booSetDelimiter)
            intLength = Len(item(strInstring, intIndex1, strDelimiter, booSetDelimiter))
            If intIndex1 = 1 Then
                intMaxWidth = intLength : intIndex = 1
            ElseIf intLength > intMaxWidth Then
                intMaxWidth = intLength : intIndex = intIndex1
            End If
        Next intIndex1
        Return (intMaxWidth)
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
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 05 13 03   Nilges         1.  Replaced msgbox by errorHandler
    '                           2.  Added null returns
    '
    '
    Public Overloads Shared Function itemPhrase(ByVal strInstring As String, _
                                                ByVal intStartIndex As Integer, _
                                                ByVal intItemCount As Integer, _
                                                ByVal strDelimiter As String, _
                                                ByVal booSetDelimiter As Boolean) _
                     As String
        Dim intIndex1 As Integer
        If intStartIndex < 1 Then
            errorHandler("Invalid intStartIndex parameter: " & intStartIndex, _
                         Name, "itemPhrase", _
                         "Returning null string")
            Return ("")
        End If
        Dim intEndIndex As Integer
        Dim intItems As Integer = items(strInstring, strDelimiter, booSetDelimiter)
        If intItemCount < -1 Then
            errorHandler("Invalid intItemCount parameter: " & intStartIndex, _
                         Name, "itemPhrase", _
                         "Returning null string")
            Return ("")
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
                         Name, "itemPhrase", _
                         Err.Number & " " & Err.Description & vbNewLine & vbNewLine & _
                         "Returning a null item phrase")
            Return ("")
        End Try
        With objStringBuilder
            For intIndex1 = intStartIndex To intEndIndex
                append(objStringBuilder, _
                       strDelimiter, _
                       item(strInstring, intIndex1, strDelimiter, booSetDelimiter))
            Next intIndex1
            Return (objStringBuilder.ToString)
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
    ' 11 30 03   Nilges         Bug: multiple-character set delimiters not
    '                           handled properly
    '
    '
    Public Overloads Shared Function items(ByVal strInstring As String) _
           As Integer
        Return (items(strInstring, " ", True))
    End Function
    Public Overloads Shared Function items(ByVal strInstring As String, _
                                           ByVal strDelimiter As String, _
                                           ByVal booSetDelimiter As Boolean) _
                     As Integer
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intItems As Integer
        Dim intLBound As Integer
        Dim strArray() As String
        Dim intStartIndex As Integer
        If Len(strDelimiter) > 1 AndAlso booSetDelimiter Then
            intIndex1 = 1
            Dim strItem As String
            Do
                strItem = item(Mid(strInstring, intIndex1), _
                               1, _
                               strDelimiter, _
                               True, _
                               intStartIndex)
                If strItem = "" Then Exit Do
                intItems += 1
                intIndex1 += intStartIndex + Len(strItem)
            Loop
            Return (intItems)
        End If
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
        Return (intCount)
    End Function

#If INCLUDE_TEST Then
    ' ----------------------------------------------------------------------
    ' Test the item methods, placing a test report in strReport
    '
    '
    Public Shared Function itemTest(ByRef strReport As String) As Boolean
        Dim booSetDelimiter As Boolean
        Dim intCount As Integer
        Dim intError As Integer
        Dim intIndex1 As Integer
        Dim intItemCount As Integer
        Dim intItemIndex As Integer
        Dim intStartIndex As Integer
        Dim strDelimiter As String
        Dim strDelimiterDisplay As String
        Dim strDelimiterSet As String = ",;:-!~/\|"
        Dim strItem As String
        Dim strTest As String
        Dim strWork As String
        intCount = CInt(Rnd() * 90) + 10
        For intIndex1 = 1 To intCount
            append(strReport, _
                    vbNewLine & vbNewLine, _
                    "Starting test " & intIndex1 & " of " & intCount)
            booSetDelimiter = (Rnd() > 0.5)
            If booSetDelimiter Then
                strDelimiter = itemTest_mkRandomSetDelimiter_(strDelimiterSet)
            Else
                strDelimiter = "[Delimiter" & _
                               copies(Mid("0123456789ABCDE", CInt(Rnd() * 14) + 1, 1), _
                                      CInt(Rnd() * 10)) & _
                               "]"
            End If
            strTest = itemTest_mkTestString_(intIndex1, _
                                             strDelimiter, _
                                             booSetDelimiter)
            Try
                strDelimiterDisplay = strDelimiter
                If booSetDelimiter Then
                    strDelimiterDisplay = "set {" & strDelimiterDisplay & "}"
                Else
                    strDelimiterDisplay = "string " & enquote(strDelimiterDisplay)
                End If
                append(strReport, _
                        vbNewLine, _
                        "Testing item using the string " & _
                        enquote(strTest) & " and " & _
                        "the delimiter " & " " & _
                        strDelimiterDisplay)
                intItemCount = items(strTest, strDelimiter, booSetDelimiter)
                intItemIndex = Math.Max(CInt(IIf(Rnd() < 0.1, _
                                        intItemCount + CInt(Rnd() * 3), _
                                        CInt(Rnd() * intItemCount))), 1)
                strWork = "item # " & _
                            intItemIndex & " " & _
                            "of " & intItemCount
                append(strReport, _
                        vbNewLine, _
                        "Attempting to get " & strWork)
                intStartIndex = -1
                If Rnd() < 0.5 Then
                    strItem = item(strTest, intItemIndex, strDelimiter, booSetDelimiter, intStartIndex)
                Else
                    strItem = item(strTest, intItemIndex, strDelimiter, booSetDelimiter)
                End If
                append(strReport, _
                        vbNewLine, _
                        strWork & " is " & _
                        enquote(strItem))
                If strItem <> "" _
                   AndAlso _
                   (word(strItem, 1) <> "Item" _
                    OrElse _
                    word(strItem, 2) <> CStr(intItemIndex) _
                    OrElse _
                    intStartIndex <> -1 _
                    AndAlso _
                    phrase(strItem, 3) <> "starts at " & intStartIndex) _
                   OrElse _
                   strItem = "" _
                   AndAlso _
                   (booSetDelimiter AndAlso intItemIndex <= intItemCount) Then
                    append(strReport, vbNewLine, "Item has an unexpected value")
                    intError += 1
                End If
            Catch objException As Exception
                append(strReport, _
                       vbNewLine, _
                       "The following error occured in test " & intCount & ": " & _
                       objException.ToString)
                intError += 1
            End Try
        Next intIndex1
        Return (intError = 0)
    End Function

    ' ----------------------------------------------------------------------
    ' Make a test string on behalf of itemTest
    '
    '
    Public Shared Function itemTest_mkTestString_(ByVal intIndex As Integer, _
                                                  ByVal strDelimiter As String, _
                                                  ByVal booSetDelimiter As Boolean) As String
        Dim intDelimiters As Integer
        Dim intIndex1 As Integer
        Dim intItemIndex As Integer = 1
        Dim strNext As String
        Dim strTest As String
        For intIndex1 = 1 To CInt(Rnd() * 10)
            intDelimiters = CInt(Rnd() * 3) + 1
            strTest &= copies(strDelimiter, intDelimiters)
            If Not booSetDelimiter Then
                intItemIndex += intDelimiters - 1
                If intIndex1 = 1 Then intItemIndex += 1
            End If
            strNext = "Item " & _
                      intItemIndex & " " & _
                      "starts at " & _
                      CStr(Len(strTest) + 1)
            strTest &= strNext
            intItemIndex += 1
        Next intIndex1
        If Rnd() < 0.25 Then
            strTest &= copies(strDelimiter, CInt(Rnd() * 5))
        End If
        Return (strTest)
    End Function

    ' ----------------------------------------------------------------------
    ' Make a random string on behalf of itemTest
    '
    '
    Public Shared Function itemTest_mkRandomSetDelimiter_(ByVal strInset As String) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strOutstring As String
        For intIndex1 = 1 To CInt(Rnd() * 31) + 1
            intIndex2 = Math.Max(CInt(Rnd() * Len(strInset)), 1)
            strOutstring &= Mid(strInset, intIndex2, 1)
            strInset = Mid(strInset, 1, intIndex2 - 1) & _
                        Mid(strInset, intIndex2 + 1)
        Next intIndex1
        Return (strOutstring)
    End Function
#End If

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
        Dim strSplit1() As String = Split(strInstring1, vbNewLine)
        Dim strSplit2() As String = Split(strInstring2, vbNewLine)
        ' --- Find maximum left column length
        For intIndex1 = 1 To UBound(strSplit1)
            intNextLength = Len(strSplit1(intIndex1))
            If intNextLength > intMaxLength Then intMaxLength = intNextLength
        Next intIndex1
        ' --- Create return value
        Dim objStringBuilder As New System.Text.StringBuilder
        For intIndex1 = 0 To Math.Max(UBound(strSplit1), UBound(strSplit2))
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
        Return (objStringBuilder.ToString)
    End Function

    ' ---------------------------------------------------------------------
    ' Replace keywords by values
    '
    '
    Public Shared Function keywordChange(ByVal strInstring As String, _
                                         ByVal strControlString As String, _
                                         ByVal ParamArray strSubstitution() As String) As String
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
    ' Return newline-separated line
    '
    '
    Public Shared Function line(ByVal strInstring As String, _
                         ByVal intLineNumber As Integer) As String
        Return (item(strInstring, intLineNumber, vbNewLine, False))
    End Function

    ' ----------------------------------------------------------------------
    ' Return number of newline-separated lines
    '
    '
    Public Shared Function lines(ByVal strInstring As String) As Integer
        Return (items(strInstring, vbNewLine, False))
    End Function

    ' ----------------------------------------------------------------------
    ' Return comma-separated list item
    '
    '
    ' Caution: the returned list item may contain leading or trailing
    ' spaces: for example, listItem("Moe, Larry, Curley", 2) will return
    ' " Larry". For best results Trim the result in many contexts.
    '
    '
    Public Shared Function listItem(ByVal strInstring As String, _
                                    ByVal intItemNumber As Integer) As String
        Return (item(strInstring, intItemNumber, ",", False))
    End Function

    ' ----------------------------------------------------------------------
    ' Return number of comma-separated list items
    '
    '
    Public Shared Function listItems(ByVal strInstring As String) As Integer
        Return (items(strInstring, ",", False))
    End Function

    ' ----------------------------------------------------------------------
    ' Convert Long number to base N
    '
    '
    ' This method converts a Long integer in lngBase10 to its representation 
    ' that uses the N distinct characters in strDigits.  The first character 
    ' in strDigits represents zero: the second represents one, and so on.
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
    Public Overloads Shared Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal intBase As Integer) As String
        Return long2BaseN(lngBase10, intBase, 0)
    End Function
    Public Overloads Shared Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal intBase As Integer, _
                                                ByVal bytWordSize As Byte) As String
        Return (long2BaseN(lngBase10, _
                          baseN2Long_base2Digits_(intBase), _
                          bytWordSize))
    End Function
    Public Overloads Shared Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal strDigits As String) As String
        Return long2BaseN(lngBase10, strDigits, 0)
    End Function
    Public Overloads Shared Function long2BaseN(ByVal lngBase10 As Long, _
                                                ByVal strDigits As String, _
                                                ByVal bytWordSize As Byte) As String
        Dim intBaseValue As Integer
        Dim lngBase10Work As Long
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strBaseN As String
        If Len(strDigits) < 2 Then
            errorHandler("length of strDigits is less than 2")
            Return ("")
        End If
        If lngBase10 = 0 Then
            objStringBuilder.Append(Mid(strDigits, 1, 1))
        Else
            intBaseValue = Len(strDigits)
            If bytWordSize <> 0 Then
                If intBaseValue <> 2 Then
                    errorHandler("Word size is specified but base is not 2")
                    Return (Mid(strDigits, 1, 1))
                End If
                If lngBase10 > 2 ^ (bytWordSize - 1) - 1 _
                   OrElse _
                   lngBase10 < -2 ^ (bytWordSize - 1) Then
                    errorHandler("Long value " & lngBase10 & " " & _
                                 "cannot be represented with word size " & _
                                 bytWordSize)
                    Return (Mid(strDigits, 1, 1))
                End If
            End If
            If lngBase10 < 0 Then
                If intBaseValue <> 2 OrElse bytWordSize = 0 Then
                    errorHandler("Input number " & lngBase10 & " is negative, " & _
                                 "but base is not 2 or word size is not specified")
                    Return ("")
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
        Return (strBaseN)
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
    Public Overloads Shared Function mkTempFile() As String
        Return (mkTempFile("C:\TEMP"))
    End Function
    Public Overloads Shared Function mkTempFile(ByVal strPath As String) As String
        Try
            MkDir(strPath)
        Catch : End Try
        Dim intIndex1 As Integer
        Dim strNext As String
        Do
            strNext = appendPath(strPath, "tmp" & intIndex1)
            If Not fileExists(strNext) Then
                Return (strNext)
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
        If InStr(strComment, vbNewLine) <> 0 AndAlso booMultipleLineEdit Then
            Dim intIndex1 As Integer
            Dim objStringBuilder As New System.Text.StringBuilder("")
            Dim strArray() As String
            strArray = Split(strComment, vbNewLine)
            For intIndex1 = LBound(strArray) To UBound(strArray)
                append(objStringBuilder, vbNewLine, mkXMLComment(strArray(intIndex1)))
            Next intIndex1
            Return (objStringBuilder.ToString)
        End If
        Return ("<!-- " & strComment & " -->")
    End Function

    ' ----------------------------------------------------------------------
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
    ' Note that this method uses the following test to see whether the
    ' value is an XML tag: if the value starts with a less than
    ' and ends with a greater than it is assumed to be a tag, and the
    ' XML meta-characters in the tag are not changed. Otherwise the
    ' XML meta-characters are changed.
    '
    ' White space is removed from the element to make the above test.
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
    ' 10 18 03     Nilges       Convert meta-characters to names
    ' 01 26 04     Nilges       Bug: <> was considered probable XML.
    ' 03 06 04     Nilges       Use Public function xmlMeta2Name
    ' 
    '
    ' 
    Public Overloads Shared Function mkXMLElement(ByVal strTag As String, _
                                                  ByVal strValue As String, _
    ByVal ParamArray strAttributes() As String) As String
        Return (mkXMLElement(strTag, strValue, 0))
    End Function
    Public Overloads Shared Function mkXMLElement(ByVal strTag As String, _
                                                  ByVal strValue As String, _
                                                  ByVal intLineLen As Integer, _
                                                  ByVal ParamArray strAttributes() As String) As String
        If intLineLen < 0 Then
            errorHandler("Internal programming error: intLineLen parameter " & intLineLen & " " & _
                            "is not valid", _
                            "mkXMLElement", "", _
                            "This parameter must be zero or positive.  A null string has been returned.")
            Return ("")
        End If
        Dim intIndex1 As Integer
        Dim intLineLenWork As Integer
        Dim strNewline As String = CStr(IIf(intLineLen = 0, "", vbNewLine))
        Dim strElement As String
        intLineLenWork = CInt(IIf(intLineLen = 0, Len(strValue), intLineLen))
        For intIndex1 = 1 To Len(strValue) Step intLineLenWork
            strElement &= strNewline & _
                          Mid(strValue, intIndex1, intLineLenWork)
        Next intIndex1
        If Not mkXMLElement_probableXML_(strElement) Then
            strElement = xmlMeta2Name(strElement)
        End If
        Return (mkXMLTag(strTag, strAttributes) & _
                strElement & _
                mkXMLTag(strTag, booEndTag:=True))
    End Function

    ' ----------------------------------------------------------------------
    ' Return True on behalf of mkXMLElement if string is probably XML
    '
    '
    ' The string is probably XML when it starts with a less than and ends with
    ' a greater than and is at least three characters in length.
    '
    '
    Private Shared Function mkXMLElement_probableXML_(ByVal strElement As String) _
            As Boolean
        Dim strWS As String = range2String(0, 32)
        Dim strWSgone As String = Trim(translate(strElement, strWS, copies(" ", Len(strWS))))
        If Len(strElement) <= 2 Then Exit Function
        Return (Mid(strWSgone, 1, 1) = "<" _
               AndAlso _
               Mid(strWSgone, Len(strWSgone)) = ">")
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
    ' 01 26 04     Nilges       Bug: produced an invalid XML name when the name
    '                           commenced with an underscore. Start such names
    '                           with an extra X
    '
    ' 01 28 04     Nilges       1.  Bug: attribute handling incorrect
    '                           2.  Need to quote attribute values as well as
    '                               convert their meta-characters
    '
    '
    Public Overloads Shared Function mkXMLTag(ByVal strTagName As String, _
                                              ByVal booEndTag As Boolean) As String

        Return ("<" & CStr(IIf(booEndTag, "/", "")) & mkXMLtag_name2XML_(strTagName) & ">")
    End Function
    Public Overloads Shared Function mkXMLTag(ByVal strTagName As String, _
                                              ByVal ParamArray strAttributes() As String) As String
        Dim strTag As String = mkXMLTag(strTagName, False)
        Dim strAttributesWork As String = Trim(mkXMLTag_attributes2String_(strAttributes))
        If strAttributesWork <> "" Then strAttributesWork = " " & strAttributesWork
        Return Mid(strTag, 1, Len(strTag) - 1) & _
               strAttributesWork & _
               ">"
    End Function

    ' ----------------------------------------------------------------------
    ' Convert attribute parameter array to string on behalf of mkXMLtag
    '
    '
    Private Shared Function mkXMLTag_attributes2String_(ByRef strAttributes() As String) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler("Can't create string builder", _
                         Name, "mkXMLTag_attributes2String_", _
                         Err.Number & Err.Description)
        End Try
        For intIndex1 = LBound(strAttributes) To UBound(strAttributes) Step 2
            intIndex2 = intIndex1 + 1
            If Not append(objStringBuilder, " ", _
                          strAttributes(intIndex1) & _
                          CStr(IIf(strAttributes(intIndex2) = "", _
                                   "", _
                                   "=" & _
                                   enquote(xmlMeta2Name(strAttributes(intIndex2)))))) Then
                errorHandler("Can't extend string builder", _
                             Name, "mkXMLTag_attributes2String_", _
                             Err.Number & Err.Description)
            End If
        Next intIndex1
        Return (objStringBuilder.ToString)
    End Function

    ' -----------------------------------------------------------------------
    ' Convert XML tag name to a string on behalf of mkXMLTag
    '
    '
    Private Shared Function mkXMLtag_name2XML_(ByVal strName As String) As String
        Dim intIndex1 As Integer
        Dim strNext As String
        Dim strOutstring As String
        For intIndex1 = 1 To Len(strName)
            strNext = char2Name(Mid(strName, intIndex1, 1))
            If Len(strNext) > 1 Then
                strNext = "_" & strNext & "_"
                If intIndex1 = 1 Then strNext = "X" & strNext
            End If
            strOutstring &= strNext
        Next intIndex1
        Return (strOutstring)
    End Function

    ' -----------------------------------------------------------------------
    ' Return class name
    '
    '
    Public Shared ReadOnly Property Name() As String
        Get
            Return ("utilities")
        End Get
    End Property

    ' ----------------------------------------------------------------------
    ' Name to character
    '
    '
    ' This method converts a character name (as returned by char2Name)
    ' back to the original character
    '
    '
    Public Shared Function name2Char(ByVal strName As String) As String
        Dim strNameWork As String = Trim(strName)
        If Len(strNameWork) = 1 Then
            If strNameWork >= "0" AndAlso strNameWork <= "9" Then Return strNameWork
            If strNameWork >= "A" AndAlso strNameWork <= "Z" Then Return strNameWork
            If strNameWork >= "a" AndAlso strNameWork <= "z" Then Return strNameWork
        End If
        strNameWork = UCase(strNameWork)
        Select Case strNameWork
            Case "TILDE" : Return "~"
            Case "ACUTEACCENT" : Return "`"
            Case "EXCLAMATIONPOINT" : Return "!"
            Case "ATSIGN" : Return "@"
            Case "POUNDSIGN" : Return "#"
            Case "DOLLARSIGN" : Return "$"
            Case "PERCENT" : Return "%"
            Case "CARAT" : Return "^"
            Case "AMPERSAND" : Return "&"
            Case "ASTERISK" : Return "*"
            Case "LEFTPARENTHESIS" : Return "("
            Case "RIGHTPARENTHESIS" : Return ")"
            Case "UNDERSCORE" : Return "_"
            Case "DASH" : Return "-"
            Case "PLUS" : Return "+"
            Case "EQUALS" : Return "="
            Case "LEFTBRACE" : Return "{"
            Case "LEFTBRACKET" : Return "["
            Case "RIGHTBRACE" : Return "}"
            Case "RIGHTBRACKET" : Return "]"
            Case "VERTICALSTROKE" : Return "|"
            Case "BACKSLASH" : Return "\"
            Case "COLON" : Return ":"
            Case "SEMICOLON" : Return ";"
            Case "DOUBLEQUOTE" : Return """"
            Case "SINGLEQUOTE" : Return "'"
            Case "LESSTHAN" : Return "<"
            Case "COMMA" : Return ","
            Case "GREATERTHAN" : Return ">"
            Case "PERIOD" : Return "."
            Case "QUESTIONMARK" : Return "?"
            Case "SLASH" : Return "/"
            Case "NUL" : Return Chr(0)
            Case "TAB" : Return vbTab
            Case "LINEFEED" : Return Chr(10)
            Case "CARRIAGERETURN" : Return Chr(13)
            Case "BLANK" : Return " "
            Case Else
                If Len(strNameWork) = 8 _
                   AndAlso _
                   Mid(strNameWork, 1, 3) = "CHR" Then
                    strNameWork = Mid(strNameWork, 4)
                    If verify(strNameWork, "0123456789") = 0 Then
                        Return ChrW(CInt(strNameWork))
                    End If
                End If
                errorHandler("Invalid character name " & _
                             enquote(strNameWork), _
                             Name, _
                             "name2Char", _
                             "Returning a null string")
                Return ""
        End Select
    End Function

    ' ----------------------------------------------------------------------
    ' Numbers to variables
    '
    '
    ' This method converts each numeric blank-delimited word to a letter.
    ' For example, it will convert "Value is 2" to "Value is x".
    '
    ' By default, each word left to right is converted to x, then y, then
    ' z, then a lower-case letter starting at the beginning of the alphabet. 
    ' The letters, starting at x, are reused if there are more than 26 numbers 
    ' in the input string.
    '
    ' The optional strLetters parameter may specify an alternate sequence
    ' of letters, such that these letters are used left to right, and letters
    ' are reused starting at the beginning when there are more numbers than
    ' letters.
    '
    ' The output string is "normalized": leading, trailing and excess internal
    ' spaces are removed.
    '
    '
    Public Shared Function numbers2Variables(ByVal strInstring As String, _
                                             Optional ByVal strLetters As String = _
                                             "xyzabdefghijklmnopqrstuvwxyz") _
           As String
        Dim strSplit() As String
        Try
            strSplit = Split(strInstring, " ")
        Catch
            errorHandler("Cannot split input string: " & _
                         Err.Number & " " & Err.Description, _
                         Name, _
                         "numbers2Variables", _
                         "Returning unchanged original string")
            Return (strInstring)
        End Try
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler("Cannot create string builder: " & _
                         Err.Number & " " & Err.Description, _
                         Name, _
                         "numbers2Variables", _
                         "Returning unchanged original string")
            Return (strInstring)
        End Try
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer = 1
        Dim strAppend As String
        For intIndex1 = 0 To UBound(strSplit)
            If IsNumeric(strSplit(intIndex1)) Then
                strAppend = Mid(strLetters, intIndex2, 1)
                intIndex2 = intIndex2 + 1
                If intIndex2 > Len(strLetters) Then intIndex2 = 1
            Else
                strAppend = strSplit(intIndex1)
            End If
            If Not append(objStringBuilder, " ", strAppend) Then Return (strInstring)
        Next intIndex1
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' Object to scalar
    '
    '
    ' This method converts objValue to the narrowest possible type of
    ' Byte, Short, Integer, Long, Single, Double or String. If
    ' objValue cannot be converted to any of these types, this method
    ' returns Null.
    '
    '
    Public Shared Function object2Scalar(ByVal objValue As Object) As Object
        Try
            Return (CByte(objValue))
        Catch
            Try
                Return (CShort(objValue))
            Catch
                Try
                    Return (CInt(objValue))
                Catch
                    Try
                        Return (CLng(objValue))
                    Catch
                        Try
                            Return (CSng(objValue))
                        Catch
                            Try
                                Return (CDbl(objValue))
                            Catch
                                Try
                                    Return (CStr(objValue))
                                Catch
                                    Return (Nothing)
                                End Try
                            End Try
                        End Try
                    End Try
                End Try
            End Try
        End Try
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
    ' 05 24 01     Nilges       1.  Don't returned condensed expressions
    '                           2.  Don't convert blanks
    '
    ' 07 05 03     Nilges       1.  If decoration is in effect and an object
    '                               does not convert to a string, return
    '                               type(value)
    '
    ' 11 30 03     Nilges       1.  Numeric check failed to catch newline:
    '                               check for characters valid only in numerics
    '
    '
    Public Overloads Shared Function object2String(ByVal objObject As Object) As String
        Return (object2String(objObject, False))
    End Function
    Public Overloads Shared Function object2String(ByVal objObject As Object, _
                                                   ByVal booDeco As Boolean) As String
        Dim strObject As String
        strObject = ""
        Try
            strObject = CStr(objObject)
            If booDeco Then
                If Not IsNumeric(strObject) OrElse verify(strObject, "0123456789eE+-.") <> 0 Then
                    strObject = string2Display(strObject, _
                                               "VBExpression", _
                                               strGraphicExclude:="""")
                End If
                strObject = objObject.GetType.ToString & "(" & strObject & ")"
            Else
                If Not IsNumeric(strObject) Then strObject = enquote(strObject)
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
            If booDeco AndAlso Not (objObject Is Nothing) Then
                strObject = objObject.GetType.ToString & "(" & strObject & ")"
            End If
        End Try
        Return (strObject)
    End Function

    ' ----------------------------------------------------------------------
    ' Object information (about info and state) to XML tag
    '
    '
    ' This method creates a single, compound XML tag from information
    ' about an object in a standardized format. The standard format consists
    ' of a comment block describing the object being documented, and zero, 
    ' one or more commented tags containing state information. Parameters 
    ' allow the comment block and/or the tag comments to be eliminated.
    '
    ' The strName parameter should be a suitable XML format name for the
    ' entire returned XML tag. Note that when the strName is null, no maintag
    ' and no endtag is generated. Instead, the indented series of object
    ' information tags is returned without enclosing tags.
    '
    ' The strAbout parameter contains information about the object.
    '
    ' The booAboutComment and the booStateComment parameters control,
    ' respectively, the generation of the About comment box and the generation
    ' of inline comments.
    '
    ' The strState parameter array should be in this format:
    '
    '
    '      <name1>,<desc1>,<value1>,<name2>,<desc2>,<value2>...
    '
    '
    ' This is a triple of values containing, for each state variable, its name
    ' (suitable for an XML tag name), its description (suitable for an XML
    ' comment) and its value, as a string.
    '
    ' A one line tag is generated for each state value, unless line breaks 
    ' occur in the state value, where a tag consists of a begin tag, a newline, 
    ' the value, a newline, and an end tag.
    '
    '
    ' C H A N G E   R E C O R D -----------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     -----------------------------------------------
    ' 05 28 03   Nilges         If the name is null, do not generate the
    '                           enclosing tags
    '
    '
    Public Shared Function objectInfo2XML(ByVal strName As String, _
                                          ByVal strAbout As String, _
                                          ByVal booAboutComment As Boolean, _
                                          ByVal booStateComment As Boolean, _
                                          ByVal ParamArray strState() As String) As String
        Dim objStringBuilder As System.Text.StringBuilder
        If (UBound(strState) + 1) Mod 3 <> 0 Then
            errorHandler("Invalid parameter array size", _
                         Name, _
                         "objectInfo2XML", _
                         "Returning a null string")
        End If
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler("Cannot create string builder: " & Err.Number & " " & Err.Description, _
                         Name, _
                         "objectInfo2XML", _
                         "Returning a null string")
            Return ("")
        End Try
        With objStringBuilder
            If booAboutComment Then
                append(objStringBuilder, _
                       vbNewLine, _
                       mkXMLComment(vbNewLine & _
                                    string2Box(vbNewLine & _
                                               soft2HardParagraph(strAbout) & _
                                               vbNewLine & vbNewLine) & _
                                    vbNewLine))
            End If
            If strName <> "" Then
                append(objStringBuilder, vbNewLine, mkXMLTag(strName, False))
            End If
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            Dim strIndent As String = "    "
            Dim strLineSep As String = vbNewLine & strIndent
            Dim strLineSep2 As String = strLineSep & strIndent
            Dim strLineSep3 As String = strLineSep & strIndent
            For intIndex1 = LBound(strState) To UBound(strState) Step 3
                If booStateComment Then
                    append(objStringBuilder, _
                           strLineSep, _
                           mkXMLComment(strState(intIndex1 + 1)))
                End If
                intIndex2 = intIndex1 + 2
                strLineSep3 = CStr(IIf(InStr(strState(intIndex2), vbNewLine) <> 0, _
                                       strLineSep2, _
                                       ""))
                append(objStringBuilder, _
                       strLineSep, _
                       mkXMLElement(strState(intIndex1), _
                                    strLineSep3 & _
                                    Replace(strState(intIndex2), _
                                            vbNewLine, _
                                            strLineSep2) & _
                                    Mid(strLineSep3, _
                                        1, _
                                        Math.Max(0, Len(strLineSep3) - Len(strIndent)))))
            Next intIndex1
        End With
        If strName <> "" Then
            append(objStringBuilder, vbNewLine, mkXMLTag(strName, True))
        End If
        Return (objStringBuilder.ToString)
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
        Return (parseXMLtag(strTag, strTagName, strTagAttributes))
    End Function
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String, _
                                                 ByRef strTagAttributes As String) As Boolean
        Dim booEndTag As Boolean
        Return (parseXMLtag(strTag, strTagName, strTagAttributes, booEndTag))
    End Function
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String, _
                                                 ByRef strTagAttributes As String, _
                                                 ByRef booEndTag As Boolean) As Boolean
        Dim booSelfContainedTag As Boolean
        Return (parseXMLtag(strTag, strTagName, strTagAttributes, booEndTag, booSelfContainedTag))
    End Function
    Public Overloads Shared Function parseXMLtag(ByVal strTag As String, _
                                                 ByRef strTagName As String, _
                                                 ByRef strTagAttributes As String, _
                                                 ByRef booEndTag As Boolean, _
                                                 ByRef booSelfContainedTag As Boolean) As Boolean
        If Len(strTag) < 3 Then Return (False)
        Dim intIndex1 As Integer = verify(strTag & " ", " >", 2, True)
        strTagName = Mid(strTag, 2, intIndex1 - 2)
        intIndex1 = verify(strTag & ">", " ", intIndex1)
        Dim intIndex2 As Integer = verify(strTag & ".", ">/", intIndex1 + 1, True)
        strTagAttributes = Mid(strTag, intIndex1, intIndex2 - intIndex1)
        booEndTag = (Mid(strTag, intIndex2, 1) = "/")
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Return a phrase of blank-delimited words
    '
    '
    Public Overloads Shared Function phrase(ByVal strInstring As String, _
                                            ByVal intStartIndex As Integer) As String
        Return (phrase(strInstring, intStartIndex, words(strInstring)))
    End Function
    Public Overloads Shared Function phrase(ByVal strInstring As String, _
                                            ByVal intStartIndex As Integer, _
                                            ByVal intWordCount As Integer) As String
        Return (itemPhrase(strInstring, intStartIndex, intWordCount, " ", True))
    End Function

    ' ----------------------------------------------------------------------
    ' Returns Proper case
    '
    '
    ' This method capitalizes the first letter of a string (if the string
    ' starts with a letter), and, it lowercases the rest of the string.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     D E S C R I P T I O N
    ' --------   ----------     --------------------------------------------
    ' 11 25 03   Nilges         Bug: string not lower cased
    '
    '
    Public Shared Function properCase(ByVal strOutstring As String) As String
        Return (UCase(Mid(strOutstring, 1, 1)) & LCase(Mid(strOutstring, 2)))
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
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 06 03 03   Nilges         Added error handling to range conversion
    '
    '
    Public Overloads Shared Function range2String(ByVal strStart As String, _
                                                  ByVal strEnd As String) As String
        If Len(strStart) <> 1 OrElse Len(strEnd) <> 1 Then
            errorHandler("Multiple-character or null strings cannot be converted", "range2String")
        End If
        Return (range2String(AscW(strStart), AscW(strEnd)))
    End Function
    Public Overloads Shared Function range2String(ByVal chrStart As Char, _
                                                  ByVal chrEnd As Char) As String
        Return (range2String(AscW(chrStart), AscW(chrEnd)))
    End Function
    Public Overloads Shared Function range2String(ByVal intStart As Integer, _
                                                  ByVal intEnd As Integer) As String
        Dim intIndex1 As Integer
        Dim objStringBuilder As New System.Text.StringBuilder("")
        For intIndex1 = intStart To intEnd Step CInt(IIf(intStart <= intEnd, 1, -1))
            Try
                objStringBuilder.Append(ChrW(intIndex1))
            Catch
                errorHandler("Cannot append character with value " & intIndex1, _
                             Name, "range2String", _
                             "Returning an incomplete output string")
                Exit For
            End Try
        Next intIndex1
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' Replace the XML metacharacters with &#nnnnn
    '
    '
    Public Overloads Shared Function replaceXMLmetaChars(ByVal strValue As String) As String
        Dim strComments As String
        Return (replaceXMLmetaChars(strValue, strComments))
    End Function
    Public Overloads Shared Function replaceXMLmetaChars(ByVal strValue As String, _
                                                  ByRef strComments As String) As String
        Return (replaceMetaChar_ _
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
        If InStr(strValue, strMetaChar) = 0 Then Return (strValue)
        Dim strXML As String = "&#" & alignRight(CStr(AscW(strMetaChar)), 5, "0") & ";"
        Dim strComment As String = "Replaced meta character " & _
                                   enquote(strMetaChar) & " " & _
                                   "with " & _
                                   enquote(strXML)
        If InStr(strComments, strComment) = 0 Then
            If strComments <> "" Then strComments &= vbNewLine
            strComments &= "*  " & strComment
        End If
        Return (Replace(strValue, strMetaChar, strXML))
    End Function

    ' ----------------------------------------------------------------------
    ' Run Length Encoding to string
    '
    '
    ' Run Length Encoding is a very simple packing method for low-intensity
    ' packing, which represents "runs" of identical characters as an escape
    ' character (known not to be part of the input string), a run length, and
    ' the repeated character. This method accepts any string which has been
    ' run length encoded and converts it back to the unpacked form.
    '
    ' See also string2RLE.
    '
    '
    Public Function RLE2String(ByVal strRLE As String) As String
        If Len(strRLE) = 0 Then
            errorHandler("Input string is null", _
                         Name, "RLE2String", _
                         "Returning null")
        End If
        Dim strSentinel As String = Mid(strRLE, 1, 1)
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch ex As Exception
            errorHandler("Can't create string builder: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "RLE2String", _
                         "Returning unconverted RLE string")
            Return (strRLE)
        End Try
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Do While intIndex1 <= Len(strRLE)
            intIndex2 = InStr(intIndex1, strRLE & strSentinel, strSentinel)
            If Not (append(objStringBuilder, _
                           "", _
                           Mid(strRLE, intIndex1, intIndex2 - intIndex1)) _
                    AndAlso _
                    append(objStringBuilder, _
                           "", _
                           copies(Mid(strRLE, intIndex2 + 2, 1), _
                                  AscW(Mid(strRLE, intIndex2 + 1, 1))))) Then
                errorHandler("append failure", _
                             Name, "RLE2String", _
                             "Returning incomplete, unpacked string")
            End If
        Loop
        Return (objStringBuilder.ToString)
    End Function

    ' -----------------------------------------------------------------
    ' Return a comma delimited list of small business names
    '
    '
    ' Useful as test data, this method celebrates the diversity of
    ' small business in the USA
    '
    ' Note: modifications other than changing or adding entries need
    ' the end may affect the results of the test method.
    '
    '
    Public Shared Function smallBusinessList() As String
        Return ("Al's Snappy Food, " & _
                "Singh Spicy Curry Shop, " & _
                "Best Way Noodle Bowl, " & _
                "You Eat You Pay You Shuddup Now Cafe, " & _
                "O'Reilly's Bar, " & _
                "Man with a Van Worldwide Moving and Storage Company, " & _
                "Joe's Gas the Heck Up 'n Git Outa Town, " & _
                "Edward's Salon de Paris Beauty Parlor (ask about our emergency entrance), " & _
                "Moon's Snappy Cab, " & _
                "Singh's Courteous Cab, " & _
                "Kwame's A Number 1 Cab, " & _
                "Giant Toad Software Development PLC, " & _
                "Lucy's Sunnyside Tap, " & _
                "Ball and Chain Bar and Grill, " & _
                "Bob's Berserko Lounge, " & _
                "Great Frog Packaged Liquors, " & _
                "Great Horned Toad Bail Bonds, " & _
                "Fester's Mortuary, " & _
                "Ramesh Patel Auto Insurance, " & _
                "Stereo Onset, " & _
                "Bed Bath Cable Chocolate Pizza and Beyond, " & _
                "Guns r Us, " & _
                "Ten Things Float in Happy Bowl Won Ton, " & _
                "Ahmed's Pita Stop, " & _
                "Letizia's Hot Breads")
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
                                              Optional ByVal strHardParaBreak As String = vbNewLine & vbNewLine, _
                                              Optional ByVal strHardLineBreak As String = vbNewLine) As String
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
                         Name, "soft2HardParagraph")
            Return ("")
        End If
        intSuffixLength = Len(strLineSuffix)
        Try
            strSplitLines = Split(strSoftParagraphs, vbNewLine)
        Catch
            errorHandler("Can't split input string", _
                         Name, "soft2HardParagraph", _
                         Err.Number & " " & Err.Description)
            Return ("")
        End Try
        For intIndex1 = 0 To UBound(strSplitLines)
            If strSplitLines(intIndex1) <> "" Then
                Try
                    strSplitWords = Split(strSplitLines(intIndex1), " ")
                Catch
                    errorHandler("Can't split input line", _
                                Name, "soft2HardParagraph", _
                                Err.Number & " " & Err.Description)
                    Return ("")
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
                strSplitLines = Split(strOutput, vbNewLine)
            Catch
                errorHandler("Can't split result string for inspection", _
                            Name, "soft2HardParagraph", _
                            Err.Number & " " & Err.Description)
                Return ("")
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
                    enquote(string2Display(strSplitLines(intIndex1), strGraphicExclude:=" ")))
            End If
        End If
        Return (strOutput)
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
                             strHardLineBreak & strPrefix)) & _
                    Trim(strOutputLine)
        If booSameWidth Then
            strOutput = strOutput & _
                        copies(Math.Max(0, bytLineWidth - Len(strOutputLine) - Len(STRsuffix)), " ")
        End If
        strOutput = strOutput & STRsuffix
        strOutputLine = strPrefix
    End Sub

    ' -----------------------------------------------------------------
    ' Sort items
    '
    '
    Public Overloads Function sortItems(ByVal strInstring As String) As String
        sortItems(strInstring, " ", False)
    End Function
    Public Overloads Function sortItems(ByVal strInstring As String, _
                                        ByVal strDelimiter As String, _
                                        ByVal booStringMode As Boolean) As String
        sortItems(strInstring, " ", False)
    End Function
    Public Overloads Function sortItems(ByVal strInstring As String, _
                                        ByVal strDelimiter As String, _
                                        ByVal booStringMode As Boolean, _
                                        ByVal booDescending As Boolean) As String
    End Function
    Public Overloads Function sortItems(ByVal strInstring As String, _
                                        ByVal strDelimiter As String, _
                                        ByVal booStringMode As Boolean, _
                                        ByVal booDescending As Boolean, _
                                        ByVal booInspect As Boolean) As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intNextIndex As Integer
        Dim intWordCount1 As Integer
        Dim intWordCountCurrent As Integer
        Dim intWords As Integer
        Dim strLast As String
        Dim strNext As String
        Dim strNextKey As String
        Dim strOutstring As String
        strOutstring = ""
        strLast = ""
        intIndex1 = 1
        intWordCount1 = words(strInstring)
        Do While strInstring <> ""
            strNext = word(strInstring, 1)
            intNextIndex = 1
            If words(strOutstring) > 0 Then strLast = word(strOutstring, words(strOutstring))
            intWordCountCurrent = words(strInstring)
            For intIndex2 = 1 To intWordCountCurrent
                strNextKey = word(strInstring, intIndex2)
                If Not booDescending _
                       AndAlso _
                       strNextKey > strLast _
                       AndAlso _
                       strNextKey < strNext _
                       OrElse _
                       booDescending AndAlso strNextKey > strNext Then
                    strNext = strNextKey
                    intNextIndex = intIndex2
                End If
            Next intIndex2
            strOutstring = append(strOutstring, strDelimiter, strNext)
            intWords = words(strInstring)
            strInstring = Trim(sortItems_itemDelete_(strInstring, intNextIndex))
            If words(strInstring) = intWords Then strInstring = phrase(strInstring, 1, intWords - 1)
            intIndex1 = intIndex1 + 1
        Loop
        If booInspect Then
            Dim booHigh As Boolean
            Dim strPrevious As String
            For intIndex1 = 1 To items(strInstring, strDelimiter, Not booStringMode)
                booHigh = (item(strInstring, intIndex1, strDelimiter, Not booStringMode) _
                           > _
                           strPrevious)
                If booDescending AndAlso Not booHigh _
                   OrElse _
                   Not booDescending AndAlso booHigh Then
                    errorHandler("Internal programming error: sort failure", _
                                 Name, _
                                 "sortItems", _
                                 "Returning incorrect result")
                    Exit For
                End If
            Next intIndex1
        End If
        Return (strOutstring)
    End Function

    ' -----------------------------------------------------------------
    ' Delete item on behalf of sortItems
    '
    '
    Private Function sortItems_itemDelete_(ByVal strInstring As String, _
                                           ByVal lngIndex As Long) As String
        Dim intIndex1 As Integer
        Dim strOutstring As String
        strOutstring = ""
        For intIndex1 = 1 To words(strInstring)
            If lngIndex <> intIndex1 Then
                strOutstring = append(strOutstring, _
                                      " ", _
                                      word(strInstring, intIndex1))
            End If
        Next intIndex1
        Return (strOutstring)
    End Function

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
    Public Overloads Shared Function spinlock(ByRef intLock As Integer) As Boolean
        Return (spinlock(intLock, _
                        SPINLOCK_TIMEOUT_DEFAULT))
    End Function
    Public Overloads Shared Function spinlock(ByRef intLock As Integer, _
                                              ByVal intTimeout As Integer) As Boolean
        Return (spinlock(intLock, _
                        intTimeout, _
                        SPINLOCK_SLEEPWAIT_DEFAULT))
    End Function
    Public Overloads Shared Function spinlock(ByRef intLock As Integer, _
                                              ByVal intTimeout As Integer, _
                                              ByVal intSleep As Integer) As Boolean
        Dim intCount As Integer
        Return (spinlock(intLock, _
                        intTimeout, _
                        SPINLOCK_SLEEPWAIT_DEFAULT, _
                        intCount))
    End Function
    Public Overloads Shared Function spinlock(ByRef intLock As Integer, _
                                              ByVal intTimeout As Integer, _
                                              ByVal intSleep As Integer, _
                                              ByRef intCount As Integer) As Boolean
        Dim intSleepWork As Integer = intSleep
        Dim intTimeoutWork As Integer = intTimeout
        If intTimeoutWork < -2 Then
            errorHandler("Invalid intTimeout value " & intTimeoutWork, _
                         Name, "spinlock")
            Return (False)
        ElseIf intTimeoutWork = -2 Then
            intTimeoutWork = SPINLOCK_TIMEOUT_DEFAULT
        End If
        If intSleepWork = -2 Then
            intSleepWork = SPINLOCK_SLEEPWAIT_DEFAULT
        ElseIf intSleepWork < 0 Then
            errorHandler("Invalid intSleep value " & intSleepWork, _
                         Name, "spinlock")
            Return (False)
        End If
        Dim datStart As Date
        If intTimeoutWork <> -1 Then datStart = Now
        Dim objThreading As Thread
        Do
            intCount += 1
            If Interlocked.Exchange(intLock, intLock) = 0 Then Return (True)
            If intSleep <> 0 Then objThreading.CurrentThread.Sleep(intSleepWork)
        Loop While intTimeoutWork = -1 _
                   OrElse _
                   intTimeoutWork > 0 AndAlso DateDiff(DateInterval.Second, datStart, Now) < intTimeoutWork
        Return (False)
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
    Public Overloads Shared Function string2Box(ByVal strInstring As String) As String
        Return string2Box(strInstring, "", "*")
    End Function
    Public Overloads Shared Function string2Box(ByVal strInstring As String, _
                                                ByVal strBoxLabel As String) As String
        Return string2Box(strInstring, strBoxLabel, "*")
    End Function
    Public Overloads Shared Function string2Box(ByVal strInstring As String, _
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
            Return ("")
        End If
        ' --- Find length of longest line, while parsing lines
        intMaxLength = 0 : intIndex1 = 1
        Do While intIndex1 <= Len(strInstring)
            intIndex2 = InStr(intIndex1, strInstring & vbNewLine, vbNewLine)
            intNextLength = intIndex2 - intIndex1
            colLines.Add(Mid(strInstring, intIndex1, intNextLength))
            If intNextLength > intMaxLength Then intMaxLength = intNextLength
            intIndex1 = intIndex2 + Len(vbNewLine)
        Loop
        intMaxLength += 4
        intMaxLength = CInt(Math.Max(intMaxLength, _
                                CInt(IIf(strBoxLabel = "", 10, Len(strBoxLabel) + 15))))
        ' --- Produce the text box
        strFence = copies(intMaxLength, strBuildChar)
        If strBoxLabel = "" Then
            strOutstring = strFence
        Else
            strOutstring = copies(5, strBuildChar) & _
                           " " & strBoxLabel & " " & _
                           copies(CInt(Math.Max(0, intMaxLength - 5 - Len(strBoxLabel) - 2)), _
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
        Return (strOutstring & vbNewLine & strFence)
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
    '      *  string2Display(string) returns the string converted to XML syntax.
    '         The graphical character set can be controlled using the following
    '         optional parameters.
    '
    '         + objGraphic: Nothing for the default or a string, containing the
    '           graphic character set. The default consists of all lower and upper
    '           case letters, digits, and the special characters commonly available
    '           on PC keyboards in the US.
    '
    '         + strGraphicInclude: characters to be added to the graphic character
    '           set.
    '
    '         + strGraphicExclude: characters to be added to the graphic character
    '           set.
    '
    '      *  string2Display(string, syntax) returns the string converted to the
    '         syntax, specified as a string:
    '
    '         + "C": converts to C syntax
    '
    '         + "XML": converts to XML syntax, such that nondisplayable characters
    '           are converted to &#nnnnn where nnnnn is the decimal value of the
    '           ASCII or wide character: note that nnnnn is always five digits with
    '           zero fill.
    '
    '           Note that this syntax is best for producing strings that need to be 
    '           converted back to their original values by display2String.
    '
    '         + "XMLcondensed": converts to XML syntax, such that nondisplayable characters
    '           are converted to &#n where n is the decimal value of the
    '           ASCII or wide character: note that n is the value of the ASCII or
    '           wide character without extension to five characters.
    '
    '           Note that this syntax may produce strings that cannot be converted back
    '           to their original values by display2String.
    '
    '         + "VBExpression": converts to Visual Basic expression syntax
    '
    '         + "VBExpressionCondensed": converts to condensed VB expression syntax
    '
    '         This overload also supports the optional parameters objGraphic, 
    '         strGraphicInclude, and strGraphicExclude described above.
    '
    '      *  string2Display(string, syntax) returns the string converted to the
    '         syntax, specified as an enumerator of the type ENUstring2DisplaySyntax:
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.C) converts to C
    '           syntax
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.XML) converts to XML
    '           syntax as described above
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.XMLcondensed) converts 
    '           to XML condensed syntax as described above
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.VBExpression) converts  
    '           to Visual Basic expression syntax
    '
    '         + string2Display(string, ENUstring2DisplaySyntax.VBExpressionCondensed) 
    '           converts to condensed Visual Basic expression syntax
    '
    '         This overload DOES NOT support the optional parameters objGraphic, 
    '         strGraphicInclude, and strGraphicExclude described above.
    '
    '      *  string2Display(string, syntax, graphic) returns the string converted to 
    '         the syntax, specified as an enumerator of the type ENUstring2DisplaySyntax,
    '         as above. It also exposes the graphic character set which is passed as a
    '         string, and the following optional parameters.
    '
    '         + strGraphicInclude: characters to be added to the graphic character
    '           set.
    '
    '         + strGraphicExclude: characters to be added to the graphic character
    '           set.
    '
    '         This overload DOES NOT support the optional parameter objGraphic 
    '         described above.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     --------------------------------------------
    ' 03 25 02   Nilges         Documented ability to specify enumeration as  
    '                           a string
    '
    ' 01 07 03   Nilges         1.  Added the strGraphic, strGraphicInclude 
    '                               and strGraphicExclude optional parameters 
    '
    '                           2.  Removed booBlankConvert and booQuoteConvert
    '
    ' 12 11 03   Nilges         1.  Added XMLcondensed option 
    '
    ' 
    ' --- Convert to use XML notation
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    Optional ByVal objGraphic As Object = Nothing, _
                                                    Optional ByVal strGraphicInclude As String = "", _
                                                    Optional ByVal strGraphicExclude As String = "") As String
        Return (string2Display(strInstring, _
                               "XML", _
                               objGraphic:=objGraphic, _
                               strGraphicInclude:=strGraphicInclude, _
                               strGraphicExclude:=strGraphicExclude))
    End Function
    ' --- Convert to use notation specified as a string
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    ByVal strSyntax As String, _
                                                    Optional ByVal objGraphic As Object = Nothing, _
                                                    Optional ByVal strGraphicInclude As String = "", _
                                                    Optional ByVal strGraphicExclude As String = "") As String
        Dim enuSyntax As ENUdisplay2StringSyntax = display2String_syntax2Enum_(strSyntax)
        Dim strGraphic As String = CStr(IIf((objGraphic Is Nothing), "", objGraphic))
        If enuSyntax <> ENUdisplay2StringSyntax.Invalid Then
            Return (string2Display(strInstring, _
                                   enuSyntax, _
                                   strGraphic, _
                                   strGraphicInclude:=strGraphicInclude, _
                                   strGraphicExclude:=strGraphicExclude))
        End If
    End Function
    ' --- Convert to use notation specified as an enumerator
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    ByVal enuSyntax As ENUdisplay2StringSyntax) As String
        Return (string2Display(strInstring, enuSyntax, ""))
    End Function
    ' --- Specify the graphic character set, and its optional inclusions and exclusions                                             
    Public Overloads Shared Function string2Display(ByVal strInstring As String, _
                                                    ByVal enuSyntax As ENUdisplay2StringSyntax, _
                                                    ByVal strGraphic As String, _
                                                    Optional ByVal strGraphicInclude As String = "", _
                                                    Optional ByVal strGraphicExclude As String = "") As String
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intLength As Integer
        Dim intTagEnd As Integer
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strNext As String
        Dim strGraphicSet As String = strGraphic
        If strGraphic = "" Then
            strGraphicSet = asciiCharsetEnum2String(ENUasciiCharset.graphic)
        End If
        strGraphicSet = strGraphicSet & strGraphicInclude
        For intIndex1 = 1 To Len(strGraphicExclude)
            strGraphicSet = Replace(strGraphicSet, Mid(strGraphicExclude, intIndex1, 1), "")
        Next intIndex1
        If enuSyntax = ENUdisplay2StringSyntax.C Then
            strGraphicSet = Replace(strGraphicSet, "\", "")
        End If
        If enuSyntax = ENUdisplay2StringSyntax.VBExpression _
           OrElse _
           enuSyntax = ENUdisplay2StringSyntax.VBExpressionCondensed Then
            strGraphicSet = Replace(strGraphicSet, ChrW(34), "")
        End If
        If strGraphicSet = "" Then
            errorHandler("After application of include and exclude, and format-related " & _
                         "changes, to the default " & _
                         "or specified set of graphic characters, the result set is null", _
                         Name, "string2Display", _
                         "Returning a null string")
            Return ("")
        End If
        If enuSyntax <> ENUdisplay2StringSyntax.C _
           And _
           enuSyntax <> ENUdisplay2StringSyntax.XML _
           And _
           enuSyntax <> ENUdisplay2StringSyntax.XMLcondensed _
           And _
           enuSyntax <> ENUdisplay2StringSyntax.VBExpression _
           And _
           enuSyntax <> ENUdisplay2StringSyntax.VBExpressionCondensed Then
            errorHandler("Invalid enuSyntax parameter of " & enuSyntax.ToString, _
                         Name, "string2Display", _
                         "Returning a null string")
            Return ("")
        End If
        If enuSyntax = ENUdisplay2StringSyntax.C _
           AndAlso _
           verify(strInstring, asciiCharsetEnum2String(ENUasciiCharset.allAscii)) _
           <> _
           0 Then
            errorHandler("Syntax is C but input string contains non-ASCII characters", _
                         "string2Display", Name, _
                         "Unlike the other formats supported by string2Display " & _
                         "(XML and VB expression) the C format is not able to handle " & _
                         "international characters outside the extended ASCII range 0..255.")
            Return ("")
        End If
        If strInstring = "" _
           AndAlso _
           (enuSyntax = ENUdisplay2StringSyntax.VBExpression _
            OrElse _
            enuSyntax = ENUdisplay2StringSyntax.VBExpressionCondensed) Then
            Return (Chr(34) & Chr(34))
        End If
        string2Display = ""
        intIndex1 = 1
        Do While intIndex1 <= Len(strInstring)
            intIndex2 = verify(strInstring & Chr(0), _
                               strGraphicSet, _
                               intIndex1)
            strNext = Mid(strInstring, intIndex1, intIndex2 - intIndex1)
            If strNext <> "" _
               AndAlso _
               (enuSyntax = ENUdisplay2StringSyntax.VBExpression _
                Or _
                enuSyntax = ENUdisplay2StringSyntax.VBExpressionCondensed) Then
                strNext = CStr(IIf(objStringBuilder.Length = 0, "", " & ")) & _
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
                               CStr(strGraphicSet), _
                               intIndex2, _
                               True)
            strNext = Mid(strInstring, intIndex2, intIndex1 - intIndex2)
            If enuSyntax = ENUdisplay2StringSyntax.C Then
                strNext = string2Display_string2C_(strNext)
            ElseIf enuSyntax = ENUdisplay2StringSyntax.XML _
                   OrElse _
                   enuSyntax = ENUdisplay2StringSyntax.XMLcondensed Then
                strNext = string2Display_string2XML_(strNext, _
                                                     enuSyntax = ENUdisplay2StringSyntax.XMLcondensed)
            Else
                strNext = CStr(IIf(objStringBuilder.Length = 0, "", " & ")) & _
                          string2Display_segment2Expression_ _
                          (strNext, _
                           enuSyntax _
                           = _
                           ENUdisplay2StringSyntax.VBExpressionCondensed, _
                           False)
            End If
            objStringBuilder.Append(strNext)
        Loop
        Return (objStringBuilder.ToString)
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
                Return (enquote(strSegment))
            Else
                Return (string2Display_nondisplayable2Expression_(strSegment, False))
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
        For intIndex1 = 1 To Len(strNonDisplayable)
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
                strNextEscape = "\n" : intIndex1 += Len(vbNewLine)
            Else
                strNextCharacter = Mid(strInstring, intIndex1, 1)
                If strNextCharacter = "\" Then
                    strNextEscape = "\\"
                Else
                    lngAsc = AscW(strNextCharacter)
                    If lngAsc < 0 OrElse lngAsc > 255 Then
                        errorHandler("International character values cannot be " & _
                                     "converted to C syntax")
                        Return ("")
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
        Return (objStringBuilder.ToString)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of string2Display, convert all chars in string to XML
    '
    '
    Private Shared Function string2Display_string2XML_ _
            (ByVal strInstring As String, _
             ByVal booCondensed As Boolean) As String
        Dim intIndex1 As Integer
        Dim objStringBuilder As New System.Text.StringBuilder
        Dim strNext As String
        For intIndex1 = 1 To Len(strInstring)
            strNext = CStr(AscW(Mid(strInstring, _
                                    intIndex1, _
                                    1)))
            If Not booCondensed Then
                strNext = align(strNext, 5, ENUalign.alignRight, "0")
            End If
            objStringBuilder.Append("&#" & strNext)
        Next intIndex1
        Return (objStringBuilder.ToString)
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
    ' C H A N G E   R E C O R D ---------------------------------------------
    '   DATE      PROGRAMMER    DESCRIPTION OF CHANGE
    ' --------    ----------    ---------------------------------------------
    ' 05 13 03    Nilges        1. Replace msgbox with error handler
    '                           2. Add error handler for file open
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
                errorHandler("Cannot erase pre-existing file: " & Err.Number & Err.Description, _
                             Name, _
                             "string2File", _
                             "Returning False")
                Return (False)
            End Try
        End If
        ' --- Open the file for stream output
        Dim objStream As IO.FileStream
        Try
            objStream = New IO.FileStream(strFileId, IO.FileMode.CreateNew)
        Catch ex As Exception
            errorHandler("Cannot create stream: " & Err.Number & " " & Err.Description, _
                         Name, _
                         "string2File", _
                         "Returning False" & _
                         vbNewLine & vbNewLine & _
                         ex.ToString)
            Return (False)
        End Try
        ' --- Write the file
        For intIndex1 = 1 To Len(strOutstring)
            Try
                objStream.WriteByte(CByte(Asc(Mid(strOutstring, intIndex1, 1))))
            Catch ex As Exception
                errorHandler("Could not write file: " & Err.Number & " " & Err.Description, _
                             Name, _
                             "string2File", _
                             "Returning False" & _
                             vbNewLine & vbNewLine & _
                             ex.ToString)
                objStream.Close()
                Return (False)
            End Try
        Next intIndex1
        ' --- Close the file
        Try
            objStream.Close()
        Catch ex As Exception
            errorHandler("Cannot close stream: " & Err.Number & Err.Description, _
                         Name, _
                         "string2File", _
                         "Returning False" & _
                         vbNewLine & vbNewLine & _
                         ex.ToString)
            Return (False)
        End Try
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' String to object
    '
    '
    ' This method converts a string containing the value of an object in the
    ' format returned by object2String back to the original value of the
    ' object, where possible.
    '
    ' The optional overload string2Object(s, OK) will set OK to True if the
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
    ' 07 05 03   Nilges         1. In string2Object_value2Type_: set object 
    '                              to unquoted string and not Nothing 
    '                           2. In string2Object_value2Type_: do not
    '                              use display2String when string is a
    '                              number
    ' 10 05 03   Nilges         Accept Integer and Long as decorated types
    '                           Prefix deco type name with SYSTEM. as needed
    ' 10 18 03   Nilges         Bug: need to add a case where a string is numeric
    '                           but has string type
    '
    Public Overloads Shared Function string2Object(ByVal strInstring As String) As Object
        Dim booOK As Boolean
        Return (string2Object(strInstring, booOK))
    End Function

    Public Overloads Shared Function string2Object(ByVal strInstring As String, _
                                                   ByRef booOK As Boolean) As Object
        Dim objValue As Object
        Dim strInstringWork As String = Trim(strInstring)
        Dim strType As String
        Dim strValue As String = strInstringWork
        string2Object_unDeco_(strInstring, strType, strValue)
        booOK = True
        If UCase(strValue) = "NOTHING" Then Return (Nothing)
        If Not string2Object_value2Type_(strValue, strType, objValue) Then
            booOK = False : Return (Nothing)
        End If
        Return canonicalTypeCast(objValue, strType)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of string2Object, undecorate string if possible
    '
    '
    Private Shared Function string2Object_unDeco_(ByVal strInstring As String, _
                                                  ByRef strType As String, _
                                                  ByRef strValue As String) As Boolean
        Dim intIndex1 As Integer = InStr(strInstring, "(")
        If intIndex1 < 2 _
           OrElse _
           intIndex1 > InStr(strInstring & Chr(34), Chr(34)) Then
            Return (False)
        End If
        If Mid(strInstring, Len(strInstring)) <> ")" Then Return (False)
        strType = Mid(strInstring, 1, intIndex1 - 1)
        strValue = Mid(strInstring, _
                       intIndex1 + 1, _
                       Len(strInstring) - intIndex1 - 1)
        If Not IsNumeric(strValue) Then
            strValue = display2String(strValue, _
                                      ENUdisplay2StringSyntax.VBExpressionCondensed)
        End If
        Return (True)
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
    Private Shared Function string2Object_value2Type_(ByVal strInstring As String, _
                                                      ByRef strType As String, _
                                                      ByRef objValue As Object) As Boolean
        If isQuoted(strInstring) Then
            objValue = dequote(strInstring)
            If Mid(strInstring, 1, 1) = "'" Then
                If Len(CStr(objValue)) <> 1 Then
                    errorHandler("String " & strInstring & " uses single quotes but is not one char")
                    Return (False)
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
                Dim strTypeWork As String = UCase(strType)
                If Len(strTypeWork) < 7 OrElse Mid(strTypeWork, 1, 7) <> "SYSTEM." Then
                    strTypeWork = "SYSTEM." & strTypeWork
                End If
                Select Case UCase(strTypeWork)
                    Case "SYSTEM.BYTE" : objValue = CByte(strInstring)
                    Case "SYSTEM.INT16" : objValue = CShort(strInstring)
                    Case "SYSTEM.INT32" : objValue = CInt(strInstring)
                    Case "SYSTEM.INT64" : objValue = CLng(strInstring)
                    Case "SYSTEM.SINGLE" : objValue = CSng(strInstring)
                    Case "SYSTEM.DOUBLE" : objValue = CDbl(strInstring)
                    Case "SYSTEM.INTEGER" : objValue = CInt(strInstring)
                    Case "SYSTEM.LONG" : objValue = CLng(strInstring)
                    Case "SYSTEM.STRING" : objValue = CStr(strInstring)
                    Case Else
                        errorHandler("Unsupported string type " & _
                                     enquote(strType), _
                                     Name, _
                                     "string2Object_value2Type_", _
                                     "Not converting the value")
                End Select
            End If
            strType = objValue.GetType.ToString
        Else
            objValue = strInstring
        End If
        Return (True)
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
        Dim strInstringWork As String = Trim(strInstring)
        Dim intValue As Integer = -1
        Try
            intValue = CInt(strInstringWork)
        Catch
            Try
                Dim sngValue As Single = CSng(strInstringWork)
                Return (sngValue)
            Catch
                Dim objRegEx As System.Text.RegularExpressions.Regex
                Try
                    objRegEx = New System.Text.RegularExpressions.Regex("^[0123456789.+-]+[ ]*%$")
                Catch
                    errorHandler("Cannot create regular expression", _
                                 Name, "string2Percent", _
                                 Err.Number & " " & Err.Description & _
                                 vbNewLine & vbNewLine & _
                                 "Return -1 to indicate non-percent value")
                    Return (-1)
                End Try
                If objRegEx.IsMatch(strInstringWork) Then
                    Try
                        Return CSng(word(strInstringWork, 1))
                    Catch
                        Return (-1)
                    End Try
                End If
            End Try
        End Try
        If intValue < 0 Then Return (-1)
        Return (CSng(intValue) / 100)
    End Function

    ' ----------------------------------------------------------------------
    ' String to Run Length Encoding
    '
    '
    ' Run Length Encoding is a very simple packing method for low-intensity
    ' packing, which represents "runs" of identical characters as an escape
    ' character (known not to be part of the input string), a run length, and
    ' the repeated character. This method accepts any string EXCEPT a string 
    ' containing all possible Unicode/DBCS characters.
    '
    ' It returns a string that commences with a sentinel character that did
    ' not occur anywhere in the input string. In the rest of the string,
    ' wherever the sentinel occurs, it is followed by:
    ' 
    '
    '      *  A character that in the input string was repeated at least three 
    '         times
    '
    '      *  The repetition count, represented as a Unicode or DBCS character.
    '
    '
    ' A null string is returned on failure to create the Run Length Encoding.
    '
    ' This method has two overloads:
    '
    '
    '      *  string2RLE(instring) returns the run length encoded string or
    '         a null string, indicating failure.
    '
    '         Just before returning the rle, this overload will check the
    '         result by comparing the return string, converted back using
    '         RLE2string, to the original string. An error occurs if these
    '         strings don't match. 
    '
    '      *  string2RLE(instring, False) returns the run length encoded 
    '         string or a null string, indicating failure. This overload does
    '         not convert the rle to match it against the input. For best
    '         results, use this overload when input strings can be very long. 
    '
    '
    ' The most suitable strings for this method are strings with long runs of
    ' blanks or other identical characters.
    '
    ' See also RLE2string.
    '
    '
    Public Overloads Function string2RLE(ByVal strInstring As String) As String
        Return (string2RLE(strInstring, True))
    End Function
    Public Overloads Function string2RLE(ByVal strInstring As String, _
                                         ByVal booCheck As Boolean) As String
        Dim strSentinel As String = string2Sentinel(strInstring)
        If strSentinel = "" Then Return ("")
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch ex As Exception
            errorHandler("Can't create string builder: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "string2RLE", _
                         "Returning a null string")
            Return ("")
        End Try
        Dim strRepeatValue As String
        Dim strNext As String
        Do While intIndex1 <= Len(strInstring)
            strNext = Mid(strInstring, intIndex1, 1)
            For intIndex2 = intIndex1 To Len(strInstring)
                If Mid(strInstring, intIndex2, 1) <> strNext Then Exit For
                Try
                    strRepeatValue = ChrW(intIndex2 - intIndex1)
                Catch
                    ' The DBCS/Unicode character cannot hold the large repetition count:
                    ' we will generate multiple escapes
                    Exit For
                End Try
            Next intIndex2
            If intIndex2 - intIndex1 >= 3 Then
                append(objStringBuilder, "", strSentinel & strRepeatValue & strNext)
            End If
            intIndex1 = intIndex2
        Loop
        Dim strOutstring As String = strSentinel & objStringBuilder.ToString
        If booCheck _
           AndAlso _
           RLE2String(strOutstring) <> strInstring Then
            errorHandler("RLE string " & _
                         enquote(ellipsis(object2String(strOutstring), 32)) & " " & _
                         "fails to match input string " & _
                         enquote(ellipsis(object2String(strInstring), 32)), _
                         Name, "string2RLE", _
                         "Returning the (possibly incorrect) RLE")
        End If
        Return (strOutstring)
    End Function

    ' ----------------------------------------------------------------------
    ' Convert string to sentinel
    '
    '
    ' This method finds a character or a string which does not occur in
    ' the input string; a common use for this method is to find precise
    ' delimiters for strings.
    '
    ' This method searches first for a one-character delimiter. If this
    ' delimiter is not found, then it will search for a two character
    ' delimiter, and so on until one is found, or the intLimit (see below)
    ' number of probes is complete, where it will return the null string,
    ' indicating that a sentinel could not be found.
    '
    ' The input string should be passed in strInstring, and the following
    ' optional parameters are supported.
    '
    '
    '      *  The strCharacters parameter can define preferred sentinel
    '         candidates as single characters, only. strCharacters
    '         defaults to comma, semicolon, and colon.
    '
    '      *  The intLimit parameter can specify the maximum number of
    '         probes for one sentinel candidate. It defaults to -1 (no
    '         limit).
    '
    '         Note that there's a possibility, for a long string 
    '         with many combinations of all characters, 
    '         that the search will take a long time,without returning
    '         a non-null string. Therefore callers of this method should
    '         always check to make sure its result is not the null string.
    '
    '
    Public Shared Function string2Sentinel(ByVal strInstring As String, _
                                            Optional ByVal strCharacters As String = ",;:", _
                                            Optional ByVal intLimit As Integer = -1) _
                    As String
        If strCharacters = "" Then
            errorHandler("strCharacters cannot be a null string", _
                         Name, "string2Sentinel", _
                         "Returning a null sentinel")
            Return ("")
        End If
        If intLimit < -1 Then
            errorHandler("intLimit is not valid", _
                         Name, "string2Sentinel", _
                         "Returning a null sentinel")
            Return ("")
        End If
        Dim intLimitWork As Integer
        Dim strSentinel As String = string2Sentinel_(strInstring, _
                                                     strCharacters, _
                                                     intLimit, _
                                                     intLimitWork)
        If strSentinel <> "" _
           OrElse _
           intLimit <> -1 AndAlso intLimitWork >= intLimit Then Return (strSentinel)
        Dim intWidth As Integer = 1
        Do
            strSentinel = string2Sentinel_(strInstring, _
                                           intWidth, _
                                           intLimit, _
                                           intLimitWork)
            If strSentinel <> "" OrElse intLimitWork >= intLimit Then Return (strSentinel)
            intWidth += 1
        Loop
    End Function

    ' -----------------------------------------------------------------
    ' Check for a specific list of characters, or a sentinel of a
    ' specific width, on behalf of string2Sentinel
    '
    '
    ' --- Check for a listed set of characters
    Private Overloads Shared Function string2Sentinel_(ByVal strInstring As String, _
                                                        ByVal strCharacters As String, _
                                                        ByVal intLimit As Integer, _
                                                        ByRef intLimitReached As Integer) _
            As String
        Dim intIndex1 As Integer
        Dim strNext As String
        For intIndex1 = 1 To Len(strCharacters)
            strNext = Mid(strCharacters, intIndex1, 1)
            If InStr(strInstring, strNext) = 0 Then Return (strNext)
            intLimitReached += 1
            If intLimitReached >= intLimit Then Return ("")
        Next intIndex1
        Return ("")
    End Function
    ' --- Check for a set of strings, of the specified width
    Private Overloads Shared Function string2Sentinel_(ByVal strInstring As String, _
                                                        ByVal intWidth As Integer, _
                                                        ByVal intLimit As Integer, _
                                                        ByRef intLimitReached As Integer) _
            As String
        Dim lngIndex1 As Long
        Dim strNext As String
        Dim strAllAscii As String = range2String(0, 255)
        For lngIndex1 = 1 To 32767
            strNext = long2BaseN(lngIndex1, strAllAscii)
            If InStr(strInstring, strNext) = 0 Then Return (strNext)
            intLimitReached += 1
            If intLimitReached >= intLimit Then Return ("")
        Next lngIndex1
        Return ("")
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
    ' sequence listed: Boolean, Short, Integer, Long, Double and String, returning
    ' the first type successfully converted.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION
    ' --------     ----------   --------------------------------------------
    ' 10 23 01     Nilges       Added strTypePreference parameter and
    '                           Char and Boolean types.
    '
    ' 06 04 03     Nilges       Corrected documentation
    '
    ' 02 04 03     Nilges       Corrected documentation
    '
    '
    Public Shared Function string2ValueObject(ByVal strInstring As String) As Object
        Try
            Return (CByte(strInstring))
        Catch
            Try
                Return (CShort(strInstring))
            Catch
                Try
                    Return (CInt(strInstring))
                Catch
                    Try
                        Return (CLng(strInstring))
                    Catch
                        Try
                            Return (CSng(strInstring))
                        Catch
                            Try
                                Return (CDbl(strInstring))
                            Catch
                                Return (strInstring)
                            End Try
                        End Try
                    End Try
                End Try
            End Try
        End Try
    End Function

    ' ---------------------------------------------------------------------
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
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE     PROGRAMMER     DESCRIPTION OF CHANGE
    ' --------   ----------     ---------------------------------------
    ' 06 01 03   Nilges         Converted from VB-6 version
    '                           Make sure filenumber starts at one
    '
    '
    Public Shared Function tempFileid(ByVal strPath As String, _
                                      Optional ByVal strFileType As String = "") As String
        Dim colDir As Collection
        Dim intIndex1 As Integer
        Dim intFileNumber As Integer
        Dim intNext As Integer
        Dim strDir As String
        Dim strFileTypeWork As String
        strFileTypeWork = Trim(strFileType)
        If strFileTypeWork = "" Then strFileTypeWork = "TMP"
        Try
            colDir = New Collection
        Catch ex As Exception
            errorHandler("Cannot create collection: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "tempFileid", _
                         "Returning null string" & vbNewLine & vbNewLine & _
                         ex.ToString)
            Return ("")
        End Try
        strDir = Dir(appendPath(strPath, "*." & strFileType))
        Do While strDir <> ""
            Try
                colDir.Add(strDir)
            Catch ex As Exception
                errorHandler("Cannot extend collection: " & _
                             Err.Number & " " & Err.Description, _
                             Name, "tempFileid", _
                             "Returning null string" & vbNewLine & vbNewLine & _
                             ex.ToString)
                Return ("")
            End Try
            strDir = Dir$()
        Loop
        intFileNumber = 0
        For intIndex1 = 1 To colDir.Count
            If tempFileid_isTempFileid_(CStr(colDir.Item(intIndex1)), _
                                        intNext, _
                                        strFileType:=strFileType) Then
                intFileNumber = Math.Max(intFileNumber, intNext)
            End If
        Next intIndex1
        intFileNumber = intFileNumber + 1
        If intFileNumber < 10000 Then
            Return (appendPath(strPath, _
                              "temp" & _
                              alignRight(CStr(intFileNumber), 4, strFill:="0") & _
                              "." & _
                              strFileTypeWork))
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Return True (and set lngNumber) when file id is in the form 
    ' tempnnnn.TMP (or tempnnnn.typ when strFileType is present and not 
    ' null or blank) on behalf of tempFileid
    '
    '
    Private Shared Function tempFileid_isTempFileid_(ByVal strFileid As String, _
                                                     ByRef intNumber As Integer, _
                                                     Optional ByVal strFileType As String = "") As Boolean
        Dim strFileTypeWork As String
        Dim strNumber As String
        Dim strSplit() As String
        Try
            strSplit = Split(strFileid, ".")
        Catch ex As Exception
            errorHandler("Cannot split: " & Err.Number & " " & Err.Description, _
                         Name, "", _
                         "Returning False" & vbNewLine & vbNewLine & _
                         ex.ToString)
            Return (False)
        End Try
        If UBound(strSplit) <> 1 Then Return (False)
        If Len(strSplit(0)) <> 8 Then Return (False)
        strFileTypeWork = Trim(UCase(strFileType))
        If strFileTypeWork = "" Then strFileTypeWork = "TMP"
        If UCase(strSplit(1)) <> strFileTypeWork Then Return (False)
        If UCase(Mid$(strSplit(0), 1, 4)) <> "TEMP" Then Return (False)
        strNumber = Mid$(strSplit(0), 5)
        If verify(strNumber, "0123456789") = 0 Then
            Try
                intNumber = CInt(strNumber)
            Catch ex As Exception
                Return (False)
            End Try
            Return (True)
        End If
    End Function

#If INCLUDE_TEST Then
    ' -----------------------------------------------------------------------------
    ' Regression tests for most methods                                           
    '
    '
    ' This method runs precoded tests on many methods, and places a test report in
    ' its reference string parameter strReport. It returns True if the tests all
    ' succeed, False otherwise.
    '
    ' This method will be generated as part of the utility library only
    ' when the compile-time symbol INCLUDE_TEST is True.
    '
    ' See also itemTest.
    '
    '
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   ---------------------------------------
    ' 06 24 03     Nilges       Added tests of collection depth and
    '                           itemMaxLength
    '
    ' 06 26 03     Nilges       Use collectionCompare in place of
    '                           string compare to test collection2String
    '                           and string2Collection, to avoid false
    '                           negatives
    '
    ' 01 31 04     Nilges       Added INCLUDE_TEST compile-time symbol
    '                           to control inclusion of this method.
    '
    '
    Public Shared Function test(ByRef strReport As String) As Boolean
        Dim booOK As Boolean
        Dim intErrCount As Integer
        strReport = "Testing utilities at " & Now & vbNewLine & vbNewLine
        test_(strReport, intErrCount, "abbrev", True, "abc", "abcd")
        test_(strReport, intErrCount, "alignLeft", "abc...", "abc", 6, ".")
        test_(strReport, intErrCount, "alignCenter", "..abc..", "abc", 7, ".")
        test_(strReport, intErrCount, "alignRight", "***abc", "abc", 6, "*")
        test_(strReport, intErrCount, "append", "c", "", " ", "c", False, -1)
        test_(strReport, intErrCount, "append", "a b", "a", " ", "b", False, -1)
        test_(strReport, intErrCount, "append", "b a", "a", " ", "b", True, -1)
        test_(strReport, intErrCount, "appendPath", "c:\temp.txt", "c:", "temp.txt")
        test_(strReport, intErrCount, "appendPath", "c:\temp\temp.txt", "c:\temp\", "temp.txt")
        test_(strReport, intErrCount, "baseN2Long", "13", "bbab", "ab", 0, False)
        test_(strReport, intErrCount, "baseN2Long", "15", "bbbb", "ab", 0, False)
        test_(strReport, intErrCount, "baseN2Long", -1, "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb", "ab", 32, False)
        test_(strReport, intErrCount, "breakLongWords", "", "", 32)
        test_(strReport, intErrCount, "breakLongWords", _
                                      "11111111 22 333333 333 4 55555555 6666\ 6666 77\ 7777777\ 777\ 77777\ 7777777\ 7777 123456 789ABC DEFGHI JKLMNO PQRSTU VWXYZ", _
                                      "11111111 22 333333333 4 55555555 6666\6666 77\7777777\777\77777\7777777\7777 123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", _
                                      8)
        test_(strReport, intErrCount, "breakLongWords", _
                                      "C:/ Microsoft/ Windows/ Flump/ Kanump/ Temp.TXT", _
                                      "C:/Microsoft/Windows/Flump/Kanump/Temp.TXT", _
                                      20)
        test_(strReport, intErrCount, "breakLongWords", "", "", 32)
        test_(strReport, intErrCount, "breakLongWords", "", "", 32)
        test_(strReport, intErrCount, "breakLongWords", "", "", 32)
        Dim strWork As String
        test_(strReport, intErrCount, "datatype", True, "32768", "unsignedLongDatatype")
        test_(strReport, intErrCount, "display2String", vbCrLf, "&#00013&#00010", "XML")
        test_(strReport, intErrCount, "display2String", vbCrLf, "\n", "C")
        test_(strReport, intErrCount, "display2String", vbCrLf, "ChrW(13) & ChrW(10)", "vbExpression")
        test_(strReport, intErrCount, "display2String", "ABCDEFG", "range2String(65, 71)", "vbExpressionCondensed")
        test_(strReport, intErrCount, "ellipsis", "This is a long...", "This is a long string", 17, "...", False)
        Dim strTestFileid As String = tempFileid(CurDir)
        If strTestFileid <> "" Then
            strReport &= vbNewLine & vbNewLine & _
                         "Testing string2File and file2String"
            Dim strTestInfo As String = "aa" & ChrW(0)
            Try
                string2File(strTestInfo, strTestFileid)
                booOK = (file2String(strTestFileid) = strTestInfo)
                strReport &= vbNewLine & _
                                "The file as written by string2File " & _
                                CStr(IIf(booOK, "matches", "does not match")) & " " & _
                                "the file as read by file2String"
                If Not booOK Then intErrCount += 1
            Catch ex As Exception
                strReport &= vbNewLine & _
                             "Error in file2String or string2File: " & _
                             Err.Number & " " & Err.Description
            End Try
            Kill(strTestFileid)
            test_(strReport, intErrCount, "fileExists", False, strTestFileid)
        Else
            intErrCount += 1
            strReport &= vbNewLine & vbNewLine & _
                         "Cannot test string2File, file2String or fileExists: tempFileid failed"
        End If
        test_(strReport, intErrCount, "findAbbrev", 2, "AB", "Al's Abdul's Joe's Jane's Moon's")
        test_(strReport, intErrCount, "findAbbrev", 4, "ja", "Al's Abdul's Joe's Jane's Moon's")
        test_(strReport, intErrCount, "findAbbrev", 0, "k", "Al's Abdul's Joe's Jane's Moon's")
        test_(strReport, intErrCount, "findAbbrev", 1, "Al's", "Al's Abdul's Joe's Jane's Moon's")
        test_(strReport, _
              intErrCount, _
              "findItem", _
              15, _
              smallBusinessList, _
              "BOB'S BERSERKO LOUNGE", _
              ", ", _
              False, _
              True)
        Dim dblIndex1 As Double
        Try
            For dblIndex1 = 0 To 10 Step 0.1
                strWork &= vbNewLine & _
                           copies("*", _
                                  CInt(histogram(Math.Abs(Math.Sin(dblIndex1)), _
                                       dblRangeMin:=1, _
                                       dblRangeMax:=32, _
                                       dblValueMax:=Math.Sin(10))))
            Next dblIndex1
        Catch ex As Exception
            strWork = "Histogram test failed"
        End Try
        strReport &= vbNewLine & vbNewLine & _
                     string2Box(strWork, "Histogram Test: should be a sine wave")
        test_(strReport, intErrCount, "itemPhrase", "You Eat You Pay You Shuddup Now Cafe, O'Reilly's Bar, Man with a Van Worldwide Moving and Storage Company", smallBusinessList, 4, 3, ", ", False)
        test_(strReport, intErrCount, "itemMaxWidth", 7, "Item 1,Item 21,,", ",", False)
        test_(strReport, intErrCount, "long2BaseN", "1100", 12, "01", 0)
        test_(strReport, intErrCount, "range2String", Chr(0) & Chr(1), 0, 1)
        test_char2Name_(strReport, intErrCount)
        test_formatOutline_(strReport, intErrCount)
        Return (intErrCount = 0)
    End Function

    ' -----------------------------------------------------------------
    ' Run one test on behalf of the test method
    '
    '
    Private Shared Sub test_(ByRef strReport As String, _
                             ByRef intErrCount As Integer, _
                             ByVal strMethod As String, _
                             ByVal objExpected As Object, _
                             ByVal ParamArray objParameters() As Object)
        Dim intIndex1 As Integer
        Dim strCall As String
        Dim strNext As String
        For intIndex1 = 0 To UBound(objParameters)
            strCall = append(strCall, _
                             ", ", _
                             object2String(objParameters(intIndex1)))
        Next intIndex1
        strCall = strMethod & "(" & strCall & ")"
        strReport = append(strReport, vbNewLine & vbNewLine, _
                           "Testing " & strMethod & " using: " & _
                           strCall)
        Dim objActual As Object
        Dim strErrInfo As String
        Try
            Select Case UBound(objParameters)
                Case -1 : objActual = utility(strMethod)
                Case 0 : objActual = utility(strMethod, _
                                             objParameters(0))
                Case 1 : objActual = utility(strMethod, _
                                             objParameters(0), _
                                             objParameters(1))
                Case 2 : objActual = utility(strMethod, _
                                            objParameters(0), _
                                            objParameters(1), _
                                            objParameters(2))
                Case 3 : objActual = utility(strMethod, _
                                            objParameters(0), _
                                            objParameters(1), _
                                            objParameters(2), _
                                            objParameters(3))
                Case 4 : objActual = utility(strMethod, _
                                            objParameters(0), _
                                            objParameters(1), _
                                            objParameters(2), _
                                            objParameters(3), _
                                            objParameters(4))
                Case 5 : objActual = utility(strMethod, _
                                            objParameters(0), _
                                            objParameters(1), _
                                            objParameters(2), _
                                            objParameters(3), _
                                            objParameters(4), _
                                            objParameters(5))
                Case 6 : objActual = utility(strMethod, _
                                            objParameters(0), _
                                            objParameters(1), _
                                            objParameters(2), _
                                            objParameters(3), _
                                            objParameters(4), _
                                            objParameters(5), _
                                            objParameters(6))
                Case Else
                    errorHandler("Unsupported parameter count " & _
                                 UBound(objParameters) + 1, _
                                 "test_", "objUtilities", "Test has failed")
            End Select
            Dim strExpected As String = object2String(objExpected)
            Dim strActual As String = object2String(objActual)
            If (strExpected <> strActual) Then
                intErrCount += 1
                strErrInfo = "Expected value " & strExpected & " " & _
                             "fails to match actual value " & strActual
            End If
        Catch ex As Exception
            intErrCount += 1
            strErrInfo = Err.Number & " " & Err.Description
        End Try
        strReport = append(strReport, vbNewLine & vbNewLine, _
                           "Test " & _
                           CStr(IIf(strErrInfo = "", "succeeded", "failed")) & ": " & _
                           test__result2String_(objExpected, _
                                                CStr(IIf(strErrInfo = "", "Expected/actual", "Expected"))))
        If strErrInfo <> "" Then
            strReport = append(strReport, vbNewLine, _
                               test__result2String_(objActual, "Actual")) & _
                        vbNewLine & vbNewLine & _
                        string2Box(soft2HardParagraph(strErrInfo), _
                                   "E R R O R   I N   T E S T")
        End If
    End Sub

    ' -----------------------------------------------------------------------------
    ' Convert test result to string on behalf of test_
    '
    '
    ' This method returns a one line string or a boxed text when the input is
    ' multiline.
    '
    '
    Private Shared Function test__result2String_(ByVal objResult As Object, _
                                                 ByVal strType As String) As String
        Dim strResult As String
        Try
            strResult = string2Display(CStr(objResult))
        Catch : End Try
        If verify(strResult, vbNewLine, True) <> 0 Then
            Return (vbNewLine & string2Box(strResult, strType & " Result"))
        Else
            Return (strType & " result: " & string2Display(object2String(objResult)))
        End If
    End Function

    ' -----------------------------------------------------------------------------
    ' Test the outline formatter
    '
    '
    Private Shared Sub test_formatOutline_(ByRef strReport As String, _
                                           ByRef intErrorCount As Integer)
        strReport &= vbNewLine & vbNewLine & _
                     "Testing formatOutline: making some random outlines: " & _
                     "for each formatted outline, it must match, word by word, " & _
                     "the unformatted outline"
        Dim booSpaceOnMultiline As Boolean
        Dim booTopLevelEdit As Boolean
        Dim intCount As Integer = CInt(Rnd() * 20)
        Dim intIndex1 As Integer
        Dim intSpacing As Integer
        Dim intWidth As Integer
        Dim strExplanation As String
        Dim strNext As String
        Dim strNextFormatted As String
        For intIndex1 = 1 To intCount
            strNext = ""
            Select Case intIndex1
                Case 1
                    test_formatOutline__mkRandomOutline_(strNext, "1", 0, 0)
                Case 2
                    test_formatOutline__mkRandomOutline_(strNext, "1", CInt(Rnd() * 1024), 1)
                Case 3
                    test_formatOutline__mkRandomOutline_(strNext, "1", CInt(Rnd() * 1024), 2)
                Case Else
                    test_formatOutline__mkRandomOutline_(strNext, "1", CInt(Rnd() * 1024))
            End Select
            intWidth = 32 + CInt(Rnd() * 40)
            booTopLevelEdit = CBool(IIf(Rnd() < 0.5, True, False))
            intSpacing = CInt(Rnd() * 2)
            booSpaceOnMultiline = CBool(IIf(Rnd() < 0.5, True, False))
            strExplanation = "The following outline was created for a text width of " & _
                             intWidth & ". " & _
                             "Top level Dewey editing is " & _
                             CStr(IIf(booTopLevelEdit, "", "not ")) & " in effect. " & _
                             "Line spacing is " & intSpacing & " extra lines per outline item " & _
                             CStr(IIf(booSpaceOnMultiline, _
                                      "but it is conditional on multiple line entries", _
                                      "and it is applied unconditionally"))
            Try
                strNextFormatted = formatOutline(strNext, _
                                                 intWidth, _
                                                 booTopLevelEdit, _
                                                 intSpacing:=intSpacing, _
                                                 booSpaceOnMultiline:=booSpaceOnMultiline)
            Catch
                strReport &= vbNewLine & vbNewLine & _
                             "Error in formatOutline: " & _
                             Err.Number & " " & Err.Description
                intErrorCount += 1
                Return
            End Try
            strReport &= vbNewLine & vbNewLine & _
                         string2Box(soft2HardParagraph(strExplanation) & _
                                    vbNewLine & vbNewLine & _
                                    strNextFormatted, _
                                    "Formatted outline " & intIndex1 & " of " & intCount)
            If phrase(Replace(Replace(strNextFormatted, ".", ""), vbNewLine, " "), 1) _
               <> _
               phrase(Replace(Replace(strNext, ".", ""), vbNewLine, " "), 1) Then
                strReport &= vbNewLine & vbNewLine & _
                             "Formatted outline fails to match unformatted after editing"
                intErrorCount += 1
            End If
        Next intIndex1
    End Sub

    ' -----------------------------------------------------------------------------
    ' Make a random outline
    '
    '
    ' --- No line limit
    Private Overloads Shared Sub test_formatOutline__mkRandomOutline_ _
                                 (ByRef strOutline As String, _
                                  ByVal strDewey As String, _
                                  ByVal intMaxLength As Integer)
        test_formatOutline__mkRandomOutline_(strOutline, strDewey, intMaxLength, -1)
    End Sub
    ' --- Restrict the line count
    Private Overloads Shared Sub test_formatOutline__mkRandomOutline_ _
                                 (ByRef strOutline As String, _
                                  ByVal strDewey As String, _
                                  ByVal intMaxLength As Integer, _
                                  ByVal intMaxLines As Integer)
        If Len(strOutline) > intMaxLength OrElse intMaxLines = 0 Then Return
        strOutline = strOutline & _
                     CStr(IIf(strOutline = "", "", vbNewLine)) & _
                     strDewey & " " & _
                     copies("This is a test ", CInt(Rnd() * 9) + 1)
        If intMaxLines <> -1 AndAlso items(strOutline, vbNewLine, False) >= intMaxLines Then
            Return
        End If
        Dim sngRnd As Single = Rnd()
        If sngRnd < 0.33 Then
            ' --- Create next parent outline entry
            Dim strDeweyParent As String = deweyParent(strDewey)
            If strDeweyParent = "" Then
                strDeweyParent = CStr(CInt(strDewey) + 1)
            Else
                strDeweyParent = incrementDewey(strDeweyParent)
            End If
            test_formatOutline__mkRandomOutline_(strOutline, _
                                                 strDeweyParent, _
                                                 intMaxLength)
        ElseIf sngRnd < 0.66 Then
            ' --- Create child outline entry
            test_formatOutline__mkRandomOutline_(strOutline, _
                                                 strDewey & ".1", _
                                                 intMaxLength)
        Else
            ' --- Create sibling outline entry
            test_formatOutline__mkRandomOutline_(strOutline, _
                                                 incrementDewey(strDewey), _
                                                 intMaxLength)
        End If
    End Sub

    ' -----------------------------------------------------------------------------
    ' Test the outline formatter
    '
    '
    Private Shared Sub test_char2Name_(ByRef strReport As String, _
                                       ByRef intErrorCount As Integer)
        strReport &= vbNewLine & vbNewLine & _
                     "Testing char2Name and name2Char using Unicode/DBCS " & _
                     "characters 0..300"
        Dim booOK As Boolean
        Dim intIndex1 As Integer
        Dim strChar As String
        Dim strName As String
        For intIndex1 = 0 To 300
            strChar = ChrW(intIndex1)
            strName = char2Name(strChar)
            booOK = (name2Char(strName) = strChar)
            strReport &= vbNewLine & _
                         "The name of the character " & intIndex1 & " is " & _
                         enquote(strName) & " " & _
                         "and this name " & _
                         CStr(IIf(booOK, "", "DOES NOT ")) & _
                         "converts back to the original character correctly"
            If Not booOK Then intErrorCount += 1
        Next intIndex1
    End Sub
#End If

    ' -----------------------------------------------------------------------------
    ' Indicate whether a test method is available
    '
    '
    Public Shared Function testAvailable() As Boolean
#If INCLUDE_TEST Then
        Return True
#End If
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
            intIndex2 = InStr(strSourceCharacterSet, Mid(strInstring, intIndex1, 1))
            If intIndex2 <> 0 AndAlso intIndex2 <= Len(strTargetCharacterSet) Then
                strOutstring &= Mid(strTargetCharacterSet, intIndex2, 1)
            Else
                strOutstring &= Mid(strInstring, intIndex1, 1)
            End If
        Next intIndex1
        Return (strOutstring)
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
    '      *  utility("display2String", str, syntax, graphic)
    '
    '         str should be a string in the displayable syntax that is
    '         returned by string2Display and identified, as a string,
    '         in syntax.  syntax should be C, XML, vbExpression,
    '         vbExpressionCondensed or determine; note that determine
    '         will try to determine the syntax from clues.
    '
    '         str is converted back to its original form and returned.
    '
    '         The graphic parameter should specify the graphic character set. 
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
    '         Searches str for the item in tgt, returning the tgt index from
    '         1 or zero.
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
    '      *  utility("itemMaxWidth", str, delim, True|False)
    '
    '         Returns the width of the longest delimited item.
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
    '      *  utility("string2Display", str, syntax, graphic)
    '
    '         str (which may be a string containing Unicode and non-graphical
    '         characters) is converted to a string containing only the
    '         characters in the string specified in the graphic parameter.
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
    ' C H A N G E   R E C O R D ---------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   ---------------------------------------
    ' 05 26 03     Nilges       Bug: wrong parameter number in call to
    '                           verify
    ' 06 01 03     Nilges       Added align
    '                           Bug: incorrect call for string2File
    ' 06 24 03     Nilges       Added itemMaxWidth
    ' 02 21 04     Nilges       Added breakLongWords
    '
    Public Shared Function utility(ByVal strUtilityName As String, _
                                   ByVal ParamArray objParameters() As Object) As Object
        Try
            Select Case UCase(Trim(strUtilityName))
                Case "ABBREV"
                    Return (abbrev(CStr(objParameters(0)), CStr(objParameters(1))))
                Case "ALIGN"
                    Return (align(CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  CType(objParameters(2), ENUalign), _
                                  CStr(objParameters(3))))
                Case "ALIGNCENTER"
                    Return (align(CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  ENUalign.alignCenter, _
                                  CStr(objParameters(2))))
                Case "ALIGNLEFT"
                    Return (align(CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  ENUalign.alignLeft, _
                                  CStr(objParameters(2))))
                Case "ALIGNRIGHT"
                    Return (align(CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  ENUalign.alignRight, _
                                  CStr(objParameters(2))))
                Case "APPEND"
                    Return (append(CStr(objParameters(0)), _
                                   CStr(objParameters(1)), _
                                   CStr(objParameters(2)), _
                                   booToStart:=CBool(objParameters(3)), _
                                   intMaxLength:=CInt(objParameters(4))))
                Case "APPENDPATH"
                    Return (appendPath(CStr(objParameters(0)), CStr(objParameters(1))))
                Case "BASEN2LONG"
                    Return (baseN2Long(CStr(objParameters(0)), _
                                       CStr(objParameters(1)), _
                                       CByte(objParameters(2)), _
                                       booIgnoreCase:=CBool(objParameters(3))))
                Case "BREAKLONGWORDS"
                    Return (breakLongWords(CStr(objParameters(0)), _
                                           CInt(objParameters(1))))
                Case "DATATYPE"
                    Return (datatype(CStr(objParameters(0)), CStr(objParameters(1))))
                Case "DEQUOTE"
                    Return (dequote(CStr(objParameters(0))))
                Case "DISPLAY2STRING"
                    Return (display2String(CStr(objParameters(0)), _
                                           display2String_syntax2Enum_(CStr(objParameters(1)))))
                Case "ELLIPSIS"
                    Return (ellipsis(CStr(objParameters(0)), _
                                    CInt(objParameters(1)), _
                                    strEllipsis:=CStr(objParameters(2)), _
                                    booLeftEllipsis:=CBool(objParameters(3))))
                Case "FILE2STRING"
                    Return (file2String(CStr(objParameters(0))))
                Case "FILEEXISTS"
                    Return (fileExists(CStr(objParameters(0))))
                Case "FINDABBREV"
                    Return (findAbbrev(CStr(objParameters(0)), CStr(objParameters(1))))
                Case "FINDITEM"
                    Return (findItem(CStr(objParameters(0)), _
                                    CStr(objParameters(1)), _
                                    CStr(objParameters(2)), _
                                    CBool(objParameters(3)), _
                                    booIgnoreCase:=CBool(objParameters(4))))
                Case "FINDWORD"
                    Return (findWord(CStr(objParameters(0)), _
                                    CStr(objParameters(1)), _
                                    booIgnoreCase:=CBool(objParameters(4))))
                Case "HISTOGRAM"
                    Return (histogram(CDbl(objParameters(0)), _
                                     CDbl(objParameters(1)), _
                                     CDbl(objParameters(2)), _
                                     CDbl(objParameters(2)), _
                                     CDbl(objParameters(2))))
                Case "INDENT"
                    If IsNumeric(objParameters(1)) Then
                        Return (indent(CStr(objParameters(0)), _
                                       CInt(objParameters(1)), _
                                       CStr(objParameters(2))))
                    Else
                        Return (indent(CStr(objParameters(0)), _
                                       CStr(objParameters(1)), _
                                       CStr(objParameters(2))))
                    End If
                Case "INT2DIGITS"
                    Return (int2Digits(CInt(objParameters(0))))
                Case "ISQUOTED"
                    Return (isQuoted(CStr(objParameters(0))))
                Case "ITEM"
                    Return (item(CStr(objParameters(0)), _
                                 CInt(objParameters(1)), _
                                 CStr(objParameters(2)), _
                                 CBool(objParameters(3))))
                Case "ITEMMAXWIDTH"
                    Return (itemMaxWidth(CStr(objParameters(0)), _
                                         CStr(objParameters(1)), _
                                         CBool(objParameters(2))))
                Case "ITEMPHRASE"
                    Return (itemPhrase(CStr(objParameters(0)), _
                                      CInt(objParameters(1)), _
                                      CInt(objParameters(2)), _
                                      CStr(objParameters(3)), _
                                      CBool(objParameters(4))))
                Case "ITEMS"
                    Return (items(CStr(objParameters(0)), CStr(objParameters(1)), CBool(objParameters(2))))
                Case "LONG2BASEN"
                    Return (long2BaseN(CLng(objParameters(0)), _
                                      CStr(objParameters(1)), _
                                      CByte(objParameters(2))))
                Case "PHRASE"
                    Return (phrase(CStr(objParameters(0)), _
                                  CInt(objParameters(1)), _
                                  CInt(objParameters(2))))
                Case "RANGE2STRING"
                    If UCase(objParameters(0).GetType.ToString) = "SYSTEM.STRING" Then
                        Return (range2String(CStr(objParameters(0)), _
                                            CStr(objParameters(1))))
                    Else
                        Return (range2String(CInt(objParameters(0)), _
                                            CInt(objParameters(1))))
                    End If
                Case "SHELL"
                    Dim objAppWinStyle As AppWinStyle
                    Dim strFocus As String = CStr(objParameters(1))
                    Select Case UCase(Trim(strFocus))
                        Case "HIDE" : objAppWinStyle = AppWinStyle.Hide
                        Case "MAXIMIZEDFOCUS" : objAppWinStyle = AppWinStyle.MaximizedFocus
                        Case "MINIMIZEDFOCUS" : objAppWinStyle = AppWinStyle.MinimizedFocus
                        Case "MINIMIZEDNOFOCUS" : objAppWinStyle = AppWinStyle.MinimizedNoFocus
                        Case "NORMALFOCUS" : objAppWinStyle = AppWinStyle.NormalFocus
                        Case "NORMALNOFOCUS" : objAppWinStyle = AppWinStyle.NormalNoFocus
                        Case Else
                            errorHandler("Invalid application focus name " & enquote(strFocus), _
                                         Name, "utility", _
                                         "Focus has been set to Hide")
                            objAppWinStyle = AppWinStyle.Hide
                    End Select
                    Return (Shell(CStr(objParameters(0)), _
                                  Style:=objAppWinStyle, _
                                  Wait:=CBool(objParameters(1))))
                Case "STRING2BOX"
                    Return (string2Box(CStr(objParameters(0)), CStr(objParameters(1)), CStr(objParameters(2))))
                Case "STRING2DISPLAY"
                    Return (string2Display(CStr(objParameters(0)), _
                                           CStr(objParameters(1)), _
                                           CObj(objParameters(2))))
                Case "STRING2FILE"
                    Return (string2File(CStr(objParameters(0)), _
                                        CStr(objParameters(1))))
                Case "TRANSLATE"
                    Return (translate(CStr(objParameters(0)), _
                                     CStr(objParameters(1)), _
                                     CStr(objParameters(2))))
                Case "VERIFY"
                    Return (verify(CStr(objParameters(0)), _
                                   CStr(objParameters(1)), _
                                   CInt(objParameters(2)), _
                                   CBool(objParameters(3))))
                Case "WORD"
                    Return (word(CStr(objParameters(0)), _
                                CInt(objParameters(1))))
                Case "WORDS"
                    Return (words(CStr(objParameters(0))))
                Case Else
                    errorHandler("Unsupported utility name " & enquote(strUtilityName), _
                                 Name, "utility", _
                                 "Returning Nothing")
            End Select
        Catch
            errorHandler("Error in calling utility " & enquote(strUtilityName) & ": " & Err.Number & " " & Err.Description, _
                         Name, "utility", _
                         "Returning Nothing")
        End Try
        Return (Nothing)
    End Function

    ' ----------------------------------------------------------------------
    ' Determine VB code type
    '
    '
    ' Returns VBNET (input string is known to contain VB .Net code, VB6
    ' (input string is known to contain VB-6 code) or a null string to
    ' indicate that the status can't be determined.
    '
    '
    ' C H A N G E   R E C O R D --------------------------------------------
    '   DATE       PROGRAMMER   DESCRIPTION OF CHANGE
    ' --------     ----------   --------------------------------------------
    ' 05 26 01     Nilges       1.  Version 1
    '
    '
    Public Shared Function vbCodeType(ByVal strInstring As String) As String
        If word(item(strInstring, 1, vbNewLine, False), 1) = "VERSION" _
           AndAlso _
           word(item(strInstring, 2, vbNewLine, False), 1) = "BEGIN" Then
            Return ("VB6")
        End If
        If item(strInstring, 1, vbNewLine, False) = "Option Strict" Then
            Return ("VBNET")
        End If
        Return ("")
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
    ' 05 13 03     Nilges       1.  Replaced msgbox by error handling
    '
    '
    Public Overloads Shared Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String) _
           As Integer
        verify = verify(strInstring, _
                        strCharacterSet, _
                        1, _
                        False)
    End Function
    Public Overloads Shared Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String, _
                                            ByVal intStartIndex As Integer) _
           As Integer
        verify = verify(strInstring, _
                        strCharacterSet, _
                        intStartIndex, _
                        False)
    End Function
    Public Overloads Shared Function verify(ByVal strInstring As String, _
                                            ByVal strCharacterSet As String, _
                                            ByVal booMatch As Boolean) _
           As Integer
        verify = verify(strInstring, _
                        strCharacterSet, _
                        1, _
                        booMatch)
    End Function
    ' --- Common overload
    Public Overloads Shared Function verify(ByVal strInstring As String, _
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
            errorHandler("Error in verify function: intStartIndex " & _
                         intStartIndex & " is not valid", _
                         Name, "verify", _
                         "Returning 0")
            Exit Function
        End If
        ' --- Degenerate cases: null input string, unity  set
        If strInstring = "" Then Exit Function
        If booMatch AndAlso intCharacterSetLength = 1 Then
            Return (InStr(intStartIndex, strInstring, strCharacterSet))
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
        Return (intVerify)
    End Function

    ' ----------------------------------------------------------------------
    ' Parse a blank-delimited word
    '
    '
    Public Shared Function word(ByVal strInstring As String, ByVal intIndex As Integer) As String
        Return (item(strInstring, intIndex))
    End Function

    ' ----------------------------------------------------------------------
    ' Count blank-delimited words
    '
    '
    Public Shared Function words(ByVal strInstring As String) As Integer
        Return (items(strInstring))
    End Function

    ' ----------------------------------------------------------------------
    ' XML meta-characters to their names
    '
    '
    Public Shared Function xmlMeta2Name(ByVal strElement As String) As String
        Dim strElementWork As String
        strElementWork = strElement
        strElementWork = Replace(strElementWork, "<", "&lt;")
        strElementWork = Replace(strElementWork, ">", "&gt;")
        strElementWork = Replace(strElementWork, "&", "&amp;")
        strElementWork = Replace(strElementWork, "'", "&apos;")
        strElementWork = Replace(strElementWork, """", "&quot;")
        xmlMeta2Name = strElementWork
    End Function

    ' ----------------------------------------------------------------------
    ' Create a Point object from its X and its Y coordinates
    '
    '
    Public Shared Function xy2Point(ByVal intX As Integer, ByVal intY As Integer) As Point
        If intX < 0 OrElse intY < 0 Then
            errorHandler("Invalid X and/or y coordinates " & _
                         intX & "," & intY, _
                         Name, "xy2Point", _
                         "These values must be zero or positive")
            Return (Nothing)
        End If
        Dim objPoint As Point
        Try
            objPoint = New Point(intX, intY)
        Catch
            errorHandler("Cannot create Point", _
                         Name, "xy2Point", _
                         Err.Number & " " & Err.Description)
            Return (Nothing)
        End Try
        Return (objPoint)
    End Function

    ' ----------------------------------------------------------------------
    ' Create a Size object from its X and its Y values
    '
    '
    Public Shared Function xy2Size(ByVal intX As Integer, ByVal intY As Integer) As Size
        If intX < 0 OrElse intY < 0 Then
            errorHandler("Invalid X and/or y values " & _
                         intX & "," & intY, _
                         Name, "xy2Size", _
                         "These values must be zero or positive")
            Return (Nothing)
        End If
        Dim objSize As Size
        Try
            objSize = New Size(intX, intY)
        Catch
            errorHandler("Cannot create Size", _
                         Name, "xy2Size", _
                         Err.Number & " " & Err.Description)
            Return (Nothing)
        End Try
        Return (objSize)
    End Function

End Class

