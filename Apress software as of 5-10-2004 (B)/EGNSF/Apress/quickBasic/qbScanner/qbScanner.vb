Option Strict On

Imports System.Threading

' *********************************************************************
' *                                                                   *
' * qbScanner     quickBasicEngine: lexical analyzer                  *
' *                                                                   *
' *                                                                   *
' * This class scans input source code and provides, on demand,       *
' * scanned source tokens and scanned lines of source code. This class*
' * uses "lazy" evaluation, scanning the source code only when        *
' * necessary, and when an unparsed token is requested.               *
' *                                                                   *
' * The rest of this comment block describes the following topics.    *
' *                                                                   *
' *                                                                   *
' *      *  The scanner data model                                    *
' *      *  The scanner state                                         *
' *      *  Properties, methods, and events of this class             *
' *      *  Multithreading considerations                             *
' *                                                                   *
' *                                                                   *
' * THE SCANNER DATA MODEL ------------------------------------------ *
' *                                                                   *
' * The state of this class consists of raw source code, and a series *
' * of qbTokens indexed commencing at the start of the input source   *
' * code, and accounting for all characters of source code except     *
' * comments and white space. See qbToken.VB for the data model of the*
' * token itself.                                                     * 
' *                                                                   *
' *                                                                   *
' * THE SCANNER STATE ----------------------------------------------- *
' *                                                                   *
' * The state of the scanner consists of the following.               *
' *                                                                   *
' *                                                                   *
' *      *  strName: the object instance name                         *
' *                                                                   *
' *      *  booUsable: the object usability switch                    *
' *                                                                   *
' *      *  strSourceCode: the source code being scanned              *
' *                                                                   *
' *      *  intLast: the index of the last token parsed               *
' *                                                                   *
' *      *  objQBtoken(): array of scanned qbTokens including the     *
' *         intLast pointer to the last used entry                    *
' *                                                                   *
' *      *  objTokenNext(): array of pending tokens                   *
' *                                                                   *
' *      *  intLineNumber: current line number                        *
' *                                                                   *
' *      *  colLineIndex: collection that provides the starting index *
' *         and length of each line                                   *
' *                                                                   *
' *                                                                   *
' * PROPERTIES, METHODS, AND EVENTS OF THIS CLASS ------------------- *
' *                                                                   *
' * Properties of this class start with an upper case letter; methods *
' * and events with a lower case letter; events end in Event.         *
' *                                                                   *
' * At any time, instance of this class are usable or nonusable. At   *
' * the end of a successful New constructor the instance becomes      *
' * usable, whereas a successful or failed execution of dispose makes *
' * the object unusable.                                              *
' *                                                                   *
' * A serious internal error: failure to create a resource: or "object*
' * abuse" (using the object after a serious error, externally        *
' * reported) makes the object not usable. The mkUnusable method may  *
' * also be used to force the object into the unusable state, and the *
' * Usable property tells the caller whether the object is usable.    *
' *                                                                   *
' * Any method, which does not otherwise return a value, will return  *
' * True on success or False on failure.                              *
' *                                                                   *
' * qbScanner objects are ICloneable and ICompareable: see the clone  *
' * and compareTo methods.                                            *
' *                                                                   *
' *                                                                   *
' * About: this shared, read-only property returns information about  *
' *      this class                                                   *
' *                                                                   *
' * checkToken(i,t): this method checks the scanned tokens for a      *
' *      type or value of token and if it finds the the type or value,*
' *      this method increments a token index.                        *
' *                                                                   *
' *      i should be an Integer, passed by reference. The scan token  *
' *      at this index will be checked. On success this integer will  *
' *      be incremented.                                              *
' *                                                                   *
' *      t should be an expected token type (passed as an enumerator, *
' *      of type ENUtokemType) or it should be an expected string     *
' *      value. This value will be compared ignoring case differences *
' *      with the value in the scanner.                               *
' *                                                                   *
' *      The optional parameter intEndIndex can be used to restrict   *
' *      the check to all tokens up to and including the token at     *
' *      the specified end index.                                     *
' *                                                                   *
' *      See also checkTokenByTypeName.                               *
' *                                                                   *
' * checkTokenByTypeName(i,t): this method checks the scanned tokens  *
' *      for a type of token and if it finds the the type, this method*
' *      increments a token index.                                    *
' *                                                                   *
' *      i should be an Integer, passed by reference. The scan token  *
' *      at this index will be checked. On success this integer will  *
' *      be incremented.                                              *
' *                                                                   *
' *      t should be the name of the expected token as one of the     *
' *      enumerator names defined by ENUtokenType, specified as a     *
' *      string.                                                      *
' *                                                                   *
' *      See also checkTokenByTypeName.                               *
' *                                                                   *
' * clear: this method clears the source code and resets the scan. See*
' *      also the reset method.                                       *
' *                                                                   *
' * clone: this method makes a clone of the scanner object. The clone *
' *      is guaranteed only to have the same source code that will    *
' *      tokenize to the same tokens and will contain the same white  *
' *      space patterns: the clone, when passed to the compareTo      *
' *      method, will return True.                                    *
' *                                                                   *
' *      The clone method implements the ICloneable interface.        *
' *                                                                   *
' * compareTo(o): this method compares the object instance with the   *
' *      scanner object, returning True when the source code in the   *
' *      instance is identical, after tokenization, to the object     *
' *      code. Note that the source code in the instance may have a   *
' *      different white space pattern from the source code in o.     * 
' *                                                                   *
' *      The compareTo method indirectly implements the ICompareable  *
' *      interface.                                                   * 
' *                                                                   *
' * dispose: this method disposes of the object, and it cleans up     *
' *      reference objects in the heap.  For best results use this    *
' *      method when you are finished using the object in code. This  *
' *      method will mark the object as unusable.                     *
' *                                                                   *
' *      The dispose(False) overlay will skip a scanner inspection    *
' *      that is conducted by default when the scanner is disposed.   *
' *                                                                   *
' * findRightParenthesis(i): at the scanner position in i, this method*
' *      searches for a balancing right parenthesis. On success this  *
' *      method returns the index of the token containing the right   *
' *      parenthesis. On failure this method returns one index past   *
' *      the last parenthesis with no other error indication.         *
' *                                                                   *
' *      The optional parameter intEndIndex can be used to restrict   *
' *      the search to all tokens up to and including the token at    *
' *      the specified end index.                                     *
' *                                                                   *
' *      On a successful return, the scanner will point to a right    *
' *      parenthesis token.                                           *
' *                                                                   *
' * findToken(i,t): this method searches the scanned tokens left to   *
' *      right starting at index i for a type or value of token and   *
' *      if it finds the the type or value, this method returns the   *
' *      scan index of the token. If findToken does not find the type *
' *      or value it returns zero.                                    * 
' *                                                                   *
' *      i should be an Integer, passed by value as the starting      *
' *      location.                                                    *
' *                                                                   *
' *      t should be an expected token type (passed as an enumerator, *
' *      of type ENUtokemType) or it should be an expected string     *
' *      value. This value will be compared ignoring case differences *
' *      with the value in the scanner.                               *
' *                                                                   *
' *      The optional parameter intEndIndex can be used to restrict   *
' *      the search to all tokens up to and including the token at    *
' *      the specified end index.                                     *
' *                                                                   *
' *      See also findTokenByTypeName.                                *
' *                                                                   *
' * findTokenByTypeName(i,t): this method searches the scanned tokens *
' *      left to right starting at index i for a token type and if it *
' *      finds the the type, this method returns the scan index of the*
' *      token. If the type cannot be found this method returns 0.    *
' *                                                                   *
' *      i should be an Integer, passed by value as the starting      *
' *      location.                                                    *
' *                                                                   *
' *      t should be the name of the expected token as one of the     *
' *      enumerator names defined by ENUtokenType, specified as a     *
' *      string.                                                      *
' *                                                                   *
' *      The optional parameter intEndIndex can be used to restrict   *
' *      the search to all tokens up to and including the token at    *
' *      the specified end index.                                     *
' *                                                                   *
' *      See also findToken.                                          *
' *                                                                   *
' * inspect(r): this method inspects the object. It checks for errors *
' *      that result from my stupid blunders in the creating the      *
' *      original source code of this class, your ham-fisted changes  *
' *      to the source code of this class, or "object abuse", the use *
' *      of this object after an error has occured. Inspect does not  *
' *      look for simple user errors: these are prevented elsewhere.  *
' *                                                                   *
' *      r should be a string, passed by reference; it is assigned an *
' *      inspection report.                                           *
' *                                                                   *
' *      The following inspection rules are used.                     *
' *                                                                   *
' *           *  The object instance must be usable                   *
' *                                                                   *
' *           *  Each token in both the array of scanned tokens, and  *
' *              the array of pending tokens, must pass the inspect   *
' *              procedure of qbToken                                 *
' *                                                                   *
' *           *  The tokens in the scanned array must be in ascending *
' *              order with gaps OK but no overlap                    *
' *                                                                   *
' *           *  No token's end index may point beyond the end of the *
' *              source code in either the scanned array, or the      *
' *              pending array                                        *
' *                                                                   *
' *           *  The line number must be greater than or equal to 0   *
' *                                                                   *
' *           *  The format of the line number index collection must  *
' *              be valid: this is a collection of three-item subcol- *
' *              lections. Item(1) must be a string containing the key*
' *              of the index entry: item(2) and item(3) must be      *
' *              positive integers: item(2) cannot be zero.           *
' *                                                                   *
' *           *  If the (nonnull) code is fully scanned, the first    *
' *              token's start index should be the same as the        *
' *              position of the first nonblank character in the      *
' *              source code. The last token's end index should be the*
' *              same as the position of the last nonblank character. *
' *                                                                   *
' *              If the code is null and indicated as fully scanned   *
' *              the scan count must be empty.                        *
' *                                                                   *
' *      If the inspection is failed the object becomes unusable.     *
' *                                                                   *
' *      An internal inspection is carried out in the constructor and *
' *      inside the dispose method.                                   *
' *                                                                   *
' * isInteger(s): this Shared method returns True when s is an        *
' *      unsigned integer in the SYNTACTICAL sense of containing no   *
' *      sign, no decimal part (including no zero decimal part as in  *
' *      1.0) and no exponent (including no meaningless exponent as in*
' *      .1e0.)                                                       *
' *                                                                   *
' * Line(L): this indexed, read-only property returns the source code *
' *      of line L.                                                   *
' *                                                                   *
' *      Note: use of this property forces the complete scan.         *
' *                                                                   *
' * LineCount: this read-only property returns the number of lines in *
' *      the source; note that continuation lines are counted as sep- *
' *      arate lines.                                                 *
' *                                                                   *
' *      Note: use of this property forces the complete scan.         *
' *                                                                   *
' * LineLength(L): this indexed, read-only property returns the       *
' *      length, from 1, of the line, in the source code, identified  *
' *      by the line number, from 1, in L.                            * 
' *                                                                   *
' *      Note: use of this property forces the complete scan.         *
' *                                                                   *
' * LineStartIndex(L): this indexed, read-only property returns the   *
' *      starting index, from 1, of the line, in the source code,     *
' *      identified by the line number, from 1, in L.                 * 
' *                                                                   *
' *      Note: use of this property forces the complete scan.         *
' *                                                                   *
' * mkUnusable: this method makes the object not usable.              *
' *                                                                   *
' * Name: this read-write property returns and can change the name of *
' *      the object instance. Name defaults to qbScannernnnn date     *
' *      time,  where nnnn is an object sequence number.              *
' *                                                                   *
' * normalize: this method returns the "normalized" form of the source*
' *      code in the object instance because it places one space      *
' *      between each token, and it removes all excess tokens at the  *
' *      start and end of the source code.                            *
' *                                                                   *
' *      This method makes no change to the source code. Instead, it  *
' *      returns the normalized code.                                 *
' *                                                                   *
' *      Normalization shouldn't be confused with prettyprinting or   *
' *      packing, although it may save some space.  The output of     *
' *      this method is not especially readable, and it actually      *
' *      inserts unneeded spaces between some tokens.                 *
' *                                                                   *
' * object2XML: this method converts the state of the object to an    *
' *      eXtended Markup Language string; note that the returned tag  *
' *      will include all source code and all parsed tokens and as    *
' *      such may be unmanageably large for large source code files.  *
' *                                                                   *
' *      The optional overload syntax object2XML(maxSource, maxTokens)*
' *      can define the maximum number of characters of source code,  *
' *      and, the maximum number of tokens, in the XML tag such that  *
' *      if either parameter is -1, there is no limit to the          *
' *      associated value.                                            *
' *                                                                   *
' *      In addition to the optional overload parameters, all         *
' *      overloads of this method support these optional parameters:  *
' *                                                                   *
' *           *  booAboutComment:=False will suppress a boxed XML     *
' *              comment at the start of the XML containing the value *
' *              of the About property of this object.                *
' *                                                                   *
' *           *  booStateComment:=False will suppress comments that   *
' *              describe each state value returned.                  *
' *                                                                   *
' * QBToken(i): this indexed, read-only property returns the indexed  *
' *      scanned token as an object, of type qbToken. Note that it    *
' *      will cause a scan, of tokens, up to and including token i,   *
' *      when the token is not available.                             *
' *                                                                   *
' * reset: this method resets the scan. Note that the reset method    *
' *      does not clear the source code; it merely undoes all parsing *
' *      done prior. See also the clear method.                       *
' *                                                                   *
' * scan: this method scans source code. It supports the following    *
' *      overloads.                                                   *
' *                                                                   *
' *      *  scan: resets the scanner object and scans all characters  *
' *         in the SourceCode.                                        *
' *                                                                   *
' *      *  scan(s): where s is source code, passed as a string, sets *
' *         SourceCode to s, resets, and scans all characters.        *
' *                                                                   *
' *      *  scan(end): end must be the Long precision end index for   *
' *         the scan (last character to be scanned from 1.)           *
' *                                                                   *
' *         The scanner is not reset. Instead, tokens are appended    *
' *         from the SourceCode starting at 1 or the end of the pre-  *
' *         vious scan until a token that ends at or after "end" is   *
' *         scanned...or the end of the SourceCode is scanned.        *
' *                                                                   *
' *         If your end value is not Long, use CLng(end) to convert   *
' *         it to the required type. To scan to the end using previous*
' *         results, use this code:                                   *
' *                                                                   *
' *              objQBscanner.scan(CLng(objQBscanner.EndIndex))       *
' *                                                                   *
' *      *  scan(start,end): start and end must be the Long precision *
' *         start and end index for the scan within the current       *
' *         source code (where characters are numbered from 1.)       *
' *                                                                   *
' *         The scanner is not reset. Instead, tokens are appended    *
' *         from the SourceCode starting at 1 or the end of the pre-  *
' *         vious scan until a token that ends at or after "end" is   *
' *         scanned...or the end of the SourceCode is scanned.        *
' *                                                                   *
' *         If your start or end values aren't Long, use CLng to      *
' *         convert them to the required type.                        *
' *                                                                   *
' *      *  scan(count): count must be the Integer precision maximum  *
' *         number of tokens to be scanned.                           *
' *                                                                   *
' *         The scanner is not reset. Instead, up to "count" tokens   *
' *         are appended from the identified substring of the source  *
' *         code.                                                     *
' *                                                                   *
' *         If your count value is not Integer, use CInt(count) to    *
' *                                                                   *
' * scanErrorEvent(msg,i,L,help): this event fires when the scanner   *
' *      cannot proceed owing to a user error. In this event:         *
' *                                                                   *
' *      *  msg is an error message                                   *
' *      *  i is the character index, at which the error was found    *
' *      *  L is the line number, at which the error was found        *
' *      *  help is additional help information                       *
' *                                                                   *
' * scanEvent(qbToken,index,length,tokens): this event fires at the   *
' *      completion of each successful scan of a token and it is      *
' *      useful for progress reporting. It passes the following to    *
' *      the delegate:                                                *
' *                                                                   *
' *      *  qbToken is the token object (of type qbToken)             *
' *                                                                   *
' *      *  index is the character index, from one, of the token      *
' *                                                                   *
' *      *  length is the length of the source code                   *
' *                                                                   *
' *      *  tokens contains the number of tokens found so far,        *
' *         including this token                                      *
' *                                                                   *
' * Scanned: this read-only property returns True when the source code*
' *      has been fully scanned, False otherwise.                     *
' *                                                                   *
' * SourceCode: this read-write property returns and may be set to    *
' *      the source code for scanning. Assigning source code clears   *
' *      the array of tokens in the object state, but does not result *
' *      in an immediate scan of the source code. Instead, scanning   *
' *      occurs when the QBToken property is called, and the token is *
' *      not available.                                               *
' *                                                                   *
' * sourceMid(start[,length]): this method returns source code start- *
' *      ing at the index in start and for a length of either one     *
' *      token (by default) or as specified by length.                *
' *                                                                   *
' *      Note that the start parameter must be the index of a token   *
' *      and not a character; likewise, the length must be the length *
' *      in tokens.                                                   *
' *                                                                   *
' * test(r): this method tests the scanner, and returns True when all *
' *      tests are passed or False when any test fails. The by-refer- *
' *      ence string parameter r is set on success or failure to a    *
' *      test report.                                                 *
' *                                                                   *
' *      Note that when the test fails, the object is marked unusable.*
' *                                                                   *
' * TestString: this Shared, read-only property returns the test      *
' *      string that is used in the test method. This string tests all*
' *      tokens for valid results.                                    *
' *                                                                   *
' * token(i): this method returns the value of the token indexed by   *
' *      i, where i is between 1 and TokenCount.                      *
' *                                                                   *
' * TokenCount: this read-only property returns the number of tokens. *
' *      Note that calling TokenCount causes a complete scan of the   *
' *      source code.                                                 *
' *                                                                   *
' * tokenEndIndex(i): this method returns the end index of the token  *
' *      at index i.                                                  *
' *                                                                   *
' * tokenLength(i): this method returns the length of the token at    *
' *      index i.                                                     *
' *                                                                   *
' * tokenLineNumber(i): this method returns the line number that      *
' *      contains the token at index i.                               *
' *                                                                   *
' * tokenStartIndex(i): this method returns the start index of the    *
' *      token at index i. See also startIndexToken which returns the *
' *      "inverse" function.                                          *
' *                                                                   *
' * tokenType(i): this method returns the type of the token at index  *
' *      (as an ENUtokenType)                                         *  
' *                                                                   *
' * tokenTypeAsString(i): this method returns the type of the token   *
' *      at index (as a string)                                       *  
' *                                                                   *
' * toString: this override method returns the tokens actually scanned*
' *      as a string in the format                                    *
' *                                                                   *
' *      <type1>@<startIdx1>..<endIdx1>:<lineNum1>:<source1><nl>      *
' *      <type2>@<startIdx2>..<endIdx2>:<lineNum2>:<source2><nl>      *
' *      .                                                            *
' *      .                                                            *
' *      .                                                            *
' *      <typen>@<startIdxn>..<endIdxn>:<lineNumn>:<sourcen><nl>      *
' *                                                                   *
' *      In the above:                                                *
' *                                                                   *
' *           *  <typei> is the ith token type as returned by the     *
' *              typeToString method of the qbToken class.            *
' *                                                                   *
' *           *  <startIdxi> is its start index                       *
' *                                                                   *
' *           *  <endIdx> is its end index (start index plus length   *
' *              minus one)                                           *
' *                                                                   *
' *           *  <lineNumi> is its line number. If the token is on    *
' *              more than one line, this is the number of the line   *
' *              where the token begins                               *
' *                                                                   *
' *           *  <sourcei> is its token value in the source code:     *
' *                                                                   *
' *              + If the token consists of graphical characters      *
' *                available on common USA keyboards, it will be a    *
' *                quoted string, with any internal quotes, doubled.  *
' *                                                                   *
' *              + If the token consists of other characters, it will *
' *                be in the form of a Visual Basic expression, con-  *
' *                sisting of quoted strings and calls to the ChrW    *
' *                function.                                          * 
' *                                                                   *
' *           *  Each token is separated from the next by new line    *
' *                                                                   *
' *      The optional overload toString(startIndex) returns tokens in *
' *      the above format commencing at startIndex. The optional over-*
' *      load toString(startIndex, count) returns "count" tokens      *
' *      commencing at startIndex.                                    *
' *                                                                   *
' * Usable: this read-only property returns True if the object is     *
' *      usable, False otherwise.                                     *
' *                                                                   *
' *                                                                   *
' * MULTITHREADING CONSIDERATIONS ----------------------------------- *
' *                                                                   *
' * Multiple, distinct instances of this object may run simultaneously*
' * in multiple threads; the same instance of this object may not run *
' * simultaneously in more than one thread.                           *
' *                                                                   *
' *                                                                   *
' *********************************************************************

Public Class qbScanner

    Implements ICloneable
    Implements IComparable

    ' ***** Utilities *****
    Private Shared _INTsequence As Integer
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJqbToken As qbToken.qbToken
    Private OBJcollectionUtilities As collectionUtilities.collectionUtilities

    ' ***** State *****
    Private Structure TYPstate
        Dim strName As String                     ' Object name
        Dim booUsable As Boolean                  ' Object usability
        Dim strSourceCode As String               ' All source
        Dim intLast As Integer                    ' Index of last token parsed
        Dim objQBtoken() As qbToken.qbToken       ' Some or all tokens
        Dim objTokenNext() As qbToken.qbToken     ' The pending tokens 
        Dim intLineNumber As Integer              ' Current line number
        Dim colLineIndex As Collection            ' Line index:
                                                  ' Key is _lineNumber
                                                  ' Data is subcollection:
                                                  ' Item(1): line number
                                                  ' Item(2): start index
                                                  ' Item(3): length
        Dim booScanned As Boolean                 ' True: indicates a completed scan
    End Structure
    Private USRstate As TYPstate

    ' ***** Constants *****
    ' --- Easter yEgg
    Private Const ABOUTINFO As String = _
        "qbScanner" & _
        vbNewLine & vbNewLine & _
        "The qbScanner class scans input source code and provides, on demand, " & _
        "scanned source tokens and lines of source code. This class uses " & _
        """lazy"" evaluation, scanning the " & _
        "source code only when necessary, and when an unparsed token is " & _
        "requested." & vbNewLine & vbNewLine & _
        "This class was developed commencing on 4/30/2003 by" & vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & _
        "http://members.screenz.com/edNilges"
    ' --- Inspection
    Private Const INSPECTION_USABLE As String = _
        "The object must be usable"
    Private Const INSPECTION_TOKENS As String = _
        "Each token in both the array of scanned tokens, and " & _
        "the array of pending tokens, must pass the inspect " & _
        "procedure of qbToken. " & _
        "The tokens in the scanned array " & _
        "must be in ascending order with gaps OK but no overlap. " & _
        "No token's end index (in either array) may point beyond the end of the " & _
        "source code in either the scanned array, or the " & _
        "pending array"
    Private Const INSPECTION_LINENUMBER As String = _
        "The line number must be greater than or equal to 0"
    Private Const INSPECTION_LINENUMBERINDEX As String = _
        "The format of the line number index collection must " & _
        "be valid: this is a collection of three-item subcollections. " & _
        "Item(1) must be a string containing the key " & _
        "of the index entry: item(2) and item(3) must be " & _
        "positive integers: item(2) can't be zero"
    Private Const INSPECTION_FULLSCAN As String = _
        "If the nonnull code is fully scanned, the first token's start " & _
        "index should be the same as the position of the first " & _
        "nonblank character in the source code. The last " & _
        "token's end index should be the same as the position " & _
        "of the last nonblank character." & _
        vbNewLine & vbNewLine & _
        "If the code is null and indicated as fully scanned " & _
        "the scan count must be zero."
    ' --- Block size of the token array
    Private Const TOKENARRAY_BLOCKSIZE As Integer = 1024
    ' --- Test
    Private Const TEST_COMPLEXSTRING1 As String = _
    """This string is """"fancy""""."""
    Private Const TEST_COMPLEXSTRING2 As String = _
    """This string is """"fancy"""". " & _
    "It """"contains"""" """"inner"""" """"""""""""strings"""""""""""""""
    Private Const TEST_SMART_STRING As String = _
    ChrW(SMARTQUOTELEFT) & "This string uses smart quotes" & ChrW(SMARTQUOTERIGHT)
    Private Const TEST_SMART_STRING_XML As String = _
    "&#08220This string uses smart quotes&#08221"
    Private Const TEST_STRING As String = _
    "' & : , identifier " & vbNewLine & _
    "+ - * / ( ) ; ""string"" 32767 -32767e-12 % ! # $ " & _
    TEST_COMPLEXSTRING1 & " " & _
    TEST_COMPLEXSTRING2 & " " & _
    ". .2 endId " & _
    TEST_SMART_STRING
    Private Const TEST_EXPECTED_RESULT As String = _
    "Apostrophe@1..1:1 '" & vbNewLine & _
    "Ampersand@3..3:1 &" & vbNewLine & _
    "Colon@5..5:1 :" & vbNewLine & _
    "Comma@7..7:1 ," & vbNewLine & _
    "Identifier@9..18:1 identifier" & vbNewLine & _
    "Newline@20..21:2 &#00013&#00010" & vbNewLine & _
    "Operator@22..22:2 +" & vbNewLine & _
    "Operator@24..24:2 -" & vbNewLine & _
    "Operator@26..26:2 *" & vbNewLine & _
    "Operator@28..28:2 /" & vbNewLine & _
    "Parenthesis@30..30:2 (" & vbNewLine & _
    "Parenthesis@32..32:2 )" & vbNewLine & _
    "Semicolon@34..34:2 ;" & vbNewLine & _
    "String@36..43:2 ""string""" & vbNewLine & _
    "UnsignedInteger@45..49:2 32767" & vbNewLine & _
    "Operator@51..51:2 -" & vbNewLine & _
    "UnsignedRealNumber@52..60:2 32767e-12" & vbNewLine & _
    "Percent@62..62:2 %" & vbNewLine & _
    "Exclamation@64..64:2 !" & vbNewLine & _
    "Pound@66..66:2 #" & vbNewLine & _
    "Currency@68..68:2 $" & vbNewLine & _
    "String@70..96:2 " & TEST_COMPLEXSTRING1 & vbNewLine & _
    "String@98..170:2 " & TEST_COMPLEXSTRING2 & vbNewLine & _
    "Period@172..172:2 ." & vbNewLine & _
    "UnsignedRealNumber@174..175:2 .2" & vbNewLine & _
    "Identifier@177..181:2 endId" & vbNewLine & _
    "String@183..213:2 " & TEST_SMART_STRING_XML
    ' --- Smart quotes are accepted in lieu of normal quotes
    Private Const SMARTQUOTELEFT As Integer = 8220
    Private Const SMARTQUOTERIGHT As Integer = 8221

    ' ***** Events *****
    Public Event scanEvent(ByVal objQBtoken As qbToken.qbToken, _
                           ByVal intCharacterIndex As Integer, _
                           ByVal intLength As Integer, _
                           ByVal intTokenCount As Integer)
    Public Event scanErrorEvent(ByVal strMsg As String, _
                                ByVal intIndex As Integer, _
                                ByVal intLineNumber As Integer, _
                                ByVal strHelp As String)

    ' ***** Object constructor ****************************************

    Public Sub New()
        With USRstate
            Interlocked.Increment(_INTsequence)
            .strName = "qbScanner" & _
                       _OBJutilities.alignRight(CStr(_INTsequence), 4, "0") & " " & _
                       Now
            Try
                ReDim .objQBtoken(TOKENARRAY_BLOCKSIZE)
                .objQBtoken(0) = New qbToken.qbToken
            Catch
                errorHandler_("Cannot allocate the token array: " & _
                              Err.Number & " " & Err.Description, _
                              "new", _
                              "Cannot construct object: object is not usable")
                Return
            End Try
            Try
                OBJcollectionUtilities = New collectionUtilities.collectionUtilities
            Catch
                errorHandler_("Cannot create collectionUtilities: " & _
                              Err.Number & " " & Err.Description, _
                              "new", _
                              "Cannot construct object: object is not usable")
                Return
            End Try
            If Not reset_() Then Return
            .booUsable = True
            inspection_()
        End With
    End Sub

    ' ***** Public procedures *****************************************
    ' *                                                               *
    ' * Also contains Private procedures that perform tasks on behalf *
    ' * of unique Public procedures.                                  *
    ' *                                                               *
    ' *****************************************************************

    ' -----------------------------------------------------------------
    ' Return about info
    '
    '
    Public Shared ReadOnly Property About() As String
        Get
            Return (ABOUTINFO)
        End Get
    End Property

    ' -----------------------------------------------------------------------
    ' Check for value or type of token and advance the pointer on success
    '
    '
    ' --- Check for token's value
    Public Overloads Function checkToken(ByRef intIndex As Integer, _
                                         ByVal strValueExpected As String, _
                                         Optional ByVal intEndIndex As Integer = 0) _
           As Boolean
        Return (checkToken_(intIndex, strValueExpected, intEndIndex))
    End Function
    ' --- Check for token's type
    Public Overloads Function checkToken(ByRef intIndex As Integer, _
                                         ByVal enuTypeExpected As qbTokenType.qbTokenType.ENUtokenType, _
                                         Optional ByVal intEndIndex As Integer = 0) _
           As Boolean
        Return (checkToken_(intIndex, enuTypeExpected, intEndIndex))
    End Function
    ' --- Common
    Private Overloads Function checkToken_(ByRef intIndex As Integer, _
                                           ByVal objExpected As Object, _
                                           ByVal intEndIndex As Integer) _
            As Boolean
        If Not checkUsable_("checkToken") Then Return (False)
        If intEndIndex < 0 Then
            errorHandler_("Invalid end index " & intEndIndex, _
                          "checkToken", _
                          "Returning false and not changing intIndex")
            Return (False)
        End If
        If intEndIndex <> 0 AndAlso intIndex > intEndIndex _
           OrElse _
           intEndIndex = 0 AndAlso intIndex > Me.TokenCount Then Return (False)
        With Me.QBToken(intIndex)
            If (TypeOf objExpected Is qbTokenType.qbTokenType.ENUtokenType) _
               AndAlso _
               .tokenTypeMatch(CType(objExpected, qbTokenType.qbTokenType.ENUtokenType)) _
               OrElse _
               UCase(.sourceCode(Me.SourceCode)) = UCase(CStr(objExpected)) Then
                intIndex += 1 : Return (True)
            End If
            Return (False)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Check token type by its name
    '
    '
    Public Overloads Function checkTokenByTypeName(ByRef intIndex As Integer, _
                                                   ByVal strTypeExpected As String, _
                                                   Optional ByVal intEndIndex As Integer = 0) _
           As Boolean
        Dim objTokenTools As qbToken.qbToken
        Return (Me.checkToken(intIndex, objTokenTools.typeToEnum(strTypeExpected), intEndIndex))
    End Function

    ' -----------------------------------------------------------------
    ' Clear the source code and reset the scan
    '
    '
    Public Function clear() As Boolean
        If Not checkUsable_("clear") Then Return (False)
        If Not Me.reset Then Return (False)
        Me.SourceCode = ""
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Clone the scanner
    '
    '
    ' --- Public interface
    Public Function clone() As qbScanner
        If Not checkUsable_("clone") Then Return (Nothing)
        Return (CType(clone_(), qbScanner))
    End Function
    ' --- Private interface
    Private Function clone_() As Object Implements ICloneable.Clone
        Dim objClone As qbScanner
        Try
            objClone = New qbScanner
        Catch
            errorHandler_("Cannot create clone: " & Err.Number & " " & Err.Description, _
                          "clone_", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        objClone.SourceCode = Me.SourceCode
        Return (objClone)
    End Function

    ' -----------------------------------------------------------------
    ' Compare the scanner to another scanner for normalized identity
    '
    '
    ' --- Public interface
    Public Function compareTo(ByVal objScanner As qbScanner) As Boolean
        If Not checkUsable_("compareTo") Then Return (False)
        Return (compareTo_(objScanner) <> 0)
    End Function
    ' --- Private implementation
    Private Function compareTo_(ByVal objScanner As Object) As Integer _
            Implements IComparable.CompareTo
        Return CInt(IIf(Me.normalize = CType(objScanner, qbScanner).normalize, 1, 0))
    End Function

    ' -----------------------------------------------------------------
    ' Disposer
    '
    '
    ' --- Inspects by default
    Public Function dispose() As Boolean
        Return (dispose(True))
    End Function
    ' --- Exposes option to inspect
    Public Function dispose(ByVal booInspect As Boolean) As Boolean
        If booInspect Then inspection_()
        Me.mkUnusable()
        Dim intIndex1 As Integer
        With USRstate
            For intIndex1 = 1 To .intLast
                If Not (.objQBtoken(intIndex1) Is Nothing) _
                   AndAlso _
                   Not tokenDispose_(.objQBtoken(intIndex1)) Then
                    errorHandler_("Cannot dispose of token " & intIndex1, _
                                  "dispose", _
                                  "Scanner object will reference unrecovered tokens")
                    Return (False)
                End If
            Next intIndex1
            For intIndex1 = 1 To UBound(.objTokenNext)
                If Not (.objTokenNext(intIndex1) Is Nothing) _
                   AndAlso _
                   Not tokenDispose_(.objTokenNext(intIndex1)) Then
                    errorHandler_("Cannot dispose of pending token " & intIndex1, _
                                  "dispose", _
                                  "Scanner object will reference unrecovered tokens")
                    Return (False)
                End If
            Next intIndex1
        End With
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Locates the right parenthesis
    '
    '
    Public Function findRightParenthesis(ByVal intIndex As Integer, _
                                         Optional ByVal intEndIndex As Integer = 0) _
           As Integer
        If intEndIndex < 0 Then
            errorHandler_("Invalid end index " & intEndIndex, _
                          "findRightParenthesis", _
                          "Returning end of source code plus one")
            Return (Me.TokenCount + 1)
        End If
        With USRstate
            Dim intEndIndexWork As Integer = intEndIndex
            If intEndIndexWork = 0 Then
                intEndIndexWork = UBound(.objQBtoken)
            End If
            Dim intIndex1 As Integer
            Dim intLevel As Integer = 1
            For intIndex1 = intIndex To intEndIndexWork
                With .objQBtoken(intIndex1)
                    If .TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeParenthesis Then
                        If .sourceCode(Me.SourceCode) = "(" Then
                            intLevel += 1
                        Else
                            intLevel -= 1
                            If intLevel = 0 Then
                                Return (intIndex1)
                            End If
                        End If
                    End If
                End With
            Next intIndex1
            Return (UBound(.objQBtoken) + 1)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Search for value or type of token and return its index or 0
    '
    '
    ' --- Search for token's value
    Public Overloads Function findToken(ByVal intIndex As Integer, _
                                        ByVal strValueExpected As String, _
                                        Optional ByVal intEndIndex As Integer = 0) _
           As Integer
        Return (findToken_(intIndex, strValueExpected, intEndIndex))
    End Function
    ' --- Search for token's type
    Public Overloads Function findToken(ByVal intIndex As Integer, _
                                        ByVal enuTypeExpected As qbTokenType.qbTokenType.ENUtokenType, _
                                        Optional ByVal intEndIndex As Integer = 0) _
           As Integer
        Return (findToken_(intIndex, enuTypeExpected, intEndIndex))
    End Function
    ' --- Common
    Private Overloads Function findToken_(ByVal intIndex As Integer, _
                                          ByVal objExpected As Object, _
                                          ByVal intEndIndex As Integer) _
            As Integer
        If Not checkUsable_("findToken") Then Return (0)
        If intIndex < 0 Then
            errorHandler_("Invalid index " & intEndIndex, _
                          "findToken", _
                          "Returning 0")
            Return (0)
        End If
        If intEndIndex < 0 Then
            errorHandler_("Invalid end index " & intEndIndex, _
                          "findToken", _
                          "Returning 0")
            Return (0)
        End If
        Dim intIndex1 As Integer = intIndex
        With Me
            Do While intIndex1 <= .TokenCount
                If (TypeOf objExpected Is String) _
                   AndAlso _
                   checkToken(intIndex1, CStr(objExpected), intEndIndex:=intEndIndex) _
                   OrElse _
                   checkToken(intIndex1, _
                              CType(objExpected, _
                                    qbTokenType.qbTokenType.ENUtokenType), _
                              intEndIndex:=intEndIndex) Then
                    Return (intIndex1 - 1)
                End If
                intIndex1 += 1
            Loop
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Find token type by its name
    '
    '
    Public Overloads Function findTokenByTypeName(ByVal intIndex As Integer, _
                                                  ByVal strTypeExpected As String, _
                                                  Optional ByVal intEndIndex As Integer = 0) _
           As Integer
        Dim objTokenTools As qbToken.qbToken
        Return (Me.findToken(intIndex, objTokenTools.typeToEnum(strTypeExpected), intEndIndex))
    End Function

    ' -----------------------------------------------------------------
    ' Object inspection
    '
    '
    Public Function inspect(ByRef strReport As String) As Boolean
        Dim booInspection As Boolean = True
        strReport = "Inspection of " & _
                    _OBJutilities.enquote(Me.Name) & " " & _
                    "at " & Now
        With USRstate
            If Not _OBJutilities.inspectionAppend(strReport, _
                                                  INSPECTION_USABLE, _
                                                  .booUsable, _
                                                  booInspection) Then Return (False)
            Dim strComment As String
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_TOKENS, _
                                           inspect_scanned_(.objQBtoken, _
                                                            strComment, _
                                                            True, _
                                                            .intLast, _
                                                            .strSourceCode) _
                                           AndAlso _
                                           inspect_scanned_(.objTokenNext, _
                                                            strComment, _
                                                            False, _
                                                            UBound(.objTokenNext), _
                                                            .strSourceCode), _
                                           booInspection, _
                                           strComment)
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_LINENUMBER, _
                                           .intLineNumber >= 0, _
                                           booInspection)
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_LINENUMBERINDEX, _
                                           inspect_lineNumberIndex_(.colLineIndex, strComment), _
                                           booInspection, _
                                           strComment)
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_FULLSCAN, _
                                           Not Me.Scanned _
                                           OrElse _
                                           Me.TokenCount = 0 _
                                           AndAlso _
                                           Trim(.strSourceCode) = "" _
                                           OrElse _
                                           Me.QBToken(1).StartIndex _
                                           = _
                                           _OBJutilities.verify(.strSourceCode, " ") _
                                           AndAlso _
                                           Me.QBToken(Me.TokenCount).EndIndex _
                                           = _
                                           inspect_findEnd_(.strSourceCode), _
                                           booInspection, _
                                           strComment)
            If Not booInspection Then Me.mkUnusable()
            Return (booInspection)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Find the nonblank end of the string on behalf of inspect
    '
    '
    Private Function inspect_findEnd_(ByVal strInstring As String) As Integer
        Dim intIndex1 As Integer
        For intIndex1 = Len(strInstring) To 1 Step -1
            If Mid(strInstring, intIndex1, 1) <> " " Then Return (intIndex1)
        Next intIndex1
        Return (0)
    End Function

    ' -----------------------------------------------------------------
    ' Inspect the line number index
    '
    '
    Private Function inspect_lineNumberIndex_(ByVal colLineIndex As Collection, _
                                              ByRef strComment As String) As Boolean
        Dim colHandle As Collection
        Dim intIndex1 As Integer
        Dim intLength As Integer
        Dim intStartIndex As Integer
        Dim strKey As String
        With colLineIndex
            For intIndex1 = 1 To .Count
                Try
                    colHandle = CType(.Item(intIndex1), Collection)
                Catch
                    strComment = "Item " & intIndex1 & " is not a collection"
                    Return (False)
                End Try
                With colHandle
                    Try
                        strKey = CStr(.Item(1))
                        intStartIndex = CInt(.Item(2))
                        intLength = CInt(.Item(3))
                    Catch
                        strComment = "Item " & intIndex1 & " " & _
                                     "does not contain a expected data types"
                        Return (False)
                    End Try
                    If key2String_(strKey) <> CStr(intIndex1) Then
                        strComment = "Item " & intIndex1 & " does not contain " & _
                                     "the expected key"
                        Return (False)
                    End If
                    If intStartIndex < 1 OrElse intLength < 0 Then
                        strComment = "Item " & intIndex1 & " does not contain " & _
                                     "a valid start index and length"
                        Return (False)
                    End If
                End With
            Next intIndex1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Inspect an array of scanned tokens, on behalf of inspect
    '
    '
    Private Function inspect_scanned_(ByVal objScanned() As qbToken.qbToken, _
                                      ByRef strComment As String, _
                                      ByVal booCheckSequencing As Boolean, _
                                      ByVal intLast As Integer, _
                                      ByVal strSourceCode As String) As Boolean
        Dim intIndex1 As Integer
        Dim intPreviousEndIndex As Integer
        For intIndex1 = 1 To intLast
            With objScanned(intIndex1)
                If Not .inspect(strComment) Then
                    strComment = _OBJutilities.string2Box(strComment, _
                                                          "Token " & intIndex1 & _
                                                          " has failed its own inspection")
                    Return (False)
                End If
                If booCheckSequencing AndAlso .StartIndex <= intPreviousEndIndex Then
                    strComment = "Token " & intIndex1 & " " & _
                                 "overlaps the end of the previous token. " & _
                                 "This token starts at " & _
                                 .StartIndex & ": " & _
                                 "the previous token ends at " & _
                                 intPreviousEndIndex
                    Return (False)
                End If
                If booCheckSequencing AndAlso .StartIndex + .Length - 1 > Len(strSourceCode) Then
                    strComment = "Token " & intIndex1 & " " & _
                                 "is outside the source code. " & _
                                 "This token starts at " & _
                                 .StartIndex & ": " & _
                                 "the length of the token is " & _
                                 .Length & ": " & _
                                 "the length of the source code is " & _
                                 Len(strSourceCode)
                    Return (False)
                End If
                intPreviousEndIndex = .EndIndex
            End With
        Next intIndex1
        strComment = ""
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Return True when string is syntactically an unsigned integer
    '
    '
    Public Shared Function isInteger(ByVal strInstring As String) As Boolean
        Dim objScanner As qbScanner
        Try
            objScanner = New qbScanner
            With objScanner
                .SourceCode = strInstring : .scan()
                Return (.TokenCount = 1 _
                       AndAlso _
                       .QBToken(1).TokenType _
                       = _
                       qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger)
            End With
        Catch
            errorHandler_("Cannot create a scanner or scan string: " & _
                          Err.Number & " " & Err.Description, _
                          "qbScanner", "isInteger", _
                          "Returning False")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Return text in indexed line
    '
    '
    Public ReadOnly Property Line(ByVal intLine As Integer) As String
        Get
            If Not checkUsable_("Line get") Then Return ("")
            If Not scanner_() Then Return ("")
            Try
                Dim colHandle As Collection = _
                    CType(USRstate.colLineIndex.Item(string2Key_(CStr(intLine))), Collection)
                With colHandle
                    Return (Mid(USRstate.strSourceCode, CInt(.Item(2)), CInt(.Item(3))))
                End With
            Catch
                errorHandler_("Invalid line number " & intLine, _
                              "LineLength get", _
                              "Returning 0")
                Return ("")
            End Try
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return count of lines in source
    '
    '
    Public ReadOnly Property LineCount() As Integer
        Get
            If Not checkUsable_("LineCount get") Then Return (0)
            If Not scanner_() Then Return (0)
            Return (USRstate.colLineIndex.Count)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return length of indexed line
    '
    '
    Public ReadOnly Property LineLength(ByVal intLine As Integer) As Integer
        Get
            If Not checkUsable_("LineLength get") Then Return (0)
            If Not scanner_() Then Return (0)
            Try
                Dim colHandle As Collection = _
                    CType(USRstate.colLineIndex.Item(string2Key_(CStr(intLine))), Collection)
                Return (CInt(colHandle.Item(3)))
            Catch
                errorHandler_("Invalid line number " & intLine, _
                              "LineLength get", _
                              "Returning 0")
                Return (0)
            End Try
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return start index of indexed line
    '
    '
    Public ReadOnly Property LineStartIndex(ByVal intLine As Integer) As Integer
        Get
            If Not checkUsable_("LineStartIndex get") Then Return (0)
            If Not scanner_() Then Return (0)
            Try
                Dim colHandle As Collection = _
                    CType(USRstate.colLineIndex.Item(string2Key_(CStr(intLine))), Collection)
                Return (CInt(colHandle.Item(2)))
            Catch
                errorHandler_("Invalid line number " & intLine, _
                              "LineStartIndex get", _
                              "Returning 0")
                Return (0)
            End Try
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Make the object not usable
    '
    '
    Public Sub mkUnusable()
        USRstate.booUsable = False
    End Sub

    ' -----------------------------------------------------------------
    ' Assigns and retrieves the name of the object
    '
    '
    Public Property Name() As String
        Get
            Return (USRstate.strName)
        End Get
        Set(ByVal strNewValue As String)
            If Not checkUsable_("Name set") Then Return
            USRstate.strName = strNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------------
    ' Normalize the source code 
    '
    '
    Public Function normalize() As String
        If Not checkUsable_("Normalize") Then Return ("")
        With Me
            If Not .Scanned Then
                If Not .scan Then Return ("")
            End If
            Dim intIndex1 As Integer
            Dim objStringBuilder As System.Text.StringBuilder
            Try
                objStringBuilder = New System.Text.StringBuilder
            Catch
                errorHandler_("Cannot create String Builder: " & _
                              Err.Number & " " & Err.Description, _
                              "normalize", _
                              "Returning null string: " & _
                              "no change made to source code")
                Return ("")
            End Try
            For intIndex1 = 1 To .TokenCount
                _OBJutilities.append(objStringBuilder, _
                                     " ", _
                                     .token(intIndex1))
            Next intIndex1
            Return (objStringBuilder.ToString)
        End With
    End Function

    ' -----------------------------------------------------------------------
    ' Convert the object to XML
    '
    '
    Public Overloads Function object2XML(Optional ByVal booAboutComment As Boolean = True, _
                                         Optional ByVal booStateComment As Boolean = True) As String
        Return (object2XML(-1, -1, _
                          booStateComment:=booStateComment, _
                          booAboutComment:=booAboutComment))
    End Function
    Public Overloads Function object2XML(ByVal intSourceTruncation As Integer, _
                                         Optional ByVal booAboutComment As Boolean = True, _
                                         Optional ByVal booStateComment As Boolean = True) As String
        Return (object2XML(intSourceTruncation, -1, _
                          booStateComment:=booStateComment, _
                          booAboutComment:=booAboutComment))
    End Function
    Public Overloads Function object2XML(ByVal intSourceTruncation As Integer, _
                                         ByVal intTokenTruncation As Integer, _
                                         Optional ByVal booAboutComment As Boolean = True, _
                                         Optional ByVal booStateComment As Boolean = True) As String
        Dim strSourceCode As String
        Dim strSourceTruncation As String
        Dim strTokensToString As String
        Dim strTokenTruncation As String
        If intSourceTruncation < -1 Then
            errorHandler_("Invalid source truncation " & intSourceTruncation, _
                            "object2XML", _
                            "Returning a null tag")
            Return ("")
        ElseIf intSourceTruncation <> -1 Then
            strSourceCode = _OBJutilities.ellipsis(Me.SourceCode, intSourceTruncation)
            strSourceTruncation = " (truncated to " & intSourceTruncation & " characters"
        Else
            strSourceCode = Me.SourceCode
        End If
        If intTokenTruncation < -1 Then
            errorHandler_("Invalid token truncation " & intTokenTruncation, _
                            "object2XML", _
                            "Returning a null tag")
            Return ("")
        ElseIf intTokenTruncation <> -1 Then
            strTokensToString = Me.ToString(1, intTokenTruncation)
            strTokenTruncation = " (truncated to " & intTokenTruncation & " token(s)"
        Else
            strTokensToString = Me.ToString
        End If
        With Me
            Return _OBJutilities.objectInfo2XML("qbScanner", _
                                                .About, _
                                                booAboutComment, _
                                                booStateComment, _
                                                "Name", _
                                                "Object instance name", _
                                                .Name, _
                                                "Usable", _
                                                "True: object is usable", _
                                                CStr(.Usable), _
                                                "SourceCode", _
                                                "Source code" & strSourceTruncation, _
                                                _OBJutilities.soft2HardParagraph _
                                                (_OBJutilities.string2Display(strSourceCode), _
                                                 bytLineWidth:=50), _
                                                "Last", _
                                                "Last token array entry in use", _
                                                CStr(USRstate.intLast), _
                                                "Tokens", _
                                                "Tokens parsed" & strTokenTruncation, _
                                                strTokensToString, _
                                                "PendingTokens", _
                                                "Pending tokens available for parse", _
                                                tokenArrayToString_(USRstate.objTokenNext, _
                                                                    1, _
                                                                    UBound(USRstate.objTokenNext)), _
                                                "TokenArraySize", _
                                                "Available size of token array", _
                                                CStr(UBound(USRstate.objQBtoken)), _
                                                "LineNumber", _
                                                "Current line number", _
                                                CStr(USRstate.intLineNumber), _
                                                "LineIndex", _
                                                "Identifies line number, start index, length of each line", _
                                                OBJcollectionUtilities.collection2String _
                                                (USRstate.colLineIndex, booReadable:=True))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return the indexed token
    '
    '
    Public ReadOnly Property QBToken(ByVal intIndex As Integer) As qbToken.qbToken
        Get
            If Not checkUsable_("QBToken get") Then Return (Nothing)
            If intIndex < 1 Then
                errorHandler_("Invalid token index " & intIndex & " " & _
                              "is less than one", _
                              "QBToken get", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            If Not scanner_(CLng(-1), intIndex) Then Return (Nothing)
            If Me.TokenCount < intIndex Then
                errorHandler_("Invalid token index " & intIndex & " " & _
                              "is beyond the end of the source code", _
                              "QBToken get", _
                              "Returning Nothing")
                Return (Nothing)
            End If
            Return (USRstate.objQBtoken(intIndex))
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Reset the scanner
    '
    '
    Public Function reset() As Boolean
        If Not checkUsable_("reset") Then Return (False)
        Return (reset_())
    End Function

    ' -----------------------------------------------------------------
    ' Do a comprehensive scan
    '
    '
    ' --- Use existing SourceCode: reset and scan everything
    Public Overloads Function scan() As Boolean
        If Not checkUsable_("scan") Then Return (False)
        If Not Me.reset Then Return (False)
        Return (scanner_())
    End Function
    ' --- Change existing SourceCode: reset and scan everything
    Public Overloads Function scan(ByVal strSourceCode As String) As Boolean
        If Not checkUsable_("scan") Then Return (False)
        Me.SourceCode = strSourceCode
        If Not Me.reset Then Return (False)
        Return (scanner_(-1, -1))
    End Function
    ' --- Use existing SourceCode: scan substring from 1 with no reset
    Public Overloads Function scan(ByVal lngEndIndex As Long) As Boolean
        If Not checkUsable_("scan") Then Return (False)
        Dim intEndIndex As Integer
        Try
            intEndIndex = CInt(lngEndIndex)
        Catch
            errorHandler_("lngEndIndex must be in positive Integer range", _
                          "scan", _
                          "Returning False")
        End Try
        Return (scanner_(intEndIndex, CInt(-1)))
    End Function
    ' --- Use existing SourceCode: scan substring from 1 with no reset
    Public Overloads Function scan(ByVal lngStartIndex As Long, _
                                   ByVal lngEndIndex As Long) As Boolean
        If Not checkUsable_("scan") Then Return (False)
        Dim intStartIndex As Integer
        Dim intEndIndex As Integer
        Try
            intStartIndex = CInt(lngStartIndex)
            intEndIndex = CInt(lngEndIndex)
        Catch
            errorHandler_("lngStartIndex and lngEndIndex must be in positive Integer range", _
                          "scan", _
                          "Returning False")
        End Try
        Return (scanner_(intEndIndex, CInt(-1), intStartingIndex:=intStartIndex))
    End Function
    ' --- Use existing SourceCode: scan up to Count with no reset
    Public Overloads Function scan(ByVal intCount As Integer) As Boolean
        If Not checkUsable_("scan") Then Return (False)
        Return (scanner_(-1, intCount))
    End Function

    ' -----------------------------------------------------------------
    ' Return True if source code has been completely scanned
    '
    '
    Public ReadOnly Property Scanned() As Boolean
        Get
            If Not checkUsable_("Scanned Get") Then Return (False)
            Return (USRstate.booScanned)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Assign and retrieve the source code
    '
    '
    Public Property SourceCode() As String
        Get
            If Not checkUsable_("SourceCode get") Then Return ("")
            Return (USRstate.strSourceCode)
        End Get
        Set(ByVal strNewValue As String)
            If Not checkUsable_("SourceCode set") Then Return
            If Not reset_() Then Return
            USRstate.strSourceCode = strNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return source code substrings
    '
    '
    ' --- Return one token
    Public Overloads Function sourceMid(ByVal intStartIndex As Integer) As String
        Return (Me.sourceMid(intStartIndex, 1))
    End Function
    ' --- Return one or more tokens    
    Public Overloads Function sourceMid(ByVal intStartIndex As Integer, _
                                        ByVal intLength As Integer) As String
        With Me
            If Not .checkUsable_("sourceMid") Then Return ""
            If intStartIndex > .TokenCount Then Return ""
            If intStartIndex < 1 OrElse intLength < 0 Then
                errorHandler_("Invalid start index and/or length: ", _
                              "sourceMid", _
                              "Returning a nul string")
                Return ""
            End If
            Return (Mid(.SourceCode, _
                        .QBToken(intStartIndex).StartIndex, _
                        .QBToken(intStartIndex + intLength - 1).EndIndex + 1 _
                        - _
                        .QBToken(intStartIndex).StartIndex))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Test the object
    '
    '
    Public Function test(ByRef strReport As String) As Boolean
        If Not test_(strReport) Then
            Me.mkUnusable() : Return (False)
        End If
        Return (True)
    End Function
    Private Function test_(ByRef strReport As String) As Boolean
        strReport = "Testing qbScanner " & Me.Name & " at " & Now & _
                    vbNewLine & vbNewLine
        Dim objScannerTest As qbScanner
        Try
            objScannerTest = New qbScanner
        Catch
            Dim strMessage As String = _
                "Could not create test scanner: " & _
                Err.Number & " " & Err.Description
            strReport &= strMessage
            errorHandler_(strMessage, "test", "Returning test failure")
            Return (False)
        End Try
        With objScannerTest
            Try
                .SourceCode = .TestString
                strReport &= "The test string is: " & _
                            _OBJutilities.string2Display(.SourceCode) & _
                            vbNewLine & vbNewLine
                If Not .scan Then
                    Dim strMessage As String = "Could not scan"
                    strReport &= strMessage
                    errorHandler_(strMessage, _
                                  "test", _
                                  "Returning test failure")
                    Return (False)
                End If
                Dim strActual As String = .ToString
                Dim booOK As Boolean = (TEST_EXPECTED_RESULT = strActual)
                strReport &= _OBJutilities.string2Box(TEST_EXPECTED_RESULT, _
                                                      "EXPECTED RESULTS") & _
                             vbNewLine & vbNewLine & _
                             _OBJutilities.string2Box(strActual, _
                                                      "ACTUAL RESULTS")
                If Not booOK Then
                    strReport &= vbNewLine & vbNewLine & _
                                 "Expected and actual results are not the same"
                    Dim intIndex1 As Integer
                    Dim strSplit1() As String
                    Dim strSplit2() As String
                    Try
                        strSplit1 = Split(TEST_EXPECTED_RESULT, vbNewLine)
                        strSplit2 = Split(strActual, vbNewLine)
                    Catch
                        errorHandler_("Cannot split: " & Err.Number & " " & Err.Description, _
                                      "test_", _
                                      "Returning False")
                    End Try
                    If UBound(strSplit1) <> UBound(strSplit2) Then
                        strReport &= vbNewLine & vbNewLine & _
                                     "Number of tokens in expected result is " & _
                                     UBound(strSplit1) & ": " & _
                                     "number of tokens in actual result is " & _
                                     UBound(strSplit2)
                    Else
                        For intIndex1 = 0 To UBound(strSplit1)
                            If strSplit1(intIndex1) <> strSplit2(intIndex1) Then
                                strReport &= vbNewLine & vbNewLine & _
                                             "At test case " & intIndex1 & ", " & _
                                             "the expected result is " & _
                                             strSplit1(intIndex1) & ": " & _
                                             "the actual result is " & _
                                             strSplit2(intIndex1)
                            End If
                        Next intIndex1
                    End If
                End If
                strReport &= vbNewLine & vbNewLine & _
                             "Testing the clone method with the compareTo method"
                Dim objClone As qbScanner
                Try
                    objClone = .clone
                    booOK = booOK AndAlso .compareTo(objClone)
                Catch
                    errorHandler_("Error in clone or compareTo methods: " & _
                                  Err.Number & " " & Err.Description, _
                                  "test_", _
                                  "Returning False")
                    Return (False)
                End Try
                strReport &= vbNewLine & vbNewLine & _
                             "Testing startIndexToken"
                Return (booOK)
            Catch
                Dim strMessage As String = _
                    "Scanner procedure failed: " & _
                    Err.Number & " " & Err.Description
                strReport &= strMessage
                errorHandler_(strMessage, "test", "Returning test failure")
                Return (False)
            End Try
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return the number of tokens...forces complete parse
    '
    '
    Public Shared ReadOnly Property TestString() As String
        Get
            Return (TEST_STRING)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return indexed token value
    '
    '
    Public Function token(ByVal intIndex1 As Integer) As String
        If Not checkUsable_("token") Then Return ("")
        With Me
            If intIndex1 < 1 OrElse intIndex1 > .TokenCount Then
                errorHandler_("Invalid token index " & intIndex1, _
                              "token", _
                              "This value must be between 1 and " & _
                              .TokenCount & _
                              vbNewLine & vbNewLine & _
                              "Returning a null string")
                Return ("")
            End If
            Return .QBToken(intIndex1).sourceCode(.SourceCode)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return the number of tokens...forces complete parse
    '
    '
    Public ReadOnly Property TokenCount() As Integer
        Get
            If Not checkUsable_("TokenCount get") Then Return (0)
            If Not Scanned AndAlso Not scanner_() Then Return (0)
            Return (USRstate.intLast)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return the token's end index
    '
    '
    Public Function tokenEndIndex(ByVal intIndex As Integer) As Integer
        If Not checkUsable_("tokenEndIndex") Then Return (0)
        If Not checkTokenIndex_(intIndex) Then Return (0)
        Return (Me.QBToken(intIndex).EndIndex)
    End Function

    ' -----------------------------------------------------------------
    ' Return the token's length
    '
    '
    Public Function tokenLength(ByVal intIndex As Integer) As Integer
        If Not checkUsable_("tokenLength") Then Return (0)
        If Not checkTokenIndex_(intIndex) Then Return (0)
        Return (Me.QBToken(intIndex).Length)
    End Function

    ' -----------------------------------------------------------------
    ' Return the token's line number
    '
    '
    Public Function tokenLineNumber(ByVal intIndex As Integer) As Integer
        If Not checkUsable_("tokenLineNumber") Then Return (0)
        If Not checkTokenIndex_(intIndex) Then Return (0)
        Return (Me.QBToken(intIndex).LineNumber)
    End Function

    ' -----------------------------------------------------------------
    ' Return the token's start index
    '
    '
    Public Function tokenStartIndex(ByVal intIndex As Integer) As Integer
        If Not checkUsable_("tokenStartIndex") Then Return (0)
        If Not checkTokenIndex_(intIndex) Then Return (0)
        Return (Me.QBToken(intIndex).StartIndex)
    End Function

    ' -----------------------------------------------------------------
    ' Return the token's type (as an enumerator)
    '
    '
    Public Function tokenType(ByVal intIndex As Integer) _
           As qbTokenType.qbTokenType.ENUtokenType
        If Not checkUsable_("tokenType") Then
            Return (qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid)
        End If
        If Not checkTokenIndex_(intIndex) Then
            Return (qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid)
        End If
        Return (Me.QBToken(intIndex).TokenType)
    End Function

    ' -----------------------------------------------------------------
    ' Return the token's type (as a string)
    '
    '
    Public Function tokenTypeAsString(ByVal intIndex As Integer) As String
        If Not checkUsable_("tokenType") Then Return ("")
        If Not checkTokenIndex_(intIndex) Then Return ("")
        Return (Me.QBToken(intIndex).TokenType.ToString)
    End Function

    ' -----------------------------------------------------------------
    ' Return token array
    '
    '
    Public Overloads Overrides Function toString() As String
        Return (Me.ToString(1, Me.TokenCount))
    End Function
    Public Overloads Function toString(ByVal intStartIndex As Integer) As String
        Return (Me.ToString(intStartIndex, Me.TokenCount))
    End Function
    Public Overloads Function toString(ByVal intStartIndex As Integer, _
                                       ByVal intCount As Integer) As String
        If Not checkUsable_("toString") Then Return ("")
        Dim intStartIndexWork As Integer
        If intStartIndex >= 1 Then
            intStartIndexWork = intStartIndex
        Else
            errorHandler_("Invalid start index " & intStartIndex, _
                          "toString", _
                          "Will start at position 1")
            intStartIndexWork = 1
        End If
        Dim intCountWork As Integer
        If intCount >= 0 Then
            intCountWork = intCount
        Else
            errorHandler_("Invalid count " & intCount, _
                          "toString", _
                          "Will return tokens to end")
            intCountWork = Me.TokenCount
        End If
        Return (tokenArrayToString_(USRstate.objQBtoken, intStartIndexWork, intCountWork))
    End Function

    ' -----------------------------------------------------------------
    ' Tell the caller if the object is usable
    '
    '
    Public ReadOnly Property Usable() As Boolean
        Get
            Return (USRstate.booUsable)
        End Get
    End Property

    ' ***** Friend procedures *****************************************

    ' -----------------------------------------------------------------
    ' Adds token to scanner
    '
    '
    Friend Function addToken(ByVal objToken As qbToken.qbToken) As Boolean
        If Not checkUsable_("addToken") Then Return (False)
        Return (addToken_(objToken))
    End Function

    ' ***** Private procedures **************************************** 

    ' -----------------------------------------------------------------
    ' Expand the token array
    '
    '
    Private Function addToken_(ByVal objQBtoken As qbToken.qbToken) As Boolean
        With USRstate
            .intLast += 1
            If .intLast > UBound(.objQBtoken) Then
                Try
                    ReDim Preserve .objQBtoken(UBound(.objQBtoken) + TOKENARRAY_BLOCKSIZE)
                Catch
                    errorHandler_("Cannot expand the token table: " & _
                                    Err.Number & " " & Err.Description, _
                                    "scanner_", _
                                    "Scanner is not usable")
                    Me.mkUnusable() : Return (False)
                End Try
            End If
            If (Not .objQBtoken(.intLast) Is Nothing) _
                AndAlso _
                Not tokenDispose_(.objQBtoken(.intLast)) Then
                errorHandler_("Cannot dispose token for recycling", _
                                "scanner_", _
                                "Scanner is not usable")
                Me.mkUnusable() : Return (False)
            End If
            .objQBtoken(.intLast) = objQBtoken
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Check token index
    '
    '
    Private Function checkTokenIndex_(ByVal intIndex As Integer) As Boolean
        If intIndex < 1 OrElse intIndex > Me.TokenCount Then
            errorHandler_("Token index " & intIndex & " " & _
                          "is less than one or greater than TokenCount " & _
                          Me.TokenCount, _
                          "checkTokenIndex_", _
                          "Returning False")
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Check object usability
    '
    '
    Private Function checkUsable_(ByVal strProcedure As String) As Boolean
        If Not USRstate.booUsable Then
            errorHandler_("Object is not usable", _
                          strProcedure, _
                          "Returning a default value to caller, " & _
                          "and making no change to object state")
            Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Dispose of the line index and character-to-line index
    '
    '
    Private Function disposeLineIndexes_() As Boolean
        With USRstate
            If Not (.colLineIndex Is Nothing) _
               AndAlso _
               Not OBJcollectionUtilities.collectionClear(.colLineIndex) Then
                errorHandler_("Cannot dispose old line indexes", _
                                "disposeLineIndexes_", _
                                "Object is unusable")
                Me.mkUnusable()
                Return (False)
            End If
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Interfaces to the error handler
    '
    '
    ' --- Unshared
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String)
        errorHandler_(strMessage, _
                      Me.Name, _
                      strProcedure, _
                      strHelp)
    End Sub
    ' --- Shared
    Private Overloads Shared Sub errorHandler_(ByVal strMessage As String, _
                                               ByVal strName As String, _
                                               ByVal strProcedure As String, _
                                               ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, _
                                   strName, _
                                   strProcedure, _
                                   strHelp)
    End Sub

    ' -----------------------------------------------------------------
    ' Set up the line index collection for the scanner
    '
    '
    Public Function initializeLineIndex_() As Boolean
        With USRstate
            If Not disposeLineIndexes_() Then Return (False)
            Try
                .colLineIndex = New Collection
            Catch
                errorHandler_("Cannot set up line index collection: " & _
                              Err.Number & " " & Err.Description, _
                              "initializePendingTokens_", _
                              "Object will not be usable")
                .booUsable = False : Return (False)
            End Try
            Return (updateLineIndex_(.colLineIndex, 1, 1, .strSourceCode))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Set up the pending token array for the scanner
    '
    '
    Public Function initializePendingTokens_() As Boolean
        With USRstate
            Dim intIndex1 As Integer
            Dim intUBound As Integer = -1
            Try
                intUBound = UBound(.objTokenNext)
            Catch : End Try
            If intUBound = -1 Then
                Try
                    ReDim .objTokenNext(_OBJqbToken.TypeCountActual)
                    For intIndex1 = 1 To UBound(.objTokenNext)
                        .objTokenNext(intIndex1) = New qbToken.qbToken
                    Next intIndex1
                Catch
                    errorHandler_("Cannot set up pending token array: " & _
                                Err.Number & " " & Err.Description, _
                                "initializePendingTokens_", _
                                "Object will not be usable")
                    .booUsable = False : Return (False)
                End Try
                .objTokenNext(_OBJqbToken.typeToIndex("AMPERSAND")).typeFromString("AMPERSAND")
                .objTokenNext(_OBJqbToken.typeToIndex("APOSTROPHE")).typeFromString("APOSTROPHE")
                .objTokenNext(_OBJqbToken.typeToIndex("COLON")).typeFromString("COLON")
                .objTokenNext(_OBJqbToken.typeToIndex("COMMA")).typeFromString("COMMA")
                .objTokenNext(_OBJqbToken.typeToIndex("CURRENCY")).typeFromString("CURRENCY")
                .objTokenNext(_OBJqbToken.typeToIndex("EXCLAMATION")).typeFromString("EXCLAMATION")
                .objTokenNext(_OBJqbToken.typeToIndex("IDENTIFIER")).typeFromString("IDENTIFIER")
                .objTokenNext(_OBJqbToken.typeToIndex("NEWLINE")).typeFromString("NEWLINE")
                .objTokenNext(_OBJqbToken.typeToIndex("OPERATOR")).typeFromString("OPERATOR")
                .objTokenNext(_OBJqbToken.typeToIndex("PARENTHESIS")).typeFromString("PARENTHESIS")
                .objTokenNext(_OBJqbToken.typeToIndex("PERCENT")).typeFromString("PERCENT")
                .objTokenNext(_OBJqbToken.typeToIndex("POUND")).typeFromString("POUND")
                .objTokenNext(_OBJqbToken.typeToIndex("SEMICOLON")).typeFromString("SEMICOLON")
                .objTokenNext(_OBJqbToken.typeToIndex("STRING")).typeFromString("STRING")
                .objTokenNext(_OBJqbToken.typeToIndex("UNSIGNEDINTEGER")).typeFromString("UNSIGNEDINTEGER")
                .objTokenNext(_OBJqbToken.typeToIndex("UNSIGNEDREALNUMBER")).typeFromString("UNSIGNEDREALNUMBER")
                .objTokenNext(_OBJqbToken.typeToIndex("PERIOD")).typeFromString("PERIOD")
            End If
            For intIndex1 = 1 To UBound(.objTokenNext)
                .objTokenNext(intIndex1).StartIndex = 0
            Next intIndex1
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' An internal inspection

    '
    '
    Private Function inspection_() As Boolean
        Dim strReport As String
        If Me.inspect(strReport) Then Return (True)
        errorHandler_("Inspection of object has failed" & _
                      vbNewLine & vbNewLine & vbNewLine & _
                      strReport, _
                      "inspection_", _
                      "Object is not usable")
    End Function

    ' -----------------------------------------------------------------
    ' Key (in the form _string) to string   
    '
    '
    Private Function key2String_(ByVal strKey As String) As String
        Return (Mid(strKey, 2))
    End Function

    ' -----------------------------------------------------------------
    ' Convert line index to entry handle
    '
    '
    Private Function lineIndex2Entry_(ByVal intIndex As Integer) As Collection
        With USRstate
            Try
                Return (CType(.colLineIndex.Item(string2Key_(CStr(intIndex))), _
                              Collection))
            Catch
                Return (Nothing)
            End Try
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Reset the scanner
    '
    '
    Private Function reset_() As Boolean
        With USRstate
            .intLast = LBound(.objQBtoken)
            .intLineNumber = 1
            .booScanned = False
            Return (initializePendingTokens_() AndAlso initializeLineIndex_())
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Our scanner
    '
    '
    ' This method updates the usrQBToken array.  The array is zero-
    ' origin, of course, but always contains no value in usrQBToken(0), so
    ' an array of one entry indicates nothing was found to scan.
    '
    '
    Private Overloads Function scanner_() As Boolean
        Return (scanner_(CLng(-1), CInt(-1)))
    End Function
    Private Overloads Function scanner_(ByVal intEndIndex As Integer, _
                                        ByVal intCount As Integer, _
                                        Optional ByVal intStartingIndex As Integer = 0) As Boolean
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim intIndex3 As Integer
        Dim intIndex4 As Integer
        Dim intLength As Integer
        Dim intNextIndex As Integer
        Dim intTokenStartIndex As Integer
        Dim strNext As String
        Dim intScanCount As Integer
        With USRstate
            ' --- Scan
            .objQBtoken(0).EndIndex = 0
            If intStartingIndex = 0 Then
                If .intLast > 0 Then
                    intIndex1 = .objQBtoken(.intLast).EndIndex + 1
                End If
            ElseIf intStartingIndex > 0 Then
                intIndex1 = intStartingIndex
            Else
                errorHandler_("Invalid intStartingIndex " & intStartingIndex, _
                              "scanner_", _
                              "Scan will start at 1")
            End If
            Do While intIndex1 <= Len(.strSourceCode) _
                     AndAlso _
                     (intEndIndex = -1 _
                      OrElse _
                      intIndex1 <= intEndIndex) _
                     AndAlso _
                     (intCount = -1 _
                      OrElse _
                      intScanCount < intCount)
                ' Find the leftmost and widest token
                intNextIndex = 0
                For intIndex2 = 1 To UBound(.objTokenNext)
                    With .objTokenNext(intIndex2)
                        intTokenStartIndex = .StartIndex : intLength = .Length
                        If .StartIndex < intIndex1 OrElse .Length = 0 Then
                            .StartIndex = _OBJutilities.verify(Me.SourceCode & "A", " ", intIndex1)
                            intTokenStartIndex = .StartIndex
                            Select Case UCase(.typeToString)
                                Case "AMPERSAND"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("&"), intTokenStartIndex, intLength)
                                Case "APOSTROPHE"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("'"), intTokenStartIndex, intLength)
                                Case "COLON"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar(":"), intTokenStartIndex, intLength)
                                Case "COMMA"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar(","), intTokenStartIndex, intLength)
                                Case "IDENTIFIER"
                                    scanner__scanIdentifier_(Me.SourceCode, intTokenStartIndex, intLength)
                                Case "NEWLINE"
                                    scanner__scanNewline_(Me.SourceCode, intTokenStartIndex, intLength)
                                Case "OPERATOR"
                                    scanner__scanOperator_(Me.SourceCode, intTokenStartIndex, intLength)
                                Case "PARENTHESIS"
                                    scanner__scanParenthesis_(Me.SourceCode, intTokenStartIndex, intLength)
                                Case "SEMICOLON"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar(";"), intTokenStartIndex, intLength)
                                Case "STRING"
                                    scanner__scanString_(Me.SourceCode, intTokenStartIndex, intLength)
                                Case "UNSIGNEDINTEGER"
                                    scanner__scanUnsignedInteger_(Me.SourceCode, _
                                                                  intTokenStartIndex, _
                                                                  intLength)
                                Case "UNSIGNEDREALNUMBER"
                                    scanner__scanUnsignedReal_(Me.SourceCode, _
                                                               intTokenStartIndex, _
                                                               intLength)
                                Case "PERCENT"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("%"), intTokenStartIndex, intLength)
                                Case "EXCLAMATION"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("!"), intTokenStartIndex, intLength)
                                Case "POUND"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("#"), intTokenStartIndex, intLength)
                                Case "CURRENCY"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("$"), intTokenStartIndex, intLength)
                                Case "PERIOD"
                                    scanner__scanSpecialChar_(Me.SourceCode, CChar("."), intTokenStartIndex, intLength)
                                Case Else
                                    errorHandler_("Internal compiler error: scan type " & _
                                                  _OBJutilities.enquote(.typeToString) & " " & _
                                                  "is not valid", _
                                                  "scanner", _
                                                  "Object is not usable")
                                    Me.mkUnusable()
                                    Return (False)
                            End Select
                        End If
                        If intNextIndex = 0 _
                           OrElse _
                           intTokenStartIndex < USRstate.objTokenNext(intNextIndex).StartIndex _
                           AndAlso _
                           intLength > 0 _
                           OrElse _
                           intTokenStartIndex = USRstate.objTokenNext(intNextIndex).StartIndex _
                           AndAlso _
                           intLength > USRstate.objTokenNext(intNextIndex).Length Then
                           intNextIndex = intIndex2
                        End If
                        .StartIndex = intTokenStartIndex
                        .Length = intLength
                    End With
                Next intIndex2
                ' Make sure that material that cannot be parsed is blank
                With .objQBtoken(.intLast)
                    If intStartingIndex = 0 OrElse intScanCount > 0 Then
                        Dim intGapStart As Integer = .EndIndex + 1
                        Dim intGapLength As Integer = USRstate.objTokenNext(intNextIndex).StartIndex - intGapStart
                        Dim strGap As String = Mid(USRstate.strSourceCode, intGapStart, intGapLength)
                        If strGap <> _OBJutilities.copies(intGapLength, " ") Then
                            RaiseEvent scanErrorEvent("The character(s) " & _
                                                      _OBJutilities.string2Display(strGap) & " " & _
                                                      "can't be recognized", _
                                                      intIndex1, _
                                                      USRstate.intLineNumber, _
                                                      "Scan will skip these characters")
                            Exit Do
                        End If
                    End If
                End With
                ' Test for end of scan
                If .objTokenNext(intNextIndex).StartIndex > Len(.strSourceCode) Then
                    .booScanned = True
                    Return (True)
                End If
                ' Attach the next token to the scan table
                Dim objClone As qbToken.qbToken = .objTokenNext(intNextIndex).clone
                .objTokenNext(intNextIndex).Length = 0
                If (objClone Is Nothing) Then
                    errorHandler_("Cannot clone pending token", _
                                  "scanner_", _
                                  "Scanner is not usable")
                    Me.mkUnusable()
                    Return (False)
                End If
                If Not addToken_(objClone) Then
                    Return (False)
                End If
                If .intLast > 0 Then
                    With .objQBtoken(.intLast)
                        If UCase(.typeToString) = "NEWLINE" Then
                            USRstate.intLineNumber += 1
                            If Not updateLineIndex_(USRstate.colLineIndex, _
                                                    USRstate.intLineNumber, _
                                                    .EndIndex + 1, _
                                                    USRstate.strSourceCode) Then
                                Return (False)
                            End If
                        End If
                    End With
                End If
                With .objQBtoken(.intLast)
                    .LineNumber = USRstate.intLineNumber
                    If UCase(.typeToString) = "AMPERSAND" _
                       AndAlso _
                       USRstate.intLast >= 2 Then
                        ' Special handling: an ampersand may indicate Long data type
                        intIndex2 = USRstate.intLast - 1
                        If UCase(USRstate.objQBtoken(intIndex2).typeToString) = "IDENTIFIER" _
                           AndAlso _
                           USRstate.objQBtoken(intIndex2).EndIndex + 1 _
                           = _
                           .StartIndex Then
                            .typeFromString("AMPERSANDSUFFIX")
                        End If
                    End If
                    intIndex1 = .StartIndex + .Length
                End With
                RaiseEvent scanEvent(.objQBtoken(.intLast), _
                                     intIndex1, _
                                     Len(.strSourceCode), _
                                     UBound(.objQBtoken))
                intScanCount += 1
            Loop
            .booScanned = (intIndex1 > Len(.strSourceCode))
            Return (True)
        End With
    End Function

    ' ----------------------------------------------------------------------
    ' Scan the next identifier; set start index to one past end of string if
    ' no identifier is available
    '
    '
    Private Sub scanner__scanIdentifier_(ByVal strInstring As String, _
                                         ByRef intStartIndex As Integer, _
                                         ByRef intLength As Integer)
        intStartIndex = _OBJutilities.verify(strInstring & "A", _
                               "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", _
                               intStartIndex, _
                               True)
        If intStartIndex > Len(strInstring) Then
            intLength = 0 : Return
        End If
        intLength = _OBJutilities.verify(strInstring & " ", _
                           "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_", _
                           intStartIndex) _
                    - _
                    intStartIndex
    End Sub

    ' ----------------------------------------------------------------------
    ' Scan the next newline
    '
    '
    Private Sub scanner__scanNewline_(ByVal strInstring As String, _
                                     ByRef intStartIndex As Integer, _
                                     ByRef intLength As Integer)
        intStartIndex = CInt(Math.Min(CDbl(InStr(intStartIndex, strInstring & ":", ":")), _
                                      Math.Min(CDbl(InStr(intStartIndex, strInstring & vbNewLine, vbNewLine)), _
                                               CDbl(InStr(intStartIndex, strInstring & Chr(10), Chr(10))))))
        If intStartIndex > Len(strInstring) Then
            intLength = 0 : Return
        End If
        If Mid(strInstring, intStartIndex, Len(vbNewLine)) = vbNewLine Then
            intLength = Len(vbNewLine)
        Else
            intLength = 1
        End If
    End Sub

    ' ----------------------------------------------------------------------
    ' Scan the next unsigned integer
    '
    '
    ' The following regular expression defines the unsigned integer: [0-9]+
    '
    '
    Private Function scanner__scanUnsignedInteger_(ByVal strInstring As String, _
                                                   ByRef intStartIndex As Integer, _
                                                   ByRef intLength As Integer) _
            As Boolean
        intStartIndex = _OBJutilities.verify(strInstring & "0", _
                                             "0123456789", _
                                             intStartIndex, _
                                             True)
        If intStartIndex > Len(strInstring) Then
            intLength = 0 : Return (False)
        End If
        intLength = _OBJutilities.verify(strInstring & " ", _
                                         "0123456789", _
                                         intStartIndex) _
                    - _
                    intStartIndex
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' Scan the next unsigned real number
    '
    '
    ' The following BNF grammar defines the unsigned real number
    '
    '
    '      unsignedReal := [ INTEGERPART ] realPart
    '      realPart := fraction
    '      realPart := exponent
    '      realPart := fraction exponent
    '      fraction := . INTEGERPART
    '      exponent := ( e|E ) [ +|- ] INTEGERPART
    '
    '
    ' Note: in fraction := . INTEGERPART, the INTEGERPART may be null as long as there
    ' is a non-null INTERGERPART to the left of the decimal point.
    '
    '
    Private Sub scanner__scanUnsignedReal_(ByVal strInstring As String, _
                                           ByRef intStartIndex As Integer, _
                                           ByRef intLength As Integer)
        Dim booIntegerPart As Boolean
        Dim intIndex0 As Integer = intStartIndex
        Dim intIndex1 As Integer
        Dim intLength1 As Integer
        intLength = 0
        Do While intIndex0 <= Len(strInstring)
            intLength = 0 : booIntegerPart = False
            intIndex1 = _OBJutilities.verify(strInstring & "0", _
                                             ".0123456789", _
                                             intIndex0, _
                                             True)
            If intIndex1 > Len(strInstring) Then Exit Do
            If Mid(strInstring, intIndex1, 1) = "." Then
                intStartIndex = intIndex1
            ElseIf scanner__scanUnsignedReal__integerPart_(strInstring, _
                                                           intIndex1, _
                                                           intLength1) Then
                intStartIndex = intIndex1 : intLength = intLength1
                booIntegerPart = True
                intIndex1 += intLength1
            End If
            If scanner__scanUnsignedReal__realPart_(strInstring, _
                                                    intIndex1, _
                                                    intLength1, _
                                                    booIntegerPart) Then
                If intStartIndex = 0 Then intStartIndex = intIndex1
                intLength += intLength1
                Exit Do
            End If
            If intLength <> 0 Then Exit Do
            intStartIndex = 0
            intLength = 0
            intIndex0 += 1
        Loop
    End Sub

    ' ----------------------------------------------------------------------
    ' On behalf of scanner_scanUnsignedReal_ parse the exponent 
    '
    '      exponent := ( e|E ) [ +|- ] INTEGERPART
    '
    '
    Private Function scanner__scanUnsignedReal__exponent_(ByVal strInstring As String, _
                                                          ByRef intStartIndex As Integer, _
                                                          ByRef intLength As Integer) _
            As Boolean
        Dim strNext As String = Mid(strInstring, intStartIndex, 1)
        If strNext <> "e" AndAlso strNext <> "E" Then Return (False)
        intLength = 1
        strNext = Mid(strInstring, intStartIndex + 1, 1)
        If strNext = "+" OrElse strNext = "-" Then
            intLength += 1
        End If
        Dim intLength1 As Integer
        If scanner__scanUnsignedReal__integerPart_(strInstring, _
                                                   intStartIndex + intLength, _
                                                   intLength1) Then
            intLength += intLength1
            Return (True)
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of scanner_scanUnsignedReal_ parse the fraction 
    '
    '      fraction := . INTEGERPART
    '
    '
    Private Function scanner__scanUnsignedReal__fraction_(ByVal strInstring As String, _
                                                          ByRef intStartIndex As Integer, _
                                                          ByRef intLength As Integer, _
                                                          ByVal booNullFractionOK As Boolean) _
            As Boolean
        If Mid(strInstring, intStartIndex, 1) <> "." Then Return (False)
        Dim intIndex1 As Integer = intStartIndex + 1
        Dim intLength1 As Integer
        If Not scanner__scanUnsignedReal__integerPart_(strInstring, _
                                                       intIndex1, _
                                                       intLength1) _
           AndAlso _
           Not booNullFractionOK Then Return (False)
        intLength = intLength1 + 1
        Return (True)
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of scanner_scanUnsignedReal_ parse the integer part 
    '
    '
    Private Function scanner__scanUnsignedReal__integerPart_(ByVal strInstring As String, _
                                                             ByVal intStartIndex As Integer, _
                                                             ByRef intLength As Integer) _
            As Boolean
        Dim intIndex1 As Integer = intStartIndex
        Dim intLength1 As Integer
        If Not scanner__scanUnsignedInteger_(strInstring, _
                                             intIndex1, _
                                             intLength1) Then Return (False)
        If intIndex1 = intStartIndex Then
            intLength = intLength1 : Return (True)
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' On behalf of scanner_scanUnsignedReal_ parse the fraction and/or
    ' exponent
    '
    '      realPart := fraction
    '      realPart := exponent
    '      realPart := fraction exponent
    '
    '
    Private Function scanner__scanUnsignedReal__realPart_(ByVal strInstring As String, _
                                                          ByRef intStartIndex As Integer, _
                                                          ByRef intLength As Integer, _
                                                          ByVal booNullFractionOK As Boolean) _
            As Boolean
        If scanner__scanUnsignedReal__fraction_(strInstring, _
                                                intStartIndex, _
                                                intLength, _
                                                booNullFractionOK) Then
            Dim intIndex1 As Integer = intStartIndex + intLength
            Dim intLength1 As Integer
            If scanner__scanUnsignedReal__exponent_(strInstring, _
                                                    intIndex1, _
                                                    intLength1) Then
                intLength += intLength1
            End If
            Return (True)
        Else
            Return (scanner__scanUnsignedReal__exponent_(strInstring, _
                                                        intStartIndex, _
                                                        intLength))
        End If
    End Function

    ' ----------------------------------------------------------------------
    ' Scan the next operator; set start index to one past end of string if
    ' no operator is available
    '
    '
    Private Sub scanner__scanOperator_(ByVal strInstring As String, _
                                      ByRef intStartIndex As Integer, _
                                      ByRef intLength As Integer)
        intStartIndex = _OBJutilities.verify(strInstring & "+", "+-*/\^<>=", intStartIndex, True)
        intLength = 1
        If intStartIndex < Len(strInstring) Then
            If Mid(strInstring, intStartIndex, 1) = "*" _
            AndAlso _
            Mid(strInstring, intStartIndex + 1, 1) = "*" _
            OrElse _
            Mid(strInstring, intStartIndex, 1) = "<" _
            AndAlso _
            (Mid(strInstring, intStartIndex + 1, 1) = "=" _
             OrElse _
             Mid(strInstring, intStartIndex + 1, 1) = ">") _
            OrElse _
            Mid(strInstring, intStartIndex, 1) = ">" _
            AndAlso _
            (Mid(strInstring, intStartIndex + 1, 1) = "=") Then
                intLength = 2
            End If
        End If
    End Sub

    ' ----------------------------------------------------------------------
    ' Scan the next parenthesis; set start index to one past end of string if
    ' no parenthesis is available
    '
    '
    Private Sub scanner__scanParenthesis_(ByVal strInstring As String, _
                                             ByRef intStartIndex As Integer, _
                                             ByRef intLength As Integer)
        intStartIndex = _OBJutilities.verify(strInstring & "(", "()", intStartIndex, True)
        intLength = 1
    End Sub

    ' ----------------------------------------------------------------------
    ' Scan the next special char; set start index to one past end of string if
    ' no special character is available
    '
    '
    Private Sub scanner__scanSpecialChar_(ByVal strInstring As String, _
                                          ByVal chrChar As Char, _
                                          ByRef intStartIndex As Integer, _
                                          ByRef intLength As Integer)
        intStartIndex = InStr(intStartIndex, strInstring & chrChar, chrChar)
        If intStartIndex > Len(strInstring) Then
            intLength = 0 : Return
        End If
        intLength = 1
        Return
    End Sub

    ' ----------------------------------------------------------------------
    ' Scan the next string
    '
    '
    ' This method sets the start index past the string end, and the length to
    ' zero, when no string is found. When called as a function, it returns
    ' the string as its function value.
    '
    '
    Private Overloads Function scanner__scanString_(ByVal strInstring As String, _
                                                    ByRef intStartIndex As Integer, _
                                                    ByRef intLength As Integer) As String
        intLength = 0
        Dim strQuotes As String = """" & ChrW(SMARTQUOTELEFT) & ChrW(SMARTQUOTERIGHT)
        intStartIndex = _OBJutilities.verify(strInstring & """", _
                                             strQuotes, _
                                             intStartIndex, _
                                             True)
        If intStartIndex > Len(strInstring) Then
            intLength = 0 : Return ("")
        End If
        Dim intIndex1 As Integer = intStartIndex + 1
        Dim intIndex2 As Integer
        Dim strStringValue As String = """"
        Do While intIndex1 <= Len(strInstring)
            intIndex2 = _OBJutilities.verify(strInstring & """", _
                                             strQuotes, _
                                             intIndex1, _
                                             True)
            strStringValue &= Mid(strInstring, intIndex1, intIndex2 - intIndex1 + 1)
            If _OBJutilities.translate(Mid(strInstring, intIndex2 + 1, 1), _
                                       strQuotes, _
                                       _OBJutilities.copies("""", Len(strQuotes))) _
               <> """" Then Exit Do
            strStringValue &= """"
            intIndex1 = intIndex2 + 2
        Loop
        intLength = Len(strStringValue)
        Return (strStringValue)
    End Function

    ' -----------------------------------------------------------------
    ' String to key
    '
    '
    Private Function string2Key_(ByVal strInstring As String) As String
        Return ("_" & strInstring)
    End Function

    ' -----------------------------------------------------------------
    ' Token array to multiline string
    '
    '
    Private Function tokenArrayToString_(ByRef objArray() As qbToken.qbToken, _
                                         ByVal intStartIndex As Integer, _
                                         ByVal intLength As Integer) As String
        Dim intIndex1 As Integer
        Dim objStringBuilder As System.Text.StringBuilder
        Try
            objStringBuilder = New System.Text.StringBuilder
        Catch
            errorHandler_("Cannot create StringBuilder: " & _
                          Err.Number & " " & Err.Description, _
                          "tokenArrayToString_", _
                          "Marking object unusable and returning a null string")
            Me.mkUnusable()
            Return ("")
        End Try
        For intIndex1 = intStartIndex To intStartIndex + intLength - 1
            If Not (objArray(intIndex1) Is Nothing) Then
                With objArray(intIndex1)
                    _OBJutilities.append(objStringBuilder, _
                                        vbNewLine, _
                                        .ToString & " " & _
                                        _OBJutilities.string2Display(.sourceCode(USRstate.strSourceCode)))
                End With
            End If
        Next intIndex1
        Return (objStringBuilder.ToString)
    End Function

    ' -----------------------------------------------------------------
    ' Dispose of the token object
    '
    '
    Private Function tokenDispose_(ByVal objToken As qbToken.qbToken) As Boolean
        With objToken
            If Not .dispose Then
                errorHandler_("Cannot dispose of token " & _OBJutilities.enquote(.Name), _
                              "tokenDispose_", _
                              "Marking the scanner object not usable")
                Return (False)
            End If
        End With
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Update line index
    '
    '
    Private Function updateLineIndex_(ByVal colLineIndex As Collection, _
                                      ByVal intLineNumber As Integer, _
                                      ByVal intStartIndex As Integer, _
                                      ByVal strSourceCode As String) As Boolean
        Dim colEntry As Collection
        Try
            colEntry = New Collection
        Catch
            errorHandler_("Can't create subcollection handle for line index: " & _
                          Err.Number & " " & Err.Description, _
                          "", _
                          "Scanner is unusable")
            Me.mkUnusable() : Return (False)
        End Try
        Dim strKey As String = string2Key_(CStr(intLineNumber))
        With colEntry
            Try
                .Add(strKey)
                .Add(intStartIndex)
                .Add(InStr(intStartIndex, strSourceCode & vbNewLine, vbNewLine) _
                     - _
                     intStartIndex)
                colLineIndex.Add(colEntry, strKey)
            Catch
                errorHandler_("Can't extend line index: " & _
                              Err.Number & " " & Err.Description, _
                              "", _
                              "Scanner is unusable")
                Me.mkUnusable() : Return (False)
            End Try
            Return (True)
        End With
    End Function

End Class
