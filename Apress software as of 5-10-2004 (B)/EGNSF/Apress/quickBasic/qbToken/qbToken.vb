Option Strict

Imports System.Threading

' *********************************************************************
' *                                                                   *
' * qbToken     quickBasicEngine: token data model                    *
' *                                                                   *
' *                                                                   *
' * This class defines one scan token as is used in the quickBasic    *
' * engine including its start index, length, type and its line       *
' * number.  The rest of this comment block defines the following     *
' * topics.                                                           *
' *                                                                   *
' *                                                                   *
' *      *  The token data model                                      *
' *      *  Properties and methods of this class                      *
' *      *  Change record                                             *
' *                                                                   *
' *                                                                   *
' * THE TOKEN DATA MODEL -------------------------------------------- *
' *                                                                   *
' * For our purposes, the token consists of the following information.*
' *                                                                   *
' *                                                                   *
' *      *  The token type: there's a comprehensive list of these     *
' *         types in the qbTokenType project.                         *
' *                                                                   *
' *      *  The start index, from one, of the token in the source     *
' *                                                                   *
' *      *  The length of the token in the source                     *
' *                                                                   *
' *      *  The line number, from one, of the token in the source     *
' *                                                                   *
' *                                                                   *
' * Note that the token data model doesn't include the value of the   *
' * token, for this would make the data structures in this class      *
' * larger, by definition, than the source code. Instead the user code*
' * is expected to use the start index and the length to get the raw  *
' * source code.                                                      *
' *                                                                   *
' *                                                                   *
' * PROPERTIES AND METHODS OF THIS CLASS ---------------------------- *
' *                                                                   *
' * Properties of this class start with an upper case letter; methods *
' * with a lower case letter.                                         *
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
' *                                                                   *
' * About: this shared, read-only property returns information about  *
' *      this class                                                   *
' *                                                                   *
' * clone: since tokens are ICloneable, this method creates a new     *
' *      and identical token object based on the instance             *
' *                                                                   *
' * dispose: this method disposes of the object: although at this     *
' *      writing, dispose is unnecessary (because no reference objects*
' *      are declared in the token's state), for best results use this*
' *      method in code when you are finished with the token. This    *
' *      method will mark the object as unusable.                     *
' *                                                                   *
' * EndIndex: this read-write property returns and can be set to the  *
' *      ending index, from one, of the token. Note that it sets the  *
' *      length of the token in the object state, and is calculated   *
' *      from the length.                                             *
' *                                                                   *
' * fromString: this method sets the object to the values in an       *
' *      expression of the format returned by toString:               *
' *                                                                   *
' *      <type>@<startIndex>..<endIndex>:<lineNumber>                 *
' *                                                                   *
' *      The line number, and the colon preceding the line number, is *
' *      optional.                                                    *
' *                                                                   *
' * inspect(r): this method inspects the object. It checks for errors *
' *      that result from blunders in the source code of this class,  *
' *      or "object abuse", not simple user errors.                   *
' *                                                                   *
' *      r should be a string, passed by reference; it is assigned an *
' *      inspection report.                                           *
' *                                                                   *
' *      The following inspection rules are used.                     *
' *                                                                   *
' *           *  The object instance must be usable                   *
' *                                                                   *
' *           *  The type must be a valid enumerator value, other than*
' *              Invalid                                              *
' *                                                                   *
' *           *  The start index must be greater than or equal to zero*                                        *
' *                                                                   *
' *           *  The length must be greater than or equal to zero     *
' *                                                                   *
' *           *  The line number must be zero or greater              *
' *                                                                   *
' *      If the inspection is failed the object becomes unusable.     *
' *                                                                   *
' *      An internal inspection is carried out in the constructor and *
' *      inside the dispose method.                                   *
' *                                                                   *
' * Length: this read-write property returns and may be set to the    *
' *      length of the token.                                         *
' *                                                                   *
' * LineNumber: this read-write property returns and may be set to the*
' *      line number of the token.                                    *
' *                                                                   *
' * mkUnusable: this method makes the object not usable.              *
' *                                                                   *
' * Name: this read-write property returns and can change a name of   *
' *      the object instance. Name defaults to qbTokennnnn date time, *
' *      where nnnn is an object sequence number.                     *
' *                                                                   *
' * object2XML: this method converts the state of the object to an    *
' *      eXtendedMarkupLanguage string.                               *
' *                                                                   *
' * StartIndex: this read-write property returns and may be set to the*
' *      start index of the token: note that this property can be set *
' *      to zero, usually to indicate a nonexistent token.            * 
' *                                                                   *
' * sourceCode(s): when provided the complete source code in s, this  *
' *      method returns the source code corresponding to the token.   *
' *                                                                   *
' * TokenType: this read-only property returns the type of the token  *
' *      as an ENUtokenType                                           *
' *                                                                   *
' * tokenTypeMatch(t): this method matches the token in the object    *
' *      instance with the ENUtokenType t. It returns True when the   *
' *      token types are identical and when the range of the instance *
' *      is a part of the range of t. For example, if the             *
' *      instance is an unsigned integer and t is "unsigned           *
' *      real number" this method will return True.                   * 
' *                                                                   *
' * toString: this override method returns the token as a string in   *             
' *      the format                                                   *
' *                                                                   *
' *      <type>@<startIndex>..<endIndex>:<lineNumber>                 *
' *                                                                   *
' *      In the above:                                                *
' *                                                                   *
' *           *  <type> is the token type as returned by the          *
' *              typeToString method.                                 *
' *                                                                   *
' *           *  <startIndex> is its start index                      *
' *                                                                   *
' *           *  <endIndex> is its end index (start index plus length *
' *              minus one)                                           *
' *                                                                   *
' *           *  <lineNumber> is the line number. If the token is on  *
' *              more than one line, this is the number of the line   *
' *              where the token begins                               *
' *                                                                   *
' * TypeCount: this shared, read-only property returns the number of  *
' *      distinct token types defined excluding null, invalid and     *
' *      ampersandSuffix.                                             *
' *                                                                   *
' * TypeCountActual: this shared, read-only property returns the      *
' *      number of distinct token types defined including null,       *
' *      invalid and ampersandSuffix.                                 *
' *                                                                   *
' * typeFromString(t): this method sets the type using the string t:  *
' *      t must be one of the types identified below (under           *
' *      typeToString) after leading and trailing blanks and case     *
' *      differences are ignored.                                     *
' *                                                                   *
' * typeToEnum: this Shared method returns the distinct ENUtokenType  *
' *      enum identified by a case-insensitive name, from which lead- *
' *      ing and trailing blanks are removed.                         * 
' *                                                                   *
' *      If the prefix "tokenType" is not present in it will be added.                                    *
' *                                                                   *
' * typeToIndex: this Shared method returns the distinct index value  *
' *      type identified by a case-insensitive name, from which lead- *
' *      ing and trailing blanks are removed.                         * 
' *                                                                   *
' * typeToString: this method returns the type as one of the following*
' *      strings: Null, Apostrophe, Ampersand, Colon, Comma,          *
' *      Identifier, Newline, Operator, Parenthesis, Semicolon,       *
' *      String, UnsignedInteger, UnsignedRealNumber, Percent,        *
' *      Exclamation, Pound, Currency, AmpersandSuffix, NullValue or  *
' *      Invalid.                                                     *
' *                                                                   *
' * Usable: this read-only property returns True if the object is     *
' *      usable, False otherwise.                                     *
' *                                                                   *
' *                                                                   *
' * C H A N G E   R E C O R D --------------------------------------- *
' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE                   *
' * --------   ----------     --------------------------------------- *
' * 05 01 03   Nilges         Version 1                               *
' *                                                                   *
' *********************************************************************

Public Class qbToken
    
    Implements ICloneable

    ' ***** Shared *****
    Private Shared _INTsequence As Integer
    Private Shared _OBJutilities As utilities.utilities

    ' ***** Scanned enum *****
    ' --- The number of scanned types: must be equal to number of ENUscanned values
    Private Const SCANTYPES As Integer = 19

    ' ***** State *****
    Private Structure TYPstate
        Dim strName As String                               ' Object name
        Dim booUsable As Boolean                            ' Object usability
        Dim enuType As qbTokenType.qbTokenType.ENUtokenType ' Type
        Dim intStartIndex As Integer                        ' Start index
        Dim intLength As Integer                            ' Length
        Dim intLineNumber As Integer                        ' Linenumber
    End Structure
    Private USRstate As TYPstate

    ' ***** Constants *****
    ' --- Easter yEgg
    Private Const ABOUTINFO As String = _
        "qbToken" & _
        vbNewLine & vbNewLine & vbNewLine & _
        "This class defines one scan token as is used in the quickBasic " & _
        "engine including its start index, length, type and its line " & _
        "number." & vbNewLine & vbNewLine & _
        "This class was developed commencing on 4/30/2003 by" & vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & _
        "http://members.screenz.com/edNilges"
    ' --- Class name
    Private Const CLASS_NAME As String = "qbToken"
    ' --- Inspection
    Private Const INSPECTION_USABLE As String = _
        "The object must be usable"
    Private Const INSPECTION_TYPE As String = _
        "The type must be a valid enumerator value, other than Invalid"
    Private Const INSPECTION_STARTINDEX As String = _
        "The token start index must be greater than or equal to zero"
    Private Const INSPECTION_LENGTH As String = _
        "The token length must be greater than or equal to zero"
    Private Const INSPECTION_LINENUMBER As String = _
        "The token line number must be zero or greater"

    ' ***** Object constructor ****************************************

    Public Sub New()
        With USRstate
            Interlocked.Increment(_INTsequence)
            .strName = "qbToken" & _
                       _OBJutilities.alignRight(CStr(_INTsequence), 4, "0") & " " & _
                       Now
            If Not reset_() Then Return
            .booUsable = True
            inspection_()
        End With
    End Sub

    ' ***** Public procedures *****************************************

    ' -----------------------------------------------------------------
    ' Return about info
    '
    '
    Public Shared ReadOnly Property About() As String
        Get
            Return (ABOUTINFO)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return class name
    '
    '
    Public Shared ReadOnly Property ClassName() As String
        Get
            Return (CLASS_NAME)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Clone the token
    '
    '
    ' --- Public implementor
    Public Function clone() As qbToken
        Dim objNew As Object = clone_()
        If (objNew Is Nothing) Then Return (Nothing)
        Return (CType(objNew, qbToken))
    End Function
    ' --- Hidden implementor
    Private Function clone_() As Object Implements ICloneable.Clone
        If Not checkUsable_("clone_") Then Return (Nothing)
        Dim objNew As qbToken
        Try
            objNew = New qbToken
        Catch
            errorHandler_("Cannot create clone: " & Err.Number & " " & Err.Description, _
                          "clone_", _
                          "Returning Nothing")
            Return (Nothing)
        End Try
        With objNew
            .Name = .Name & " (clones " & _OBJutilities.enquote(Me.Name) & ")"
            .typeFromString(Me.typeToString)
            .StartIndex = Me.StartIndex
            .EndIndex = Me.EndIndex
            Return (objNew)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Disposer
    '
    '
    ' --- Inspect by default
    Public Function dispose() As Boolean
        Return (Me.dispose(True))
    End Function
    ' --- Exposes an inspection option
    Public Function dispose(ByVal booInspect As Boolean) As Boolean
        If booInspect Then inspection_()
        Me.mkUnusable()
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Return and set end index
    '
    '
    Public Property EndIndex() As Integer
        Get
            If Not checkUsable_("EndIndex get") Then Return (0)
            With Me
                Return (.StartIndex + .Length - 1)
            End With
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("EndIndex set") Then Return
            With Me
                If intNewValue < .StartIndex Then
                    errorHandler_("End index " & intNewValue & " precedes " & _
                                  "start index " & .StartIndex, _
                                  "EndIndex set", _
                                  "Object state has not been changed")
                    Return
                End If
                .Length = intNewValue - .StartIndex + 1
            End With
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Convert the object from a string in the format 
    ' <type>@<startIndex>..<endIndex>:<lineNumber>
    '
    '
    Public Function fromString(ByVal strToString As String) As Boolean
        If Not checkUsable_("fromString") Then Return (False)
        Dim strSplit() As String
        Try
            strSplit = Split(strToString, "@")
        Catch
            errorHandler_("Cannot split around at sign: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End Try
        If UBound(strSplit) <> 1 Then
            errorHandler_("Cannot split around at sign", _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End If
        Dim enuType As qbTokenType.qbTokenType.ENUtokenType = typeFromString_(strSplit(0))
        If enuType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid Then
            errorHandler_("Invalid type " & _OBJutilities.enquote(strSplit(0)), _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End If
        Try
            strSplit = Split(strSplit(1), ":")
        Catch
            errorHandler_("Cannot split around colon: " & _
                          Err.Number & " " & Err.Description, _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End Try
        If UBound(strSplit) < 0 Then
            errorHandler_("Cannot split around colon", _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End If
        Dim strSplit2() As String
        Try
            strSplit2 = Split(strSplit(0), "..")
        Catch
            errorHandler_("Cannot split around "".."": " & _
                          Err.Number & " " & Err.Description, _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End Try
        If UBound(strSplit2) <> 1 Then
            errorHandler_("Cannot split around ""..""", _
                          "fromString", _
                          "Object state will not change")
            Return (False)
        End If
        With Me
            Dim intStartIndexSave As Integer = .StartIndex
            Try
                .StartIndex = CInt(strSplit2(0))
                .EndIndex = CInt(strSplit2(1))
            Catch
                .StartIndex = intStartIndexSave
                errorHandler_("Cannot set start index and/or length: " & _
                              Err.Number & " " & Err.Description, _
                              "fromString", _
                              "Object state will not change")
                Return (False)
            End Try
            .LineNumber = 0
            If UBound(strSplit) < 1 Then Return (True)
            Try
                .LineNumber = CInt(strSplit(1))
            Catch
                errorHandler_("Cannot set line number: " & Err.Number & " " & Err.Description, _
                              "fromString", _
                              "Object state will not change")
                Return (False)
            End Try
            Return (True)
        End With
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
            Dim strType As String = UCase(type2String_(.enuType))
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_TYPE, _
                                           strType <> "INVALID", _
                                           booInspection)
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_STARTINDEX, _
                                           .intStartIndex >= 0, _
                                           booInspection)
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_LENGTH, _
                                           .intLength >= 0, _
                                           booInspection)
            _OBJutilities.inspectionAppend(strReport, _
                                           INSPECTION_LINENUMBER, _
                                           .intLineNumber >= 0, _
                                           booInspection)
            If Not booInspection Then Me.mkUnusable()
            Return (booInspection)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return and change the token length
    '
    '
    Public Property Length() As Integer
        Get
            If Not checkUsable_("Length get") Then Return (0)
            Return (USRstate.intLength)
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("Length set") Then Return
            If intNewValue < 0 Then
                errorHandler_("Invalid length " & intNewValue, _
                              "Length set", _
                              "Length won't be changed")
                Return
            End If
            USRstate.intLength = intNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return and change the token's line number
    '
    '
    Public Property LineNumber() As Integer
        Get
            If Not checkUsable_("Linenumber get") Then Return (0)
            Return (USRstate.intLineNumber)
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("LineNumber set") Then Return
            If intNewValue < 0 Then
                errorHandler_("Invalid line number " & intNewValue, _
                              "LineNumber set", _
                              "LineNumber won't be changed")
                Return
            End If
            USRstate.intLineNumber = intNewValue
        End Set
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
    ' Convert the object to XML
    '
    '
    Public Function object2XML() As String
        With _OBJutilities
            Return .mkXMLElement("qbToken", _
                                 vbNewLine & _
                                 .mkXMLComment(vbNewLine & .string2Box(Me.About) & vbNewLine) & _
                                 vbNewLine & _
                                 "    " & .mkXMLComment("Name of the object instance") & vbNewLine & _
                                 "    " & .mkXMLElement("Name", Me.Name) & vbNewLine & _
                                 "    " & .mkXMLComment("Usability") & vbNewLine & _
                                 "    " & .mkXMLElement("Usable", CStr(Me.Usable)) & vbNewLine & _
                                 "    " & .mkXMLComment("Token type") & vbNewLine & _
                                 "    " & .mkXMLElement("TokenType", USRstate.enuType.ToString) & vbNewLine & _
                                 "    " & .mkXMLComment("Token start index") & vbNewLine & _
                                 "    " & .mkXMLElement("StartIndex", CStr(USRstate.intStartIndex)) & vbNewLine & _
                                 "    " & .mkXMLComment("Token length") & vbNewLine & _
                                 "    " & .mkXMLElement("Length", CStr(USRstate.intLength)) & vbNewLine & _
                                 "    " & .mkXMLComment("Line (if any) at which the token begins") & vbNewLine & _
                                 "    " & .mkXMLElement("LineNumber", CStr(USRstate.intLineNumber)) & vbNewLine & _
                                 "    " & .mkXMLComment("toString information") & vbNewLine & _
                                 "    " & .mkXMLElement("toString", Me.ToString) & vbNewLine)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return source code
    '
    '
    Public Function sourceCode(ByVal strSource As String) As String
        If Not checkUsable_("sourceCode") Then Return ("")
        If Me.StartIndex = 0 Then Return ("")
        Return (Mid(strSource, Me.StartIndex, Me.Length))
    End Function

    ' -----------------------------------------------------------------
    ' Return and change the start index
    '
    '
    Public Property StartIndex() As Integer
        Get
            If Not checkUsable_("StartIndex get") Then Return (0)
            Return (USRstate.intStartIndex)
        End Get
        Set(ByVal intNewValue As Integer)
            If Not checkUsable_("StartIndex set") Then Return
            If intNewValue < 0 Then
                errorHandler_("Invalid start index " & intNewValue, _
                              "StartIndex set", _
                              "Start index won't be changed")
                Return
            End If
            USRstate.intStartIndex = intNewValue
        End Set
    End Property

    ' -----------------------------------------------------------------
    ' Return token type (as enumerator)
    '
    '
    Public ReadOnly Property TokenType() As qbTokenType.qbTokenType.ENUtokenType
        Get
            If Not checkUsable_("TokenType Get") Then
                Return (qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid)
            End If
            Return (USRstate.enuType)
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Match token types
    '
    '
    Public Function tokenTypeMatch(ByVal enuType As qbTokenType.qbTokenType.ENUtokenType) As Boolean
        If Not checkUsable_("tokenTypeMatch") Then Return (False)
        With Me
            If .TokenType = enuType Then Return (True)
            If .TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger _
               AndAlso _
               enuType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedRealNumber Then
                Return (True)
            End If
            Return (False)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Convert object state to <type>@<startIndex>..<endIndex>:
    ' <lineNumber>
    '
    '
    Public Overrides Function toString() As String
        If Not checkUsable_("toString") Then Return ("")
        With Me
            Return (type2String_(USRstate.enuType) & "@" & _
                   .StartIndex & ".." & .EndIndex & _
                   CStr(IIf(.LineNumber = 0, "", ":" & .LineNumber)))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Return number of distinct token types including null, invalid,
    ' and ampersandSuffix
    '
    '
    Public Shared ReadOnly Property TypeCount() As Integer
        Get
            Return qbTokenType.qbTokenType.TOKENTYPECOUNT
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Return number of distinct token types including null, invalid,
    ' and ampersandSuffix
    '
    '
    Public Shared ReadOnly Property TypeCountActual() As Integer
        Get
            Return qbTokenType.qbTokenType.TOKENTYPECOUNT_ACTUAL
        End Get
    End Property

    ' -----------------------------------------------------------------
    ' Sets the type from a string
    '
    '
    Public Function typeFromString(ByVal strType As String) As Boolean
        If Not checkUsable_("typeFromString") Then Return (False)
        Dim enuType As qbTokenType.qbTokenType.ENUtokenType = typeFromString_(strType)
        If enuType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid Then
            errorHandler_("Cannot set type to invalid", _
                          "typeFromString", _
                          "Object state has not been changed")
            Return (False)
        End If
        USRstate.enuType = enuType
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Converts a type name to its index value
    '
    '
    Public Shared Function typeToEnum(ByVal strType As String) _
           As qbTokenType.qbTokenType.ENUtokenType
        Dim strTypeWork As String = UCase(Trim(strType))
        If Len(strTypeWork) < 9 OrElse Mid(strTypeWork, 1, 9) <> "TOKENTYPE" Then
            strTypeWork = "TOKENTYPE" & strTypeWork
        End If
        Select Case strTypeWork
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersand.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersand
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersandSuffix.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersandSuffix
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeApostrophe.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeApostrophe
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeCurrency.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeCurrency
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeExclamation.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeExclamation
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeNull.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeNull
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeOperator.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeOperator
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeParenthesis.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeParenthesis
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypePercent.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypePercent
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypePeriod.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypePeriod
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypePound.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypePound
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeSemicolon.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeSemicolon
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeString.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeString
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger
            Case UCase(qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedRealNumber.ToString)
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedRealNumber
            Case Else
                errorHandlerShared_("Invalid type name " & _OBJutilities.enquote(strType), _
                                    "typeToEnum", _
                                    "Returning the invalid type")
                Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Converts a type name to its index value
    '
    '
    Public Shared Function typeToIndex(ByVal strType As String) As Integer
        Dim objQBtoken As qbToken
        Try
            objQBtoken = New qbToken
        Catch ex As Exception
            errorHandlerShared_("Can't create qbToken: " & _
                                Err.Number & " " & Err.Description, _
                                "typeToIndex", _
                                "Returning zero.  Debugging information follows." & _
                                vbNewLine & vbNewLine & _
                                ex.ToString)
        End Try
        With objQBtoken
            .typeFromString(strType)
            Dim intIndex1 As Integer = CInt(.TokenType)
            .dispose()
            Return (intIndex1)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Returns the type as a string
    '
    '
    Public Function typeToString() As String
        If Not checkUsable_("typeToString") Then Return ("Invalid")
        Return type2String_(USRstate.enuType)
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

    ' ***** Private procedures **************************************** 

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
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Interface to the error handler
    '
    '
    ' --- Stateful version passes instance name
    Private Overloads Sub errorHandler_(ByVal strMessage As String, _
                                        ByVal strProcedure As String, _
                                        ByVal strHelp As String)
        errorHandler__(strMessage, Me.Name, strProcedure, strHelp)
    End Sub
    ' --- Stateless version passes class name
    Private Shared Sub errorHandlerShared_(ByVal strMessage As String, _
                                           ByVal strProcedure As String, _
                                           ByVal strHelp As String)
        errorHandler__(strMessage, "qbToken", strProcedure, strHelp)
    End Sub
    ' --- Common error handler
    Private Shared Sub errorHandler__(ByVal strMessage As String, _
                                      ByVal strObjectName As String, _
                                      ByVal strProcedure As String, _
                                      ByVal strHelp As String)
        _OBJutilities.errorHandler(strMessage, _
                                   strObjectName, _
                                   strProcedure, _
                                   strHelp)
    End Sub

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
    ' Reset the token
    '
    '
    Private Function reset_() As Boolean
        With USRstate
            .enuType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeNull
            .intStartIndex = 0
            .intLength = 0
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Converts the token's type to a string
    '
    '
    Private Function type2String_(ByVal enuType As qbTokenType.qbTokenType.ENUtokenType) As String
        Try
            Return (Mid(enuType.ToString, 10))
        Catch
            Return ("INVALID")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Converts the token's type from a string
    '
    '
    Private Function typeFromString_(ByVal strType As String) As qbTokenType.qbTokenType.ENUtokenType
        Select Case UCase(Trim(strType))
            Case "AMPERSAND" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersand
            Case "AMPERSANDSUFFIX" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersandSuffix
            Case "APOSTROPHE" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeApostrophe
            Case "COLON" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon
            Case "COMMA" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma
            Case "CURRENCY" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeCurrency
            Case "EXCLAMATION" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeExclamation
            Case "IDENTIFIER" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier
            Case "INVALID" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid
            Case "NEWLINE" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline
            Case "NULL" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeNull
            Case "OPERATOR" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeOperator
            Case "PARENTHESIS" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeParenthesis
            Case "PERCENT" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypePercent
            Case "PERIOD" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypePeriod
            Case "POUND" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypePound
            Case "SEMICOLON" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeSemicolon
            Case "STRING" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeString
            Case "UNSIGNEDINTEGER" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger
            Case "UNSIGNEDREALNUMBER" : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedRealNumber
            Case Else : Return qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid
        End Select
    End Function

End Class
