Option Strict On

Imports utilities

' **************************************************************************
' *                                                                        *
' * regex2Explanation     Regular expression to explanation                *
' *                                                                        *
' *                                                                        *
' * This class exposes one Public method, regex2Explanation, which produces*
' * the explanation of many regular expressions in an outline format suit- *
' * able for display in a monospaced font such as Courier New. See the     *
' * comment block of regex2Explanation for details.                        *
' *                                                                        *
' *                                                                        *
' * "Spacing as writing is the becoming-absent and the becoming-unconscious*
' * of the subject. By the movement of its drift/derivation, the emancipa- *
' * tion of the sign constitutes in return the desire of presence. That    *
' * becoming-or that drift/derivation-does not befall the subject which    *
' * would choose it or would passively let itself be drawn along by it.    *
' * As the subject's relation with its own death, this becoming is the     *
' * constitution of subjectivity. On all levels of life's organization,    *
' * that is to say, of the *economy of death*. All graphemes are of a      *
' * testamentary essence. And the original absence of the subject in writ- *
' * ing is also the absence of the thing or referent".                     *
' *                                                                        *
' *                     - Jacques Derrida, OF GRAMMATOLOGY                 *
' *                                                                        *
' *                                                                        *
' **************************************************************************
Public Class regex2Explanation

    ' ----------------------------------------------------------------------
    ' Explain regular expression
    '
    '
    ' This method produces an English explanation of many generic
    ' regular expressions by parsing the regex according to generic rules.
    '
    ' The following generic syntax is used to parse the regular expression.
    '
    '
    '      regex := sequenceFactor [ regex ]
    '      sequenceFactor := alternationFactor alternationRHS
    '      alternationFactor := ( postfixFactor [ postfixOp ] ) | zeroOperandOp
    '      alternationRHS := STROKE sequenceFactor
    '      postfixFactor := string | charset | ( regex )
    '      string := logicalChar [ string ]
    '      postfixOp := ASTERISK | PLUS | repeater
    '      repeater := LEFT_BRACE [ INTEGER ] [ COMMA ] [ INTEGER ] RIGHT_BRACE
    '      zeroOperandOp := CARAT | DOLLAR_SIGN
    '      charset := LEFT_BRACKET charsetExpression RIGHT_BRACKET
    '      charsetExpression := charsetRange charsetExpression
    '      charsetRange := logicalChar [ DASH logicalChar ]
    '      logicalChar := ordinaryChar | hexSequence | escapeSequence
    '      ordinaryChar := nonDigit | digit ' Where an ORDINARY CHAR is [^0-9\*\+\^\$\\\-\{\}\[\]\|] 
    '      nonDigit := NONDIGIT ' Where a NONDIGIT is [^0-9\*\+\^\$\\\-\{\}\[\]\|] 
    '      digit := DIGIT ' Where a DIGIT is [0-9] 
    '      hexSequence :=  HEXSEQUENCE ' Where a hex sequence is \\x[0123456789ABCDEFabcdef]+
    '      escapeSequence := ESCAPESEQUENCE ' Where an esc sequence is  \\[\*\+\^\$\\\-\{\}\[\]\|]                 
    '
    '
    ' The explanation will be a multiple-line outline of the effect of the
    ' regular expression that will be suitable for display using a monospace
    ' font such as Courier New.
    '
    '
    Private Enum ENUregexTokenType
        tokenType_stroke
        tokenType_asterisk
        tokenType_plus
        tokenType_leftBrace
        tokenType_integer
        tokenType_comma
        tokenType_rightBrace
        tokenType_escSequence
    End Enum
    Private Structure TYPregexScanned
        Dim strTokenValue As String
        Dim enuTokenType As ENUregexTokenType
        Dim intTokenIndex As Integer
        Dim intTokenLength As Integer
    End Structure
    Public Shared Function regex2Explanation(ByVal strRegex As String) As String
        Dim usrRegexScanned() As TYPregexScanned
        If Not regex2Explanation_scan_(strRegex, usrRegexScanned) Then
            Dim strOutstring As String
            strOutstring = "Error in scanning the regular expression"
            Return strOutstring
        End If
        Dim intIndex1 As Integer = 1
        Return utilities.formatOutline _
               (regex2Explanation_parseExplain_(strRegex, _
                                                usrRegexScanned, _
                                                intIndex1, _
                                                UBound(usrRegexScanned)))
    End Function

    ' ----------------------------------------------------------------------
    ' Parse and explain the regular expression on behalf of regex2Explanation
    '
    '
    Private Shared Function regex2Explanation_parseExplain_ _
                            (ByVal strRegex As String, _
                             ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer) _
            As String
        Dim strOutstring As String = _
            "The regular expression " & _
            regex2Explanation_parseExplain__regex2Display_(strRegex) & " " & _
            "is matched by a string that follows these rules."
        Dim strError As String
        If Not regex2Explanation_parseExplain__regex_(usrScanned, _
                                                      intIndex, _
                                                      intEndIndex, _
                                                      strOutstring, _
                                                      "") Then
            strOutstring &= vbNewLine & vbNewLine & "No explanation is available"
        End If
        Return strOutstring
    End Function

    ' ----------------------------------------------------------------------
    ' alternationFactor := ( postfixFactor [ postfixOp ] ) | zeroOperandOp
    '
    '
    ' This method changes the order in which the zeroOperandOp and the
    ' sequence "postfixFactor postfixOp" are evaluated, since it's slightly
    ' more efficient to test for the zeroOperandOp first, and its left handle
    ' (carat or dollar) is distinct from that of the postfixFactor: the left
    ' handle of a postfixFactor cannot include a carat or dollar sign.
    '
    '
    Private Shared Function regex2Explanation_parseExplain__alternationFactor_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String, _
                             ByRef strOutlineNumber As String) As Boolean
        If regex2Explanation_parseExplain__zeroOperandOp_(usrScanned, _
                                                          intIndex, _
                                                          intEndIndex, _
                                                          strOutstring, _
                                                          strOutlineNumber) Then
            Return True
        End If
        If Not regex2Explanation_parseExplain__postfixFactor_(usrScanned, _
                                                              intIndex, _
                                                              intEndIndex, _
                                                              strOutstring) Then
            strOutstring &= vbNewLine & vbNewLine & _
                            regex2Explanation_parseExplain__mkErrorMessage_ _
                            (usrScanned(intIndex).intTokenIndex, _
                             "Expected but did not find a carat, dollar sign, " & _
                             "string or character set")
            Return False
        End If
        regex2Explanation_parseExplain__postfixOp_(usrScanned, _
                                                   intIndex, _
                                                   intEndIndex, _
                                                   strOutstring)
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' alternationRHS := STROKE sequenceFactor
    '
    '
    Private Shared Function regex2Explanation_parseExplain__alternationRHS_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String, _
                             ByRef strOutlineNumber As String) As Boolean
        If Not regex2Explanation_parseExplain__postfixFactor_(usrScanned, _
                                                              intIndex, _
                                                              intEndIndex, _
                                                              strOutstring) Then
            strOutstring &= vbNewLine & vbNewLine & _
                            regex2Explanation_parseExplain__mkErrorMessage_ _
                            (usrScanned(intIndex).intTokenIndex)
            Return False
        End If
        regex2Explanation_parseExplain__postfixOp_(usrScanned, _
                                                   intIndex, _
                                                   intEndIndex, _
                                                   strOutstring)
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' Attach output
    '
    '
    Private Overloads Shared Sub regex2Explanation_parseExplain__attachOutput_ _
                                 (ByRef strLHS As String, _
                                  ByVal strRHS As String, _
                                  ByRef strDewey As String)

    End Sub
    Private Overloads Shared Sub regex2Explanation_parseExplain__attachOutput_ _
                                 (ByRef strLHS As String, _
                                  ByVal strRHS As String, _
                                  ByRef strDewey As String, _
                                  ByVal booNewlevel As Boolean)
        Dim booLHSmultiline As Boolean = (InStr(strLHS, vbNewLine) <> 0)
        Dim booLHSnull As Boolean = (strLHS = "")
        Dim booRHSmultiline As Boolean = (InStr(strRHS, vbNewLine) <> 0)
        Dim booRHSnull As Boolean = (strLHS = "")
        If booRHSnull Then
            errorHandler("Internal programming error: RHS is a null string", _
                         Name, "regex2Explanation_parseExplain__attachOutput_", _
                         "No change made")
            Return
        End If
        If booRHSmultiline OrElse booNewlevel Then
            ' --- Attach as new paragraph group 
            If strLHS <> "" Then strLHS &= vbNewLine
            Dim strNewDewey As String = "1"
            If strLHS <> "" Then strNewDewey = word(line(strLHS, lines(strLHS)), 1)
            If booNewlevel Then
                If strNewDewey <> "" Then strNewDewey &= "."
                strNewDewey = strNewDewey & "1"
                strLHS &= strNewDewey & _
                          vbNewLine & _
                          strNewDewey & "." & _
                          Replace(strRHS, vbNewLine, vbNewLine & strNewDewey & ".")
            Else
                Dim intIndex1 As Integer
                Dim intIndex2 As Integer = items(strNewDewey)
                Dim strRHSnew As String
                For intIndex1 = 1 To lines(strRHS)
                    strRHSnew = incrementDewey(word(line(strRHS, intIndex1), 1), intIndex2)
                Next intIndex1
            End If
        End If
        Return
    End Sub

    ' ----------------------------------------------------------------------
    ' Check for token on behalf of regex2Explanation_parseExplain__
    '
    '
    Private Shared Function regex2Explanation_parseExplain__checkToken_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByVal enuType As ENUregexTokenType, _
                             ByVal strExpected As String) As Boolean

    End Function

    ' ----------------------------------------------------------------------
    ' Make the error message on behalf of regex2Explanation_parseExplain__
    '
    '
    ' --- Basic call
    Private Overloads Shared Function regex2Explanation_parseExplain__mkErrorMessage_ _
                                      (ByVal intIndex As Integer) _
            As String
        Return regex2Explanation_parseExplain__mkErrorMessage_(intIndex, "")
    End Function
    ' --- Additional comments are available
    Private Overloads Shared Function regex2Explanation_parseExplain__mkErrorMessage_ _
                                      (ByVal intIndex As Integer, _
                                       ByVal strComment As String) _
            As String
        Return "Error in the regular expression at character " & intIndex & _
               CStr(IIf(strComment <> "", ": ", "")) & _
               strComment
    End Function

    ' ----------------------------------------------------------------------
    ' Tell caller whether one string is in the input string
    '
    '
    Private Shared Function regex2Explanation_parseExplain__oneString_ _
                            (ByVal strInstring As String) _
            As Boolean

    End Function

    ' ----------------------------------------------------------------------
    ' postfixFactor := STRING | charset | ( regex )
    '
    '
    Private Shared Function regex2Explanation_parseExplain__postfixFactor_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String) As Boolean

    End Function

    ' ----------------------------------------------------------------------
    ' postfixOp := ASTERISK | PLUS | repeater
    '
    '
    Private Shared Function regex2Explanation_parseExplain__postfixOp_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String) As Boolean

    End Function

    ' ----------------------------------------------------------------------
    ' regex := sequenceFactor [ regex ]
    '
    '
    Private Shared Function regex2Explanation_parseExplain__regex_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String, _
                             ByRef strOutlineNumber As String) As Boolean
        If Not regex2Explanation_parseExplain__sequenceFactor_(usrScanned, _
                                                                intIndex, _
                                                                intEndIndex, _
                                                                strOutstring, _
                                                                strOutlineNumber) Then
            strOutstring &= vbNewLine & vbNewLine & _
                            regex2Explanation_parseExplain__mkErrorMessage_ _
                            (usrScanned(intIndex).intTokenIndex)
            Return False
        End If
        Dim strOutstringNext As String
        If regex2Explanation_parseExplain__regex_(usrScanned, _
                                                  intIndex, _
                                                  intEndIndex, _
                                                  strOutstringNext, _
                                                  strOutlineNumber) Then
            regex2Explanation_parseExplain__attachOutput_(strOutstring, _
                                                          strOutstringNext, _
                                                          strOutlineNumber, _
                                                          True)
        End If
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' Convert the regular expression to displayable format on behalf of
    ' regex2Explanation_parseExplain_
    '
    '
    Private Shared Function regex2Explanation_parseExplain__regex2Display_ _
                            (ByVal strRegex As String) _
            As String
        If Len(strRegex) > 32 Then
            Return vbNewLine & vbNewLine & _
                   strRegex & _
                   vbNewLine & vbNewLine
        End If
        Return enquote(strRegex)
    End Function

    ' ----------------------------------------------------------------------
    ' sequenceFactor := alternationFactor alternationRHS
    '
    '
    ' This method generates:
    '
    '
    '      <number> The following alternatives: <list>
    '
    '
    Private Shared Function regex2Explanation_parseExplain__sequenceFactor_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String, _
                             ByRef strOutlineNumber As String) As Boolean
        If Not regex2Explanation_parseExplain__alternationFactor_(usrScanned, _
                                                                  intIndex, _
                                                                  intEndIndex, _
                                                                  strOutstring, _
                                                                  strOutlineNumber) Then
            strOutstring &= vbNewLine & vbNewLine & _
                            regex2Explanation_parseExplain__mkErrorMessage_ _
                            (usrScanned(intIndex).intTokenIndex)
            Return False
        End If
        If regex2Explanation_parseExplain__alternationRHS_(usrScanned, _
                                                           intIndex, _
                                                           intEndIndex, _
                                                           strOutstring, _
                                                           strOutlineNumber) Then
            strOutstring = "The following alternatives" & strOutstring
        End If
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' zeroOperandOp := CARAT | DOLLAR_SIGN
    '
    '
    Private Shared Function regex2Explanation_parseExplain__zeroOperandOp_ _
                            (ByVal usrScanned() As TYPregexScanned, _
                             ByRef intIndex As Integer, _
                             ByVal intEndIndex As Integer, _
                             ByRef strOutstring As String, _
                             ByRef strOutlineNumber As String) As Boolean
        If regex2Explanation_parseExplain__checkToken_(usrScanned, _
                                                       intIndex, _
                                                       intEndIndex, _
                                                       ENUregexTokenType.tokenType_metaCharacter, _
                                                       "^") Then
            regex2Explanation_parseExplain__attachOutput_(strOutlineNumber, _
                                                          "The start of a line", _
                                                          strOutlineNumber)
        End If
        If regex2Explanation_parseExplain__alternationRHS_(usrScanned, _
                                                           intIndex, _
                                                           intEndIndex, _
                                                           strOutstring, _
                                                           strOutlineNumber) Then
            strOutstring = "The following alternatives" & strOutstring
        End If
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' Scan the regular expression on behalf of regex2Explanation
    '
    '
    Private Shared Function regex2Explanation_scan_(ByVal strRegex As String, _
                                                    ByRef usrScanned() As TYPregexScanned) _
            As Boolean
        Try
            ReDim usrScanned(0)
        Catch
            errorHandler("Not able to allocate the scan table: " & _
                         Err.Number & " " & Err.Description, _
                         "utilities", _
                         "regex2Explanation_scan_", _
                         "Returning False")
            Return False
        End Try
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim intLength As Integer = Len(strRegex)
        Dim intTokenLength As Integer
        Do While intIndex1 <= intLength
            intIndex2 = intIndex1
            Do While intIndex2 <= intLength
                intIndex2 = verify(strRegex & "*", "\*+-^$(){}[]", intIndex2, True)
                If intIndex2 < intLength AndAlso Mid(strRegex, intIndex2, 1) = "\" Then
                    intIndex2 += 2
                End If
            Loop
            If intIndex2 > intIndex1 Then
                intTokenLength = intIndex2 - intIndex1
                If Not regex2Explanation_scan__expandScan_(usrScanned, _
                                                           Mid(strRegex, _
                                                               intIndex1, _
                                                               intTokenLength), _
                                                           ENUregexTokenType.tokenType_string, _
                                                           intIndex1, _
                                                           intTokenLength) Then
                    Return False
                End If
            End If
            If intIndex2 > intLength Then Exit Do
            If Not regex2Explanation_scan__expandScan_(usrScanned, _
                                                       Mid(strRegex, intIndex2, 1), _
                                                       ENUregexTokenType.tokenType_metaCharacter, _
                                                       intIndex2, _
                                                       1) Then
                Return False
            End If
            intIndex1 += 1
        Loop
        Return True
    End Function

    ' ----------------------------------------------------------------------
    ' Expand the scan table on behalf of regex2Explanation_scan_
    '
    '
    Private Shared Function regex2Explanation_scan__expandScan_ _
                            (ByRef usrScanned() As TYPregexScanned, _
                             ByVal strTokenValue As String, _
                             ByVal enuTokenType As ENUregexTokenType, _
                             ByVal intTokenIndex As Integer, _
                             ByVal intTokenLength As Integer) _
            As Boolean
        Try
            ReDim Preserve usrScanned(UBound(usrScanned) + 1)
        Catch
            errorHandler("Not able to expand the scan table: " & _
                         Err.Number & " " & Err.Description, _
                         "utilities", _
                         "regex2Explanation_scan_", _
                         "Returning False")
            Return False
        End Try
        With usrScanned(UBound(usrScanned))
            .enuTokenType = enuTokenType
            .strTokenValue = strTokenValue
            .intTokenIndex = intTokenIndex
            .intTokenLength = intTokenLength
        End With
        Return True
    End Function

End Class
