Option Strict

' *********************************************************************
' *                                                                   *
' * qbTokenType     Define the scan token types                       *
' *                                                                   *
' *********************************************************************

Public Class qbTokenType

    Public Const TOKENTYPECOUNT As Integer = 20             ' Number of tokens  
    Public Const TOKENTYPECOUNT_ACTUAL As Integer = 17      ' Number of "real" tokens
    Public Enum ENUtokenType
        ' --- The actual tokens are those found by the scanner
        tokenTypeAmpersand = 1                    ' Ampersand  
        tokenTypeApostrophe = 2                   ' Single quote
        tokenTypeColon = 3                        ' Colon
        tokenTypeComma = 4                        ' Comma
        tokenTypeIdentifier = 5                   ' Identifier
        tokenTypeNewline =  6                     ' New line
        tokenTypeOperator = 7                     ' Operator
        tokenTypeParenthesis = 8                  ' Left or right parenthesis
        tokenTypeSemicolon = 9                    ' Semicolon
        tokenTypeString = 10                      ' String
        tokenTypeUnsignedInteger = 11             ' Unsigned integer (sign is always an op)
        tokenTypeUnsignedRealNumber = 12          ' Unsigned real number (sign is always an op)
        tokenTypePercent = 13                     ' Percent
        tokenTypeExclamation = 14                 ' Exclamation
        tokenTypePound = 15                       ' Pound sign
        tokenTypeCurrency = 16                    ' Dollar sign
        tokenTypePeriod = 17                      ' Dollar sign
        ' --- These tokens are special values or synthesized by the scanner
        tokenTypeNull = 18                        ' Null value
        tokenTypeInvalid = 19                     ' Invalid
        tokenTypeAmpersandSuffix = 20             ' Ampersand, preceded by an identifier             
    End Enum

End Class
