Option Strict On

Friend Class parseRule

    ' *****************************************************************
    ' *                                                               *
    ' * parseRule                                                     *
    ' *                                                               *
    ' *                                                               *
    ' * This stateless class provides some simple tools for parsing   *
    ' * the rule format of condition,action:comments.                 *
    ' *                                                               *
    ' *                                                               *
    ' *****************************************************************

    ' -----------------------------------------------------------------
    ' Get comment part of the rule
    '
    '
    Public Shared Function getComment(ByVal strRule As String) As String
        Dim strCondition As String
        Dim objAction As Object
        Dim strComment As String
        parseRule(strRule, strCondition, objAction, strComment)
        Return (strComment)
    End Function

    ' -----------------------------------------------------------------
    ' Get condition part of the rule
    '
    '
    Public Shared Function getCondition(ByVal strRule As String) As String
        Dim strCondition As String
        Dim objAction As Object
        Dim strComment As String
        parseRule(strRule, strCondition, objAction, strComment)
        Return (strCondition)
    End Function

    ' -----------------------------------------------------------------
    ' Parse the rule
    '
    '
    ' This method parses the displayed format
    '
    '
    '      condition,action: comment
    '
    '
    ' --- Get condition
    Public Overloads Shared Sub parseRule(ByVal strRule As String, _
                                    ByRef strCondition As String)
        Dim objAction As Object
        Dim strComment As String
        parseRule(strRule, strCondition, objAction, strComment)
    End Sub
    ' --- Get condition and action
    Public Overloads Shared Sub parseRule(ByVal strRule As String, _
                                    ByRef strCondition As String, _
                                    ByRef objAction As Object)
        Dim strComment As String
        parseRule(strRule, strCondition, objAction, strComment)
    End Sub
    ' --- Get condition, action and comment
    Public Overloads Shared Sub parseRule(ByVal strRule As String, _
                                          ByRef strCondition As String, _
                                          ByRef objAction As Object, _
                                          ByRef strComment As String)
        Dim objScanner As qbScanner.qbScanner
        parseRule(strRule, strCondition, objAction, strComment, objScanner)
        If Not (objScanner Is Nothing) Then objScanner.dispose()
    End Sub
    ' --- Get condition, action and comment: return scanner
    Public Overloads Shared Sub parseRule(ByVal strRule As String, _
                                          ByRef strCondition As String, _
                                          ByRef objAction As Object, _
                                          ByRef strComment As String, _
                                          ByRef objScanner As qbScanner.qbScanner)
        strCondition = ""
        objAction = Nothing
        strComment = ""
        Try
            objScanner = New qbScanner.qbScanner
            With objScanner
                .SourceCode = strRule
                .scan()
                Dim intIndex1 As Integer
                For intIndex1 = 1 To .TokenCount
                    If .sourceMid(intIndex1, 1) = "," Then Exit For
                Next intIndex1
                strCondition = .sourceMid(1, intIndex1 - 1)
                intIndex1 += 1
                Dim intIndex2 As Integer = intIndex1
                For intIndex1 = intIndex1 To .TokenCount
                    If .sourceMid(intIndex1, 1) = ":" Then Exit For
                Next intIndex1
                objAction = ""
                If intIndex1 > intIndex2 Then
                    objAction = .sourceMid(intIndex2, intIndex1 - intIndex2)
                End If
                Try
                    objAction = CSng(objAction)
                Catch : End Try
                strComment = .sourceMid(intIndex1 + 1, .TokenCount - intIndex1)
            End With
        Catch
            utilities.utilities.errorHandler("Can't create a scanner or scan rule: " & _
                                             Err.Number & " " & Err.Description, _
                                             "parseRule", "parseRule")
            Return
        End Try
    End Sub

End Class
