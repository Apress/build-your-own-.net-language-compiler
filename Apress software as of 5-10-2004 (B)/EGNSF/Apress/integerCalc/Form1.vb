Option Strict

Imports utilities  
Imports windowsUtilities  

' *********************************************************************
' *                                                                   *
' * integerCalc                                                       *
' *                                                                   *
' *********************************************************************

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub
    Friend WithEvents txtExpression As System.Windows.Forms.TextBox
    Friend WithEvents cmdEvaluate As System.Windows.Forms.Button
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.Container

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.txtValue = New System.Windows.Forms.TextBox
Me.cmdEvaluate = New System.Windows.Forms.Button
Me.txtExpression = New System.Windows.Forms.TextBox
Me.cmdClose = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'txtValue
'
Me.txtValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtValue.Location = New System.Drawing.Point(10, 92)
Me.txtValue.Name = "txtValue"
Me.txtValue.Size = New System.Drawing.Size(480, 30)
Me.txtValue.TabIndex = 1
Me.txtValue.Text = "0"
Me.txtValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'cmdEvaluate
'
Me.cmdEvaluate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdEvaluate.Location = New System.Drawing.Point(10, 46)
Me.cmdEvaluate.Name = "cmdEvaluate"
Me.cmdEvaluate.Size = New System.Drawing.Size(480, 37)
Me.cmdEvaluate.TabIndex = 1
Me.cmdEvaluate.Text = "Evaluate"
'
'txtExpression
'
Me.txtExpression.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtExpression.Location = New System.Drawing.Point(10, 9)
Me.txtExpression.Name = "txtExpression"
Me.txtExpression.Size = New System.Drawing.Size(480, 30)
Me.txtExpression.TabIndex = 0
Me.txtExpression.Text = "0"
Me.txtExpression.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(10, 129)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(480, 37)
Me.cmdClose.TabIndex = 1
Me.cmdClose.Text = "Close"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
Me.ClientSize = New System.Drawing.Size(496, 198)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.txtValue)
Me.Controls.Add(Me.txtExpression)
Me.Controls.Add(Me.cmdEvaluate)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "integerCalc"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    ' ***** Operators *****
    Private Enum ENUoperator
        push
        add
        subtract
        multiply
        divide
    End Enum        
        
    ' ***** Scanner's data structure *****
    Private Enum ENUscannedType
        number
        operator
        leftParenthesis
        rightParenthesis
    End Enum    
    Private Structure TYPscanned
        Dim enuType As ENUscannedType   ' Token type
        Dim objValue As Object          ' Token value: operator enum or integer
        Dim intStartIndex As Integer    ' Token start index
        Dim intLength As Integer        ' Length
    End Structure    
    
    ' ***** Reverse-polish notation code *****
    Private Structure TYPrpn
        Dim enuOp As ENUoperator
        Dim intOperand As Integer       ' Only present for the push operator
    End Structure    
    
    ' ***** Trace controls *****
    Private WithEvents CMDmore As Button
    Private CTLstatusDisplay As statusDisplay
    
    ' ***** Utilities *****
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities
    Private Shared OBJcollectionUtilities As collectionUtilities.collectionUtilities

    ' ***** Constants *****
    Private Const WIDTH_TOLERANCE As Single = 0.9
    Private Const HEIGHT_TOLERANCE As Single = 0.5

#End Region ' Form data

#Region " Form events "

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Dispose
        End
    End Sub

    Private Sub cmdEvaluate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEvaluate.Click
        With txtValue
            .Text = "Evaluating": .Refresh
            .Text = evaluateExpression(txtExpression.Text)
        End With        
    End Sub
    
    Private Sub cmdMore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With CMDmore
            If .Text = "More" Then
                showMore
            ElseIf .Text = "Less" Then
                showLess
            Else        
                MsgBox("Programming error: unexpected caption"): Dispose: End                        
            End If            
        End With        
    End Sub    

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            OBJcollectionUtilities = New collectionUtilities.collectionUtilities
        Catch
            _OBJutilities.errorHandler("Cannot create collection utilities: " & _
                                       Err.Number & " " & Err.Description, _
                                       Name, "Form1_Load")
            End
        End Try
        registry2Form()
        With cmdClose
            .Width = (.Width \ 2) - _OBJwindowsUtilities.Grid
            .Left = .Width + _OBJwindowsUtilities.Grid * 3
        End With
        mkMoreButton()
        ClientSize = New Size(cmdClose.Left + cmdClose.Width + 8, _
                              cmdClose.Top + cmdClose.Height + 8)
        CenterToScreen()
    End Sub

#End Region ' Form events 
    
#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Evaluate the expression
    '
    '
    Private Function evaluateExpression(ByVal strExpression As String) As String
        Dim usrScanned() As TYPscanned
        If statusDisplayInEffect Then
            With CTLstatusDisplay
                .clear: .Expression = strExpression
            End With            
        End If        
        If Not scanExpression(strExpression, usrScanned) Then 
            Return("Could not scan expression")
        End If            
        Dim usrRPN() As TYPrpn
        If Not parseExpression(usrScanned, usrRPN) Then 
            Return("Could not parse expression")
        End If            
        Return(interpretExpression(usrRPN))
    End Function    
    
    ' ----------------------------------------------------------------
    ' Restore form settings from the Registry
    '
    '
    Private Sub form2Registry
        Try
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        txtExpression.Name, _
                        txtExpression.Text)
        Catch
            MsgBox("Could not save Registry setting: " & _
                   Err.Number & " " & Err.Description)
        End Try        
    End Sub   
    
    ' ----------------------------------------------------------------
    ' Make the More button
    '
    '
    Private Sub mkMoreButton
        Try
            CMDmore = New Button
            Controls.Add(CMDmore)
            AddHandler CMDmore.Click, AddressOf cmdMore_Click
            With CMDmore
                .Top = cmdClose.Top
                .Width = cmdClose.Width \ 2
                .Left = txtValue.Left
                .Height = cmdClose.Height
                .Text = "More"
            End With        
        Catch
            _OBJutilities.errorHandler("Cannot create a More button: " & _
                                       Err.Number & " " & Err.Description)
        End Try        
    End Sub    
    
    ' ----------------------------------------------------------------
    ' Restore form settings from the Registry
    '
    '
    Private Sub registry2Form
        Try
            txtExpression.Text = GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            txtExpression.Name, _
                                            txtExpression.Text)
        Catch
            MsgBox("Could not obtain Registry setting: " & _
                   Err.Number & " " & Err.Description)
        End Try        
    End Sub 
    
    ' ----------------------------------------------------------------
    ' Show simple screen
    '
    '
    Private Sub showLess
        Dim booSave As Boolean = Visible
        Visible = False: Refresh
        If Not (CTLstatusDisplay Is Nothing) Then
            CTLstatusDisplay.Visible = False
        End If        
        CMDmore.Text = "More"
        Width = _OBJwindowsUtilities.controlRight(cmdClose) _  
                + _
                _OBJwindowsUtilities.Grid * 2
        Height = cmdClose.Top + cmdClose.Height + 32
        CenterToScreen
        Visible = booSave: Refresh
    End Sub    

    ' ----------------------------------------------------------------
    ' Show complex screen
    '
    '
    Private Sub showMore
        Dim booSave As Boolean = Visible
        Visible = False: Refresh
        CMDmore.Text = "Less"
        ClientSize = New Size(CInt(_OBJwindowsUtilities.screenWidth * WIDTH_TOLERANCE), _
                              CInt(_OBJwindowsUtilities.screenHeight * HEIGHT_TOLERANCE))
        If (CTLstatusDisplay Is Nothing) Then
            CTLstatusDisplay = New statusDisplay(Me, _
                                                 0, _
                                                 _OBJwindowsUtilities.controlBottom(CMDmore) _
                                                 + _
                                                 _OBJwindowsUtilities.Grid, _
                                                 ClientSize.Width - _OBJwindowsUtilities.Grid * 2, _
                                                 ClientSize.Height)
        Else
            CTLstatusDisplay.Visible = True
        End If
        Height = CTLstatusDisplay.Top + CTLstatusDisplay.Height + 32
        CenterToScreen()
        Visible = booSave : Refresh()
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Return True if a status display is active
    '
    '
    Private Function statusDisplayInEffect As Boolean
        If (CTLstatusDisplay Is Nothing) Then Return(False)
        Return(CTLstatusDisplay.Visible)
    End Function    
   
   
#Region " Scanner (lexical analyzer) "

    ' -----------------------------------------------------------------
    ' Scan the expression into individual tokens
    '
    '
    Private Function scanExpression(ByVal strExpression As String, _
                                    ByRef usrScanned() As TYPscanned) As Boolean
        Try
            Redim usrScanned(0) 
        Catch
            MsgBox("Cannot allocate the scan table: " & _
                   Err.Number & " " & Err.Description)
            Return(False)
        End Try        
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer  
        Dim intLength As Integer = Len(strExpression)
        Dim intNext As Integer
        Dim objValue As Object
        dim strnext As String
        Do While intIndex1 <= intLength
            ' --- Find next token start
            For intIndex1 = intIndex1 To intLength
                If Mid(strExpression, intIndex1, 1) <> " " Then Exit For
            Next intIndex1
            If intIndex1 > intLength Then Exit Do
            ' --- Parse next token
            Try
                Redim Preserve usrScanned(UBound(usrScanned) + 1)
            Catch
                MsgBox("Cannot expand the scan table: " & _
                       Err.Number & " " & Err.Description)
                Return(False)
            End Try            
            With usrScanned(UBound(usrScanned))
                objValue = Nothing
                .enuType = ENUscannedType.operator ' Assume
                .intStartIndex = intIndex1
                .intLength = 1
                Select Case Mid(strExpression, intIndex1, 1)
                    Case "+": objValue = ENUoperator.add
                    Case "-": objValue = ENUoperator.subtract
                    Case "*": objValue = ENUoperator.multiply
                    Case "/": objValue = ENUoperator.divide
                    Case "(": .enuType = ENUscannedType.leftParenthesis
                              objValue = "("
                    Case ")": .enuType = ENUscannedType.rightParenthesis
                              objValue = ")"
                    Case Else:
                        .enuType = ENUscannedType.number  
                        objValue = ""
                        .intLength = 0
                        For intIndex2 = intIndex1 To intLength
                            strNext = Mid(strExpression, intIndex2, 1)
                            intNext = AscW(strNext)
                            If intNext < AscW("0") OrElse intNext > AscW("9") Then Exit For
                            .intLength += 1 
                            objValue = CStr(objValue) & strNext
                        Next intIndex2      
                        If .intLength = 0 Then
                            Msgbox("Cannot parse input expression starting at position " & intIndex1)
                            Return(False)
                        End If                                  
                End Select                        
                .objValue = objValue
                If statusDisplayInEffect Then
                    CTLstatusDisplay.scanUpdate(.intStartIndex, .intLength, .enuType.ToString)
                End If                
                intIndex1 += .intLength
            End With            
        Loop                 
        Return(True)
    End Function  
     
#End Region  ' Scanner (lexical analyzer)       
    
#Region " Parser "

    ' -----------------------------------------------------------------
    ' Parse the expression and compile to reverse Polish notation
    '
    '
    Private Function parseExpression(ByRef usrScanned() As TYPscanned, _
                                     ByRef usrRPN() As TYPrpn) As Boolean
        Try
            Redim usrRPN(0)  
        Catch
            MsgBox("Can't allocate RPN object code array: " & _
                   Err.Number & " " & Err.Description)
            Return(False)                   
        End Try                                             
        Dim intIndex1 As Integer = 1
        If Not expression(intIndex1, _
                          usrScanned, _
                          usrRPN, _
                          UBound(usrScanned), _
                          0) Then Return(False)
        If intIndex1 <= UBound(usrScanned) Then
            MsgBox("Cannot parse complete expression"): Return(False)
        End If                                  
        Return(True)
    End Function    
    
    ' -----------------------------------------------------------------
    ' expression := addFactor [ expressionRHS ]
    '
    '
    Private Function expression(ByRef intIndex As Integer, _
                                ByRef usrScanned() As TYPscanned, _
                                ByRef usrRPN() As TYPrpn, _
                                ByVal intEndIndex As Integer, _
                                ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex                                
        If Not addFactor(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1) Then Return(False)
        expressionRHS(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1)     
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "expression", _
                                         intLevel)
        End If        
        Return(True)                           
    End Function                                
    
    ' ----------------------------------------------------------------
    ' expressionRHS = addOp addFactor [ expressionRHS ]
    '
    '
    Private Function expressionRHS(ByRef intIndex As Integer, _
                                   ByRef usrScanned() As TYPscanned, _
                                   ByRef usrRPN() As TYPrpn, _
                                   ByVal intEndIndex As Integer, _
                                   ByVal intLevel As Integer) As Boolean
        Dim enuOp As ENUoperator                                  
        Dim intIndex1 As Integer = intIndex
        If Not addOp(intIndex, usrScanned, enuOp, intEndIndex, intLevel + 1) Then Return(False)
        If Not addFactor(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1) Then 
            MsgBox("Syntax error at or near " & intIndex)
            Return(False)
        End If            
        If Not genCode(usrRPN, enuOp, 0, "Factor") Then Return(False)
        expressionRHS(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1)
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "expressionRHS", _
                                         intLevel)
        End If        
        Return(True)
    End Function    
    
    ' ----------------------------------------------------------------
    ' addOp := +|-
    '
    '
    Private Function addOp(ByRef intIndex As Integer, _
                           ByRef usrScanned() As TYPscanned, _
                           ByRef enuOp As ENUoperator, _
                           ByVal intEndIndex As Integer, _
                           ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex                           
        If checkToken(intIndex, usrScanned, "add", intEndIndex) Then
            enuOp = ENUoperator.add 
        ElseIf checkToken(intIndex, usrScanned, "subtract", intEndIndex) Then
            enuOp = ENUoperator.subtract 
        Else
            Return(False)            
        End If        
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "addOp", _
                                         intLevel)
        End If     
        Return(True)   
    End Function 
    
    ' ----------------------------------------------------------------
    ' addFactor = term [addFactorRHS]
    '
    '
    Private Function addFactor(ByRef intIndex As Integer, _
                               ByRef usrScanned() As TYPscanned, _
                               ByRef usrRPN() As TYPrpn, _
                               ByVal intEndIndex As Integer, _
                               ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex                               
        If Not term(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1) Then Return(False)
        addFactorRHS(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1)       
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "addFactor", _
                                         intLevel)
        End If        
        Return(True)                         
    End Function    
    
    ' ----------------------------------------------------------------
    ' addFactorRHS = mulOp term [ addFactorRHS ] 
    '
    '
    Private Function addFactorRHS(ByRef intIndex As Integer, _
                                  ByRef usrScanned() As TYPscanned, _
                                  ByRef usrRPN() As TYPrpn, _
                                  ByVal intEndIndex As Integer, _
                                  ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex                                  
        Dim enuOp As ENUoperator                                  
        If Not mulOp(intIndex, usrScanned, enuOp, intEndIndex, intLevel + 1) Then Return(False)
        If Not term(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1) Then Return(False)
        If Not genCode(usrRPN, enuOp, 0, "Term") Then Return(False)
        addFactorRHS(intIndex, usrScanned, usrRPN, intEndIndex, intLevel + 1)                      
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "addFactorRHS", _
                                         intLevel)
        End If        
        Return(True)   
    End Function    
    
    ' ----------------------------------------------------------------
    ' mulOp := *|/
    '
    '
    Private Function mulOp(ByRef intIndex As Integer, _
                           ByRef usrScanned() As TYPscanned, _
                           ByRef enuOp As ENUoperator, _
                           ByVal intEndIndex As Integer, _
                           ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex                           
        If checkToken(intIndex, usrScanned, "multiply", intEndIndex) Then
            enuOp = ENUoperator.multiply 
        ElseIf checkToken(intIndex, usrScanned, "divide", intEndIndex) Then
            enuOp = ENUoperator.divide 
        Else
            Return(False)            
        End If        
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "mulOp", _
                                         intLevel)
        End If        
        Return(True)
    End Function 
    
    ' ----------------------------------------------------------------
    ' term := INTEGER | ( expression )
    '
    '
    Private Function term(ByRef intIndex As Integer, _
                          ByRef usrScanned() As TYPscanned, _
                          ByRef usrRPN() As TYPrpn, _
                          ByVal intEndIndex As Integer, _
                          ByVal intLevel As Integer) As Boolean
        Dim intIndex1 As Integer = intIndex                          
        Dim intIntegerValue As Integer                               
        If checkToken(intIndex, usrScanned, intIntegerValue, intEndIndex) Then  
            If Not genCode(usrRPN, _
                           ENUoperator.push, _
                           intIntegerValue, _
                           "Push term to stack") Then Return(False)
            If statusDisplayInEffect Then
                CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                             usrScanned(intIndex - 1).intStartIndex _
                                             + _
                                             usrScanned(intIndex - 1).intLength _
                                             - _
                                             usrScanned(intIndex1).intStartIndex, _
                                             "term", _
                                             intLevel)
            End If        
            Return(True)
        End If         
        If Not checkToken(intIndex, usrScanned, "(", intEndIndex) Then
            MsgBox("Invalid term found at " & intIndex)
            Return(False)
        End If           
        Dim intRPindex As Integer = findRightParenthesis(intIndex, _
                                                         usrScanned, _
                                                         intEndIndex)
        If intRPindex = 0 Then
            MsgBox("Cannot find the right parenthesis that corresponds " & _
                   "to the left parenthesis at " & intIndex)
            Return(False)                   
        End If                                                                 
        If Not expression(intIndex, _
                          usrScanned, _
                          usrRPN, _
                          intRPIndex - 1, _
                          intLevel + 1) Then Return(False)            
        intIndex = intRPindex + 1
        If statusDisplayInEffect Then
            CTLstatusDisplay.parseUpdate(usrScanned(intIndex1).intStartIndex, _
                                         usrScanned(intIndex - 1).intStartIndex _
                                         + _
                                         usrScanned(intIndex - 1).intLength _
                                         - _
                                         usrScanned(intIndex1).intStartIndex, _
                                         "expression in parentheses", _
                                         intLevel)
        End If        
        Return(True)                          
    End Function                    
    
    ' ---------------------------------------------------------------
    ' Find balanced right parenthesis
    '
    '
    Private Function findRightParenthesis(ByVal intIndex As Integer, _
                                          ByRef usrScanned() As TYPscanned, _
                                          ByVal intEndIndex As Integer) As Integer
        Dim intLevel As Integer = 1
        Dim intIndex1 As Integer = intIndex
        Do While intIndex1 <= intEndIndex 
            If checkToken(intIndex1, usrScanned, "(", intEndIndex) Then
                intLevel += 1
            ElseIf checkToken(intIndex1, usrScanned, ")", intEndIndex) Then
                intLevel -= 1
                If intLevel = 0 Then Return(intIndex1 - 1)
            Else
                intIndex1 += 1                
            End If            
        Loop        
        Return(0)
    End Function                                          

    ' ---------------------------------------------------------------
    ' Check for a token: increment index if found
    '
    '
    ' --- Check for an integer
    Private Overloads Function checkToken(ByRef intIndex As Integer, _
                                          ByRef usrScanned() As TYPscanned, _
                                          ByRef intIntegerValue As Integer, _
                                          ByVal intEndIndex As Integer) _
            As Boolean   
        If intIndex > intEndIndex Then Return(False)            
        With usrScanned(intIndex)            
            If .enuType <> ENUscannedType.number Then 
                Return(False)
            End If
            Try
                intIntegerValue = CInt(.objValue)
            Catch
                Msgbox("Token at " & intIndex & " is not a valid integer")
                Return(False)
            End Try
            intIndex += 1
            Return(True)   
        End With        
    End Function       
    ' --- Check for a string op or parenthesis    
    Private Overloads Function checkToken(ByRef intIndex As Integer, _
                                          ByRef usrScanned() As TYPscanned, _
                                          ByVal strToken As String, _
                                          ByVal intEndIndex As Integer) _
            As Boolean   
        If intIndex > intEndIndex Then Return(False)            
        With usrScanned(intIndex) 
            Dim strTokenWork As String
            Try
                strTokenWork = .objValue.ToString
            Catch
                strTokenWork = CStr(.objValue)                
            End Try             
            If .enuType <> ENUscannedType.operator _
               AndAlso _
               .enuType <> ENUscannedType.leftParenthesis _
               AndAlso _
               .enuType <> ENUscannedType.rightParenthesis _
               OrElse _
               LCase(strTokenWork) <> LCase(strToken) Then 
                Return(False)
            End If
            intIndex += 1
            Return(True)   
        End With        
    End Function      
    
    ' -----------------------------------------------------------------
    ' Code generator
    '
    '
    Private Function genCode(ByRef usrRPN() As TYPrpn, _
                             ByVal enuOp As ENUoperator, _
                             ByVal intOperand As Integer, _
                             ByVal strComment As String) As Boolean
        Try
            Redim Preserve usrRPN(UBound(usrRPN) + 1)
        Catch
            MsgBox("Can't expand RPN code table: " & _
                   Err.Number & " " & Err.Description)
            Return(False)                   
        End Try        
        With usrRPN(UBound(usrRPN))
            .enuOp = enuOp: .intOperand = intOperand
        End With        
        If statusDisplayInEffect Then
            CTLstatusDisplay.codeUpdate(enuOp.ToString, CStr(intOperand), strComment)
        End If        
        Return(True)
    End Function                                   
    
#End Region  ' Parser     
 
#Region " Interpreter "

    ' -----------------------------------------------------------------
    ' Scan the expression into individual tokens
    '
    '
    Private Function interpretExpression(ByRef usrRPN() As TYPrpn) As String
        Dim objStack As Stack       ' Of integers
        Try
            objStack = New Stack
        Catch
            MsgBox("Can't create stack: " & Err.Number & " " & Err.Description)
            Return("Cannot interpret expression")
        End Try        
        Dim intIndex1 As Integer
        Dim intOperand1 As Integer
        Dim intOperand2 As Integer
        For intIndex1 = 1 To UBound(USRrpn)
            With USRrpn(intIndex1)
                Try
                    Select Case .enuOp
                        Case ENUoperator.add:
                            pushStack(objStack, popStack(objStack) + popStack(objStack))
                        Case ENUoperator.subtract:
                            intOperand2 = popStack(objStack)
                            intOperand1 = popStack(objStack)
                            pushStack(objStack, intOperand1 - intOperand2)
                        Case ENUoperator.multiply:
                            pushStack(objStack, popStack(objStack) * popStack(objStack))
                        Case ENUoperator.divide:
                            intOperand2 = popStack(objStack)
                            intOperand1 = popStack(objStack)
                            If intOperand2 = 0 Then
                                Return("Division by zero")
                            End If                            
                            pushStack(objStack, intOperand1 \ intOperand2)
                        Case ENUoperator.push:
                            pushStack(objStack, .intOperand)
                        Case Else
                            Return("Invalid op was compiled")
                    End Select   
                    If statusDisplayInEffect Then
                        CTLstatusDisplay.interpreterDisplayUpdate(intIndex1, objStack)
                    End If                    
                Catch
                    Return("Interpreter error: " & Err.Number & " " & Err.Description)
                End Try                                 
            End With            
        Next intIndex1        
        If objStack.Count < 1 Then
            Return("Stack empty")
        End If        
        Return(CStr(popStack(objStack)))
    End Function   
    
    ' -----------------------------------------------------------------
    ' Pop the stack
    '
    '
    Private Function popStack(ByRef objStack As Stack) As Integer
        Try
            Return(CInt(objStack.Pop))
        Catch
            MsgBox("Error in popStack: " & Err.Number & " " & Err.Description)
            Return(0)
        End Try        
    End Function     
    
    ' -----------------------------------------------------------------
    ' Push the stack
    '
    '
    Private Sub pushStack(ByRef objStack As Stack, _
                          ByVal intValue As Integer)  
        Try
            objStack.Push(intValue)
        Catch
            MsgBox("Error in pushStack: " & Err.Number & " " & Err.Description)
        End Try        
    End Sub     

#End Region ' Interpreter   

#End Region ' General procedures 

#Region " Private class for status information display "

    ' *****************************************************************
    ' *                                                               *
    ' * statusDisplay                                                 *
    ' *                                                               *
    ' *                                                               *
    ' * This control displays the progress of scanning, parsing and   *
    ' * interpretation visually.  It contains the following constitu- *
    ' * ent controls.                                                 *    
    ' *                                                               *
    ' *                                                               *
    ' *      *  A text box shows the expression, and is highlighted   *
    ' *         during scanning, parsing, and when the parse outline  *
    ' *         (below) is clicked.                                   *
    ' *                                                               *
    ' *      *  A list box shows the scan results                     *
    ' *                                                               *
    ' *      *  A list box presents an outline of the parse           *
    ' *         tree; when this box is clicked the expression is      *
    ' *         highlighted.                                          *
    ' *                                                               *
    ' *      *  A list box shows the individual Polish instructions,  *
    ' *         highlighting the current instruction.                 *
    ' *                                                               *
    ' *      *  A list box shows the interpretation stack             *
    ' *                                                               *
    ' *      *  A check box indicates whether "instant replay" should *
    ' *         be available                                          *
    ' *                                                               *
    ' *      *  Buttons are provided for "instant replay":            *
    ' *                                                               *
    ' *         + The Replay button replays the display               *
    ' *         + The Reset button resets the replay to step 1        *
    ' *         + The Step button advances one step                   *
    ' *         + The Backup button goes back                         *
    ' *                                                               *
    ' *                                                               *
    ' * This control exposes and overrides the following properties   *
    ' * and methods in addition to the base properties and methods of *
    ' * controls.                                                     *
    ' *                                                               *
    ' *                                                               *
    ' *      *  clear: this method resets the control                 *
    ' *                                                               *
    ' *      *  codeUpdate(op,operand, comment): this method adds a   *
    ' *         line of compiled reverse-Polish code to the code      *
    ' *         display. This method requires two parameters.         *
    ' *                                                               *
    ' *         + op: the string form of an RPN operator              *
    ' *         + operand: the string form of an RPN operand          *
    ' *         + comment: comment                                    *
    ' *                                                               *
    ' *      *  dispose: this method gets rid of reference objects    *
    ' *         associated with the statusDisplay; for best results   *
    ' *         use this method when you no longer need this control. *
    ' *                                                               *
    ' *      *  Expression: this write-only property can be set to    *
    ' *         the expression being evaluated.                       *
    ' *                                                               *
    ' *      *  interpreterDisplayUpdate(ip,stack): this method       *
    ' *         updates the interpreter display.                      *
    ' *                                                               *
    ' *         + ip should be the instruction pointer                *
    ' *         + stack should be the interpreter stack               *
    ' *                                                               *
    ' *      *  new([parent[,left[,top[,width[,height[,backColor,     *
    ' *         foreColor]]]]]]): the constructor can specify the     *
    ' *         parent control and the control geometry as the left   *
    ' *         side, top, width and height of the statusDisplay.     *
    ' *         Optionally it can also specify a control color scheme *
    ' *         as the background and foreground colors of labels.    *
    ' *                                                               *
    ' *         As shown above, all parameters may be omitted.  The   *
    ' *         parent defaults to Nothing.  The default geometry uses*
    ' *         the size of a default group box and a Location of     *
    ' *         0, 0.                                                 *
    ' *                                                               *
    ' *         To obtain the corresponding default value of any      *
    ' *         geometry parameter, use a value of -1.                *
    ' *                                                               *
    ' *      *  parseUpdate(index,length,name,level): this method     *
    ' *         updates the parse display.  The parse display consists*
    ' *         of a visual parse outline: clicking on this outline   *
    ' *         highlights text belonging to grammar symbols.         *
    ' *                                                               *
    ' *         + index: the parse index (first character of grammar  *
    ' *           symbol.)                                            *
    ' *                                                               *
    ' *         + length: the length of the parsed grammar symbol     *
    ' *                                                               *
    ' *         + name: the name of the parsed grammar symbol         *
    ' *                                                               *
    ' *         + level: the depth of the parsed grammar symbol       *
    ' *                                                               *
    ' *      *  scanUpdate(index,length,type): this method updates the*
    ' *         scan display.  The scan display consists of highlights*
    ' *         to the expression and a listbox of scan information.  *
    ' *                                                               *
    ' *         + index: the scan index                               *
    ' *                                                               *
    ' *         + length: the length of the token being scanned       *
    ' *                                                               *
    ' *         + type: an enumerator of type ENUscannedType and the  *
    ' *           type of the token being scanned                     *
    ' *                                                               *
    ' *                                                               *
    ' * C H A N G E   R E C O R D ----------------------------------- *
    ' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE               *
    ' * --------   ----------     ----------------------------------- *
    ' * 01 27 03   Nilges         Version 1.0                         *
    ' *                                                               *
    ' *                                                               *
    ' *****************************************************************
    Private Class statusDisplay
        Inherits Control

        ' ***** Shared *****
        Private Shared _INTsequenceNumber As Integer
        Private Shared _OBJutilities As utilities.utilities
        Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities

        ' ***** Object state *****
        ' --- The state's structure
        Private Structure TYPstate
            Dim booUsable As Boolean                ' True: object is usable
            Dim gbxContainer As GroupBox            ' Container
            Dim lblExpression As Label              ' Expression
            Dim txtExpression As TextBox
            Dim lblParseOutline As Label            ' The parse outline
            Dim lstParseOutline As ListBox
            Dim lblScanned As Label                 ' Scanned code     
            Dim lstScanned As ListBox
            Dim lblRPN As Label                     ' RPN code
            Dim lstRPN As ListBox
            Dim lblStack As Label                   ' Stack
            Dim lstStack As ListBox
            Dim chkReplay As CheckBox               ' Instant replay
            Dim cmdReplay As Button                 ' Instant replay
            Dim cmdReset As Button                  ' Instant replay reset
            Dim cmdStep As Button                   ' Instant replay step
            Dim cmdScanZoom As Button               ' Zoom buttons
            Dim cmdParseZoom As Button
            Dim cmdRPNZoom As Button
            Dim cmdstackZoom As Button
        End Structure
        ' --- The state 
        Private USRstate As TYPstate
        ' --- State replay
        Private COLmacro As Collection              ' Of String
        Private INTmacroIndex As Integer            ' Replay index  
        ' --- With-events controls also have handles in the state
        Private WithEvents LSTscanned As ListBox
        Private WithEvents LSTparseOutline As ListBox
        Private WithEvents CMDscanZoom As Button
        Private WithEvents CMDparseZoom As Button
        Private WithEvents CMDrpnZoom As Button
        Private WithEvents CMDstackZoom As Button
        Private WithEvents CHKreplay As CheckBox
        Private WithEvents CMDreplay As Button
        Private WithEvents CMDreset As Button
        Private WithEvents CMDstep As Button

        ' ***** Constants *****
        ' --- The following constants should sum to 1
        Private Const LISTBOX_PROPORTION_SCANNED As Single = 0.25
        Private Const LISTBOX_PROPORTION_PARSE As Single = 0.38
        Private Const LISTBOX_PROPORTION_RPN As Single = 0.21
        Private Const LISTBOX_PROPORTION_STACK As Single = 0.16

        ' ***** Public procedures **************************************  

        ' --------------------------------------------------------------
        ' Reset the control
        '
        '
        Public Function clear() As Boolean
            If Not checkUsable_("clear") Then Return (False)
            With USRstate
                .txtExpression.Text = " "
                clearListBoxes_()
                replayErase_()
                Return (True)
            End With
        End Function

        ' --------------------------------------------------------------
        ' Update the RPN code display
        '
        '
        Public Function codeUpdate(ByVal strOp As String, _
                                   ByVal strOperand As String, _
                                   ByVal strComment As String) As Boolean
            If Not checkUsable_("codeUpdate") Then Return (False)
            With USRstate.lstRPN
                .Items.Add(strOp & " " & strOperand & " " & strComment)
            End With
            If Not replayStore_("codeUpdate", strOp, strOperand, strComment) Then
                Return (False)
            End If
            Return (True)
        End Function

        ' --------------------------------------------------------------
        ' Dispose of the object
        '
        '
        Private Overloads Sub dispose()
            If Not checkUsable_("dispose") Then Return
            With USRstate
                .booUsable = False
                _OBJwindowsUtilities.disposeControl(.gbxContainer)
                replayErase_()
                .gbxContainer = Nothing
            End With
        End Sub

        ' --------------------------------------------------------------
        ' Set the expression
        '
        '
        Public WriteOnly Property Expression() As String
            Set(ByVal strNewValue As String)
                If Not checkUsable_("Expression set") Then Return
                With USRstate.txtExpression
                    .Text = strNewValue : .Refresh()
                End With
            End Set
        End Property

        ' --------------------------------------------------------------
        ' Interpreter display update
        '
        '
        Public Function interpreterDisplayUpdate(ByVal intIP As Integer, _
                                                 ByRef objStack As Stack) As Boolean
            If Not checkUsable_("interpreterDisplay") Then Return (False)
            With USRstate
                .gbxContainer.Visible = False
                ' --- Check the instruction pointer
                If intIP < 1 OrElse intIP > .lstRPN.Items.Count Then
                    errorHandler_("Invalid instruction pointer " & intIP, _
                                  "interpreterDisplay", _
                                  "No change is being made to the display")
                    Return (False)
                End If
                ' --- Highlight current op
                .lstRPN.SelectedIndex = intIP - 1
                ' --- Display the stack
                ' Stack to collection (and a string if we're doing a replay)
                Dim colSave As Collection
                Dim intIndex1 As Integer
                Try
                    colSave = New Collection
                    For intIndex1 = 1 To objStack.Count
                        colSave.Add(objStack.Pop)
                    Next intIndex1
                Catch
                    errorHandler_("Cannot save stack: " & _
                                  Err.Number & " " & Err.Description, _
                                  "interpreterDisplay", _
                                  "Returning false")
                    Return (False)
                End Try
                ' Collection to display
                With colSave
                    USRstate.lstStack.Items.Clear()
                    For intIndex1 = 1 To colSave.Count
                        USRstate.lstStack.Items.Add(.Item(intIndex1))
                    Next intIndex1
                End With
                ' Restore the stack
                If Not collection2Stack_(colSave, objStack) Then
                    errorHandler_("The stack passed to me has been damaged", _
                                  "interpreterDisplayUpdate", _
                                  "Object is unusable")
                    USRstate.booUsable = False
                    Return (False)
                End If
                .gbxContainer.Visible = True
                replayStore_("interpreterDisplayUpdate", _
                             CStr(intIP), _
                             OBJcollectionUtilities.collection2String(colSave))
                Return (True)
            End With
        End Function

        ' --------------------------------------------------------------
        ' Object constructor
        '
        '
        Public Sub New()
            new_(Nothing, -1, -1, -1, -1, Color.Red, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control)
            new_(ctlParent, -1, -1, -1, -1, Color.Red, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control, _
                       ByVal intLeft As Integer)
            new_(ctlParent, intLeft, -1, -1, -1, Color.Red, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer)
            new_(ctlParent, intLeft, intTop, -1, -1, Color.Red, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer)
            new_(ctlParent, intLeft, intTop, intWidth, -1, Color.Red, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer, _
                       ByVal intHeight As Integer)
            new_(ctlParent, intLeft, intTop, intWidth, intHeight, Color.Red, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal objBackColor As Color)
            new_(ctlParent, intLeft, intTop, intWidth, intHeight, objBackColor, Color.White)
        End Sub
        Public Sub New(ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal objBackColor As Color, _
                       ByVal objForeColor As Color)
            new_(ctlParent, intLeft, intTop, intWidth, intHeight, objBackColor, objForeColor)
        End Sub
        Private Sub new_(ByVal ctlParent As Control, _
                         ByVal intLeft As Integer, _
                         ByVal intTop As Integer, _
                         ByVal intWidth As Integer, _
                         ByVal intHeight As Integer, _
                         ByVal objBackColor As Color, _
                         ByVal objForeColor As Color)
            If intLeft < -1 _
               OrElse _
               intTop < -1 Then
                errorHandler_("Invalid left or top", _
                              "new_ constructor", _
                              "Object will not be usable")
                Return
            End If
            With USRstate
                Try
                    ' --- Create group box and position control
                    If Not _OBJwindowsUtilities.setControlToGroupBox(Me, .gbxContainer, _
                                                                     intWidth, intHeight) Then
                        errorHandler_("Cannot assign a group box to base control: " & _
                                      Err.Number & " " & Err.Description, _
                                      "new_", _
                                      "Object is not usable")
                        Return
                    End If
                    With Me
                        If intLeft <> -1 Then MyBase.Left = intLeft
                        If intTop <> -1 Then MyBase.Top = intTop
                    End With
                    If Not (ctlParent Is Nothing) Then
                        ctlParent.Controls.Add(Me)
                    End If
                    ' --- Create expression display
                    ' Create label                  
                    .lblExpression = New Label
                    .gbxContainer.Controls.Add(.lblExpression)
                    With .lblExpression
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.Grid
                        .Top = _OBJwindowsUtilities.Grid * 3
                        .Width = USRstate.gbxContainer.Width _
                                 - _
                                 _OBJwindowsUtilities.Grid * 2
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Expression"
                        .TextAlign = ContentAlignment.MiddleCenter
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create text box                  
                    .txtExpression = New TextBox
                    .gbxContainer.Controls.Add(.txtExpression)
                    With .txtExpression
                        .Left = USRstate.lblExpression.Left
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblExpression)
                        .Width = USRstate.lblExpression.Width
                        .ReadOnly = True
                    End With
                    ' --- Create scan display below expression and almost flush with left              
                    Dim intUnderExpression As Integer = _
                        _OBJwindowsUtilities.controlBottom(.txtExpression)
                    ' Create label   
                    .lblScanned = New Label
                    .gbxContainer.Controls.Add(.lblScanned)
                    With .lblScanned
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = USRstate.lblExpression.Left
                        .Top = intUnderExpression + _OBJwindowsUtilities.Grid
                        .Width = CInt(USRstate.lblExpression.Width * LISTBOX_PROPORTION_SCANNED)
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Scanned Tokens"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box with events            
                    LSTscanned = New ListBox
                    .lstScanned = LSTscanned
                    AddHandler LSTscanned.SelectedIndexChanged, AddressOf lstScanned_SelectedIndexChanged_
                    .gbxContainer.Controls.Add(.lstScanned)
                    With .lstScanned
                        .Left = USRstate.lblExpression.Left
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblScanned)
                        .Width = USRstate.lblScanned.Width
                        .Height = USRstate.gbxContainer.Height _
                                  - _
                                  new__buttonRowHeight_() _
                                  - _
                                  _OBJwindowsUtilities.Grid * 2 _
                                  - _
                                  .Top
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                    End With
                    ' Create zoomer
                    .cmdScanZoom = new__createZoom_(.lblScanned, _
                                                    CMDscanZoom, _
                                                    .gbxContainer, _
                                                    .lstScanned)
                    ' --- Create parse display below expression and in the middle
                    ' Create label
                    .lblParseOutline = New Label
                    .gbxContainer.Controls.Add(.lblParseOutline)
                    With .lblParseOutline
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.lstScanned)
                        .Top = intUnderExpression + _OBJwindowsUtilities.Grid
                        .Width = CInt(USRstate.lblExpression.Width * LISTBOX_PROPORTION_PARSE)
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Parse Outline"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box with events
                    LSTparseOutline = New ListBox
                    .lstParseOutline = LSTparseOutline
                    AddHandler LSTparseOutline.SelectedIndexChanged, _
                               AddressOf lstParseOutline_SelectedIndexChanged_
                    .gbxContainer.Controls.Add(.lstParseOutline)
                    With LSTparseOutline
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblParseOutline)
                        .Left = USRstate.lblParseOutline.Left
                        .Width = USRstate.lblParseOutline.Width
                        .Height = USRstate.lstScanned.Height
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                    End With
                    ' Create zoomer
                    .cmdParseZoom = new__createZoom_(.lblParseOutline, _
                                                     CMDparseZoom, _
                                                     .gbxContainer, _
                                                     .lstParseOutline)
                    ' --- Create object code display below expression and to right of the parse display   
                    ' Create label           
                    .lblRPN = New Label
                    .gbxContainer.Controls.Add(.lblRPN)
                    With .lblRPN
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.lstParseOutline)
                        .Width = CInt(USRstate.lblExpression.Width * LISTBOX_PROPORTION_RPN)
                        .Top = intUnderExpression + _OBJwindowsUtilities.Grid
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "RPN"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box              
                    .lstRPN = New ListBox
                    .gbxContainer.Controls.Add(.lstRPN)
                    With .lstRPN
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblRPN)
                        .Left = USRstate.lblRPN.Left
                        .Width = USRstate.lblRPN.Width
                        .Height = USRstate.lstScanned.Height
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                    End With
                    ' Create zoomer
                    .cmdRPNZoom = new__createZoom_(.lblRPN, _
                                                   CMDrpnZoom, _
                                                   .gbxContainer, _
                                                   .lstRPN)
                    ' --- Create stack display below expression and to right of object code display    
                    ' Create label
                    .lblStack = New Label
                    .gbxContainer.Controls.Add(.lblStack)
                    With .lblStack
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.lstRPN)
                        .Width = CInt(USRstate.lblExpression.Width * LISTBOX_PROPORTION_STACK)
                        .Top = intUnderExpression + _OBJwindowsUtilities.Grid
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Stack"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box          
                    .lstStack = New ListBox
                    .gbxContainer.Controls.Add(.lstStack)
                    With .lstStack
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblStack)
                        .Left = USRstate.lblStack.Left
                        .Width = USRstate.lblStack.Width
                        .Height = USRstate.lstScanned.Height
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                    End With
                    ' Create zoomer
                    .cmdstackZoom = new__createZoom_(.lblStack, _
                                                     CMDstackZoom, _
                                                     .gbxContainer, _
                                                     .lstStack)
                    ' --- Create replay commands under other controls
                    Dim intBottom As Integer = Math.Max(_OBJwindowsUtilities.controlBottom(.lstScanned), _
                                                        Math.Max(_OBJwindowsUtilities.controlBottom(.lstParseOutline), _
                                                                 _OBJwindowsUtilities.controlBottom(.lstRPN))) _
                                               + _
                                               _OBJwindowsUtilities.Grid
                    ' Replay check box                                                                 
                    CHKreplay = New CheckBox
                    AddHandler CHKreplay.CheckedChanged, AddressOf chkReplay_CheckedChanged_
                    .chkReplay = CHKreplay
                    .gbxContainer.Controls.Add(.chkReplay)
                    With .chkReplay
                        .Left = USRstate.lblExpression.Left
                        .Top = intBottom
                        .Text = "Replay"
                    End With
                    ' Replay command button                                                                 
                    CMDreplay = New Button
                    AddHandler CMDreplay.Click, AddressOf cmdReplay_Click_
                    .cmdReplay = CMDreplay
                    .gbxContainer.Controls.Add(.cmdReplay)
                    With .cmdReplay
                        .Visible = False
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.chkReplay) _
                                + _
                                _OBJwindowsUtilities.Grid
                        .Top = intBottom
                        .Text = "Replay"
                    End With
                    ' Reset command button                                                                 
                    CMDreset = New Button
                    AddHandler CMDreset.Click, AddressOf cmdReset_Click_
                    .cmdReset = CMDreset
                    .gbxContainer.Controls.Add(.cmdReset)
                    With .cmdReset
                        .Visible = False
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.cmdReplay) _
                                + _
                                _OBJwindowsUtilities.Grid
                        .Top = intBottom
                        .Text = "Reset"
                    End With
                    ' Step command button                                                                 
                    CMDstep = New Button
                    AddHandler CMDstep.Click, AddressOf cmdStep_Click_
                    .cmdStep = CMDstep
                    .gbxContainer.Controls.Add(.cmdStep)
                    With .cmdStep
                        .Visible = False
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.cmdReset) _
                                + _
                                _OBJwindowsUtilities.Grid
                        .Top = intBottom
                        .Text = "Step"
                    End With
                Catch
                    errorHandler_("Cannot create statusDisplay: " & _
                                  Err.Number & " " & Err.Description, _
                                  "new_", _
                                  "Object is not usable")
                    Return
                End Try
                .booUsable = True
                .booUsable = Me.clear
            End With
        End Sub

        ' ---------------------------------------------------------------
        ' On behalf of the new constructor, return the height of the
        ' area at the bottom of the display for replay buttons
        '
        '
        Private Function new__buttonRowHeight_() As Integer
            Dim chkBox As CheckBox
            Dim cmdButton As Button
            Try
                chkBox = New CheckBox : cmdButton = New Button
            Catch
                errorHandler_("Cannot create example controls", _
                              "new__buttonRowHeight_", _
                              "Returning zero")
                Return (0)
            End Try
            Dim intHeight As Integer = Math.Max(chkBox.Height, cmdButton.Height) _
                                       + _
                                       _OBJwindowsUtilities.Grid * 2
            chkBox.Dispose() : cmdButton.Dispose()
            Return (intHeight)
        End Function

        ' ---------------------------------------------------------------
        ' On behalf of the new constructor, create and position a zoomer
        '
        '
        ' Note: the zoom button is tagged with the control it zooms.
        '
        '
        Private Function new__createZoom_(ByVal lblContainer As Label, _
                                          ByVal cmdWithEvents As Button, _
                                          ByVal gbxContainer As GroupBox, _
                                          ByVal ctlZoomed As Control) As Button
            Try
                cmdWithEvents = New Button
                AddHandler cmdWithEvents.Click, AddressOf cmdZoom_Click_
                Dim cmdNew As Button = cmdWithEvents
                With cmdNew
                    .Text = "Zoom"
                    .Left = _OBJwindowsUtilities.controlRight(lblContainer) _
                            - _
                            .Width
                    .Top = lblContainer.Top
                    lblContainer.Width -= .Width
                    .Tag = ctlZoomed
                End With
                gbxContainer.Controls.Add(cmdNew)
                Return (cmdNew)
            Catch
                errorHandler_("Could not create zoom button associated with label " & _
                              lblContainer.Name & ": " & _
                              Err.Number & " " & Err.Description, _
                              "new__createZoom_", _
                              "Returning Nothing: no zoom button shall appear")
                Return (Nothing)
            End Try
        End Function

        ' ---------------------------------------------------------------
        ' Update the parse outline
        '
        '
        ' This method adds a string in the form
        '
        '
        '      <level><grammarSymbol> "<text>" at n-m
        '
        '
        ' to the parse outline list box, where:
        '
        '
        '      *  <level> is a string of L spaces where L is the parse
        '         level
        '
        '      *  <grammarSymbol> is a grammar symbol
        '
        '      *  <text> is the text corresponding to the grammar
        '         symbol
        '
        '      *  n is the start index of <text> (from 1)
        '
        '      *  m is the end index
        '
        '
        ' See parseDisplay2Selection, which parses the above format, and
        ' may need to be updated if it changes.
        '
        '
        Public Function parseUpdate(ByVal intIndex As Integer, _
                                    ByVal intLength As Integer, _
                                    ByVal strName As String, _
                                    ByVal intLevel As Integer) As Boolean
            If Not checkUsable_("parseUpdate") Then Return (False)
            If intIndex < 1 OrElse intLength < 0 OrElse intLevel < 0 Then
                errorHandler_("Invalid parameters", _
                              "parseUpdate", _
                              "Returning False")
                Return (False)
            End If
            With USRstate.lstParseOutline
                Try
                    Dim intIndex1 As Integer
                    Dim colSave As New Collection
                    For intIndex1 = 0 To .Items.Count - 1
                        colSave.Add(.Items(intIndex1))
                    Next intIndex1
                    .Items.Clear()
                    .Items.Add(_OBJutilities.copies(" ", intLevel * 5) & _
                                strName & " " & _
                                _OBJutilities.enquote(Mid(USRstate.txtExpression.Text, _
                                                            intIndex, intLength)) & " " & _
                                "at " & _
                                intIndex & "-" & intIndex + intLength - 1)
                    For intIndex1 = 1 To colSave.Count
                        .Items.Add(colSave.Item(intIndex1))
                    Next intIndex1
                Catch
                    errorHandler_("Cannot extend the parse outline: " & _
                                  Err.Number & " " & Err.Description, _
                                  "parseUpdate", _
                                  "Making object unusable")
                    USRstate.booUsable = False
                    Return (False)
                End Try
                .SelectedIndex = 0
                .Refresh()
            End With
            replayStore_("parseUpdate", _
                         CStr(intIndex), _
                         CStr(intLength), _
                         strName, _
                         CStr(intLevel))
            Return (True)
        End Function

        ' ---------------------------------------------------------------
        ' Update the scan display and highlight scanned token
        '
        '
        ' This method adds a line in the form
        '
        '
        '      tokenType at start-end
        '
        '
        Public Function scanUpdate(ByVal intIndex As Integer, _
                                   ByVal intLength As Integer, _
                                   ByVal strTokenType As String) As Boolean
            If Not checkUsable_("scanUpdate") Then Return (False)
            If intIndex < 1 OrElse intLength < 0 Then
                errorHandler_("Invalid parameters", _
                              "scanUpdate", _
                              "Returning False")
                Return (False)
            End If
            With USRstate.lstScanned
                Try
                    .Items.Add(strTokenType & " at " & _
                               intIndex & "-" & intIndex + intLength - 1)
                Catch
                    errorHandler_("Cannot extend the parse outline: " & _
                                  Err.Number & " " & Err.Description, _
                                  "parseUpdate", _
                                  "Making object unusable")
                    USRstate.booUsable = False
                    Return (False)
                End Try
                .SelectedIndex = .Items.Count - 1
            End With
            replayStore_("scanUpdate", _
                         CStr(intIndex), _
                         CStr(intLength), _
                         strTokenType)
            Return (True)
        End Function

        ' ***** Private procedures **************************************

        ' ---------------------------------------------------------------     
        ' Check control usability
        '
        '
        Private Function checkUsable_(ByVal strProcedure As String) As Boolean
            If Not USRstate.booUsable Then
                errorHandler_("Object is not usable", _
                              strProcedure, _
                              "Returning a default value")
                Return (False)
            End If
            Return (True)
        End Function

        ' ---------------------------------------------------------------
        ' Clears "my" list boxes
        '
        '
        Private Sub clearListBoxes_()
            With USRstate
                .lstScanned.Items.Clear()
                .lstParseOutline.Items.Clear()
                .lstRPN.Items.Clear()
                .lstStack.Items.Clear()
            End With
        End Sub

        ' ---------------------------------------------------------------
        ' Collection to stack
        '
        '
        Private Function collection2NewStack_(ByVal colCollection As Collection) As Stack
            Dim objNewStack As Stack
            Try
                objNewStack = New Stack
            Catch
                errorHandler_("Can't create a new stack: " & _
                              Err.Number & " " & Err.Description, _
                              "collection2NewStack_", _
                              "Making object unusable")
                USRstate.booUsable = False
                Return (Nothing)
            End Try
            If Not collection2Stack_(colCollection, objNewStack) Then Return (Nothing)
            Return (objNewStack)
        End Function

        ' ---------------------------------------------------------------
        ' Collection to existing stack
        '
        '
        Private Function collection2Stack_(ByVal colCollection As Collection, _
                                           ByRef objStack As Stack) As Boolean
            Dim intIndex1 As Integer
            With colCollection
                For intIndex1 = .Count To 1 Step -1
                    Try
                        objStack.Push(.Item(intIndex1))
                    Catch
                        errorHandler_("Failure to push collection member on stack" & _
                                      Err.Number & " " & Err.Description, _
                                      "collection2Stack_", _
                                      "Returning False")
                        Return (False)
                    End Try
                Next intIndex1
                Return (True)
            End With
        End Function

        ' ---------------------------------------------------------------
        ' Error handling interface
        '
        '
        Private Sub errorHandler_(ByVal strMessage As String, _
                                  ByVal strProcedure As String, _
                                  ByVal strHelp As String)
            _OBJutilities.errorHandler(strMessage, _
                                       "statusDisplay", _
                                       strProcedure, _
                                       strHelp)
        End Sub

        ' ---------------------------------------------------------------
        ' Extract selection start and length from parse or scan display string
        '
        '
        ' Expects to find start-end at end of either the parse or the scan
        ' display, where start-end indicates the starting and ending position
        ' of a grammar category or token.
        '
        '
        Private Function parseDisplay2Selection_(ByVal strDisplay As String) _
                As Boolean
            Dim strLast As String = _OBJutilities.word(strDisplay, _
                                                       _OBJutilities.words(strDisplay))
            If _OBJutilities.items(strLast, "-", False) = 2 Then
                Dim strStartIndex As String = _OBJutilities.item(strLast, 1, "-", False)
                Dim strEndIndex As String = _OBJutilities.item(strLast, 2, "-", False)
                If _OBJutilities.datatype(strStartIndex, "unsignedIntegerDatatype") _
                   AndAlso _
                   _OBJutilities.datatype(strEndIndex, "unsignedIntegerDatatype") Then
                    Dim intStartIndex As Integer = CInt(strStartIndex)
                    USRstate.txtExpression.Select(intStartIndex - 1, _
                                                  CInt(strEndIndex) - intStartIndex + 1)
                    USRstate.txtExpression.Focus()
                    Return (True)
                End If
            End If
            errorHandler_("Programming error: invalid parse display string " & _
                          _OBJutilities.enquote(strDisplay), _
                          "parseDisplay2Selection_", _
                          "Marking object unusable")
            USRstate.booUsable = False
            Return (False)
        End Function

        ' ---------------------------------------------------------------
        ' Replay
        '
        '
        Private Overloads Function replay_() As Boolean
            Return (replay_(1, COLmacro.Count))
        End Function
        Private Overloads Function replay_(ByVal intStartIndex As Integer, _
                                           ByVal intCount As Integer) As Boolean
            If (COLmacro Is Nothing) Then Return (True)
            Dim intIndex1 As Integer
            Dim booReplay As Boolean
            With USRstate.chkReplay
                booReplay = .Checked
                RemoveHandler CHKreplay.CheckedChanged, AddressOf chkReplay_CheckedChanged_
                .Checked = False
                .Refresh()
            End With
            For intIndex1 = intStartIndex To Math.Min(intStartIndex + intCount - 1, COLmacro.Count)
                Try
                    Dim strSplit() As String
                    strSplit = Split(CStr(COLmacro.Item(intIndex1)), ChrW(0))
                    Select Case UCase(Trim(strSplit(0)))
                        Case "CODEUPDATE"
                            Me.codeUpdate(strSplit(1), strSplit(2), strSplit(3))
                        Case "INTERPRETERDISPLAYUPDATE"
                            Me.interpreterDisplayUpdate(CInt(strSplit(1)), _
                                                        collection2NewStack_(OBJcollectionUtilities.string2Collection(strSplit(2))))
                        Case "PARSEUPDATE"
                            Me.parseUpdate(CInt(strSplit(1)), _
                                           CInt(strSplit(2)), _
                                           strSplit(3), _
                                           CInt(strSplit(4)))
                        Case "SCANUPDATE"
                            Me.scanUpdate(CInt(strSplit(1)), _
                                          CInt(strSplit(2)), _
                                          strSplit(3))
                        Case Else
                            errorHandler_("Programming error: unexpected procedure name " & _
                                          _OBJutilities.enquote(strSplit(0)) & " " & _
                                          "occured in macro record", _
                                          "replay_", _
                                          "Making object unusable")
                            USRstate.booUsable = False
                            Return (False)
                    End Select
                Catch
                    errorHandler_("Cannot replay: " & _
                                  Err.Number & " " & Err.Description, _
                                  "replay_", _
                                  "Replay has failed")
                    Exit For
                End Try
                Me.Refresh()
            Next intIndex1
            With USRstate.chkReplay
                .Checked = booReplay : .Refresh()
            End With
            AddHandler CHKreplay.CheckedChanged, AddressOf chkReplay_CheckedChanged_
            INTmacroIndex = intIndex1
            Return (intIndex1 > intCount)
        End Function

        ' ---------------------------------------------------------------
        ' Erase the replay
        '
        '
        Private Function replayErase_() As Boolean
            If (COLmacro Is Nothing) Then Return (True)
            COLmacro = Nothing
            Return (True)
        End Function

        ' ---------------------------------------------------------------
        ' Store clone in the replay collection
        '
        '
        Private Function replayStore_(ByVal strProcedure As String, _
                                      ByVal ParamArray strOperands() As String) As Boolean
            If (COLmacro Is Nothing) Then
                Try
                    COLmacro = New Collection
                Catch
                    errorHandler_("Cannot create the replay collection: " & _
                                  Err.Number & " " & Err.Description, _
                                  "", _
                                  "Object will be unusable")
                    USRstate.booUsable = False
                    Return (False)
                End Try
            End If
            With COLmacro
                Try
                    Dim intIndex1 As Integer
                    Dim objStringBuilder As System.Text.StringBuilder
                    objStringBuilder = New System.Text.StringBuilder
                    With objStringBuilder
                        If InStr(strProcedure, ChrW(0)) <> 0 Then
                            errorHandler_("Procedure name contains the null character", _
                                          "replayStore_", _
                                          "Object will be unusable")
                            USRstate.booUsable = False
                            Return (False)
                        End If
                        .Append(strProcedure)
                        For intIndex1 = 0 To UBound(strOperands)
                            If InStr(strOperands(intIndex1), ChrW(0)) <> 0 Then
                                errorHandler_("Procedure operand contains the null character", _
                                              "replayStore_", _
                                              "Object will be unusable")
                                USRstate.booUsable = False
                                Return (False)
                            End If
                            _OBJutilities.append(objStringBuilder, ChrW(0), strOperands(intIndex1))
                        Next intIndex1
                    End With
                    .Add(objStringBuilder.ToString)
                Catch
                    errorHandler_("Cannot extend the replay collection: " & _
                                  Err.Number & " " & Err.Description, _
                                  "replayStore_", _
                                  "Object will be unusable")
                    USRstate.booUsable = False
                    Return (False)
                End Try
            End With
        End Function

        ' ---------------------------------------------------------------
        ' Toggle instant replay
        '
        '
        Private Function replayToggle_(ByVal booValue As Boolean) As Boolean
            With USRstate
                .cmdReplay.Visible = booValue
                .cmdStep.Visible = booValue
                .cmdReset.Visible = booValue
                Me.Refresh()
            End With
            If booValue Then
                resetReplay_()
            Else
                replayErase_()
            End If
            INTmacroIndex = 0
            Return (True)
        End Function

        ' -----------------------------------------------------------------
        ' Reset replay
        '
        '
        Private Sub resetReplay_()
            INTmacroIndex = 0
            If (COLmacro Is Nothing) Then Return
            INTmacroIndex = Math.Min(1, COLmacro.Count)
            clearListBoxes_()
        End Sub

        ' ***** Event handlers ********************************************

        ' -----------------------------------------------------------------
        ' Replay check box clicked
        '
        '
        Private Sub chkReplay_CheckedChanged_(ByVal objSender As Object, _
                                              ByVal objEventArgs As System.EventArgs)
            replayToggle_(USRstate.chkReplay.Checked)
        End Sub

        ' -----------------------------------------------------------------
        ' Replay command button clicked
        '
        '
        Private Sub cmdReplay_Click_(ByVal objSender As Object, _
                                     ByVal objEventArgs As System.EventArgs)
            clearListBoxes_()
            If (COLmacro Is Nothing) Then
                MsgBox("Nothing is available for replay: use Evaluate first")
                Return
            End If
            replay_()
        End Sub

        ' -----------------------------------------------------------------
        ' Reset command button clicked
        '
        '
        Private Sub cmdReset_Click_(ByVal objSender As Object, _
                                    ByVal objEventArgs As System.EventArgs)
            resetReplay_()
        End Sub

        ' -----------------------------------------------------------------
        ' Step command button clicked
        '
        '
        Private Sub cmdStep_Click_(ByVal objSender As Object, _
                                   ByVal objEventArgs As System.EventArgs)
            If INTmacroIndex = 0 _
               OrElse _
               Not (COLmacro Is Nothing) AndAlso INTmacroIndex > COLmacro.Count Then
                clearListBoxes_()
                INTmacroIndex = 1
            End If
            replay_(INTmacroIndex, 1)
        End Sub

        ' -----------------------------------------------------------------
        ' Zoom button clicked
        '
        '
        Private Sub cmdZoom_Click_(ByVal objSender As Object, _
                                   ByVal objEventArgs As System.EventArgs)
            Try
                Dim ctlZoom As zoom.zoom
                Dim ctlHandle As Control = CType(CType(objSender, Control).Tag, Control)
                ctlZoom = New zoom.zoom(ctlHandle, CDbl(3), CDbl(3))
                With ctlZoom
                    .ZoomTextBox.Font = New Font(.ZoomTextBox.Font.FontFamily, _
                                                 CInt(.ZoomTextBox.Font.SizeInPoints * 1.5), _
                                                 FontStyle.Bold, _
                                                 System.Drawing.GraphicsUnit.Point)
                    .showZoom()
                    .dispose()
                End With
            Catch
                errorHandler_("Failed to zoom control: " & Err.Number & " " & Err.Description, _
                              "cmdZoom_Click_", _
                              "Continuing execution")
            End Try
        End Sub

        ' -----------------------------------------------------------------
        ' Parse outline clicked
        '
        '
        Private Sub lstParseOutline_SelectedIndexChanged_(ByVal objSender As Object, _
                                                          ByVal objEventArgs As System.EventArgs)
            Dim intLength As Integer
            Dim intStartIndex As Integer
            With USRstate.lstParseOutline
                If Not parseDisplay2Selection_(CStr(.Items(.SelectedIndex))) Then
                    Return
                End If
            End With
        End Sub

        ' -----------------------------------------------------------------
        ' Scanned list box clicked
        '
        '
        Private Sub lstScanned_SelectedIndexChanged_(ByVal objSender As Object, _
                                                     ByVal objEventArgs As System.EventArgs)
            With USRstate.lstScanned
                parseDisplay2Selection_(CStr(.Items(.SelectedIndex)))
            End With
        End Sub

    End Class

#End Region

End Class

