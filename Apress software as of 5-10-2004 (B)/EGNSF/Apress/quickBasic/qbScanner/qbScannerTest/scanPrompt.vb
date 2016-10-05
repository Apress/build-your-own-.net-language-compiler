Option Strict

' *********************************************************************
' *                                                                   *
' * Scan mode prompter                                                *
' *                                                                   *
' *                                                                   *
' * This form supports selection of the scan mode. Three scan modes   *
' * are supported.                                                    *
' *                                                                   *
' *                                                                   *
' *      *  ALL: the scanner object is reset (cleared) and all text   *
' *         is fully scanned.                                         *
' *                                                                   *
' *      *  NEXTTOKEN: one token is obtained and appended to the      *
' *         existing tokens in the scanner object.                    *
' *                                                                   *
' *      *  SUBSTRING: the tokens in a substring of the source code   *
' *         (indicated by text box selection on the main form) are    *
' *         obtained and appended to existing tokens.                 *
' *                                                                   *
' *                                                                   *
' * This form exposes the following Friendly properties and methods.  *
' *                                                                   *
' *                                                                   *
' *      *  Always: this read-only property returns True when the     *
' *         ScanMode returned should always be displayed, and the     *
' *         prompt should not be shown.                               *
' *                                                                   *
' *      *  Cancel: this read-only property returns True when the     *
' *         form isn't dismissed by clicking the Scan button.         *
' *                                                                   *
' *      *  CloseText: this write-only property may assign text to the*
' *         cmdScan button. By default this button has the label Scan;*
' *         the CloseText should be Close when the form won't start a *
' *         scan.                                                     *  
' *                                                                   *
' *      *  ScanMode: this read-only property returns the scan mode,  *
' *         as one of the strings ALL, NEXTTOKEN or SUBSTRING         *
' *                                                                   *
' *      *  SelectionEnd: this read-only property returns the last    *
' *         index (from 1) to be scanned based on SelectionStart and  *
' *         SelectionLength.  If SelectionLength was set to a null    *
' *         string and is unknown, SelectionEnd returns Nothing.      *
' *                                                                   *
' *      *  SelectionLength: this write-only property can set this    *
' *         form's copy of the length of the text selected on the main*
' *         form.                                                     *
' *                                                                   *
' *         Note: SelectionLength can be a positive or zero integer,  *
' *         or a null string. A null string indicates selection to the*
' *         end of the source code.                                   *
' *                                                                   *
' *      *  SelectionStart: this write-only property can set this     *
' *         form's copy of the start (from zero) of the text selected *
' *         on the main form.                                         *
' *                                                                   *
' *********************************************************************

Public Class scanPrompt
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        initializeMe

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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents radScanPromptAll As System.Windows.Forms.RadioButton
    Friend WithEvents radScanPromptNext As System.Windows.Forms.RadioButton
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdScan As System.Windows.Forms.Button
    Friend WithEvents radScanPromptSubstring As System.Windows.Forms.RadioButton
    Friend WithEvents chkAlways As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.radScanPromptAll = New System.Windows.Forms.RadioButton()
Me.radScanPromptNext = New System.Windows.Forms.RadioButton()
Me.cmdCancel = New System.Windows.Forms.Button()
Me.cmdScan = New System.Windows.Forms.Button()
Me.radScanPromptSubstring = New System.Windows.Forms.RadioButton()
Me.chkAlways = New System.Windows.Forms.CheckBox()
Me.SuspendLayout()
'
'radScanPromptAll
'
Me.radScanPromptAll.Checked = True
Me.radScanPromptAll.Location = New System.Drawing.Point(8, 8)
Me.radScanPromptAll.Name = "radScanPromptAll"
Me.radScanPromptAll.Size = New System.Drawing.Size(424, 16)
Me.radScanPromptAll.TabIndex = 0
Me.radScanPromptAll.TabStop = True
Me.radScanPromptAll.Text = "Reset and scan all source code"
'
'radScanPromptNext
'
Me.radScanPromptNext.Location = New System.Drawing.Point(7, 32)
Me.radScanPromptNext.Name = "radScanPromptNext"
Me.radScanPromptNext.Size = New System.Drawing.Size(425, 16)
Me.radScanPromptNext.TabIndex = 1
Me.radScanPromptNext.Text = "Obtain the next token without reset commencing at %STARTINDEX"
'
'cmdCancel
'
Me.cmdCancel.Location = New System.Drawing.Point(8, 80)
Me.cmdCancel.Name = "cmdCancel"
Me.cmdCancel.Size = New System.Drawing.Size(80, 24)
Me.cmdCancel.TabIndex = 2
Me.cmdCancel.Text = "Cancel"
'
'cmdScan
'
Me.cmdScan.Location = New System.Drawing.Point(352, 80)
Me.cmdScan.Name = "cmdScan"
Me.cmdScan.Size = New System.Drawing.Size(80, 24)
Me.cmdScan.TabIndex = 3
Me.cmdScan.Text = "Scan"
'
'radScanPromptSubstring
'
Me.radScanPromptSubstring.Location = New System.Drawing.Point(7, 56)
Me.radScanPromptSubstring.Name = "radScanPromptSubstring"
Me.radScanPromptSubstring.Size = New System.Drawing.Size(425, 16)
Me.radScanPromptSubstring.TabIndex = 4
Me.radScanPromptSubstring.Text = "Scan from %STARTINDEX to %ENDINDEX without reset"
'
'chkAlways
'
Me.chkAlways.Location = New System.Drawing.Point(8, 112)
Me.chkAlways.Name = "chkAlways"
Me.chkAlways.Size = New System.Drawing.Size(424, 40)
Me.chkAlways.TabIndex = 5
Me.chkAlways.Text = "Always perform selected action: %ACTION  "
'
'scanPrompt
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(438, 155)
Me.ControlBox = False
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.chkAlways, Me.radScanPromptSubstring, Me.cmdScan, Me.cmdCancel, Me.radScanPromptNext, Me.radScanPromptAll})
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "scanPrompt"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "scanPrompt"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private Shared _OBJutilities As utilities.utilities
    Private INTstartIndex As Integer 
    Private OBJlength As Object 
    Private BOOcancel As Boolean
    
#End Region ' " Form state "

#Region " Properties and methods "

    Friend ReadOnly Property Always As Boolean
        Get
            Return(chkAlways.Checked)
        End Get        
    End Property    

    Friend ReadOnly Property Cancel As Boolean
        Get
            Return(BOOcancel)
        End Get        
    End Property    

    Friend WriteOnly Property CloseText As String
        Set(ByVal strNewValue As String)
            cmdScan.Text = strNewValue
        End Set        
    End Property    

    Friend ReadOnly Property ScanMode As String
        Get
            Return(radioButtons2Mode)
        End Get        
    End Property    
    
    Friend ReadOnly Property SelectionEnd As Object
        Get
            If (OBJlength Is Nothing) Then Return(Nothing)
            Return(INTstartIndex + CInt(OBJlength) - 1)
        End Get        
    End Property    
    
    Friend WriteOnly Property SelectionStart As Integer
        Set(ByVal intNewValue As Integer)
            If intNewValue < 0 Then
                MsgBox("Internal programming error: " & _
                       "value sent to SelectionStart cannot be negative")
                Return                       
            End If            
            INTstartIndex = intNewValue + 1
            With radScanPromptNext
                .Text = Replace(CStr(.Tag), _
                                "%STARTINDEX", _
                                "character " & INTstartIndex)
            End With                                             
        End Set        
    End Property    
    
    Friend WriteOnly Property SelectionLength As Object
        Set(ByVal objNewValue As Object)
            Dim strEndIndex As String
            If (TypeOf objNewValue Is System.String) AndAlso CStr(objNewValue) = "" Then
                OBJlength = Nothing
                strEndIndex = "end of string" 
            ElseIf (TypeOf objNewValue Is System.Int32) Then
                OBJlength = CInt(objNewValue)
                strEndIndex = "character " & INTstartIndex + CInt(OBJlength)
            End If            
            If strEndIndex = "" Then
                MsgBox("Invalid SelectionLength"): Return
            End If            
            With radScanPromptSubstring
                .Text = Replace(CStr(.Tag), _
                                "%ENDINDEX", _
                                strEndIndex)
                .Text = Replace(.Text, _
                                "%STARTINDEX", _
                                "character " & CStr(INTstartIndex))
            End With                                             
        End Set        
    End Property    

#End Region ' " Properties and methods "

#Region " Form events "

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        registry2Form
        Hide
    End Sub

    Private Sub cmdScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScan.Click
        form2Registry
        If Me.Always Then
            MsgBox("To display this form you can use Tools..Scan Prompt on the main form")
        End If        
        BOOcancel = False
        Hide
    End Sub

    Private Sub radScanPromptAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radScanPromptAll.CheckedChanged
        text2Always(radScanPromptAll.Text)
    End Sub

    Private Sub radScanPromptNext_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radScanPromptNext.CheckedChanged
        text2Always(radScanPromptNext.Text)
    End Sub

    Private Sub radScanPromptSubstring_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radScanPromptSubstring.CheckedChanged
        text2Always(radScanPromptSubstring.Text)
    End Sub

    Private Sub scanPrompt_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        BOOcancel = True
    End Sub

    Private Sub scanPrompt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub
    
#End Region ' " Form events "    

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Save Registry values
    '
    '
    Private Sub form2Registry
        Try
            With radScanPromptAll
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked)) 
            End With            
            With radScanPromptNext
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked)) 
            End With            
            With radScanPromptSubstring
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked)) 
            End With            
            With chkAlways
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked)) 
            End With            
        Catch
            MsgBox("Can't save Registry values in " & Me.Name & ": " & _
                   Err.Number & " " & Err.Description)
        End Try        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Initialize component extensions
    '
    '
    Private Sub initializeMe
        radScanPromptNext.Tag = radScanPromptNext.Text
        radScanPromptSubstring.Tag = radScanPromptSubstring.Text
        chkAlways.Tag = chkAlways.Text
        Me.SelectionStart = 1
        Me.SelectionLength = ""
        Select Case Me.ScanMode
            Case "ALL": text2Always(radScanPromptAll.Text)
            Case "NEXTTOKEN": text2Always(radScanPromptNext.Text)
            Case "SUBSTRING": text2Always(radScanPromptSubstring.Text)
            Case Else:
                MsgBox("Internal programming error: invalid ScanMode")
        End Select        
        registry2Form
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Convert radio selections to ALL, NEXTTOKEN or SUBSTRING: repair
    ' erroneous selections
    '
    '
    Private Function radioButtons2Mode As String
        If radScanPromptAll.Checked Then Return("ALL")
        If radScanPromptNext.Checked Then Return("NEXTTOKEN")
        If radScanPromptSubstring.Checked Then Return("SUBSTRING")
        radScanPromptAll.Checked = True
        Return(radioButtons2Mode)
    End Function    

    ' -----------------------------------------------------------------
    ' Retrieve Registry values
    '
    '
    Private Sub registry2Form
        Try
            With radScanPromptAll
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With            
            With radScanPromptNext
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With            
            With radScanPromptSubstring
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With            
            With chkAlways
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With            
        Catch
            MsgBox("Can't obtain Registry values in " & Me.Name & ": " & _
                   Err.Number & " " & Err.Description)
        End Try        
    End Sub 
    
    ' -----------------------------------------------------------------
    ' Radio button text to Always   
    '
    '
    Private Sub text2Always(ByVal strText As String)
        With chkAlways
            chkAlways.Text = Replace(CStr(.Tag), "%ACTION", strText)
            chkAlways.Text = Replace(chkAlways.Text, _
                                     "end of string", _
                                     "9999", _
                                     CompareMethod.Text)
            chkAlways.Text = _OBJutilities.numbers2Variables(chkAlways.Text)
        End With            
    End Sub    

#End Region ' " General procedures "

End Class
