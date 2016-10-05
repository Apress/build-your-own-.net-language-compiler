Option Strict On

' *********************************************************************
' *                                                                   *
' * Display the event log                                             *
' *                                                                   *
' * This form exposes the write-only QuickBasicEngine property which  *
' * should be set to the Quick Basic engine whose event log needs to  *
' * be displayed.                                                     *
' *                                                                   *
' *********************************************************************

Public Class eventLogFormat
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
        MyBase.dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblStartIndex As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdDisplay As System.Windows.Forms.Button
    Friend WithEvents hscStartIndex As System.Windows.Forms.HScrollBar
    Friend WithEvents hscCount As System.Windows.Forms.HScrollBar
    Friend WithEvents lblCount As System.Windows.Forms.Label
    Friend WithEvents txtEventLog As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblStartIndex = New System.Windows.Forms.Label
Me.hscStartIndex = New System.Windows.Forms.HScrollBar
Me.hscCount = New System.Windows.Forms.HScrollBar
Me.lblCount = New System.Windows.Forms.Label
Me.cmdClose = New System.Windows.Forms.Button
Me.txtEventLog = New System.Windows.Forms.TextBox
Me.cmdDisplay = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblStartIndex
'
Me.lblStartIndex.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
Me.lblStartIndex.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStartIndex.ForeColor = System.Drawing.Color.White
Me.lblStartIndex.Location = New System.Drawing.Point(8, 8)
Me.lblStartIndex.Name = "lblStartIndex"
Me.lblStartIndex.Size = New System.Drawing.Size(192, 16)
Me.lblStartIndex.TabIndex = 0
Me.lblStartIndex.Text = "Start Index: %STARTINDEX"
Me.lblStartIndex.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'hscStartIndex
'
Me.hscStartIndex.Location = New System.Drawing.Point(8, 24)
Me.hscStartIndex.Name = "hscStartIndex"
Me.hscStartIndex.Size = New System.Drawing.Size(192, 16)
Me.hscStartIndex.TabIndex = 1
'
'hscCount
'
Me.hscCount.Location = New System.Drawing.Point(200, 24)
Me.hscCount.Name = "hscCount"
Me.hscCount.Size = New System.Drawing.Size(192, 16)
Me.hscCount.TabIndex = 3
'
'lblCount
'
Me.lblCount.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
Me.lblCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblCount.ForeColor = System.Drawing.Color.White
Me.lblCount.Location = New System.Drawing.Point(200, 8)
Me.lblCount.Name = "lblCount"
Me.lblCount.Size = New System.Drawing.Size(192, 16)
Me.lblCount.TabIndex = 2
Me.lblCount.Text = "Count: %COUNT"
Me.lblCount.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(696, 8)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(88, 32)
Me.cmdClose.TabIndex = 4
Me.cmdClose.Text = "Close"
'
'txtEventLog
'
Me.txtEventLog.BackColor = System.Drawing.SystemColors.Control
Me.txtEventLog.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtEventLog.Location = New System.Drawing.Point(8, 48)
Me.txtEventLog.Multiline = True
Me.txtEventLog.Name = "txtEventLog"
Me.txtEventLog.ReadOnly = True
Me.txtEventLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
Me.txtEventLog.Size = New System.Drawing.Size(776, 368)
Me.txtEventLog.TabIndex = 5
Me.txtEventLog.Text = ""
'
'cmdDisplay
'
Me.cmdDisplay.Location = New System.Drawing.Point(400, 8)
Me.cmdDisplay.Name = "cmdDisplay"
Me.cmdDisplay.Size = New System.Drawing.Size(88, 32)
Me.cmdDisplay.TabIndex = 6
Me.cmdDisplay.Text = "Display"
'
'eventLogFormat
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(790, 421)
Me.ControlBox = False
Me.Controls.Add(Me.cmdDisplay)
Me.Controls.Add(Me.txtEventLog)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.hscCount)
Me.Controls.Add(Me.lblCount)
Me.Controls.Add(Me.hscStartIndex)
Me.Controls.Add(Me.lblStartIndex)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "eventLogFormat"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "Event Log"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private OBJqbe As QuickBasicEngine.qbQuickBasicEngine

#End Region ' " Form data "

#Region " Public properties "

    Public WriteOnly Property QuickBasicEngine() As QuickBasicEngine.qbQuickBasicEngine
        Set(ByVal objNewValue As QuickBasicEngine.qbQuickBasicEngine)
            OBJqbe = objNewValue
        End Set
    End Property

#End Region ' " Public properties "

#Region " Form events "

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Hide()
    End Sub

    Private Sub cmdDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDisplay.Click
        displayEventLog()
    End Sub

    Private Sub eventLogFormat_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If (OBJqbe Is Nothing) Then
            utilities.utilities.errorHandler("Quick Basic engine is not assigned", _
                                             Me.Name, _
                                             "eventLogFormat_Load")
            Hide()
        End If
        If Not OBJqbe.EventLogging Then
            utilities.utilities.errorHandler("No event log exists", _
                                             Me.Name, _
                                             "eventLogFormat_Load")
            Hide()
        End If
        If OBJqbe.EventLog.Count = 0 Then
            hscStartIndex.Visible = False
            lblStartIndex.Visible = False
            hscCount.Visible = False
            lblCount.Visible = False
            cmdDisplay.Visible = False
            txtEventLog.Text = "Event log is empty"
        Else
            hscStartIndex.Visible = True
            lblStartIndex.Visible = True
            hscCount.Visible = True
            lblCount.Visible = True
            cmdDisplay.Visible = True
            With hscStartIndex
                .Minimum = 1
                .Maximum = Math.Max(100, OBJqbe.EventLog.Count)
                If .Maximum > .Minimum Then .Value = 1
            End With
            With hscCount
                .Minimum = 0
                .Maximum = hscStartIndex.Maximum
                If .Maximum > .Minimum Then .Value = .Maximum
            End With
            scrollBox2Label(hscStartIndex, "%STARTINDEX", lblStartIndex)
            scrollBox2Label(hscCount, "%COUNT", lblCount)
        End If
    End Sub

    Private Sub hscCount_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscCount.Scroll
        scrollBox2Label(hscCount, "%COUNT", lblCount)
    End Sub

    Private Sub hscStartIndex_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscStartIndex.Scroll
        scrollBox2Label(hscStartIndex, "%STARTINDEX", lblStartIndex)
    End Sub

#End Region ' " Form events "

#Region " General procedures "

    ' ---------------------------------------------------------------------
    ' Display the event log
    '
    '
    Private Sub displayEventLog()
        AddHandler OBJqbe.loopEvent, AddressOf loopEventHandler
        With txtEventLog
            .Text = ""
            .Refresh()
            .Text = OBJqbe.eventLogFormat(hscStartIndex.Value, _
                                          Math.Min(hscCount.Value, _
                                                   OBJqbe.EventLog.Count))
            .Refresh()
        End With
        RemoveHandler OBJqbe.loopEvent, AddressOf loopEventHandler
    End Sub

    ' ---------------------------------------------------------------------
    ' Scroll box value to label 
    '
    '
    Private Sub scrollBox2Label(ByVal hscScrollBox As HScrollBar, _
                                ByVal strKeyword As String, _
                                ByVal lblLabel As Label)
        With lblLabel
            If (.Tag Is Nothing) Then .Tag = .Text
            With hscScrollBox
                If .Maximum - .Value < (.Maximum - .Minimum) * 0.1 Then
                    .Value = .Maximum
                End If
            End With
            .Text = Replace(CStr(.Tag), strKeyword, CStr(hscScrollBox.Value))
        End With
    End Sub


#End Region ' " General procedures "

#Region " Event handlers "

    ' -------------------------------------------------------------------------
    ' loopEvent
    '
    '
    Private Sub loopEventHandler(ByVal objQBsender As QuickBasicEngine.qbQuickBasicEngine, _
                                 ByVal strActivity As String, _
                                 ByVal strEntity As String, _
                                 ByVal intNumber As Integer, _
                                 ByVal intCount As Integer, _
                                 ByVal intLevel As Integer, _
                                 ByVal strComment As String)
        With txtEventLog
            Dim intSelectionStart As Integer = Len(.Text)
            Dim strMessage As String = _
                    CStr(IIf(.Text = "", "", vbNewLine)) & _
                    utilities.utilities.copies(intLevel, " ") & _
                    strActivity & " at " & strEntity & intNumber & " of " & intCount & ": " & _
                    strComment
            .Text = .Text & _
                    CStr(IIf(.Text = "", "", vbNewLine)) & _
                    strMessage
            .SelectionStart = intSelectionStart
            .SelectionLength = Len(strMessage)
            .ScrollToCaret()
            .Focus()
            .Refresh()
            If intNumber = intCount Then
                .Text = "" : .Refresh()
            End If
        End With
    End Sub

#End Region

End Class
