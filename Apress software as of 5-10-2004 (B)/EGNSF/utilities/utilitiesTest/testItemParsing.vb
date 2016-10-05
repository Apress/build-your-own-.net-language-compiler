Option Strict On

' *********************************************************************
' *                                                                   *
' * Test item parsing on behalf of utilitiesTest                      *
' *                                                                   *
' *                                                                   *
' * Tag usage: the txtDelimiter control is tagged with its original   *
' * color. The lblDelimiter control is tagged with its original text. *
' *                                                                   *
' *********************************************************************

Public Class testItemParsing
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
    Friend WithEvents lblInputString As System.Windows.Forms.Label
    Friend WithEvents txtInputString As System.Windows.Forms.TextBox
    Friend WithEvents lblItemNumber As System.Windows.Forms.Label
    Friend WithEvents hscItemNumber As System.Windows.Forms.HScrollBar
    Friend WithEvents gbxDelimiter As System.Windows.Forms.GroupBox
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lstDelimiter As System.Windows.Forms.ListBox
    Friend WithEvents txtDelimiter As System.Windows.Forms.TextBox
    Friend WithEvents chkDelimiterSet As System.Windows.Forms.CheckBox
    Friend WithEvents cmdClearRegistry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdInputStringRandom As System.Windows.Forms.Button
    Friend WithEvents chkFindStartIndex As System.Windows.Forms.CheckBox
    Friend WithEvents cmdStressTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblInputString = New System.Windows.Forms.Label
Me.txtInputString = New System.Windows.Forms.TextBox
Me.lblItemNumber = New System.Windows.Forms.Label
Me.hscItemNumber = New System.Windows.Forms.HScrollBar
Me.gbxDelimiter = New System.Windows.Forms.GroupBox
Me.chkFindStartIndex = New System.Windows.Forms.CheckBox
Me.lstDelimiter = New System.Windows.Forms.ListBox
Me.txtDelimiter = New System.Windows.Forms.TextBox
Me.chkDelimiterSet = New System.Windows.Forms.CheckBox
Me.lblStatus = New System.Windows.Forms.Label
Me.cmdTest = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.cmdClearRegistry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdInputStringRandom = New System.Windows.Forms.Button
Me.cmdStressTest = New System.Windows.Forms.Button
Me.gbxDelimiter.SuspendLayout()
Me.SuspendLayout()
'
'lblInputString
'
Me.lblInputString.BackColor = System.Drawing.Color.Red
Me.lblInputString.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblInputString.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblInputString.ForeColor = System.Drawing.Color.White
Me.lblInputString.Location = New System.Drawing.Point(8, 8)
Me.lblInputString.Name = "lblInputString"
Me.lblInputString.Size = New System.Drawing.Size(512, 20)
Me.lblInputString.TabIndex = 0
Me.lblInputString.Text = "Input String"
Me.lblInputString.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'txtInputString
'
Me.txtInputString.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtInputString.Location = New System.Drawing.Point(8, 28)
Me.txtInputString.Multiline = True
Me.txtInputString.Name = "txtInputString"
Me.txtInputString.Size = New System.Drawing.Size(512, 152)
Me.txtInputString.TabIndex = 1
Me.txtInputString.Text = ""
'
'lblItemNumber
'
Me.lblItemNumber.BackColor = System.Drawing.Color.Red
Me.lblItemNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblItemNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblItemNumber.ForeColor = System.Drawing.Color.White
Me.lblItemNumber.Location = New System.Drawing.Point(8, 192)
Me.lblItemNumber.Name = "lblItemNumber"
Me.lblItemNumber.Size = New System.Drawing.Size(512, 16)
Me.lblItemNumber.TabIndex = 2
Me.lblItemNumber.Text = "Item Number = %ITEM"
Me.lblItemNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'hscItemNumber
'
Me.hscItemNumber.Location = New System.Drawing.Point(8, 208)
Me.hscItemNumber.Minimum = 1
Me.hscItemNumber.Name = "hscItemNumber"
Me.hscItemNumber.Size = New System.Drawing.Size(512, 16)
Me.hscItemNumber.TabIndex = 3
Me.hscItemNumber.Value = 1
'
'gbxDelimiter
'
Me.gbxDelimiter.Controls.Add(Me.chkFindStartIndex)
Me.gbxDelimiter.Controls.Add(Me.lstDelimiter)
Me.gbxDelimiter.Controls.Add(Me.txtDelimiter)
Me.gbxDelimiter.Controls.Add(Me.chkDelimiterSet)
Me.gbxDelimiter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxDelimiter.Location = New System.Drawing.Point(8, 232)
Me.gbxDelimiter.Name = "gbxDelimiter"
Me.gbxDelimiter.Size = New System.Drawing.Size(408, 184)
Me.gbxDelimiter.TabIndex = 4
Me.gbxDelimiter.TabStop = False
Me.gbxDelimiter.Text = "Delimiter"
'
'chkFindStartIndex
'
Me.chkFindStartIndex.Checked = True
Me.chkFindStartIndex.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkFindStartIndex.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkFindStartIndex.Location = New System.Drawing.Point(8, 48)
Me.chkFindStartIndex.Name = "chkFindStartIndex"
Me.chkFindStartIndex.Size = New System.Drawing.Size(120, 16)
Me.chkFindStartIndex.TabIndex = 3
Me.chkFindStartIndex.Text = "Find Start Index"
'
'lstDelimiter
'
Me.lstDelimiter.BackColor = System.Drawing.SystemColors.Control
Me.lstDelimiter.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstDelimiter.ItemHeight = 14
Me.lstDelimiter.Location = New System.Drawing.Point(136, 48)
Me.lstDelimiter.Name = "lstDelimiter"
Me.lstDelimiter.Size = New System.Drawing.Size(264, 130)
Me.lstDelimiter.TabIndex = 2
'
'txtDelimiter
'
Me.txtDelimiter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtDelimiter.Location = New System.Drawing.Point(136, 24)
Me.txtDelimiter.Name = "txtDelimiter"
Me.txtDelimiter.Size = New System.Drawing.Size(264, 20)
Me.txtDelimiter.TabIndex = 1
Me.txtDelimiter.Text = ""
'
'chkDelimiterSet
'
Me.chkDelimiterSet.Checked = True
Me.chkDelimiterSet.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkDelimiterSet.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkDelimiterSet.Location = New System.Drawing.Point(8, 24)
Me.chkDelimiterSet.Name = "chkDelimiterSet"
Me.chkDelimiterSet.Size = New System.Drawing.Size(120, 16)
Me.chkDelimiterSet.TabIndex = 0
Me.chkDelimiterSet.Text = "Set Delimiter"
'
'lblStatus
'
Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatus.Location = New System.Drawing.Point(8, 424)
Me.lblStatus.Name = "lblStatus"
Me.lblStatus.Size = New System.Drawing.Size(512, 80)
Me.lblStatus.TabIndex = 5
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(424, 232)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.Size = New System.Drawing.Size(96, 24)
Me.cmdTest.TabIndex = 6
Me.cmdTest.Text = "Test"
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(424, 392)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(96, 24)
Me.cmdClose.TabIndex = 7
Me.cmdClose.Text = "Close"
'
'cmdClearRegistry
'
Me.cmdClearRegistry.Location = New System.Drawing.Point(424, 328)
Me.cmdClearRegistry.Name = "cmdClearRegistry"
Me.cmdClearRegistry.Size = New System.Drawing.Size(96, 24)
Me.cmdClearRegistry.TabIndex = 25
Me.cmdClearRegistry.Text = "Clear Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(424, 296)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(96, 24)
Me.cmdRegistry2Form.TabIndex = 24
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(424, 264)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(96, 24)
Me.cmdForm2Registry.TabIndex = 23
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdInputStringRandom
'
Me.cmdInputStringRandom.Location = New System.Drawing.Point(416, 8)
Me.cmdInputStringRandom.Name = "cmdInputStringRandom"
Me.cmdInputStringRandom.Size = New System.Drawing.Size(99, 20)
Me.cmdInputStringRandom.TabIndex = 26
Me.cmdInputStringRandom.Text = "Random"
'
'cmdStressTest
'
Me.cmdStressTest.Location = New System.Drawing.Point(424, 360)
Me.cmdStressTest.Name = "cmdStressTest"
Me.cmdStressTest.Size = New System.Drawing.Size(96, 24)
Me.cmdStressTest.TabIndex = 27
Me.cmdStressTest.Text = "Stress Test"
'
'testItemParsing
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(526, 507)
Me.ControlBox = False
Me.Controls.Add(Me.cmdStressTest)
Me.Controls.Add(Me.cmdInputStringRandom)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.lblStatus)
Me.Controls.Add(Me.gbxDelimiter)
Me.Controls.Add(Me.hscItemNumber)
Me.Controls.Add(Me.lblItemNumber)
Me.Controls.Add(Me.txtInputString)
Me.Controls.Add(Me.lblInputString)
Me.Controls.Add(Me.cmdClearRegistry)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "testItemParsing"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "testItemParsing"
Me.gbxDelimiter.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form Data "

    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJzoomInterface As zoomInterface
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities

    Private Const BLANK_DELIMITER As String = "(Blank)"
    Private Const NEWLINE_DELIMITER As String = "(Newline)"
    Private Const WHITESPACE_DELIMITER As String = "(White space)"
    Private Const NEWLINE_DELIMITER_XML As String = "&#00013&#00010"

#End Region ' " Form Data "

#Region " Form events "

    Private Sub cmdClearRegistry_Click(ByVal objSender As System.Object, ByVal objEvent As System.EventArgs) Handles cmdClearRegistry.Click
        clearRegistry()
    End Sub

    Private Sub cmdClose_Click(ByVal objSender As System.Object, _
                               ByVal objEvent As System.EventArgs) _
            Handles cmdClose.Click
        form2Registry()
        Hide()
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal objSender As System.Object, ByVal objEvent As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry()
    End Sub

    Private Sub cmdInputStringRandom_Click(ByVal objSender As System.Object, ByVal objEvent As System.EventArgs) Handles cmdInputStringRandom.Click
        mkRandomInputString()
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal objSender As System.Object, ByVal objEvent As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form()
    End Sub

    Private Sub cmdStressTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStressTest.Click
        stressTest()
    End Sub

    Private Sub cmdTest_Click(ByVal objSender As System.Object, _
                              ByVal objEvent As System.EventArgs) _
            Handles cmdTest.Click
        test()
    End Sub

    Private Sub hscItemNumber_Scroll(ByVal objSender As System.Object, _
                                     ByVal e As System.Windows.Forms.ScrollEventArgs) _
            Handles hscItemNumber.Scroll
        itemNumber2Label()
    End Sub

    Private Sub lstDelimiter_DoubleClick(ByVal objSender As System.Object, _
                                         ByVal objEvent As System.EventArgs) _
                Handles lstDelimiter.DoubleClick
        With lstDelimiter
            txtDelimiter.Text = CStr(.Items.Item(.SelectedIndex))
        End With
    End Sub

    Private Sub testItemParsing_Load(ByVal objSender As System.Object, _
                                     ByVal objEvent As System.EventArgs) _
            Handles MyBase.Load
        With txtDelimiter
            .Tag = .BackColor
        End With
        With txtInputString
            .SelectionStart = 0 : .SelectionLength = 0
        End With
        registry2Form()
        With txtInputString
            .SelectionStart = 0 : .SelectionLength = 0
        End With
    End Sub

    Private Sub txtDelimiter_Leave(ByVal objSender As System.Object, _
                                   ByVal objEvent As System.EventArgs) _
            Handles txtDelimiter.Leave
        If _OBJwindowsUtilities.searchListBox(lstDelimiter, _
                                              txtDelimiter.Text) _
           = _
           -1 Then
            lstDelimiter.Items.Add(txtDelimiter.Text)
        End If
    End Sub

#End Region ' " Form events "

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Adjust settings based on the chkDelimiterSet
    '
    '
    Private Sub adjustDelimiterSet()
        If Not chkDelimiterSet.Checked Then
            selectDelimiter(BLANK_DELIMITER)
        End If
    End Sub

    ' -----------------------------------------------------------------
    ' Clear the Registry
    '
    '
    Private Function clearRegistry() As Boolean
        Try
            DeleteSetting(Application.ProductName, Me.Name)
        Catch
            Return (False)
        End Try
        Return (registry2Form())
    End Function

    ' -----------------------------------------------------------------
    ' Save form settings in the Registry
    '
    '
    Private Function form2Registry() As Boolean
        Try
            With txtInputString
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            .Text)
            End With
            With txtDelimiter
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            .Text)
            End With
            With chkDelimiterSet
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            With chkFindStartIndex
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            With txtDelimiter
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            .Text)
            End With
            With hscItemNumber
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Value))
            End With
            _OBJwindowsUtilities.listBox2Registry(lstDelimiter, _
                                                  strApplication:=Application.ProductName, _
                                                  strSection:=Me.Name)
        Catch
            Return (False)
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Initialize delimiter list
    '
    '
    Private Sub initializeDelimiters()
        With lstDelimiter
            .Items.Add(BLANK_DELIMITER)
            .Items.Add(NEWLINE_DELIMITER)
            .Items.Add(WHITESPACE_DELIMITER)
            .Items.Add(NEWLINE_DELIMITER_XML)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Return True if delimiter is one of the original set
    '
    '
    Private Function isInitialDelimiter(ByVal strDelimiter As String) _
            As Boolean
        Select Case strDelimiter
            Case BLANK_DELIMITER : Return (True)
            Case NEWLINE_DELIMITER : Return (True)
            Case WHITESPACE_DELIMITER : Return (True)
            Case NEWLINE_DELIMITER_XML : Return (True)
        End Select
        Return (False)
    End Function

    ' -----------------------------------------------------------------
    ' Return True if delimiter is one of the original set
    '
    '
    Private Function isParenthesizedInitialDelimiter(ByVal strDelimiter As String) _
            As Boolean
        Select Case strDelimiter
            Case BLANK_DELIMITER : Return (True)
            Case NEWLINE_DELIMITER : Return (True)
            Case WHITESPACE_DELIMITER : Return (True)
        End Select
        Return (False)
    End Function

    ' -----------------------------------------------------------------
    ' Item number to label
    '
    '
    Private Sub itemNumber2Label()
        With lblItemNumber
            If (.Tag Is Nothing) Then
                .Tag = .Text
            End If
            .Text = Replace(CStr(.Tag), "%ITEM", CStr(hscItemNumber.Value))
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Make a random input string
    '
    '
    Private Sub mkRandomInputString()
        Dim strOutstring As String
        Dim intIndex1 As Integer
        Dim strAscii As String
        With _OBJutilities
            strAscii = .copies("ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                "ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
                                "0123456789", _
                               32) & _
                       .range2String(0, 255) & _
                       .copies(" ", 255)
            For intIndex1 = 1 To CInt(Rnd() * 255)
                strOutstring &= Mid(strAscii, Math.Max(CInt(Rnd() * Len(strAscii)), 1), 1)
            Next intIndex1
            txtInputString.Text = .string2Display(strOutstring)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Set form settings from the Registry
    '
    '
    Private Function registry2Form() As Boolean
        Try
            With txtInputString
                .Text = CStr(GetSetting(Application.ProductName, _
                                        Me.Name, _
                                        .Name, _
                                        _OBJutilities.smallBusinessList))
                adjustDelimiterSet()
            End With
            With txtDelimiter
                txtDelimiter.Text = CStr(GetSetting(Application.ProductName, _
                                                    Me.Name, _
                                                    .Name, _
                                                    .Text))
            End With
            With chkDelimiterSet
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
                adjustDelimiterSet()
            End With
            With chkFindStartIndex
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With hscItemNumber
                .Value = CInt(GetSetting(Application.ProductName, _
                                          Me.Name, _
                                          .Name, _
                                          CStr(.Value)))
                itemNumber2Label()
            End With
            _OBJwindowsUtilities.registry2ListBox(lstDelimiter, _
                                                  strApplication:=Application.ProductName, _
                                                  strSection:=Me.Name)
            If lstDelimiter.Items.Count = 0 Then
                initializeDelimiters()
            End If
        Catch
            Return (False)
        End Try
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Select the specified delimiter
    '
    '
    Private Function selectDelimiter(ByVal strDelimiter As String) As Boolean
        With lstDelimiter
            .SelectedIndex = _OBJwindowsUtilities.searchListBox(lstDelimiter, _
                                                                strDelimiter)
            If .SelectedIndex >= 0 Then
                txtDelimiter.Text = CStr(.Items.Item(.SelectedIndex))
            End If
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Stress test
    '
    '
    Private Function stressTest() As Boolean
        Dim strReport As String
        Dim booOK As Boolean = _OBJutilities.itemTest(strReport)
        Select Case MsgBox("Stress test has " & _
                            CStr(IIf(booOK, "succeeded", "failed")) & ": " & _
                            "Click OK to see report: click Cancel to return to form", _
                            MsgBoxStyle.OKCancel)
            Case MsgBoxResult.OK
                _OBJzoomInterface.zoomInterface(Me, strReport, dblHeight:=1.0)
            Case MsgBoxResult.Cancel
            Case MsgBoxResult.OK
                _OBJutilities.errorHandler("Unexpected reply to MsgBox", _
                                           Me.Name, _
                                           "stressTest")
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Test the item parser
    '
    '
    Private Sub test()
        If txtDelimiter.Text = "" Then
            MsgBox("No delimiter has been selected")
            Return
        End If
        With _OBJutilities
            Dim booEnabled As Boolean = Me.Enabled
            Me.Enabled = False
            Dim intStartIndex As Integer
            Dim strError As String
            Dim strItem As String
            Dim strResult As String
            Dim strInstring As String = .display2String(txtInputString.Text, "XML")
            Dim strDelimiter As String = textBox2Delimiter()
            Try
                If chkFindStartIndex.Checked Then
                    strItem = _OBJutilities.item(strInstring, _
                                                    hscItemNumber.Value, _
                                                    strDelimiter, _
                                                    chkDelimiterSet.Checked, _
                                                    intStartIndex)
                    With txtInputString
                        If intStartIndex = 0 Then
                            .SelectionStart = 0 : .SelectionLength = 0
                        Else
                            .SelectionStart = intStartIndex - 1
                            .SelectionLength = Len(strItem)
                        End If
                        .Refresh()
                        .Focus()
                    End With
                Else
                    strItem = _OBJutilities.item(strInstring, _
                                                 hscItemNumber.Value, _
                                                 strDelimiter, _
                                                 chkDelimiterSet.Checked)
                End If
                strResult = .enquote(.string2Display(strItem)) & _
                            CStr(IIf(chkFindStartIndex.Checked, _
                                     vbNewLine & vbNewLine & _
                                     "The item occurs at " & intStartIndex, _
                                     "")) & _
                            vbNewLine & vbNewLine & _
                            "The number of items is " & _
                            .items(strInstring, _
                                   strDelimiter, _
                                   chkDelimiterSet.Checked)
            Catch objException As Exception
                strResult = "the error: " & Err.Number & " " & Err.Description
            End Try
            lblStatus.Text = _
                "The " & _
                CStr(IIf(chkDelimiterSet.Checked, _
                         "set-based", _
                         "string-based")) & " " & _
                "parse of the string " & _
                _OBJutilities.enquote _
                (_OBJutilities.ellipsis(_OBJutilities.string2Display(txtInputString.Text), _
                                        32)) & " " & _
                "using the item number " & _
                hscItemNumber.Value & " is " & _
                strResult
            Me.Enabled = booEnabled
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Convert text box to delimiter value
    '
    '
    Private Function textBox2Delimiter() As String
        Select Case txtDelimiter.Text
            Case BLANK_DELIMITER : Return (" ")
            Case NEWLINE_DELIMITER : Return (vbNewLine)
            Case WHITESPACE_DELIMITER : Return (_OBJutilities.range2String(0, 32))
            Case NEWLINE_DELIMITER_XML : Return (_OBJutilities.string2Display(vbNewLine))
        End Select
        Return (_OBJutilities.display2String(txtDelimiter.Text, "XML"))
    End Function

#End Region ' " General procedures "

End Class
