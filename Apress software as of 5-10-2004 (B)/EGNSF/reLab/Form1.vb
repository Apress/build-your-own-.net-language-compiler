Option Strict

' *********************************************************************
' *                                                                   *
' * reLab     Regular expression laboratory                           *
' *                                                                   *
' *                                                                   *
' * The following describes the uses of tags in this form.            *
' *                                                                   *
' *                                                                   *
' *      *  cmdREavailableDelete.Tag: String: contains the original   *
' *         Text of the command button including the substitution     *
' *         keyword.                                                  *
' *                                                                   *
' *      *  cmdTestDataAvailableDelete.Tag: String: contains the      *
' *         original Text of the command button including the         *
' *         substitution keyword.                                     *
' *                                                                   *
' *      *  lstREavailable.Tag: Collection: contains a collection of  *
' *         the standard regular expressions, that are added on       *
' *         initial load and which cannot be removed.                 *
' *                                                                   *
' *      *  lstTestData.Tag: Collection: contains a collection of     *
' *         the standard test strings, that are added on initial load *
' *         and which cannot be removed.                              *
' *                                                                   *
' *      *  txtRE.Tag: Boolean: set to False on entering the control: *
' *         set to True on leaving the control: tested to see whether *
' *         the regular expression list box should be updated.        *
' *                                                                   *
' *      *  txtTestData.Tag: Boolean: set to False on entering the    *
' *         control: set to True on leaving the control: tested to see*
' *         whether the test data list box should be updated.         *
' *                                                                   *
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblRE As System.Windows.Forms.Label
    Friend WithEvents txtRE As System.Windows.Forms.TextBox
    Friend WithEvents lblREavailable As System.Windows.Forms.Label
    Friend WithEvents lblTestData As System.Windows.Forms.Label
    Friend WithEvents txtTestData As System.Windows.Forms.TextBox
    Friend WithEvents lblTestDataAvailable As System.Windows.Forms.Label
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lstREavailable As System.Windows.Forms.ListBox
    Friend WithEvents lstTestData As System.Windows.Forms.ListBox
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    Friend WithEvents cmdREavailableDelete As System.Windows.Forms.Button
    Friend WithEvents cmdTestDataAvailableDelete As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsPromptCustomizationRE As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsPromptCustomizationSTR As System.Windows.Forms.MenuItem
    Friend WithEvents cmdREsave As System.Windows.Forms.Button
    Friend WithEvents cmdTestDataSave As System.Windows.Forms.Button
    Friend WithEvents lblVB As System.Windows.Forms.Label
    Friend WithEvents txtVB As System.Windows.Forms.TextBox
    Friend WithEvents cmdVBzoom As System.Windows.Forms.Button
    Friend WithEvents cmdTestDataZoom As System.Windows.Forms.Button
    Friend WithEvents cmdTestDataNext As System.Windows.Forms.Button
    Friend WithEvents lblREexplanation As System.Windows.Forms.Label
    Friend WithEvents lblTestDataExplanation As System.Windows.Forms.Label
    Friend WithEvents cmdCommonRETest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblRE = New System.Windows.Forms.Label
Me.txtRE = New System.Windows.Forms.TextBox
Me.lblREavailable = New System.Windows.Forms.Label
Me.lstREavailable = New System.Windows.Forms.ListBox
Me.lblTestData = New System.Windows.Forms.Label
Me.txtTestData = New System.Windows.Forms.TextBox
Me.lstTestData = New System.Windows.Forms.ListBox
Me.lblTestDataAvailable = New System.Windows.Forms.Label
Me.cmdTest = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.cmdAbout = New System.Windows.Forms.Button
Me.cmdREavailableDelete = New System.Windows.Forms.Button
Me.cmdTestDataAvailableDelete = New System.Windows.Forms.Button
Me.MainMenu1 = New System.Windows.Forms.MainMenu
Me.mnuFile = New System.Windows.Forms.MenuItem
Me.mnuFileExit = New System.Windows.Forms.MenuItem
Me.mnuTools = New System.Windows.Forms.MenuItem
Me.mnuToolsPromptCustomizationRE = New System.Windows.Forms.MenuItem
Me.mnuToolsPromptCustomizationSTR = New System.Windows.Forms.MenuItem
Me.mnuHelp = New System.Windows.Forms.MenuItem
Me.mnuHelpAbout = New System.Windows.Forms.MenuItem
Me.cmdREsave = New System.Windows.Forms.Button
Me.cmdTestDataSave = New System.Windows.Forms.Button
Me.lblVB = New System.Windows.Forms.Label
Me.txtVB = New System.Windows.Forms.TextBox
Me.cmdVBzoom = New System.Windows.Forms.Button
Me.cmdTestDataZoom = New System.Windows.Forms.Button
Me.cmdTestDataNext = New System.Windows.Forms.Button
Me.lblREexplanation = New System.Windows.Forms.Label
Me.lblTestDataExplanation = New System.Windows.Forms.Label
Me.cmdCommonRETest = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblRE
'
Me.lblRE.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(144, Byte))
Me.lblRE.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblRE.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblRE.Location = New System.Drawing.Point(10, 9)
Me.lblRE.Name = "lblRE"
Me.lblRE.Size = New System.Drawing.Size(990, 23)
Me.lblRE.TabIndex = 0
Me.lblRE.Text = "Regular Expression"
Me.lblRE.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'txtRE
'
Me.txtRE.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtRE.Location = New System.Drawing.Point(10, 92)
Me.txtRE.Name = "txtRE"
Me.txtRE.Size = New System.Drawing.Size(990, 23)
Me.txtRE.TabIndex = 1
Me.txtRE.Text = ""
'
'lblREavailable
'
Me.lblREavailable.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(144, Byte))
Me.lblREavailable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblREavailable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblREavailable.Location = New System.Drawing.Point(10, 120)
Me.lblREavailable.Name = "lblREavailable"
Me.lblREavailable.Size = New System.Drawing.Size(518, 23)
Me.lblREavailable.TabIndex = 2
Me.lblREavailable.Text = "Regular Expressions Available: double-click to select"
Me.lblREavailable.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'lstREavailable
'
Me.lstREavailable.BackColor = System.Drawing.SystemColors.Control
Me.lstREavailable.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstREavailable.ItemHeight = 17
Me.lstREavailable.Location = New System.Drawing.Point(10, 143)
Me.lstREavailable.Name = "lstREavailable"
Me.lstREavailable.Size = New System.Drawing.Size(518, 72)
Me.lstREavailable.TabIndex = 3
'
'lblTestData
'
Me.lblTestData.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(144, Byte))
Me.lblTestData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblTestData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblTestData.Location = New System.Drawing.Point(10, 231)
Me.lblTestData.Name = "lblTestData"
Me.lblTestData.Size = New System.Drawing.Size(518, 23)
Me.lblTestData.TabIndex = 4
Me.lblTestData.Text = "Test Data"
Me.lblTestData.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'txtTestData
'
Me.txtTestData.BackColor = System.Drawing.Color.FromArgb(CType(232, Byte), CType(232, Byte), CType(232, Byte))
Me.txtTestData.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtTestData.Location = New System.Drawing.Point(10, 295)
Me.txtTestData.Multiline = True
Me.txtTestData.Name = "txtTestData"
Me.txtTestData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
Me.txtTestData.Size = New System.Drawing.Size(518, 93)
Me.txtTestData.TabIndex = 5
Me.txtTestData.Text = ""
'
'lstTestData
'
Me.lstTestData.BackColor = System.Drawing.SystemColors.Control
Me.lstTestData.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstTestData.ItemHeight = 17
Me.lstTestData.Location = New System.Drawing.Point(8, 420)
Me.lstTestData.Name = "lstTestData"
Me.lstTestData.Size = New System.Drawing.Size(992, 72)
Me.lstTestData.TabIndex = 7
'
'lblTestDataAvailable
'
Me.lblTestDataAvailable.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(144, Byte))
Me.lblTestDataAvailable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblTestDataAvailable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblTestDataAvailable.Location = New System.Drawing.Point(8, 397)
Me.lblTestDataAvailable.Name = "lblTestDataAvailable"
Me.lblTestDataAvailable.Size = New System.Drawing.Size(992, 23)
Me.lblTestDataAvailable.TabIndex = 6
Me.lblTestDataAvailable.Text = "Test Data Strings Available: double-click to select"
Me.lblTestDataAvailable.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(8, 528)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.Size = New System.Drawing.Size(128, 32)
Me.cmdTest.TabIndex = 8
Me.cmdTest.Text = "Test"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(144, 528)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(128, 32)
Me.cmdForm2Registry.TabIndex = 9
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(280, 528)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(128, 32)
Me.cmdRegistry2Form.TabIndex = 10
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(872, 528)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(128, 32)
Me.cmdClose.TabIndex = 11
Me.cmdClose.Text = "Close"
'
'cmdAbout
'
Me.cmdAbout.Location = New System.Drawing.Point(416, 528)
Me.cmdAbout.Name = "cmdAbout"
Me.cmdAbout.Size = New System.Drawing.Size(128, 32)
Me.cmdAbout.TabIndex = 12
Me.cmdAbout.Text = "About"
'
'cmdREavailableDelete
'
Me.cmdREavailableDelete.Location = New System.Drawing.Point(365, 120)
Me.cmdREavailableDelete.Name = "cmdREavailableDelete"
Me.cmdREavailableDelete.Size = New System.Drawing.Size(163, 23)
Me.cmdREavailableDelete.TabIndex = 13
Me.cmdREavailableDelete.Text = "Delete %ENTRY"
Me.cmdREavailableDelete.Visible = False
'
'cmdTestDataAvailableDelete
'
Me.cmdTestDataAvailableDelete.Location = New System.Drawing.Point(840, 398)
Me.cmdTestDataAvailableDelete.Name = "cmdTestDataAvailableDelete"
Me.cmdTestDataAvailableDelete.Size = New System.Drawing.Size(160, 23)
Me.cmdTestDataAvailableDelete.TabIndex = 14
Me.cmdTestDataAvailableDelete.Text = "Delete %ENTRY"
Me.cmdTestDataAvailableDelete.Visible = False
'
'MainMenu1
'
Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuTools, Me.mnuHelp})
'
'mnuFile
'
Me.mnuFile.Index = 0
Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileExit})
Me.mnuFile.Text = "&File"
'
'mnuFileExit
'
Me.mnuFileExit.Index = 0
Me.mnuFileExit.Text = "E&xit"
'
'mnuTools
'
Me.mnuTools.Index = 1
Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuToolsPromptCustomizationRE, Me.mnuToolsPromptCustomizationSTR})
Me.mnuTools.Text = "Tools"
'
'mnuToolsPromptCustomizationRE
'
Me.mnuToolsPromptCustomizationRE.Index = 0
Me.mnuToolsPromptCustomizationRE.Text = "Add Prompt customization for regular expressions..."
'
'mnuToolsPromptCustomizationSTR
'
Me.mnuToolsPromptCustomizationSTR.Index = 1
Me.mnuToolsPromptCustomizationSTR.Text = "Add Prompt customization for test data strings..."
'
'mnuHelp
'
Me.mnuHelp.Index = 2
Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpAbout})
Me.mnuHelp.Text = "Help"
'
'mnuHelpAbout
'
Me.mnuHelpAbout.Index = 0
Me.mnuHelpAbout.Text = "About..."
'
'cmdREsave
'
Me.cmdREsave.Location = New System.Drawing.Point(928, 9)
Me.cmdREsave.Name = "cmdREsave"
Me.cmdREsave.Size = New System.Drawing.Size(72, 23)
Me.cmdREsave.TabIndex = 15
Me.cmdREsave.Text = "Save"
'
'cmdTestDataSave
'
Me.cmdTestDataSave.Location = New System.Drawing.Point(461, 231)
Me.cmdTestDataSave.Name = "cmdTestDataSave"
Me.cmdTestDataSave.Size = New System.Drawing.Size(67, 23)
Me.cmdTestDataSave.TabIndex = 16
Me.cmdTestDataSave.Text = "Save"
'
'lblVB
'
Me.lblVB.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(128, Byte), CType(144, Byte))
Me.lblVB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblVB.Location = New System.Drawing.Point(528, 120)
Me.lblVB.Name = "lblVB"
Me.lblVB.Size = New System.Drawing.Size(472, 23)
Me.lblVB.TabIndex = 17
Me.lblVB.Text = "Visual Basic code"
Me.lblVB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'txtVB
'
Me.txtVB.BackColor = System.Drawing.SystemColors.Control
Me.txtVB.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtVB.Location = New System.Drawing.Point(528, 143)
Me.txtVB.Multiline = True
Me.txtVB.Name = "txtVB"
Me.txtVB.Size = New System.Drawing.Size(472, 245)
Me.txtVB.TabIndex = 18
Me.txtVB.Text = ""
'
'cmdVBzoom
'
Me.cmdVBzoom.Location = New System.Drawing.Point(928, 120)
Me.cmdVBzoom.Name = "cmdVBzoom"
Me.cmdVBzoom.Size = New System.Drawing.Size(72, 23)
Me.cmdVBzoom.TabIndex = 19
Me.cmdVBzoom.Text = "Zoom"
'
'cmdTestDataZoom
'
Me.cmdTestDataZoom.Location = New System.Drawing.Point(394, 231)
Me.cmdTestDataZoom.Name = "cmdTestDataZoom"
Me.cmdTestDataZoom.Size = New System.Drawing.Size(67, 23)
Me.cmdTestDataZoom.TabIndex = 20
Me.cmdTestDataZoom.Text = "Zoom"
'
'cmdTestDataNext
'
Me.cmdTestDataNext.Location = New System.Drawing.Point(326, 231)
Me.cmdTestDataNext.Name = "cmdTestDataNext"
Me.cmdTestDataNext.Size = New System.Drawing.Size(68, 23)
Me.cmdTestDataNext.TabIndex = 21
Me.cmdTestDataNext.Text = "Next"
'
'lblREexplanation
'
Me.lblREexplanation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblREexplanation.Location = New System.Drawing.Point(10, 32)
Me.lblREexplanation.Name = "lblREexplanation"
Me.lblREexplanation.Size = New System.Drawing.Size(990, 51)
Me.lblREexplanation.TabIndex = 22
'
'lblTestDataExplanation
'
Me.lblTestDataExplanation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblTestDataExplanation.Location = New System.Drawing.Point(10, 256)
Me.lblTestDataExplanation.Name = "lblTestDataExplanation"
Me.lblTestDataExplanation.Size = New System.Drawing.Size(518, 41)
Me.lblTestDataExplanation.TabIndex = 23
'
'cmdCommonRETest
'
Me.cmdCommonRETest.Location = New System.Drawing.Point(552, 528)
Me.cmdCommonRETest.Name = "cmdCommonRETest"
Me.cmdCommonRETest.Size = New System.Drawing.Size(264, 32)
Me.cmdCommonRETest.TabIndex = 24
Me.cmdCommonRETest.Text = "Test the common regular expressions"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
Me.ClientSize = New System.Drawing.Size(1006, 561)
Me.Controls.Add(Me.cmdCommonRETest)
Me.Controls.Add(Me.lblTestDataExplanation)
Me.Controls.Add(Me.lblREexplanation)
Me.Controls.Add(Me.cmdTestDataNext)
Me.Controls.Add(Me.cmdTestDataZoom)
Me.Controls.Add(Me.cmdVBzoom)
Me.Controls.Add(Me.txtVB)
Me.Controls.Add(Me.lblVB)
Me.Controls.Add(Me.cmdTestDataSave)
Me.Controls.Add(Me.cmdREsave)
Me.Controls.Add(Me.cmdTestDataAvailableDelete)
Me.Controls.Add(Me.cmdREavailableDelete)
Me.Controls.Add(Me.cmdAbout)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.lstTestData)
Me.Controls.Add(Me.lblTestDataAvailable)
Me.Controls.Add(Me.txtTestData)
Me.Controls.Add(Me.lblTestData)
Me.Controls.Add(Me.lstREavailable)
Me.Controls.Add(Me.lblREavailable)
Me.Controls.Add(Me.txtRE)
Me.Controls.Add(Me.lblRE)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Menu = Me.MainMenu1
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "reLab"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    ' ***** Shared utilities *****
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities

    ' ***** Constants *****
    ' --- Separator
    Private Const LISTBOX_SEPARATOR As String = ": "
    ' --- About info  
    Private Const ABOUTINFO As String = _
    "reLab   Regular expression testing laboratory" & vbNewLine & vbNewLine & _
    "This application and form allows the user to test, save and document regular expressions. " & _
    "It provides access to a set of prewritten and pretested regular expressions for common tasks " & _
    "including parsing Visual Basic code and platform-independent newline detection." & _
    vbNewLine & vbNewLine & _
    "This application helps to overcome a problem of regular expressions, and this is their " & _
    "gnomic and obscure appearance to programmers without a unix or perl background. Even unix and " & _
    "perl programmers, however, will find this tool useful when using Visual Basic regular expressions, " & _
    "because it allows the programmer to organize and test complex expressions." & _
    vbNewLine & vbNewLine & _
    "This class was developed commencing on 5/18/2003 by" & _
    vbNewLine & vbNewLine & _
    "Edward G. Nilges" & vbNewLine & _
    "spinoza1111@yahoo.COM" & vbNewLine & _
    "http://members.screenz.com/edNilges"
    ' --- Form tolerance
    Private Const HEIGHT_TOLERANCE As Single = 0.75
    Private Const WIDTH_TOLERANCE As Single = 0.9

    ' ***** Add policy forms *****
    Private FRMaddPromptRE As addPrompt
    Private FRMaddPromptSTR As addPrompt
    
    ' ***** Common regular expressions *****
    Private COLcre As Collection

    ' ***** Tooltips *****
    Private OBJtooltip As ToolTip

#End Region

#Region " Form events "

    Private Sub cmdAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbout.Click
        showAbout
    End Sub

    Private Sub cmdCommonRETest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCommonRETest.Click
        commonRETest
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer(True)
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry
    End Sub

    Private Sub cmdREavailableDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdREavailableDelete.Click
        deleteLstEntry(lstREavailable, cmdREavailableDelete, "regular expression")
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form
    End Sub

    Private Sub cmdREsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdREsave.Click
        conditionalSave(txtRE, frmAddPromptRE, lstREavailable)
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click
        testRE
    End Sub

    Private Sub cmdTestDataAvailableDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestDataAvailableDelete.Click
        deleteLstEntry(lstTestData, cmdTestDataAvailableDelete, "test data string")
    End Sub

    Private Sub cmdTestDatasave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestDataSave.Click
        conditionalSave(txtTestData, FRMaddPromptSTR, lstTestData)                             
    End Sub

    Private Sub cmdTestDataZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestDataZoom.Click
        zoomInterface(txtTestData)
    End Sub

    Private Sub cmdTestDataNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTestDataNext.Click
        With txtTestData
            .SelectionStart = .SelectionStart + .SelectionLength
        End With        
        testRE
    End Sub

    Private Sub cmdVBzoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVBzoom.Click
        zoomInterface(txtVB)
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        closer(True)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not CBool(GetSetting(Application.ProductName, _
                                Me.Name, _
                                "showAbout", _
                                "False")) Then
            showAbout()
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "showAbout", _
                        "True")
        End If
        Try
            OBJtooltip = New ToolTip
        Catch
            _OBJutilities.errorHandler("Cannot create my tooltip object: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "Form1_load", _
                                       "Continuing: tool tips won't be available")
        End Try
        mkToolTips(OBJtooltip)
        FRMaddPromptRE = mkAddPrompt() : FRMaddPromptRE.BackColor = Color.SkyBlue
        FRMaddPromptSTR = mkAddPrompt() : FRMaddPromptSTR.BackColor = Color.SpringGreen
        If (FRMaddPromptRE Is Nothing) OrElse (FRMaddPromptRE Is Nothing) Then
            errorHandler("Cannot create prompt forms: " & _
                         Err.Number & " " & Err.Description, _
                         "Form1_Load", _
                         "Can't continue")
            closer(False)
        End If
        With FRMaddPromptRE
            .ItemDescription = "Regular expression"
            .Value = ""
        End With
        With FRMaddPromptSTR
            .ItemDescription = "Test string"
            .Value = ""
        End With
        Dim intOldHeight As Integer = Height
        Dim intOldWidth As Integer = Width
        Width = CInt(_OBJwindowsUtilities.screenWidth * WIDTH_TOLERANCE)
        Height = CInt(_OBJwindowsUtilities.screenHeight * HEIGHT_TOLERANCE)
        _OBJwindowsUtilities.resizeConstituentControls(Me, intOldWidth, intOldHeight)
        ClientSize = New Size(ClientSize.Width, _
                              cmdClose.Bottom + _OBJwindowsUtilities.Grid)
        CenterToScreen()
        Opacity = 0.5
        Show()
        Refresh()
        registry2Form()
        COLcre = _OBJutilities.commonRegularExpressions
        If (COLcre Is Nothing) Then closer(False)
        addStandardStrings(lstREavailable, COLcre, FRMaddPromptRE)
        addStandardStrings(lstTestData, _
                           mkStandardTestStrings, _
                           FRMaddPromptSTR)
        With lstREavailable
            If .SelectedIndex >= 0 Then
                If getInfo(CStr(.Items(.SelectedIndex)), booDisplay2String:=True) _
                   = _
                   txtRE.Text Then
                    selectString(lstREavailable, txtRE, lblRE, lblREexplanation, False)
                Else
                    .SelectedIndex = -1
                End If
            End If
        End With
        With lstTestData
            If .SelectedIndex >= 0 Then
                If getInfo(CStr(.Items(.SelectedIndex)), booDisplay2String:=True) _
                   = _
                   txtTestData.Text Then
                    selectString(lstTestData, txtTestData, lblTestData, lblTestDataExplanation, True)
                Else
                    .SelectedIndex = -1
                End If
            End If
        End With
        Opacity = 1
        Refresh()
    End Sub

    Private Sub lstREavailable_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstREavailable.DoubleClick
        selectString(lstREavailable, txtRE, lblRE, lblREexplanation, False)
    End Sub

    Private Sub lstREavailable_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstREavailable.SelectedIndexChanged
        selectedIndexChange(lstREavailable, cmdREavailableDelete)
        With lstREavailable
            If .SelectedIndex >= 0 Then
                txtVB.Text = re2VBcode(CStr(lstREavailable.Items(lstREavailable.SelectedIndex)))
            End If
        End With
    End Sub

    Private Sub lstTestData_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstTestData.DoubleClick
        selectString(lstTestData, txtTestData, lblTestData, lblTestDataExplanation, True)
    End Sub

    Private Sub lstTestData_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstTestData.SelectedIndexChanged
        selectedIndexChange(lstTestData, cmdTestDataAvailableDelete)
    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        closer(True)
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        showAbout()
    End Sub

    Private Sub mnuToolsPromptCustomizationRE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsPromptCustomizationRE.Click
        showAddPrompt(FRMaddPromptRE)
    End Sub

    Private Sub mnuToolsPromptCustomizationSTR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsPromptCustomizationSTR.Click
        showAddPrompt(FRMaddPromptSTR)
    End Sub

    Private Sub txtRE_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRE.Enter
        unsetChangeIndicator(txtRE)
    End Sub

    Private Sub txtRE_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRE.Leave
        If CBool(txtRE.Tag) Then
            updateListBox(string2ListboxItem(txtRE.Text, FRMaddPromptRE, False), lstREavailable)
        End If
    End Sub

    Private Sub txtRE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRE.TextChanged
        setChangeIndicator(txtRE)
    End Sub

    Private Sub txtTestData_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTestData.Enter
        unsetChangeIndicator(txtTestData)
    End Sub

    Private Sub txtTestData_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTestData.Leave
        If CBool(txtTestData.Tag) Then
            updateListBox(string2ListboxItem(txtTestData.Text, FRMaddPromptSTR, True), lstTestData)
        End If
    End Sub

    Private Sub txtTestData_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTestData.TextChanged
        setChangeIndicator(txtTestData)
    End Sub

#End Region ' " Form events "

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Add standard strings
    '
    '
    ' Creates a collection and places this collection in the Tag of the
    ' regular expression or test string list box: this collection, for 
    ' each standard regular expression or test string, contains an entry 
    ' with the key _i (where i is the index of the standard re or string) 
    ' and the value True.
    '
    ' This collection is used to prevent deletion of standard entries.
    '
    '
    Private Sub addStandardStrings(ByRef lstBox As ListBox, _
                                   ByVal colStandardStrings As Collection, _
                                   ByRef frmAddPrompt As addPrompt)
        If (colStandardStrings Is Nothing) Then Return
        Dim booStringDisplay As Boolean
        If (frmAddPrompt Is FRMaddPromptSTR) Then
            booStringDisplay = True
        ElseIf Not (frmAddPrompt Is FRMaddPromptRE) Then            
            errorHandler("Unsupported form", "addStandardStrings", "Not making any additions")
            Return
        End If        
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strNext As String
        Dim intSelectedIndex As Integer = lstBox.SelectedIndex
        With colStandardStrings
            For intIndex1 = 1 To .Count
                With CType(.Item(intIndex1), Collection)
                    strNext = CStr(.Item(2))
                    If booStringDisplay Then
                        strNext = _OBJutilities.string2Display(strNext)
                    End If                    
                    strNext = string2ListboxItem(strNext, CStr(.Item(3)), booStringDisplay)
                    intIndex2 = _OBJwindowsUtilities.searchListBox(lstBox, strNext)
                    If intIndex2 = -1 Then
                        lstBox.Items.Add(strNext)
                        intIndex2 = lstBox.Items.Count - 1
                    End If     
                    Try
                        colStandardStrings.Add(True, "_" & intIndex2)
                    Catch: End Try                    
                End With           
            Next intIndex1     
        End With 
        With lstBox  
            .Tag = colStandardStrings
            .SelectedIndex = intSelectedIndex
        End With
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Interface to the common regular expression tester
    '
    '
    Private Sub commonRETest
        Dim strReport As String
        Dim strSuccess As String
        If _OBJutilities.commonRegularExpressionTester(strReport) Then
            strSuccess = "succeeded"
        Else
            strSuccess = "failed"
        End If        
        If MsgBox("Common regular expression test has " & strSuccess & ": " & _
                  "click Yes to view report: click No to go back to main form", _
                  MsgBoxStyle.YesNo) _
           = _
           MsgBoxResult.No Then Return
        zoomInterface(strReport)                             
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Save text box, unless it has already been saved
    '
    '
    Private Sub conditionalSave(ByVal txtBox As TextBox, _
                                ByVal frmAddPrompt As addPrompt, _
                                ByVal lstBox As ListBox)
        If Not updateListbox(string2ListboxItem_(txtBox.Text, _
                                                 Iif(FRMaddPrompt.Policy = addPrompt.ENUaddPromptPolicy.alwaysAdd, _
                                                     FRMaddPrompt.ItemDescription, _
                                                     FRMaddPrompt), _
                                                 True), _
                             lstBox) Then
            Msgbox("This " & frmAddPrompt.ItemDescription &  " has already been saved")
        End If                             
    End Sub    

    ' -----------------------------------------------------------------
    ' Terminate the application
    '
    '
    Private Sub closer(ByVal booOK As Boolean)
        Opacity = .5: Refresh
        If booOK Then form2Registry
        Dispose
        End
    End Sub

    ' -----------------------------------------------------------------
    ' Interface to the error handler
    '
    '
    Private Sub errorHandler(ByVal strMessage As String, _
                             ByVal strProcedure As String, _
                             ByVal strHelp As String, _
                             Optional ByVal booInfo As Boolean = False)
        Dim strFullMessage As String
        _OBJutilities.errorHandler(strMessage, _
                                   Me.Name, _
                                   strProcedure, _
                                   strHelp, _
                                   strFullMessage, _
                                   booInfo:=booInfo)
    End Sub
    
    ' -----------------------------------------------------------------
    ' Delete list box entry, and hide its associated delete button
    '
    '
    Private Sub deleteLstEntry(ByVal lstBox As ListBox, _
                               ByVal cmdDelete As Button, _
                               ByVal strDescription As String)
        With lstBox      
            Dim strMessage As String = "Since the " & strDescription & " " & _
                                       _OBJutilities.enquote(CStr(.Items(.SelectedIndex))) & " " & _
                                       "is one of the added items, it will be " & _
                                       "permanently deleted."   
            If Not (.Tag Is Nothing) Then
                Dim colHandle As Collection = CType(.Tag, Collection)
                Dim booStdEntry As Boolean  
                Try
                    booStdEntry = CBool(colHandle.Item("_" & .SelectedIndex))
                Catch: End Try                
                If booStdEntry Then
                    strMessage = "Since the " & strDescription & " " & _
                                 _OBJutilities.enquote(CStr(.Items(.SelectedIndex))) & " " & _
                                 "is one of the standard items, it will not be " & _
                                 "permanently deleted. Instead, it will be restored " & _
                                 "the next time this application runs."   
                End If   
            End If   
            strMessage &= "  Click Yes to proceed, No to cancel the deletion."                               
            If MsgBox(strMessage, MsgBoxStyle.YesNo) = MsgBoxResult.No Then Return                              
            .Items.RemoveAt(.SelectedIndex)
            cmdDelete.Visible = False
        End With
    End Sub                               
    
    ' -----------------------------------------------------------------
    ' Find the explanation of a regular expression or a test string
    '
    '
    Private Function findExplanation(ByVal lstBox As ListBox, _
                                     ByVal strString As String) As String
        With lstBox
            Dim intIndex1 As Integer
            Dim strREnext As String  
            Dim strREwork As String = UCase(Trim(strString))
            For intIndex1 = 0 To .Items.Count - 1
                strREnext = CStr(.Items.Item(intIndex1))
                If strREwork = UCase(Trim(getInfo(strREnext))) Then
                    Exit For
                End If                
            Next intIndex1            
            If intIndex1 < .Items.Count Then Return(getInfo(strREnext))
            Return("")
        End With                        
    End Function                             

    ' -----------------------------------------------------------------
    ' Save Registry values
    '
    '
    Private Function form2Registry() As Boolean
        Dim strErrInfo As String
        If Not _OBJwindowsUtilities.listBox2Registry(lstREavailable, _
                                                     strApplication:=Application.ProductName, _
                                                     strSection:=Me.Name) _
           OrElse _                                                 
           Not _OBJwindowsUtilities.listBox2Registry(lstTestData, _
                                                     strApplication:=Application.ProductName, _
                                                     strSection:=Me.Name) Then
            strErrInfo = "Cannot store list boxes in the registry"
        Else            
            Try
                With txtRE
                    SaveSetting(Application.ProductName, Me.Name, .Name, .Text)
                End With            
                With txtTestData
                    SaveSetting(Application.ProductName, Me.Name, .Name, .Text)
                End With            
            Catch ex As Exception
                strErrInfo = ex.ToString
            End Try        
        End If                                                     
        If strErrInfo <> "" Then
            _OBJutilities.errorHandler("Cannot save Registry settings", _
                                       Me.Name, _
                                       "form2Registry", _
                                       strErrInfo & vbNewline & vbNewline & _
                                       "Continuing without saving any settings")
            Return (False)
        End If
        Return (True)
    End Function
    
    ' -----------------------------------------------------------------
    ' Get the explanation (to the left of the colon) from the list box entry
    '
    '
    Private Function getExplanation(ByVal strListBoxEntry As String) As String
        Return(_OBJutilities.item(strListBoxEntry, 1, LISTBOX_SEPARATOR, False))
    End Function    
    
    ' -----------------------------------------------------------------
    ' Get the information (to the right of the colon) from the list box  
    '
    '
    Private Function getInfo(ByVal strListBoxEntry As String, _
                             Optional ByVal booDisplay2String As Boolean = False) As String
        With _OBJutilities
            Dim strInfo As String = .itemPhrase(strListBoxEntry, _
                                                2, _
                                                .items(strListBoxEntry, LISTBOX_SEPARATOR, False), _
                                                LISTBOX_SEPARATOR, _
                                                False) 
            If booDisplay2String Then strInfo = _OBJutilities.display2String(strInfo, "XML") 
            Return(strInfo)                                               
        End With                                        
    End Function    
    
    ' -----------------------------------------------------------------
    ' Item information to delete button
    '
    '
    Private Sub itemInfo2DeleteButton(ByVal strInfo As String, _
                                      ByVal cmdDelete As Button)
        With cmdDelete
            If (.Tag Is Nothing) Then
                .Tag = .Text
            End If          
            .Text = Replace(CStr(.Tag), _
                            "%ENTRY", _
                            _OBJutilities.enquote(_OBJutilities.ellipsis(strInfo, 16)))   
        End With        
    End Sub                                      
    
    ' -----------------------------------------------------------------
    ' Make the addprompt form 
    '
    '
    Private Function mkAddPrompt As addPrompt
        Dim frmAddPrompt As addPrompt
        Try
            frmAddPrompt = New addPrompt
            frmAddPrompt.StartPosition = FormStartPosition.CenterScreen
        Catch ex As Exception
            _OBJutilities.errorHandler("Can't create add prompt form: " & _
                                       Err.Number & " " & Err.Description, _
                                       Name, _
                                       "mkAddPrompt", _
                                       "Returning Nothing")
            Return(Nothing)                                       
        End Try        
        Return(frmAddPrompt)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Make standard test strings
    '
    '
    ' This method creates a collection containing standard test strings.
    ' It is the same structure as the collection that is returned by the
    ' utilities method, commonRegularExpressions. Each entry is a
    ' collection containing three items: the test string name, the
    ' test string, and a comment that describes what the test string 
    ' contains. The key of each entry is the test string name.
    '
    '
    Private Function mkStandardTestStrings As Collection
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch ex As Exception
            errorHandler("Cannot create test string collection: " & _
                         Err.Number & " " & Err.Description, _
                         "mkStandardTestStrings", _
                         "Returning Nothing and continuing")
            Return(Nothing)                         
        End Try        
        If Not mkStandardTestStrings_mk(colNew, "null", "", "Null string") Then
            colNew = Nothing
            Return(Nothing)
        End If        
        If Not mkStandardTestStrings_mk(colNew, "blank", " ", "One blank character") Then
            colNew = Nothing
            Return(Nothing)
        End If        
        If Not mkStandardTestStrings_mk(colNew, _
                                        "multiline", _
                                        "Line 1 of 3" & vbNewline & "Line 2 of 3" & vbNewline & "Line 3 of 3", _
                                        "Multiple lines") Then
            colNew = Nothing
            Return(Nothing)
        End If        
        If Not mkStandardTestStrings_mk(colNew, _
                                        "vb6", _
                                        "' ***** Visual Basic 6 procedures *****" & vbNewline & _
                                        "Public Sub vbSub6" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub" & vbNewline & _
                                        "Public Property Get vbPropertyGet6 As String" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub" & vbNewline & _
                                        "Public Property Let vbPropertyLet6(ByVal strNewValue As String)" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub" & vbNewline & _
                                        "Property Set vbPropertyLet6(ByVal objNewValue As Object)" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub" & vbNewline & _
                                        "Private Function vbFunc6" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub", _
                                        "Visual Basic 6 procedures") Then
            colNew = Nothing
            Return(Nothing)
        End If        
        If Not mkStandardTestStrings_mk(colNew, _
                                        "vbNet", _
                                        "' ***** Visual Basic .Net procedures *****" & vbNewline & _
                                        "Friend Overloads Sub vbSub6" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub" & vbNewline & _
                                        "Public Property vbPropertyGet6 As String" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub" & vbNewline & _
                                        "Private Shared Overloads Function vbFunc6" & vbNewline & _
                                        "   Msgbox(""Hello world"")" & vbNewline & _
                                        "End Sub", _
                                        "Visual Basic .Net procedures") Then
            colNew = Nothing
            Return(Nothing)
        End If        
        Return(colNew)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Make one standard test string on behalf of mkStandardTestStrings
    '
    '
    Private Function mkStandardTestStrings_mk(ByVal colNew As Collection, _
                                              ByVal strName As String, _
                                              ByVal strValue As String, _
                                              ByVal strDescription As String) As Boolean
        Try
            Dim colNewEntry As Collection
            colNewEntry = New Collection
            With colNewEntry
                .Add(strName): .Add(strValue): .Add(strDescription)
            End With        
            colNew.Add(colNewEntry)
        Catch ex As Exception
            errorHandler("Cannot create test string collection entry, or add it to main collection: " & _
                         Err.Number & " " & Err.Description, _
                         "mkStandardTestStrings_mk", _
                         "Returning False")
            Return(Nothing)                         
        End Try  
        Return(True)      
    End Function

    ' -----------------------------------------------------------------
    ' Make tool tips
    '
    '
    Private Sub mkToolTips(ByVal objToolTip As ToolTip)
        With objToolTip
            .SetToolTip(cmdAbout, "Get general information ""about"" this application")
            .SetToolTip(cmdClose, "Save form settings in the Registry and exit")
            .SetToolTip(cmdCommonRETest, "Test the common regular expressions")
            .SetToolTip(cmdForm2Registry, "Save form selections in the Registry")
            .SetToolTip(cmdREavailableDelete, "Deletes the currently selected regular expression")
            .SetToolTip(cmdRegistry2Form, "Restores the form to values currently in the Registry")
            .SetToolTip(cmdREsave, "Saves the current regular expression in your set of regular expressions")
            .SetToolTip(cmdTest, "Tests the current regular expression and applies it to the current string")
            .SetToolTip(cmdTestDataAvailableDelete, "Deletes the current test string")
            .SetToolTip(cmdTestDataNext, "Sets the selection in the test string past the current point")
            .SetToolTip(cmdTestDataSave, "Saves the current test string in the collection of test strings")
            .SetToolTip(cmdTestDataZoom, "Provides a larger view of the test strings")
            .SetToolTip(cmdVBzoom, "Provides a larger view of the Visual Basic definition of the regular expression")
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Regular expression to Visual Basic code
    '
    '
    Private Function re2VBcode(ByVal strREinfo As String) As String
        Dim strREvalue As String = getInfo(strREinfo)
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strOutstring As String
        For intIndex1 = 1 To Len(strREvalue) Step 50
            strOutstring &= _
                _OBJutilities.string2Display(Mid(strREvalue, intIndex1, 50), _
                                             "vbExpression")
            If intIndex1 + 50 <= Len(strREvalue) Then
                strOutstring &= " & _" & vbNewLine & "        "
            End If
        Next intIndex1
        Return ("    ' " & getExplanation(strREinfo) & vbNewLine & _
               "    Dim objRE As New System.Text.RegularExpression.RegEx _" & vbNewLine & _
               "    (" & strOutstring & ")")
    End Function

    ' -----------------------------------------------------------------
    ' Obtain Registry values
    '
    '
    Private Function registry2Form() As Boolean
        Dim strErrInfo As String
        If Not _OBJwindowsUtilities.registry2ListBox(lstREavailable, _
                                                     strApplication:=Application.ProductName, _
                                                     strSection:=Me.Name) _
           OrElse _
           Not _OBJwindowsUtilities.registry2ListBox(lstTestData, _
                                                     strApplication:=Application.ProductName, _
                                                     strSection:=Me.Name) Then
            strErrInfo = "Cannot load list boxes from the registry"
        Else
            Try
                With txtRE
                    .Text = GetSetting(Application.ProductName, _
                                       Me.Name, _
                                       .Name, _
                                       .Text)
                    lblREexplanation.Text = findExplanation(lstREavailable, .Text)
                End With
                With txtTestData
                    .Text = GetSetting(Application.ProductName, _
                                    Me.Name, _
                                    .Name, _
                                    .Text)
                    lblTestDataExplanation.Text = findExplanation(lstTestData, .Text)
                End With
            Catch ex As Exception
                strErrInfo = ex.ToString
            End Try
        End If
        If strErrInfo <> "" Then
            _OBJutilities.errorHandler("Cannot save Registry settings", _
                                       Me.Name, _
                                       "form2Registry", _
                                       strErrInfo & vbNewLine & vbNewLine & _
                                       "Continuing without saving any settings")
            Return (False)
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Selected index change event support
    '
    '
    Private Sub selectedIndexChange(ByVal lstBox As ListBox, _
                                    ByVal cmdDelete As Button)
        With lstBox
            If .SelectedIndex >= 0 Then
                itemInfo2DeleteButton(getInfo(CStr(.Items(.SelectedIndex))), cmdDelete)
                cmdDelete.Visible = True
            Else
                cmdDelete.Visible = False
            End If
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Transfer string from list box to text box and labels
    '
    '
    Private Sub selectString(ByVal lstBox As ListBox, _
                             ByVal txtBox As TextBox, _
                             ByVal lblLabel As Label, _
                             ByVal lblExplanation As Label, _
                             ByVal booDisplay2String As Boolean)
        With lstBox
            Dim strEntry As String = CStr(.Items(.SelectedIndex))
            lblExplanation.Text = getExplanation(strEntry)
            txtBox.Text = getInfo(strEntry, booDisplay2String:=booDisplay2String)
            unsetChangeIndicator(txtBox)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Set the change indication for the text box
    '
    '
    Private Sub setChangeIndicator(ByVal txtBox As TextBox)
        txtBox.Tag = True
    End Sub

    ' -----------------------------------------------------------------
    ' Show the Easter Egg
    '
    '
    Private Sub showAbout()
        MsgBox(ABOUTINFO)
    End Sub

    ' -----------------------------------------------------------------
    ' Show the add prompt
    '
    '
    Private Sub showAddPrompt(ByVal frmAddPrompt As addPrompt)
        frmAddPrompt.showForCustomization()
    End Sub

    ' -----------------------------------------------------------------
    ' Convert the string to the list box item
    '
    '
    ' --- May prompt
    Private Overloads Function string2ListboxItem(ByVal strInstring As String, _
                                                  ByVal frmAddPrompt As addPrompt, _
                                                  ByVal booString2Display As Boolean) _
            As String
        Return (string2ListboxItem_(strInstring, frmAddPrompt, booString2Display))
    End Function
    ' --- Doesn't prompt           
    Private Overloads Function string2ListboxItem(ByVal strInstring As String, _
                                                  ByVal strDescription As String, _
                                                  ByVal booString2Display As Boolean) _
            As String
        Return (string2ListboxItem_(strInstring, strDescription, booString2Display))
    End Function
    ' --- Core logic    
    Private Overloads Function string2ListboxItem_(ByVal strInstring As String, _
                                                   ByVal objDescriptionSource As Object, _
                                                   ByVal booString2Display As Boolean) _
            As String
        Dim strInstringWork As String = strInstring
        If booString2Display Then
            strInstringWork = _OBJutilities.string2Display(strInstring, strGraphicInclude:=" ")
        End If
        If (TypeOf objDescriptionSource Is addPrompt) Then
            Dim frmAddPrompt As addPrompt = CType(objDescriptionSource, addPrompt)
            If frmAddPrompt.Policy = addPrompt.ENUaddPromptPolicy.alwaysShow Then
                frmAddPrompt.ShowDialog()
            End If
            With frmAddPrompt
                strInstringWork = Replace(.ItemDescription, ":", ";") & " " & _
                                  .Description & LISTBOX_SEPARATOR & _
                                  strInstringWork
            End With
        ElseIf (TypeOf objDescriptionSource Is System.String) Then
            strInstringWork = Replace(CStr(objDescriptionSource), ":", ";") & _
                              LISTBOX_SEPARATOR & _
                              strInstringWork
        End If
        Return (strInstringWork)
    End Function

    ' -----------------------------------------------------------------
    ' Test the regular expression
    '
    '
    Private Sub testRE()
        Dim objRE As System.Text.RegularExpressions.Regex
        Try
            objRE = New System.Text.RegularExpressions.Regex(txtRE.Text)
        Catch ex As Exception
            MsgBox("Can't create regular expression object based on " & _
                   _OBJutilities.enquote(txtRE.Text) & ": " & _
                   Err.Number & Err.Description)
            Return
        End Try
        Try
            With objRE
                Dim objMatch As System.Text.RegularExpressions.Match
                Dim intIndex1 As Integer = txtTestData.SelectionStart + 1
                objMatch = .Match(txtTestData.Text, intIndex1 - 1)
                With objMatch
                    If .Length = 0 Then
                        MsgBox("No match was found." & _
                               CStr(IIf(txtTestData.SelectionStart <> 0, _
                                        vbNewLine & vbNewLine & _
                                        "Note that scanning started at index " & intIndex1, _
                                        "")))
                    Else
                        txtTestData.SelectionStart = .Index
                        txtTestData.SelectionLength = .Length
                        With txtTestData
                            .ScrollToCaret()
                            .Focus()
                            .Refresh()
                        End With
                    End If
                End With
            End With
        Catch ex As Exception
            MsgBox("Can't apply regular expression object based on " & _
                   _OBJutilities.enquote(txtRE.Text) & ": " & _
                   Err.Number & Err.Description)
            Return
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Clear the change indication for the text box
    '
    '
    Private Sub unsetChangeIndicator(ByVal txtBox As TextBox)
        txtBox.Tag = False
    End Sub

    ' -----------------------------------------------------------------
    ' Add new item to list box
    '
    '
    ' Returns False if the item is already in the list box, True when 
    ' the item is new.
    '
    '
    Private Function updateListBox(ByVal strNew As String, _
                                   ByVal lstBox As ListBox) As Boolean
        With lstBox
            Dim intIndex1 As Integer = _OBJwindowsUtilities.searchListBox(lstBox, strNew)
            If intIndex1 >= 0 Then Return (False)
            .Items.Add(strNew)
            .SelectedIndex = .Items.Count - 1
            .Refresh()
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Zoom control
    '
    '
    Private Overloads Sub zoomInterface(ByVal objZoomed As Object)
        Dim ctlZoom As zoom.zoom
        Dim ctlZoomed As Control
        Dim strZoomed As String
        If (TypeOf objZoomed Is Control) Then
            ctlZoomed = CType(objZoomed, Control)
        Else
            Try
                strZoomed = CStr(objZoomed)
            Catch : End Try
            If (strZoomed Is Nothing) Then
                errorHandler("Internal programming error in calling zoomInterface: " & _
                             "invalid object " & _OBJutilities.object2String(objZoomed), _
                             "zoomInterface", _
                             "Returning to caller")
                Return
            End If
            ctlZoomed = Me
        End If
        Try
            ctlZoom = New zoom.zoom
        Catch ex As Exception
            errorHandler("Cannot create zoom control: " & _
                         Err.Number & " " & Err.Description, _
                         "zoomInterface", _
                         ex.ToString)
            Return
        End Try
        With ctlZoom
            .setSize(CDbl(2), CDbl(1))
            If (strZoomed Is Nothing) Then
                .ZoomTextBox.Text = ctlZoomed.Text
            Else
                .ZoomTextBox.Text = strZoomed
            End If
            .showZoom()
            .dispose()
        End With
    End Sub

#End Region

End Class
