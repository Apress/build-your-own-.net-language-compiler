Option Strict On

Imports utilities.utilities
Imports windowsUtilities.windowsUtilities

' *********************************************************************
' *                                                                   *
' * Scanner test form                                                 *
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
    Friend WithEvents lblSourceCode As System.Windows.Forms.Label
    Friend WithEvents cmdSourceCodeLoad As System.Windows.Forms.Button
    Friend WithEvents cmdScanAll As System.Windows.Forms.Button
    Friend WithEvents lblScanned As System.Windows.Forms.Label
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lstScanned As System.Windows.Forms.ListBox
    Friend WithEvents cmdInspect As System.Windows.Forms.Button
    Friend WithEvents cmdXML As System.Windows.Forms.Button
    Friend WithEvents lblXMLmaxTokens As System.Windows.Forms.Label
    Friend WithEvents chkXMLmaxSourceNoLimit As System.Windows.Forms.CheckBox
    Friend WithEvents txtXMLmaxTokens As System.Windows.Forms.TextBox
    Friend WithEvents txtXMLmaxSource As System.Windows.Forms.TextBox
    Friend WithEvents chkXMLmaxTokensNoLimit As System.Windows.Forms.CheckBox
    Friend WithEvents lblXMLmaxSource As System.Windows.Forms.Label
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdScannedZoom As System.Windows.Forms.Button
    Friend WithEvents gbxXML As System.Windows.Forms.GroupBox
    Friend WithEvents chkXMLabout As System.Windows.Forms.CheckBox
    Friend WithEvents chkXMLcomments As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSourceCodeSave As System.Windows.Forms.Button
    Friend WithEvents txtSourceCode As System.Windows.Forms.TextBox
    Friend WithEvents cmdScanNext As System.Windows.Forms.Button
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileLoad As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsScanPrompting As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents cmdScanReset As System.Windows.Forms.Button
    Friend WithEvents ofdSourceCodeLoad As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfdSourceCodeSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuToolsTest As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsRegistryClear As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsForm2Registry As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsRegistry2Form As System.Windows.Forms.MenuItem
    Friend WithEvents cmdCloseNoSave As System.Windows.Forms.Button
    Friend WithEvents mnuToolsSep2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents cmdSourceCodeLoadTestString As System.Windows.Forms.Button
    Friend WithEvents cmdClearSettings As System.Windows.Forms.Button
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblSourceCode = New System.Windows.Forms.Label
Me.cmdSourceCodeLoad = New System.Windows.Forms.Button
Me.cmdScanAll = New System.Windows.Forms.Button
Me.lblScanned = New System.Windows.Forms.Label
Me.cmdClose = New System.Windows.Forms.Button
Me.lstScanned = New System.Windows.Forms.ListBox
Me.cmdInspect = New System.Windows.Forms.Button
Me.gbxXML = New System.Windows.Forms.GroupBox
Me.chkXMLcomments = New System.Windows.Forms.CheckBox
Me.chkXMLabout = New System.Windows.Forms.CheckBox
Me.chkXMLmaxTokensNoLimit = New System.Windows.Forms.CheckBox
Me.chkXMLmaxSourceNoLimit = New System.Windows.Forms.CheckBox
Me.txtXMLmaxTokens = New System.Windows.Forms.TextBox
Me.lblXMLmaxTokens = New System.Windows.Forms.Label
Me.txtXMLmaxSource = New System.Windows.Forms.TextBox
Me.lblXMLmaxSource = New System.Windows.Forms.Label
Me.cmdXML = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdScannedZoom = New System.Windows.Forms.Button
Me.ofdSourceCodeLoad = New System.Windows.Forms.OpenFileDialog
Me.cmdSourceCodeSave = New System.Windows.Forms.Button
Me.txtSourceCode = New System.Windows.Forms.TextBox
Me.cmdScanNext = New System.Windows.Forms.Button
Me.mnuMain = New System.Windows.Forms.MainMenu
Me.mnuFile = New System.Windows.Forms.MenuItem
Me.mnuFileLoad = New System.Windows.Forms.MenuItem
Me.mnuFileSave = New System.Windows.Forms.MenuItem
Me.mnuFileSep1 = New System.Windows.Forms.MenuItem
Me.mnuFileExit = New System.Windows.Forms.MenuItem
Me.mnuTools = New System.Windows.Forms.MenuItem
Me.mnuToolsScanPrompting = New System.Windows.Forms.MenuItem
Me.mnuToolsSep1 = New System.Windows.Forms.MenuItem
Me.mnuToolsTest = New System.Windows.Forms.MenuItem
Me.mnuToolsSep2 = New System.Windows.Forms.MenuItem
Me.mnuToolsForm2Registry = New System.Windows.Forms.MenuItem
Me.mnuToolsRegistry2Form = New System.Windows.Forms.MenuItem
Me.mnuToolsRegistryClear = New System.Windows.Forms.MenuItem
Me.mnuHelp = New System.Windows.Forms.MenuItem
Me.mnuHelpAbout = New System.Windows.Forms.MenuItem
Me.cmdScanReset = New System.Windows.Forms.Button
Me.sfdSourceCodeSave = New System.Windows.Forms.SaveFileDialog
Me.cmdCloseNoSave = New System.Windows.Forms.Button
Me.cmdSourceCodeLoadTestString = New System.Windows.Forms.Button
Me.cmdClearSettings = New System.Windows.Forms.Button
Me.cmdTest = New System.Windows.Forms.Button
Me.gbxXML.SuspendLayout()
Me.SuspendLayout()
'
'lblSourceCode
'
Me.lblSourceCode.BackColor = System.Drawing.Color.Navy
Me.lblSourceCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblSourceCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblSourceCode.ForeColor = System.Drawing.Color.White
Me.lblSourceCode.Location = New System.Drawing.Point(10, 9)
Me.lblSourceCode.Name = "lblSourceCode"
Me.lblSourceCode.Size = New System.Drawing.Size(777, 23)
Me.lblSourceCode.TabIndex = 0
Me.lblSourceCode.Text = "Source Code"
Me.lblSourceCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'cmdSourceCodeLoad
'
Me.cmdSourceCodeLoad.Location = New System.Drawing.Point(557, 9)
Me.cmdSourceCodeLoad.Name = "cmdSourceCodeLoad"
Me.cmdSourceCodeLoad.Size = New System.Drawing.Size(115, 23)
Me.cmdSourceCodeLoad.TabIndex = 2
Me.cmdSourceCodeLoad.Text = "Load from file"
'
'cmdScanAll
'
Me.cmdScanAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdScanAll.Location = New System.Drawing.Point(10, 249)
Me.cmdScanAll.Name = "cmdScanAll"
Me.cmdScanAll.Size = New System.Drawing.Size(115, 28)
Me.cmdScanAll.TabIndex = 3
Me.cmdScanAll.Text = "Scan"
'
'lblScanned
'
Me.lblScanned.BackColor = System.Drawing.Color.Navy
Me.lblScanned.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblScanned.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblScanned.ForeColor = System.Drawing.Color.White
Me.lblScanned.Location = New System.Drawing.Point(10, 397)
Me.lblScanned.Name = "lblScanned"
Me.lblScanned.Size = New System.Drawing.Size(777, 23)
Me.lblScanned.TabIndex = 4
Me.lblScanned.Text = "Scanned Source Code (Type@start..end:linenumber value)"
Me.lblScanned.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(701, 249)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(86, 46)
Me.cmdClose.TabIndex = 5
Me.cmdClose.Text = "Close"
'
'lstScanned
'
Me.lstScanned.BackColor = System.Drawing.SystemColors.Control
Me.lstScanned.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstScanned.ItemHeight = 17
Me.lstScanned.Location = New System.Drawing.Point(10, 420)
Me.lstScanned.Name = "lstScanned"
Me.lstScanned.Size = New System.Drawing.Size(777, 72)
Me.lstScanned.TabIndex = 6
'
'cmdInspect
'
Me.cmdInspect.Location = New System.Drawing.Point(134, 286)
Me.cmdInspect.Name = "cmdInspect"
Me.cmdInspect.Size = New System.Drawing.Size(116, 28)
Me.cmdInspect.TabIndex = 7
Me.cmdInspect.Text = "Inspect"
'
'gbxXML
'
Me.gbxXML.Controls.Add(Me.chkXMLcomments)
Me.gbxXML.Controls.Add(Me.chkXMLabout)
Me.gbxXML.Controls.Add(Me.chkXMLmaxTokensNoLimit)
Me.gbxXML.Controls.Add(Me.chkXMLmaxSourceNoLimit)
Me.gbxXML.Controls.Add(Me.txtXMLmaxTokens)
Me.gbxXML.Controls.Add(Me.lblXMLmaxTokens)
Me.gbxXML.Controls.Add(Me.txtXMLmaxSource)
Me.gbxXML.Controls.Add(Me.lblXMLmaxSource)
Me.gbxXML.Controls.Add(Me.cmdXML)
Me.gbxXML.Location = New System.Drawing.Point(259, 249)
Me.gbxXML.Name = "gbxXML"
Me.gbxXML.Size = New System.Drawing.Size(432, 129)
Me.gbxXML.TabIndex = 9
Me.gbxXML.TabStop = False
'
'chkXMLcomments
'
Me.chkXMLcomments.Checked = True
Me.chkXMLcomments.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkXMLcomments.Location = New System.Drawing.Point(259, 65)
Me.chkXMLcomments.Name = "chkXMLcomments"
Me.chkXMLcomments.Size = New System.Drawing.Size(163, 18)
Me.chkXMLcomments.TabIndex = 18
Me.chkXMLcomments.Text = "Include comments"
'
'chkXMLabout
'
Me.chkXMLabout.Checked = True
Me.chkXMLabout.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkXMLabout.Location = New System.Drawing.Point(259, 28)
Me.chkXMLabout.Name = "chkXMLabout"
Me.chkXMLabout.Size = New System.Drawing.Size(163, 18)
Me.chkXMLabout.TabIndex = 17
Me.chkXMLabout.Text = "Include About info"
'
'chkXMLmaxTokensNoLimit
'
Me.chkXMLmaxTokensNoLimit.BackColor = System.Drawing.Color.Navy
Me.chkXMLmaxTokensNoLimit.ForeColor = System.Drawing.Color.White
Me.chkXMLmaxTokensNoLimit.Location = New System.Drawing.Point(106, 92)
Me.chkXMLmaxTokensNoLimit.Name = "chkXMLmaxTokensNoLimit"
Me.chkXMLmaxTokensNoLimit.Size = New System.Drawing.Size(76, 16)
Me.chkXMLmaxTokensNoLimit.TabIndex = 16
Me.chkXMLmaxTokensNoLimit.Text = "No limit"
'
'chkXMLmaxSourceNoLimit
'
Me.chkXMLmaxSourceNoLimit.BackColor = System.Drawing.Color.Navy
Me.chkXMLmaxSourceNoLimit.ForeColor = System.Drawing.Color.White
Me.chkXMLmaxSourceNoLimit.Location = New System.Drawing.Point(106, 60)
Me.chkXMLmaxSourceNoLimit.Name = "chkXMLmaxSourceNoLimit"
Me.chkXMLmaxSourceNoLimit.Size = New System.Drawing.Size(76, 15)
Me.chkXMLmaxSourceNoLimit.TabIndex = 15
Me.chkXMLmaxSourceNoLimit.Text = "No limit"
'
'txtXMLmaxTokens
'
Me.txtXMLmaxTokens.Location = New System.Drawing.Point(192, 92)
Me.txtXMLmaxTokens.Name = "txtXMLmaxTokens"
Me.txtXMLmaxTokens.Size = New System.Drawing.Size(58, 22)
Me.txtXMLmaxTokens.TabIndex = 14
Me.txtXMLmaxTokens.Text = "100"
Me.txtXMLmaxTokens.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'lblXMLmaxTokens
'
Me.lblXMLmaxTokens.BackColor = System.Drawing.Color.Navy
Me.lblXMLmaxTokens.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblXMLmaxTokens.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblXMLmaxTokens.ForeColor = System.Drawing.Color.White
Me.lblXMLmaxTokens.Location = New System.Drawing.Point(10, 92)
Me.lblXMLmaxTokens.Name = "lblXMLmaxTokens"
Me.lblXMLmaxTokens.Size = New System.Drawing.Size(182, 23)
Me.lblXMLmaxTokens.TabIndex = 13
Me.lblXMLmaxTokens.Text = "Max tokens"
Me.lblXMLmaxTokens.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'txtXMLmaxSource
'
Me.txtXMLmaxSource.Location = New System.Drawing.Point(192, 55)
Me.txtXMLmaxSource.Name = "txtXMLmaxSource"
Me.txtXMLmaxSource.Size = New System.Drawing.Size(58, 22)
Me.txtXMLmaxSource.TabIndex = 12
Me.txtXMLmaxSource.Text = "100"
Me.txtXMLmaxSource.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'lblXMLmaxSource
'
Me.lblXMLmaxSource.BackColor = System.Drawing.Color.Navy
Me.lblXMLmaxSource.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblXMLmaxSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblXMLmaxSource.ForeColor = System.Drawing.Color.White
Me.lblXMLmaxSource.Location = New System.Drawing.Point(10, 55)
Me.lblXMLmaxSource.Name = "lblXMLmaxSource"
Me.lblXMLmaxSource.Size = New System.Drawing.Size(182, 23)
Me.lblXMLmaxSource.TabIndex = 11
Me.lblXMLmaxSource.Text = "Max source"
Me.lblXMLmaxSource.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'cmdXML
'
Me.cmdXML.Location = New System.Drawing.Point(10, 18)
Me.cmdXML.Name = "cmdXML"
Me.cmdXML.Size = New System.Drawing.Size(240, 28)
Me.cmdXML.TabIndex = 9
Me.cmdXML.Text = "Object to XML"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(134, 323)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(116, 28)
Me.cmdForm2Registry.TabIndex = 10
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(134, 360)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(116, 28)
Me.cmdRegistry2Form.TabIndex = 11
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdScannedZoom
'
Me.cmdScannedZoom.Location = New System.Drawing.Point(701, 397)
Me.cmdScannedZoom.Name = "cmdScannedZoom"
Me.cmdScannedZoom.Size = New System.Drawing.Size(86, 23)
Me.cmdScannedZoom.TabIndex = 12
Me.cmdScannedZoom.Text = "Zoom"
'
'ofdSourceCodeLoad
'
Me.ofdSourceCodeLoad.CheckFileExists = False
Me.ofdSourceCodeLoad.DefaultExt = "BAS"
'
'cmdSourceCodeSave
'
Me.cmdSourceCodeSave.Location = New System.Drawing.Point(672, 9)
Me.cmdSourceCodeSave.Name = "cmdSourceCodeSave"
Me.cmdSourceCodeSave.Size = New System.Drawing.Size(115, 23)
Me.cmdSourceCodeSave.TabIndex = 14
Me.cmdSourceCodeSave.Text = "Save to file"
'
'txtSourceCode
'
Me.txtSourceCode.BackColor = System.Drawing.SystemColors.Control
Me.txtSourceCode.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtSourceCode.Location = New System.Drawing.Point(10, 32)
Me.txtSourceCode.Multiline = True
Me.txtSourceCode.Name = "txtSourceCode"
Me.txtSourceCode.Size = New System.Drawing.Size(777, 208)
Me.txtSourceCode.TabIndex = 15
Me.txtSourceCode.Text = ""
'
'cmdScanNext
'
Me.cmdScanNext.Location = New System.Drawing.Point(134, 249)
Me.cmdScanNext.Name = "cmdScanNext"
Me.cmdScanNext.Size = New System.Drawing.Size(116, 28)
Me.cmdScanNext.TabIndex = 16
Me.cmdScanNext.Text = "Scan next"
'
'mnuMain
'
Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuTools, Me.mnuHelp})
'
'mnuFile
'
Me.mnuFile.Index = 0
Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileLoad, Me.mnuFileSave, Me.mnuFileSep1, Me.mnuFileExit})
Me.mnuFile.Text = "File"
'
'mnuFileLoad
'
Me.mnuFileLoad.Index = 0
Me.mnuFileLoad.Text = "Load..."
'
'mnuFileSave
'
Me.mnuFileSave.Index = 1
Me.mnuFileSave.Text = "Save..."
'
'mnuFileSep1
'
Me.mnuFileSep1.Index = 2
Me.mnuFileSep1.Text = "-"
'
'mnuFileExit
'
Me.mnuFileExit.Index = 3
Me.mnuFileExit.Text = "E&xit"
'
'mnuTools
'
Me.mnuTools.Index = 1
Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuToolsScanPrompting, Me.mnuToolsSep1, Me.mnuToolsTest, Me.mnuToolsSep2, Me.mnuToolsForm2Registry, Me.mnuToolsRegistry2Form, Me.mnuToolsRegistryClear})
Me.mnuTools.Text = "Tools"
'
'mnuToolsScanPrompting
'
Me.mnuToolsScanPrompting.Index = 0
Me.mnuToolsScanPrompting.Text = "Scan prompting..."
'
'mnuToolsSep1
'
Me.mnuToolsSep1.Index = 1
Me.mnuToolsSep1.Text = "-"
'
'mnuToolsTest
'
Me.mnuToolsTest.Index = 2
Me.mnuToolsTest.Text = "Test the scanner object"
'
'mnuToolsSep2
'
Me.mnuToolsSep2.Index = 3
Me.mnuToolsSep2.Text = "-"
'
'mnuToolsForm2Registry
'
Me.mnuToolsForm2Registry.Index = 4
Me.mnuToolsForm2Registry.Text = "Save Settings"
'
'mnuToolsRegistry2Form
'
Me.mnuToolsRegistry2Form.Index = 5
Me.mnuToolsRegistry2Form.Text = "Restore Settings"
'
'mnuToolsRegistryClear
'
Me.mnuToolsRegistryClear.Index = 6
Me.mnuToolsRegistryClear.Text = "Clear Settings"
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
'cmdScanReset
'
Me.cmdScanReset.Location = New System.Drawing.Point(10, 286)
Me.cmdScanReset.Name = "cmdScanReset"
Me.cmdScanReset.Size = New System.Drawing.Size(115, 28)
Me.cmdScanReset.TabIndex = 17
Me.cmdScanReset.Text = "Reset"
'
'cmdCloseNoSave
'
Me.cmdCloseNoSave.Location = New System.Drawing.Point(701, 305)
Me.cmdCloseNoSave.Name = "cmdCloseNoSave"
Me.cmdCloseNoSave.Size = New System.Drawing.Size(86, 83)
Me.cmdCloseNoSave.TabIndex = 19
Me.cmdCloseNoSave.Text = "Close-don't save settings"
'
'cmdSourceCodeLoadTestString
'
Me.cmdSourceCodeLoadTestString.Location = New System.Drawing.Point(442, 9)
Me.cmdSourceCodeLoadTestString.Name = "cmdSourceCodeLoadTestString"
Me.cmdSourceCodeLoadTestString.Size = New System.Drawing.Size(115, 23)
Me.cmdSourceCodeLoadTestString.TabIndex = 20
Me.cmdSourceCodeLoadTestString.Text = "Load test string"
'
'cmdClearSettings
'
Me.cmdClearSettings.Location = New System.Drawing.Point(10, 360)
Me.cmdClearSettings.Name = "cmdClearSettings"
Me.cmdClearSettings.Size = New System.Drawing.Size(115, 28)
Me.cmdClearSettings.TabIndex = 21
Me.cmdClearSettings.Text = "Clear Settings"
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(10, 323)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.Size = New System.Drawing.Size(115, 28)
Me.cmdTest.TabIndex = 22
Me.cmdTest.Text = "Test"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
Me.ClientSize = New System.Drawing.Size(794, 529)
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.cmdClearSettings)
Me.Controls.Add(Me.cmdSourceCodeLoadTestString)
Me.Controls.Add(Me.cmdCloseNoSave)
Me.Controls.Add(Me.cmdScanReset)
Me.Controls.Add(Me.cmdScanNext)
Me.Controls.Add(Me.txtSourceCode)
Me.Controls.Add(Me.cmdSourceCodeSave)
Me.Controls.Add(Me.cmdScannedZoom)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.gbxXML)
Me.Controls.Add(Me.cmdInspect)
Me.Controls.Add(Me.lstScanned)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.lblScanned)
Me.Controls.Add(Me.cmdScanAll)
Me.Controls.Add(Me.cmdSourceCodeLoad)
Me.Controls.Add(Me.lblSourceCode)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Menu = Me.mnuMain
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "qbScanner"
Me.gbxXML.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private WithEvents OBJscanner As qbScanner.qbScanner
    Private Const ABOUTINFO As String = _
        "qbScanner Test" & _
        vbNewLine & vbNewLine & _
        "This form and application tests the qbScanner"
    Private Const SCREEN_WIDTH_TOLERANCE As Single = 0.9
    Private Const SCREEN_HEIGHT_TOLERANCE As Single = 0.9

#End Region ' " Form data "

#Region " Form events "

    Private Sub chkXMLmaxSourceNoLimit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkXMLmaxSourceNoLimit.CheckedChanged
        txtXMLmaxSource.Enabled = Not chkXMLmaxSourceNoLimit.Checked
    End Sub

    Private Sub chkXMLmaxTokensNoLimit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkXMLmaxTokensNoLimit.CheckedChanged
        txtXMLmaxTokens.Enabled = Not chkXMLmaxTokensNoLimit.Checked
    End Sub

    Private Sub cmdClearSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearSettings.Click
        clearRegistry
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer(True)
    End Sub

    Private Sub cmdCloseNoSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCloseNoSave.Click
        closer(True, booForm2Registry:=False)
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry
    End Sub

    Private Sub cmdInspect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInspect.Click
        inspectInterface
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form
    End Sub

    Private Sub cmdScanAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScanAll.Click
        scanInterface
    End Sub

    Private Sub cmdScanNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScanNext.Click
        scanNext
    End Sub

    Private Sub cmdScanReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScanReset.Click
        scanReset
    End Sub

    Private Sub cmdSourceCodeLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceCodeLoad.Click
        loadFile
    End Sub

    Private Sub cmdSourceCodeLoadTestString_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceCodeLoadTestString.Click
        txtSourceCode.Text = OBJscanner.TestString
    End Sub

    Private Sub cmdSourceCodeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceCodeSave.Click
        saveFile 
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click
        testInterface()
    End Sub

    Private Sub cmdXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdXML.Click
        object2XMLinterface()
    End Sub

    Private Sub cmdScannedZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScannedZoom.Click
        zoomInterface(lstScanned, 1, 4)
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        closer(True)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim booFirstTime As Boolean = True
        Try
            booFirstTime = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            "firstTime", _
                                            "True"))
        Catch : End Try
        If booFirstTime Then
            showAboutInfo()
            Try
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            "firstTime", _
                            "False")
            Catch : End Try
        End If
        Dim intOldWidth As Integer = Width
        Dim intOldHeight As Integer = Height
        Width = CInt(screenWidth() * SCREEN_WIDTH_TOLERANCE)
        Height = CInt(screenHeight() * SCREEN_HEIGHT_TOLERANCE)
        resizeConstituentControls(Me, intOldWidth, intOldHeight)
        CenterToScreen()
        createToolTips()
        ofdSourceCodeLoad.InitialDirectory = Application.StartupPath
        mkSourceCodeTag()
        Try
            OBJscanner = New qbScanner.qbScanner
            AddHandler OBJscanner.scanEvent, AddressOf scanEventDelegate
        Catch ex As Exception
            errorHandler("Cannot create the scanner: " & _
                         Err.Number & " " & Err.Description, _
                         "Form1_Load", _
                         "Terminating" & _
                         vbNewLine & vbNewLine & _
                         ex.ToString, _
                         True)
        End Try
        registry2Form()
    End Sub

    Private Sub lstScanned_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstScanned.MouseUp
        highlightSelectedToken()
    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        closer(True)
    End Sub

    Private Sub mnuFileLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileLoad.Click
        loadFile()
    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        saveFile()
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        showAboutInfo()
    End Sub

    Private Sub mnuToolsForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsForm2Registry.Click
        form2Registry()
    End Sub

    Private Sub mnuToolsRegistryClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsRegistryClear.Click
        registryClear()
    End Sub

    Private Sub mnuToolsRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsRegistry2Form.Click
        registry2Form()
    End Sub

    Private Sub mnuToolsScanPrompting_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsScanPrompting.Click
        showScanPrompt()
    End Sub

    Private Sub mnuToolsTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsTest.Click
        testInterface()
    End Sub

    Private Sub txtSourceCode_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSourceCode.Leave
        OBJscanner.SourceCode = txtSourceCode.Text
    End Sub

    Private Sub txtSourceCode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSourceCode.KeyPress
        updateSourceChange(True)
    End Sub

    Private Sub txtXMLmaxSource_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtXMLmaxSource.LostFocus
        checkIntegerEntry(txtXMLmaxSource)
    End Sub

    Private Sub txtXMLmaxTokens_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtXMLmaxTokens.LostFocus
        checkIntegerEntry(txtXMLmaxTokens)
    End Sub

#End Region

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Check integer entry
    '
    '
    Private Sub checkIntegerEntry(ByVal txtBox As TextBox)
        With txtBox
            Dim colTag As Collection
            If (.Tag Is Nothing) Then
                Try
                    colTag = New Collection
                    colTag.Add(.BackColor)
                    colTag.Add(.ForeColor)
                    colTag.Add(.Font)
                    .Tag = colTag
                Catch
                    errorHandler("Can't create text box tag: " & _
                                 Err.Number & " " & Err.Description, _
                                 "checkIntegerEntry", _
                                 "Terminating application", _
                                 True)
                    closer(False)                                               
                End Try                
            End If            
            If verify(.Text, "0123456789") = 0 Then
                colTag = CType(.Tag, Collection)
                .BackColor = CType(colTag.Item(1), Color)
                .ForeColor = CType(colTag.Item(2), Color)
                .Font = CType(colTag.Item(3), Font)
                Return
            End If
            .SelectionStart = 0
            .SelectionLength = Len(.Text)
            .BackColor = Color.Red : .ForeColor = Color.White
            .Font = New Font(.Font, FontStyle.Bold)
            .Focus()
        End With
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Clear settings
    '
    '
    Private Function clearRegistry As Boolean
        Try
            DeleteSetting(Application.ProductName, Me.Name)
        Catch  
            errorHandler("Cannot clear Registry: " & _
                         Err.Number & " " & Err.Description, _
                         "clearRegistry", _
                         "Continuing", _
                         False)
            Return(False)                         
        End Try        
        Return(True)
    End Function    

    ' -----------------------------------------------------------------
    ' Close application 
    '
    '
    Private Sub closer(ByVal booNormal As Boolean, _
                       Optional ByVal booForm2Registry As Boolean = True)
        If booNormal Then
            If saveFileIfNecessary = "CANCEL" Then Return
            If booForm2Registry Then form2Registry
            If Not OBJscanner.Usable Then MsgBox("Test object is unusable")
            Try
                OBJscanner.dispose
            Catch: End Try                
        End If
        Try
            Dispose
            End
        Catch: End Try            
    End Sub

    ' ------------------------------------------------------------------
    ' Create all tool tips
    '
    '
    Private Sub createToolTips()
        Dim objToolTip As ToolTip
        Try
            objToolTip = New ToolTip
        Catch
            errorHandler("Cannot create tool tip: " & _
                         Err.Number & " " & Err.Description, _
                         "createToolTips", _
                         "Continuing without tool tip support", _
                         False)
            Return
        End Try
        With objToolTip
            .SetToolTip(cmdClearSettings, _
                        "Clear saved form settings to their defaults: click this, and then use " & _
                        "Close: don't save settings, and restart, to start with initial settings")
            .SetToolTip(cmdClose, _
                        "Save your form selections and end this application")
            .SetToolTip(cmdCloseNoSave, _
                        "End this application: form selections won't be saved")
            .SetToolTip(cmdForm2Registry, _
                        "Save current form selections to the Registry")
            .SetToolTip(cmdInspect, _
                        "Inspects the state of the test scanner object and reports internal " & _
                        "errors. If the source code is working correctly, this should produce " & _
                        "an error free inspection.")
            .SetToolTip(cmdRegistry2Form, _
                        "Sets form selections to the most recently saved selections")
            .SetToolTip(cmdScanAll, _
                        "Scans the complete contents of the text box containing source code, " & _
                        "and displays the scan results in the list box at the bottom of the form")
            .SetToolTip(cmdScannedZoom, _
                        "Displays the scan results in a read-only and scrollable text box")
            .SetToolTip(cmdScanNext, _
                        "Starting at the cursor position in the source code text box, gets the " & _
                        "next token")
            .SetToolTip(cmdScanReset, _
                        "Resets the scan position to the beginning of the test data")
            .SetToolTip(cmdSourceCodeLoad, _
                        "Loads source code from a file")
            .SetToolTip(cmdSourceCodeSave, _
                        "Saves source code in a file")
            .SetToolTip(cmdTest, _
                        "Executes a test of the qbScanner object")
            .SetToolTip(cmdXML, _
                        "Converts the state of the scanner to eXtended Markup Language")
            .SetToolTip(chkXMLabout, _
                        "Normally, a leading comment block is included in the XML state, " & _
                        "containing general information: uncheck this box to suppress this box")
            .SetToolTip(chkXMLcomments, _
                        "Normally, comments are included in the XML state that describe state variables: " & _
                        "uncheck this box to suppress these comments")
            .SetToolTip(chkXMLmaxSourceNoLimit, _
                        "Normally, up to 100 characters of input source code are displayed in XML state: " & _
                        "check this box for no limit or enter the limit in the text box at right")
            .SetToolTip(chkXMLmaxTokensNoLimit, _
                        "Normally, up to 100 tokens of input source code are displayed in XML state: " & _
                        "check this box for no limit or enter the limit in the text box at right")
            .SetToolTip(txtXMLmaxSource, _
                        "Normally, up to 100 characters of input source code are displayed in XML state: " & _
                        "enter a different maximum here or check the box at left for no limit")
            .SetToolTip(txtXMLmaxTokens, _
                        "Normally, up to 100 tokens of input source code are displayed in XML state: " & _
                        "enter a different maximum here or check the box at left for no limit")
        End With
    End Sub

    ' ------------------------------------------------------------------
    ' Error handling
    '
    '
    Private Sub errorHandler(ByVal strMessage As String, _
                             ByVal strProcedure As String, _
                             ByVal strHelp As String, _
                             ByVal booFatal As Boolean)
        Dim strHelpWork As String = strHelp & CStr(IIf(Trim(strHelp) = "", "", ": "))
        If booFatal Then
            ' Use our error handling utility and cancel
            strHelpWork &= "this error is fatal"
            utilities.utilities.errorHandler(strMessage, _
                                             Me.Name, strProcedure, _
                                             strHelpWork)
            closer(False)
        Else
            ' Display a message box
            strHelpWork &= "continuing"
            MsgBox(Now & " Error from procedure " & strProcedure & " in " & Me.Name & _
                   vbNewLine & vbNewLine & _
                   strMessage & _
                   vbNewLine & vbNewLine & _
                   strHelpWork)
        End If
    End Sub

    ' -----------------------------------------------------------------
    ' Form settings to Registry
    '
    '
    Private Function form2Registry() As Boolean
        Try
            With chkXMLmaxSourceNoLimit
                SaveSetting(Application.ProductName, Me.Name, .Name, CStr(.Checked))
            End With
            With chkXMLmaxTokensNoLimit
                SaveSetting(Application.ProductName, Me.Name, .Name, CStr(.Checked))
            End With
            With txtXMLmaxSource
                SaveSetting(Application.ProductName, Me.Name, .Name, .Text)
            End With
            With txtXMLmaxTokens
                SaveSetting(Application.ProductName, Me.Name, .Name, .Text)
            End With
            With chkXMLabout
                SaveSetting(Application.ProductName, Me.Name, .Name, CStr(.Checked))
            End With
            With chkXMLcomments
                SaveSetting(Application.ProductName, Me.Name, .Name, CStr(.Checked))
            End With
            With ofdSourceCodeLoad
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            "ofdSourceCodeLoad_InitialDirectory", _
                            .InitialDirectory)
                Try
                    SaveSetting(Application.ProductName, _
                                Me.Name, _
                                "ofdSourceCodeLoad_FileName", _
                                .FileName)
                Catch : End Try
            End With
            With txtSourceCode
                If Len(.Text) < 1024 Then
                    SaveSetting(Application.ProductName, _
                                Me.Name, _
                                .Name, _
                                .Text)
                Else
                    SaveSetting(Application.ProductName, _
                                Me.Name, _
                                .Name, _
                                "")
                End If
            End With
        Catch
            errorHandler("Cannot save settings to Registry: " & _
                         Err.Number & " " & Err.Description, _
                         "form2Registry", _
                         "Settings not saved", _
                         False)
            Return (False)
        End Try
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Get the directory from the file identifier
    '
    '
    Private Function getDirectory(ByVal strFileid As String) As String
        Dim intIndex1 As Integer = Math.Max(InStrRev(strFileid, "\"), _
                                            InStrRev(strFileid, ":"))
        If intIndex1 = 0 Then Return ("")
        Return (Mid(strFileid, 1, intIndex1 - 1))
    End Function

    ' -----------------------------------------------------------------
    ' Highlight clicked token in listbox, in text box
    '
    '
    Private Sub highlightSelectedToken()
        With lstScanned
            If .SelectedIndex >= 0 Then
                highlightToken(word(CStr(.Items(.SelectedIndex)), 1))
            End If
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Hightlight one token in the source code
    '
    '
    Private Sub highlightToken(ByVal strFromString As String)
        Dim objToken As qbToken.qbToken
        Try
            objToken = New qbToken.qbToken
            objToken.fromString(strFromString)
        Catch
            errorHandler("Cannot create a token object: " & _
                         Err.Number & " " & Err.Description, _
                         "highlightToken_", _
                         "", _
                         True)
        End Try
        With txtSourceCode
            .Select(objToken.StartIndex - 1, objToken.Length)
            .Focus()
            .Refresh()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Object inspection
    '
    '
    Private Sub inspectInterface()
        Dim strReport As String
        Select Case MsgBox("Inspection of the test scanner has " & _
                           CStr(IIf(OBJscanner.inspect(strReport), _
                                    "succeeded", _
                                    "failed")) & ": " & _
                           "click Yes to view the report: " & _
                           "click No to return to the main form", _
                           MsgBoxStyle.YesNo)
            Case MsgBoxResult.Yes : zoomInterface(strReport, 3, 2)
            Case MsgBoxResult.No
            Case Else
                errorHandler("Unexpected reply from MsgBox", _
                             "inspectInterface", _
                             "Select statement changed or very serious problem", _
                             True)
        End Select
    End Sub

    ' -----------------------------------------------------------------
    ' Load file
    '
    '
    Private Sub loadFile()
        With ofdSourceCodeLoad
            .ShowDialog()
            If .FileName <> "" Then
                txtSourceCode.Text = file2String(.FileName)
                .InitialDirectory = getDirectory(.FileName)
                updateSourceChange(False)
                updateSourceLabel(.FileName, False)
            Else
                MsgBox("No file selected")
            End If
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Create the scan prompter
    '
    '
    Private Function mkScanPrompt() As scanPrompt
        Dim frmScanPrompt As scanPrompt
        Try
            frmScanPrompt = New scanPrompt
        Catch
            errorHandler("Cannot create scanPrompt form: " & _
                            Err.Number & " " & Err.Description, _
                            "scanInterface", _
                            "", _
                            False)
            Return (Nothing)
        End Try
        Return (frmScanPrompt)
    End Function

    ' -----------------------------------------------------------------
    ' Create source code tag
    '
    '
    ' We tag the source code text box with a Collection, that contains
    ' three items:
    '
    '
    '      *  The original foreground color of the text box
    '      *  The original Font of the text box
    '      *  A flag indicating any changes
    '
    '
    Private Function mkSourceCodeTag() As Boolean
        Dim colTag As Collection
        Try
            colTag = New Collection
            With colTag
                .Add(txtSourceCode.BackColor)
                .Add(txtSourceCode.Font)
                .Add(False)
            End With
            txtSourceCode.Tag = colTag
        Catch
            errorHandler("Cannot create collection for tag: " & _
                         Err.Number & " " & Err.Description, _
                         "scanInterface", _
                         "", _
                         True)
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Object to XML interface
    '
    '
    Private Sub object2XMLinterface()
        Dim strXML As String
        Try
            strXML = OBJscanner.object2XML(CInt(IIf(chkXMLmaxSourceNoLimit.Checked, _
                                                    -1, txtXMLmaxSource.Text)), _
                                           CInt(IIf(chkXMLmaxTokensNoLimit.Checked, _
                                                    -1, txtXMLmaxTokens.Text)), _
                                           booAboutComment:=chkXMLabout.Checked, _
                                           booStateComment:=chkXMLcomments.Checked)
        Catch
            strXML = "Cannot convert object to XML: " & _
                     Err.Number & " " & Err.Description
            errorHandler(strXML, _
                         "object2XMLinterface", _
                         "Will display above in zoom box", _
                         False)
        End Try
        zoomInterface(strXML, 3, 2)
    End Sub

    ' -----------------------------------------------------------------
    ' Registry clear
    '
    '
    Private Function registryClear() As Boolean
        Select Case MsgBox("Do you want to clear the saved Registry " & _
                           "values associated with this form and this " & _
                           "application? This will set form values to " & _
                           "default values.", _
                           MsgBoxStyle.YesNo)
            Case MsgBoxResult.Yes
            Case MsgBoxResult.No : Return (False)
            Case Else
                errorHandler("Unexpected reply from MsgBox", _
                             "registryClear", _
                             "Continuing although this indicates, probably, " & _
                             "a serious problem", _
                             False)
                Return (False)
        End Select
        Try
            DeleteSetting(Application.ProductName)
        Catch
            errorHandler("Cannot clear Registry settings associated with this form: " & _
                         Err.Number & " " & Err.Description, _
                         "form2Registry", _
                         "Continuing: settings not cleared", _
                         False)
            Return (False)
        End Try
        registry2Form()
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Registry to form settings
    '
    '
    Private Function registry2Form() As Boolean
        Try
            With chkXMLmaxSourceNoLimit
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With chkXMLmaxTokensNoLimit
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With txtXMLmaxSource
                .Text = GetSetting(Application.ProductName, _
                                   Me.Name, _
                                   .Name, _
                                   .Text)
            End With
            checkIntegerEntry(txtXMLmaxSource)
            With txtXMLmaxTokens
                .Text = GetSetting(Application.ProductName, _
                                   Me.Name, _
                                   .Name, _
                                   .Text)
            End With
            checkIntegerEntry(txtXMLmaxTokens)
            With chkXMLabout
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With chkXMLcomments
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With ofdSourceCodeLoad
                .InitialDirectory = GetSetting(Application.ProductName, _
                                               Me.Name, _
                                               "ofdSourceCodeLoad_InitialDirectory", _
                                               .InitialDirectory)
                Try
                    .FileName = GetSetting(Application.ProductName, _
                                           Me.Name, _
                                           "ofdSourceCodeLoad_FileName", _
                                           .FileName)
                Catch : End Try
            End With
            With txtSourceCode
                .Text = GetSetting(Application.ProductName, _
                                   Me.Name, _
                                   txtSourceCode.Name, _
                                   OBJscanner.TestString)
                OBJscanner.SourceCode = .Text
            End With
        Catch
            errorHandler("Cannot get settings from Registry: " & _
                         Err.Number & " " & Err.Description, _
                         "registry2Form", _
                         "Continuing: settings not loaded", _
                         False)
            Return (False)
        End Try
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Reset source code display visuals
    '
    '
    Private Sub resetSourceVisuals()
        Dim colTag As Collection = CType(txtSourceCode.Tag, Collection)
        With colTag
            txtSourceCode.BackColor = CType(.Item(1), Color)
            txtSourceCode.Font = CType(.Item(2), Font)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Save source code to file
    '
    '
    ' Returns SAVED, NOT_SAVED, or CANCEL when the file exists already,
    ' and the user clicks Cancel in response to the file exists prompt.
    '
    '
    Private Sub saveFile()
        With sfdSourceCodeSave
            .Title = "Save"
            .ShowDialog()
            If Not string2File(txtSourceCode.Text, .FileName) Then
                errorHandler("Cannot write file", "saveFile", "", False)
            End If
            updateSourceChange(False)
            updateSourceLabel(.FileName, False)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Check status of file and save it's been changed
    '
    '
    ' Returns SAVED, NOT_SAVED or CANCEL
    '
    '
    Private Function saveFileIfNecessary() As String
        If sourceCodeChange() Then
            Select Case MsgBox("Source code has changed. " & _
                               "Click Yes to save the changes: " & _
                               "click No to continue without saving (changes may be lost): " & _
                               "click Cancel to cancel.", _
                               MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                Case MsgBoxResult.No : Return ("NOT_SAVED")
                Case MsgBoxResult.Cancel : Return ("CANCEL")
                Case Else
                    errorHandler("Unexpected Case", "saveFileIfNecessary", "", True)
            End Select
            saveFile()
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Test the scan
    '
    '
    Private Overloads Sub scanInterface()
        With OBJscanner
            Dim booOK As Boolean = True
            Dim frmScanPrompt As scanPrompt = mkScanPrompt()
            With frmScanPrompt
                .SelectionStart = txtSourceCode.SelectionStart
                .SelectionLength = IIf(txtSourceCode.SelectionLength = 0, _
                                    "", _
                                    txtSourceCode.SelectionLength)
                If Not .Always Then
                    .ShowDialog()
                    If .Cancel Then Return
                End If
                Try
                    Select Case .ScanMode
                        Case "ALL"
                            With lstScanned
                                .Items.Clear()
                                .Refresh()
                                OBJscanner.scan()
                                .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
                                highlightSelectedToken()
                            End With
                        Case "NEXTTOKEN"
                            OBJscanner.scan(CInt(1))
                        Case "SUBSTRING"
                            Dim lngSelectionEnd As Long = Len(txtSourceCode.Text)
                            If Not (.SelectionEnd Is Nothing) Then
                                lngSelectionEnd = CLng(.SelectionEnd)
                            End If
                            With OBJscanner
                                .scan(CLng(txtSourceCode.SelectionStart + 1), _
                                      lngSelectionEnd)
                            End With
                        Case Else
                            errorHandler("Programming error: unexpected case", _
                                         "scanInterface", _
                                         "", _
                                         True)
                    End Select
                Catch
                    errorHandler("Cannot scan: " & Err.Number & " " & Err.Description, _
                                 "scanInterface", _
                                 "", _
                                 False)
                End Try
            End With
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Scan next token
    '
    '
    Private Sub scanNext()
        With OBJscanner
            If .Scanned Then
                MsgBox("Scan has been completed. Use the Reset button " & _
                       "to clear the scan")
                Return
            End If
            Try
                .scan(CInt(1))
            Catch
                errorHandler("Can't scan next: " & Err.Number & " " & Err.Description, _
                            "scanNext", _
                            "", _
                            False)
            End Try
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Resets the scan object and clears the list box
    '
    '
    Private Sub scanReset()
        OBJscanner.reset()
        With lstScanned
            .Items.Clear() : .Refresh()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Show information about form and scanner object
    '
    '
    Private Sub showAboutInfo()
        MsgBox(ABOUTINFO & vbNewLine & vbNewLine & OBJscanner.About)
    End Sub

    ' -----------------------------------------------------------------
    ' Show the scan prompter for standalone use
    '
    '
    Private Sub showScanPrompt()
        Dim frmScanPrompt As scanPrompt = mkScanPrompt()
        If (frmScanPrompt Is Nothing) Then Return
        With frmScanPrompt
            .CloseText = "Close"
            .ShowDialog()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Tests the change indicator for the source code
    '
    '
    Private Function sourceCodeChange() As Boolean
        Return (CType(CType(txtSourceCode.Tag, Collection).Item(3), Boolean))
    End Function

    ' -----------------------------------------------------------------
    ' Tests scanner
    '
    '
    Private Sub testInterface()
        Dim strReport As String
        Dim booTest As Boolean = OBJscanner.test(strReport)
        Select Case MsgBox("Test has " & _
                            CStr(IIf(booTest, "succeeded", "failed")) & ": " & _
                            "click Yes to see report: click No to return to main form", _
                            MsgBoxStyle.YesNo)
            Case MsgBoxResult.Yes : zoomInterface(strReport, 3, 2)
            Case MsgBoxResult.No
            Case Else
                errorHandler("Unexpected reply from MsgBox", _
                             "testInterface", _
                             "Continuing although this error indicates a probably " & _
                             "significant problem", _
                             False)
        End Select
    End Sub

    ' -----------------------------------------------------------------
    ' Set the flag and visual indicators indicating change to the source 
    ' code
    '
    '
    Private Sub updateSourceChange(ByVal booChange As Boolean)
        With txtSourceCode
            With CType(.Tag, Collection)
                Try
                    .Remove(3) : .Add(booChange)
                Catch
                    errorHandler("Cannot update source code Tag: " & _
                                 Err.Number & " " & Err.Description, _
                                 "updateSourceChange", _
                                 "", _
                                 False)
                End Try
            End With
            If booChange Then
                .BackColor = Color.White
                .Font = New Font(.Font, FontStyle.Bold)
            Else
                resetSourceVisuals()
            End If
        End With
        cmdSourceCodeSave.Enabled = booChange
    End Sub

    ' -----------------------------------------------------------------
    ' Update the source file label
    '
    '
    ' Note that the Tag of the label is set to the original text of the
    ' label.
    '
    '
    Private Sub updateSourceLabel(ByVal strFileid As String, _
                                  ByVal booIsChanged As Boolean)
        With lblSourceCode
            If (.Tag Is Nothing) Then
                .Tag = .Text
            End If
            If strFileid = "" Then
                .Text = CStr(.Tag) : Return
            End If
            .Text = CStr(.Tag) & " " & _
                    "from " & _
                    CStr(IIf(booIsChanged, "modified version of ", "")) & " " & _
                    enquote(strFileid)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Zoom a string or list box
    '
    '
    Private Sub zoomInterface(ByVal objZoomed As Object, _
                              ByVal dblWidth As Double, _
                              ByVal dblHeight As Double)
        Dim ctlZoom As zoom.zoom    ' Ka boom
        Try
            ctlZoom = New zoom.zoom
        Catch
            errorHandler("objZoomed has unsupported type", _
                         "zoomInterface", _
                         "", _
                         True)
        End Try
        With ctlZoom
            .ZoomTextBox.Font = New Font(.ZoomTextBox.Font, FontStyle.Bold)
            .setSize(dblWidth, dblHeight)
            If (TypeOf objZoomed Is String) Then
                .ZoomTextBox.Text = CStr(objZoomed)
            ElseIf (TypeOf objZoomed Is Control) Then
                .Control = CType(objZoomed, Control)
            Else
                errorHandler("objZoomed has unsupported type", _
                             "zoomInterface", _
                             "", _
                             True)
            End If
            .showZoom()
            .dispose()
        End With
    End Sub

#End Region

#Region " Event delegates "

    ' -----------------------------------------------------------------
    ' Scan event handler
    '
    '
    Private Sub scanEventDelegate(ByVal objToken As qbToken.qbToken, _
                                    ByVal intCharacterIndex As Integer, _
                                    ByVal intLength As Integer, _
                                    ByVal intTokenCount As Integer)
        With lstScanned
            .Items.Add(objToken.ToString & " " & _
                       string2Display(objToken.sourceCode(OBJscanner.SourceCode)))
            .SelectedIndex = .Items.Count - 1
            .Refresh
        End With 
        With txtSourceCode
            .SelectionStart = objToken.StartIndex - 1
            .SelectionLength = objToken.Length    
            .Focus
        End With     
    End Sub    

#End Region ' " Event delegates "

End Class
