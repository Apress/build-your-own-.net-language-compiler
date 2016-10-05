Option Strict

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
    Friend WithEvents dlgSourceCodeLoad As System.Windows.Forms.OpenFileDialog
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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblSourceCode = New System.Windows.Forms.Label()
Me.cmdSourceCodeLoad = New System.Windows.Forms.Button()
Me.cmdScanAll = New System.Windows.Forms.Button()
Me.lblScanned = New System.Windows.Forms.Label()
Me.cmdClose = New System.Windows.Forms.Button()
Me.lstScanned = New System.Windows.Forms.ListBox()
Me.cmdInspect = New System.Windows.Forms.Button()
Me.gbxXML = New System.Windows.Forms.GroupBox()
Me.chkXMLcomments = New System.Windows.Forms.CheckBox()
Me.chkXMLabout = New System.Windows.Forms.CheckBox()
Me.chkXMLmaxTokensNoLimit = New System.Windows.Forms.CheckBox()
Me.chkXMLmaxSourceNoLimit = New System.Windows.Forms.CheckBox()
Me.txtXMLmaxTokens = New System.Windows.Forms.TextBox()
Me.lblXMLmaxTokens = New System.Windows.Forms.Label()
Me.txtXMLmaxSource = New System.Windows.Forms.TextBox()
Me.lblXMLmaxSource = New System.Windows.Forms.Label()
Me.cmdXML = New System.Windows.Forms.Button()
Me.cmdForm2Registry = New System.Windows.Forms.Button()
Me.cmdRegistry2Form = New System.Windows.Forms.Button()
Me.cmdScannedZoom = New System.Windows.Forms.Button()
Me.dlgSourceCodeLoad = New System.Windows.Forms.OpenFileDialog()
Me.cmdSourceCodeSave = New System.Windows.Forms.Button()
Me.txtSourceCode = New System.Windows.Forms.TextBox()
Me.cmdScanNext = New System.Windows.Forms.Button()
Me.mnuMain = New System.Windows.Forms.MainMenu()
Me.mnuFile = New System.Windows.Forms.MenuItem()
Me.mnuFileLoad = New System.Windows.Forms.MenuItem()
Me.mnuFileSave = New System.Windows.Forms.MenuItem()
Me.mnuFileSep1 = New System.Windows.Forms.MenuItem()
Me.mnuFileExit = New System.Windows.Forms.MenuItem()
Me.mnuTools = New System.Windows.Forms.MenuItem()
Me.mnuToolsScanPrompting = New System.Windows.Forms.MenuItem()
Me.mnuHelp = New System.Windows.Forms.MenuItem()
Me.mnuHelpAbout = New System.Windows.Forms.MenuItem()
Me.gbxXML.SuspendLayout()
Me.SuspendLayout()
'
'lblSourceCode
'
Me.lblSourceCode.BackColor = System.Drawing.Color.Navy
Me.lblSourceCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblSourceCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblSourceCode.ForeColor = System.Drawing.Color.White
Me.lblSourceCode.Location = New System.Drawing.Point(8, 8)
Me.lblSourceCode.Name = "lblSourceCode"
Me.lblSourceCode.Size = New System.Drawing.Size(648, 20)
Me.lblSourceCode.TabIndex = 0
Me.lblSourceCode.Text = "Source Code"
Me.lblSourceCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'cmdSourceCodeLoad
'
Me.cmdSourceCodeLoad.Location = New System.Drawing.Point(464, 8)
Me.cmdSourceCodeLoad.Name = "cmdSourceCodeLoad"
Me.cmdSourceCodeLoad.Size = New System.Drawing.Size(96, 20)
Me.cmdSourceCodeLoad.TabIndex = 2
Me.cmdSourceCodeLoad.Text = "Load from file"
'
'cmdScanAll
'
Me.cmdScanAll.Location = New System.Drawing.Point(8, 216)
Me.cmdScanAll.Name = "cmdScanAll"
Me.cmdScanAll.Size = New System.Drawing.Size(96, 24)
Me.cmdScanAll.TabIndex = 3
Me.cmdScanAll.Text = "Scan"
'
'lblScanned
'
Me.lblScanned.BackColor = System.Drawing.Color.Navy
Me.lblScanned.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblScanned.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblScanned.ForeColor = System.Drawing.Color.White
Me.lblScanned.Location = New System.Drawing.Point(8, 320)
Me.lblScanned.Name = "lblScanned"
Me.lblScanned.Size = New System.Drawing.Size(648, 20)
Me.lblScanned.TabIndex = 4
Me.lblScanned.Text = "Scanned Source Code"
Me.lblScanned.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(584, 216)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(72, 24)
Me.cmdClose.TabIndex = 5
Me.cmdClose.Text = "Close"
'
'lstScanned
'
Me.lstScanned.BackColor = System.Drawing.SystemColors.Control
Me.lstScanned.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstScanned.ItemHeight = 14
Me.lstScanned.Location = New System.Drawing.Point(8, 340)
Me.lstScanned.Name = "lstScanned"
Me.lstScanned.Size = New System.Drawing.Size(648, 102)
Me.lstScanned.TabIndex = 6
'
'cmdInspect
'
Me.cmdInspect.Location = New System.Drawing.Point(8, 248)
Me.cmdInspect.Name = "cmdInspect"
Me.cmdInspect.Size = New System.Drawing.Size(96, 24)
Me.cmdInspect.TabIndex = 7
Me.cmdInspect.Text = "Inspect"
'
'gbxXML
'
Me.gbxXML.Controls.AddRange(New System.Windows.Forms.Control() {Me.chkXMLcomments, Me.chkXMLabout, Me.chkXMLmaxTokensNoLimit, Me.chkXMLmaxSourceNoLimit, Me.txtXMLmaxTokens, Me.lblXMLmaxTokens, Me.txtXMLmaxSource, Me.lblXMLmaxSource, Me.cmdXML})
Me.gbxXML.Location = New System.Drawing.Point(216, 216)
Me.gbxXML.Name = "gbxXML"
Me.gbxXML.Size = New System.Drawing.Size(360, 96)
Me.gbxXML.TabIndex = 9
Me.gbxXML.TabStop = False
'
'chkXMLcomments
'
Me.chkXMLcomments.Checked = True
Me.chkXMLcomments.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkXMLcomments.Location = New System.Drawing.Point(216, 56)
Me.chkXMLcomments.Name = "chkXMLcomments"
Me.chkXMLcomments.Size = New System.Drawing.Size(136, 16)
Me.chkXMLcomments.TabIndex = 18
Me.chkXMLcomments.Text = "Include comments"
'
'chkXMLabout
'
Me.chkXMLabout.Checked = True
Me.chkXMLabout.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkXMLabout.Location = New System.Drawing.Point(216, 24)
Me.chkXMLabout.Name = "chkXMLabout"
Me.chkXMLabout.Size = New System.Drawing.Size(136, 16)
Me.chkXMLabout.TabIndex = 17
Me.chkXMLabout.Text = "Include About info"
'
'chkXMLmaxTokensNoLimit
'
Me.chkXMLmaxTokensNoLimit.BackColor = System.Drawing.Color.Navy
Me.chkXMLmaxTokensNoLimit.ForeColor = System.Drawing.Color.White
Me.chkXMLmaxTokensNoLimit.Location = New System.Drawing.Point(88, 75)
Me.chkXMLmaxTokensNoLimit.Name = "chkXMLmaxTokensNoLimit"
Me.chkXMLmaxTokensNoLimit.Size = New System.Drawing.Size(64, 14)
Me.chkXMLmaxTokensNoLimit.TabIndex = 16
Me.chkXMLmaxTokensNoLimit.Text = "No limit"
'
'chkXMLmaxSourceNoLimit
'
Me.chkXMLmaxSourceNoLimit.BackColor = System.Drawing.Color.Navy
Me.chkXMLmaxSourceNoLimit.ForeColor = System.Drawing.Color.White
Me.chkXMLmaxSourceNoLimit.Location = New System.Drawing.Point(88, 52)
Me.chkXMLmaxSourceNoLimit.Name = "chkXMLmaxSourceNoLimit"
Me.chkXMLmaxSourceNoLimit.Size = New System.Drawing.Size(64, 13)
Me.chkXMLmaxSourceNoLimit.TabIndex = 15
Me.chkXMLmaxSourceNoLimit.Text = "No limit"
'
'txtXMLmaxTokens
'
Me.txtXMLmaxTokens.Location = New System.Drawing.Point(160, 72)
Me.txtXMLmaxTokens.Name = "txtXMLmaxTokens"
Me.txtXMLmaxTokens.Size = New System.Drawing.Size(48, 20)
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
Me.lblXMLmaxTokens.Location = New System.Drawing.Point(8, 72)
Me.lblXMLmaxTokens.Name = "lblXMLmaxTokens"
Me.lblXMLmaxTokens.Size = New System.Drawing.Size(152, 20)
Me.lblXMLmaxTokens.TabIndex = 13
Me.lblXMLmaxTokens.Text = "Max tokens"
Me.lblXMLmaxTokens.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'txtXMLmaxSource
'
Me.txtXMLmaxSource.Location = New System.Drawing.Point(160, 48)
Me.txtXMLmaxSource.Name = "txtXMLmaxSource"
Me.txtXMLmaxSource.Size = New System.Drawing.Size(48, 20)
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
Me.lblXMLmaxSource.Location = New System.Drawing.Point(8, 48)
Me.lblXMLmaxSource.Name = "lblXMLmaxSource"
Me.lblXMLmaxSource.Size = New System.Drawing.Size(152, 20)
Me.lblXMLmaxSource.TabIndex = 11
Me.lblXMLmaxSource.Text = "Max source"
Me.lblXMLmaxSource.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'cmdXML
'
Me.cmdXML.Location = New System.Drawing.Point(8, 16)
Me.cmdXML.Name = "cmdXML"
Me.cmdXML.Size = New System.Drawing.Size(200, 24)
Me.cmdXML.TabIndex = 9
Me.cmdXML.Text = "Object to XML"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(8, 288)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(96, 24)
Me.cmdForm2Registry.TabIndex = 10
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(112, 288)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(96, 24)
Me.cmdRegistry2Form.TabIndex = 11
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdScannedZoom
'
Me.cmdScannedZoom.Location = New System.Drawing.Point(584, 320)
Me.cmdScannedZoom.Name = "cmdScannedZoom"
Me.cmdScannedZoom.Size = New System.Drawing.Size(72, 20)
Me.cmdScannedZoom.TabIndex = 12
Me.cmdScannedZoom.Text = "Zoom"
'
'dlgSourceCodeLoad
'
Me.dlgSourceCodeLoad.CheckFileExists = False
Me.dlgSourceCodeLoad.DefaultExt = "BAS"
'
'cmdSourceCodeSave
'
Me.cmdSourceCodeSave.Location = New System.Drawing.Point(560, 8)
Me.cmdSourceCodeSave.Name = "cmdSourceCodeSave"
Me.cmdSourceCodeSave.Size = New System.Drawing.Size(96, 20)
Me.cmdSourceCodeSave.TabIndex = 14
Me.cmdSourceCodeSave.Text = "Save to file"
'
'txtSourceCode
'
Me.txtSourceCode.BackColor = System.Drawing.SystemColors.Control
Me.txtSourceCode.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtSourceCode.Location = New System.Drawing.Point(8, 28)
Me.txtSourceCode.Multiline = True
Me.txtSourceCode.Name = "txtSourceCode"
Me.txtSourceCode.Size = New System.Drawing.Size(648, 180)
Me.txtSourceCode.TabIndex = 15
Me.txtSourceCode.Text = ""
'
'cmdScanNext
'
Me.cmdScanNext.Location = New System.Drawing.Point(112, 216)
Me.cmdScanNext.Name = "cmdScanNext"
Me.cmdScanNext.Size = New System.Drawing.Size(96, 24)
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
Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuToolsScanPrompting})
Me.mnuTools.Text = "Tools"
'
'mnuToolsScanPrompting
'
Me.mnuToolsScanPrompting.Index = 0
Me.mnuToolsScanPrompting.Text = "Scan prompting..."
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
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(662, 459)
Me.ControlBox = False
Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdScanNext, Me.txtSourceCode, Me.cmdSourceCodeSave, Me.cmdScannedZoom, Me.cmdRegistry2Form, Me.cmdForm2Registry, Me.gbxXML, Me.cmdInspect, Me.lstScanned, Me.cmdClose, Me.lblScanned, Me.cmdScanAll, Me.cmdSourceCodeLoad, Me.lblSourceCode})
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

    Private Shared _OBJutilities As utilities
    Private WithEvents OBJscanner As qbScanner
    Private Const ABOUTINFO As String = _
        "qbScanner Test" & _
        vbNewline & vbNewline & _
        "This form and application tests the qbScanner"

#End Region ' " Form data "

#Region " Form events "

    Private Sub chkXMLmaxSourceNoLimit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkXMLmaxSourceNoLimit.CheckedChanged
        txtXMLmaxSource.Enabled = Not chkXMLmaxSourceNoLimit.Checked
    End Sub

    Private Sub chkXMLmaxTokensNoLimit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkXMLmaxTokensNoLimit.CheckedChanged
        txtXMLmaxTokens.Enabled = Not chkXMLmaxTokensNoLimit.Checked
    End Sub
    
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer(True)
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

    Private Sub cmdSourceCodeLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceCodeLoad.Click
        loadFile
    End Sub

    Private Sub cmdSourceCodeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceCodeSave.Click
        saveFile 
    End Sub

    Private Sub cmdXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdXML.Click
        object2XMLinterface
    End Sub

    Private Sub cmdScannedZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdScannedZoom.Click
        zoomInterface(lstScanned)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim booFirstTime As Boolean = True
        Try
            booFirstTime = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            "firstTime", _
                                            "True")) 
        Catch: End Try                    
        If booFirstTime Then
            showAboutInfo
            Try
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            "firstTime", _
                            "False")
            Catch: End Try            
        End If                    
        dlgSourceCodeLoad.InitialDirectory = Application.StartupPath
        mkSourceCodeTag  
        Try
            OBJscanner = New qbScanner
            AddHandler OBJscanner.scanEvent, AddressOf scanEventDelegate
        Catch            
            errorHandler("Cannot create the scanner: " & _
                         Err.Number & " " & Err.Description, _
                         "Form1_Load", _
                         "Terminating", _
                         True)
        End Try        
        registry2Form
    End Sub

    Private Sub lstScanned_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstScanned.MouseUp
        With lstScanned
            If .SelectedIndex >= 0 Then 
                highLightToken(_OBJutilities.word(CStr(.Items(.SelectedIndex)), 1))
            End If                
        End With            
    End Sub

    Private Sub mnuFileLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileLoad.Click
        loadFile 
    End Sub

    Private Sub mnuFileSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        saveFile 
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        showAboutInfo
    End Sub

    Private Sub mnuToolsScanPrompting_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuToolsScanPrompting.Click
        showScanPrompt
    End Sub

    Private Sub txtSourceCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSourceCode.TextChanged
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
            If _OBJutilities.verify(.Text, "0123456789") = 0 Then 
                colTag = CType(.Tag, Collection)
                .BackColor = CType(colTag.Item(1), Color)
                .ForeColor = CType(colTag.Item(2), Color)
                .Font = CType(colTag.Item(3), Font)
                Return
            End If                
            .SelectionStart = 0
            .SelectionLength = Len(.Text)
            .BackColor = Color.Red: .ForeColor = Color.White
            .Font = New Font(.Font, FontStyle.Bold)
            .Focus
        End With            
    End Sub    

    ' -----------------------------------------------------------------
    ' Close application 
    '
    '
    Private Sub closer(ByVal booNormal As Boolean)
        If booNormal Then
            If saveFileIfNecessary = "CANCEL" Then Return
            form2Registry
            OBJscanner.dispose
        End If
        Try
            Dispose
            End
        Catch: End Try            
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
            _OBJutilities.errorHandler(strMessage, _
                                       Me.Name, strProcedure, _
                                       strHelpWork)
            closer(False)
        Else
            ' Display a message box
            strHelpWork &= "continuing
            MsgBox(Now & " Error from procedure " & strProcedure & " in " & Me.Name & _
                   vbNewline & vbNewline & _
                   strMessage & _
                   vbNewline & vbNewline & _
                   strHelpWork)
        End If        
    End Sub                             

    ' -----------------------------------------------------------------
    ' Form settings to Registry
    '
    '
    Private Function form2Registry As Boolean
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
            With dlgSourceCodeLoad
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            "dlgSourceCodeLoad_InitialDirectory", _
                            .InitialDirectory)
                Try                                               
                    SaveSetting(Application.ProductName, _
                                Me.Name, _
                                "dlgSourceCodeLoad_FileName", _
                                .FileName)
                Catch: End Try                                                
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
            Return(False)                                       
        End Try        
        Return(True)
    End Function  
    
    ' -----------------------------------------------------------------
    ' Get the directory from the file identifier
    '
    '
    Private Function getDirectory(ByVal strFileid As String) As String
        Dim intIndex1 As Integer = Math.Max(Instrrev(strFileid, "\"), _
                                            InStrRev(strFileid, ":"))
        If intIndex1 = 0 Then Return("")
        Return(Mid(strFileid, 1, intIndex1 - 1))                                            
    End Function      
    
    ' -----------------------------------------------------------------
    ' Hightlight one token in the source code
    '
    '
    Private Sub highlightToken(ByVal strFromString As String)
        Dim objToken As qbToken
        Try
            objToken = New qbToken
            objToken.fromString(strFromstring)
        Catch
            errorHandler("Cannot create a token object: " & _
                         Err.Number & " " & Err.Description, _
                         "highlightToken_", _
                         "", _
                         True)
        End Try        
        With txtSourceCode
            .Select(objToken.StartIndex - 1, objToken.Length)
            .Focus
            .Refresh
        End With        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Object inspection
    '
    '
    Private Sub inspectInterface
        Dim strReport As String
        Select Case MsgBox("Inspection of the test scanner has " & _
                           CStr(Iif(OBJscanner.inspect(strReport), _
                                    "succeeded", _
                                    "failed")) & ": " & _
                           "click Yes to view the report: " & _
                           "click No to return to the main form", _
                           MsgBoxStyle.YesNo)                         
            Case MsgBoxResult.Yes: zoomInterface(strReport)
            Case MsgBoxResult.No: 
            Case Else: 
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
    Private Sub loadFile
        With dlgSourceCodeLoad
            .ShowDialog 
            If .FileName <> "" Then
                txtSourceCode.Text = _OBJutilities.file2String(.FileName)
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
    Private Function mkScanPrompt As scanPrompt
        Dim frmScanPrompt As scanPrompt
        Try
            frmScanPrompt = New scanPrompt
        Catch
            errorHandler("Cannot create scanPrompt form: " & _
                            Err.Number & " " & Err.Description, _
                            "scanInterface", _
                            "", _
                            False)
            Return(Nothing)                                        
        End Try 
        Return(frmScanPrompt)
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
    Private Function mkSourceCodeTag As Boolean
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
    ' Show information about form and scanner object
    '
    '
    Private Sub showAboutInfo 
        MsgBox(ABOUTINFO & vbNewline & vbNewline & OBJscanner.About)
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Show the scan prompter for standalone use
    '
    '
    Private Sub showScanPrompt 
        Dim frmScanPrompt As scanPrompt = mkScanPrompt
        If (frmScanPrompt Is Nothing) Then Return
        With frmScanPrompt
            .CloseText = "Close"
            .ShowDialog
        End With        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Object to XML interface
    '
    '
    Private Sub object2XMLinterface
        Dim strXML As String
        Try
            strXML = OBJscanner.object2XML(CInt(Iif(chkXMLmaxSourceNoLimit.Checked, _
                                                    -1, txtXMLmaxSource)), _
                                           CInt(Iif(chkXMLmaxTokensNoLimit.Checked, _
                                                    -1, txtXMLmaxTokens)), _
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
        zoomInterface(strXML)  
    End Sub    

    ' -----------------------------------------------------------------
    ' Registry to form settings
    '
    '
    Private Function registry2Form As Boolean
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
            With dlgSourceCodeLoad
                .InitialDirectory = GetSetting(Application.ProductName, _
                                               Me.Name, _
                                               "dlgSourceCodeLoad_InitialDirectory", _
                                               .InitialDirectory)
                Try                                               
                    .FileName = GetSetting(Application.ProductName, _
                                           Me.Name, _
                                           "dlgSourceCodeLoad_FileName", _
                                           .FileName)
                Catch: End Try                                                
            End With            
            With txtSourceCode
                .Text = GetSetting(Application.ProductName, _
                                   Me.Name, _
                                   .Name, _
                                   .Text)
            End With            
        Catch
            errorHandler("Cannot save settings to Registry: " & _
                         Err.Number & " " & Err.Description, _
                         "form2Registry", _
                         "Continuing: settings not saved", _
                         False)
            Return(False)                                       
        End Try        
        Return(True)
    End Function  
    
    ' -----------------------------------------------------------------
    ' Reset source code display visuals
    '
    '
    Private Sub resetSourceVisuals
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
    Private Sub saveFile   
        With dlgSourceCodeLoad
            .Title = "Save"
            .ShowDialog 
            If Not _OBJutilities.string2File(txtSourceCode.Text, .FileName) Then
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
    Private Function saveFileIfNecessary As String
        If sourceCodeChange Then 
            Select Case MsgBox("Source code has changed. " & _
                               "Click Yes to save the changes: " & _
                               "click No to continue without saving (changes may be lost): " & _
                               "click Cancel to cancel.", _
                               MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes: 
                Case MsgBoxResult.No: Return("NOT_SAVED")
                Case MsgBoxResult.Cancel: Return("CANCEL")
                Case Else:
                    errorHandler("Unexpected Case", "saveFileIfNecessary", "", True)
            End Select                               
            saveFile 
        End If            
    End Function         

    ' -----------------------------------------------------------------
    ' Test the scan
    '
    '
    Private Overloads Sub scanInterface
        With OBJscanner
            Dim booOK As Boolean = True
            Dim frmScanPrompt As scanPrompt = mkScanPrompt
            With frmScanPrompt
                .SelectionStart = txtSourceCode.SelectionStart
                .SelectionLength = Iif(txtSourceCode.SelectionLength = 0, _
                                    "", _
                                    txtSourceCode.SelectionLength)
                If Not .Always Then
                    .ShowDialog
                    If .Cancel Then Return 
                End If 
                Try
                    Select Case .ScanMode  
                        Case "ALL": 
                            lstScanned.Items.Clear
                            lstScanned.Refresh
                            OBJscanner.scan
                        Case "NEXTTOKEN": 
                            OBJscanner.scan(CInt(1))
                        Case "SUBSTRING": 
                            Dim objSelectionEnd As Object = .SelectionEnd
                            If (objSelectionEnd Is Nothing) Then
                                objSelectionEnd = Len(txtSourceCode)
                            End If                        
                            OBJscanner.scan(CLng(.SelectionEnd))
                        Case Else:
                            errorHandler("Programming error: unexpected case", _
                                         "scanInterface", _
                                         "", _
                                         True)                            
                    End Select   
                Catch
                    errorHandler("Cannot scan: " & Err.Number & " " & Err.Description,  _
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
    Private Sub scanNext
        Try
            OBJscanner.scan(CInt(1))
        Catch
            errorHandler("Can't scan next: " & Err.Number & " " & Err.Description, _
                         "scanNext", _
                         "", _
                         False)
        End Try        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Tests the change indicator for the source code
    '
    '
    Private Function sourceCodeChange As Boolean
        Return(CType(CType(txtSourceCode.Tag, Collection).Item(3), Boolean))
    End Function    
    
    ' -----------------------------------------------------------------
    ' Set the flag and visual indicators indicating change to the source 
    ' code
    '
    '
    Private Sub updateSourceChange(ByVal booChange As Boolean)
        With txtSourceCode
            With CType(.Tag, Collection)
                Try
                    .Remove(3): .Add(booChange)
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
                resetSourceVisuals
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
                .Text = CStr(.Tag): Return
            End If              
            .Text = CStr(.Tag) & " " & _
                    "from " & _
                    CStr(Iif(booIsChanged, "modified version of ", "")) & " " & _
                    _OBJutilities.enquote(strFileid)
        End With        
    End Sub                                  
        
    ' -----------------------------------------------------------------
    ' Zoom a string or list box
    '
    '
    Private Sub zoomInterface(ByVal objZoomed As Object)
        Dim ctlZoom As zoom
        Try
            ctlZoom = New zoom
        Catch
            errorHandler("objZoomed has unsupported type", _
                         "zoomInterface", _
                         "", _
                         True)
        End Try        
        With ctlZoom
            .ZoomTextBox.Font = New Font(.ZoomTextBox.Font, FontStyle.Bold)
            .setSize(CDbl(2.25), CDbl(2))
            If (TypeOf objZoomed Is String) Then
                .ZoomTextBox.Text = CStr(objZoomed) 
            ElseIf (Typeof objZoomed Is Control) Then
                .Control = CType(objZoomed, Control)
            Else
                errorHandler("objZoomed has unsupported type", _
                             "zoomInterface", _
                             "", _
                             True)
            End If
            .showZoom
            .dispose
        End With
    End Sub    

#End Region

#Region " Event delegates "

    ' -----------------------------------------------------------------
    ' Scan event handler
    '
    '
    Private Sub scanEventDelegate(ByVal objToken As qbToken)
        With lstScanned
            .Items.Add(objToken.toString & " " & _
                       _OBJutilities.string2Display(objToken.sourceCode(OBJscanner.SourceCode)))
            .SelectedIndex = .Items.Count - 1
            .Refresh
        End With              
    End Sub    

#End Region ' " Event delegates "

End Class
