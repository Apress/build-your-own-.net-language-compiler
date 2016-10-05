Option Strict On

Imports quickBasicEngine.qbQuickBasicEngine

' *********************************************************************
' *                                                                   *
' * quickBasicEngine: one possible GUI                                *
' *                                                                   *
' *                                                                   *
' * This form and application is a "glass-box" view of the quickBasic *
' * engine, showing its detailed operations in scanning, parsing,     *
' * assembling and interpreting the code.                             *
' *                                                                   *
' * See also the readme.TXT file included with this project. This     *
' * block addresses these topics:                                     *
' *                                                                   *
' *                                                                   *
' *      *  Guidelines for adding new form controls                   *
' *      *  Tag usage                                                 *
' *      *  Compile-time symbols                                      *
' *                                                                   *
' *                                                                   *
' * GUIDELINES FOR ADDING NEW FORM CONTROLS ------------------------- *
' *                                                                   *
' * Since this form should provide the user a consistent experience   *
' * including Tooltips and persistence of user selections, follow the *
' * following procedure to add new form controls.                     *
' *                                                                   *
' *                                                                   *
' *      1.  Define the control as usual                              *
' *                                                                   *
' *      2.  If the control has a value that should persist between   *
' *          activations of this application, add a Select Case       *
' *          block that gets its value to the registry2Form           *
' *          method, and add a Select Case block that writes its      *
' *          value to form2Registry, in both cases following the      *
' *          pattern of the existing code.                            *
' *                                                                   *
' *      3.  If it makes sense to associate a Tool Tip with the con-  *
' *          trol, add a Select Case block to the setToolTips         *
' *          method with the tool tip.                                *
' *                                                                   *
' *                                                                   *
' * TAG USAGE ------------------------------------------------------- *
' *                                                                   *
' * Tags are untyped objects associated with controls and other       *
' * objects to contain user extension data. Here is how this form uses*
' * the Tags of a few controls. The tags of these controls should not *
' * be assumed available for any other purpose.                       *
' *                                                                   *
' *                                                                   *
' *      *  The Tag of the lstStatus display contains the level of    *
' *         status reporting.                                         *
' *                                                                   *
' *      *  The Tag of the rtbScreen display contains a five-item     *
' *         collection:                                               *
' *                                                                   *
' *         + Item(1) is True when the screen mode is source code,    *
' *           False when the screen mode is output                    *
' *                                                                   *
' *         + Item(2) contains the current source code                *
' *                                                                   *
' *         + Item(3) contains the current output                     *
' *                                                                   *
' *         + Item(4) contains the original screen background color   *
' *                                                                   *
' *         + Item(5) contains the original screen foreground color   *
' *                                                                   *
' *         + Item(6) contains a Nothing or a subcollection, which    *
' *           contains zero, one or more sub-sub-collections. Item(1) *
' *           of the sub-sub-collection is the start index of text    *
' *           highlighted as an error. Item(2) is its length. Item(3) *
' *           is its font prior to error highlights. Item(4) is its   *
' *           Rich Text selection color.                              * 
' *                                                                   *
' *      *  When the run form is shown, the command buttons on this   *
' *         form are greyed out. The run form itself is Tagged with   *
' *         a collection created by windowsUtilities.controlChange    *
' *         that shows how to restore all controls.                   *
' *                                                                   *
' *      *  The Tag of the cmdRun button is used by runInterface. It  *
' *         contains the total amount of time spent in handling IO    *
' *         events.                                                   *
' *                                                                   *
' *                                                                   *
' * COMPILE-TIME SYMBOLS -------------------------------------------- *
' *                                                                   *
' * This form exposes the following compile-time symbols:             *
' *                                                                   *
' *                                                                   *
' *      *  QBGUI_EXTENSIONS                                          *
' *      *  QBGUI_POPCHECK                                            *
' *                                                                   *
' *                                                                   *
' * The QBGUI_EXTENSIONS compile-time symbol may be set to True, to   *
' * obtain extensions to qbGUI facilities beyond what are documented  *
' * in the Apress book.                                               *
' *                                                                   *
' * If QBGUI_EXTENSIONS is True, then an option to expand arrays as   *
' * displayed in the storage view, and include their values, is       *
' * exposed in the Options screen.                                    *
' *                                                                   *
' * If QBGUI_EXTENSIONS is True, then an option check stack frames    *
' * during interpretation is exposed in the Options screen.           *
' *                                                                   *
' * The QBGUI_POPCHECK compile-time symbol must be set to True when   *
' * the quickBasicEngine referenced by this form was compiled with    *
' * its QUICKBASICENGINE_POPCHECK symbol set to True: QBGUI_POPCHECK  *
' * must otherwise be set to False.                                   *
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
        MyBase.dispose(disposing)
    End Sub
    Friend WithEvents cmdClose As System.Windows.Forms.Button

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.Container

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdRun As System.Windows.Forms.Button
    Friend WithEvents cmdEvaluate As System.Windows.Forms.Button
    Friend WithEvents gbxCE As System.Windows.Forms.GroupBox
    Friend WithEvents txtCE As System.Windows.Forms.TextBox
    Friend WithEvents cmdCEobject2XML As System.Windows.Forms.Button
    Friend WithEvents cmdCEinspector As System.Windows.Forms.Button
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents chkViewOutput As System.Windows.Forms.CheckBox
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSourceCodeLoad As System.Windows.Forms.MenuItem
    Friend WithEvents mnu As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFileSourceCodeSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSourceCodeClear As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTools As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsCompile As System.Windows.Forms.MenuItem
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuFileExitNoSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsForm2Registry As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsRegistry2Form As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsScan As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsSep2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsRun As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsSep3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsScanForced As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsAssemble As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsCompileForced As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsAssembleForced As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuToolsOptions As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsOptionsCloneTest As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsViewEventLog As System.Windows.Forms.MenuItem
    Friend WithEvents lblStatusLbl As System.Windows.Forms.Label
    Friend WithEvents cmdStatusZoom As System.Windows.Forms.Button
    Friend WithEvents cmdCEtest As System.Windows.Forms.Button
    Friend WithEvents mnuRegistryClear As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsDisplayFormControls As System.Windows.Forms.MenuItem
    Friend WithEvents mnuToolsMSILrun As System.Windows.Forms.MenuItem
    Friend WithEvents chkCEtestEventLog As System.Windows.Forms.CheckBox
    Friend WithEvents rtbScreen As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.cmdClose = New System.Windows.Forms.Button
Me.cmdRun = New System.Windows.Forms.Button
Me.cmdEvaluate = New System.Windows.Forms.Button
Me.gbxCE = New System.Windows.Forms.GroupBox
Me.chkCEtestEventLog = New System.Windows.Forms.CheckBox
Me.cmdCEtest = New System.Windows.Forms.Button
Me.cmdCEinspector = New System.Windows.Forms.Button
Me.cmdCEobject2XML = New System.Windows.Forms.Button
Me.txtCE = New System.Windows.Forms.TextBox
Me.lstStatus = New System.Windows.Forms.ListBox
Me.lblProgress = New System.Windows.Forms.Label
Me.chkViewOutput = New System.Windows.Forms.CheckBox
Me.mnu = New System.Windows.Forms.MainMenu
Me.mnuFile = New System.Windows.Forms.MenuItem
Me.mnuFileSourceCodeLoad = New System.Windows.Forms.MenuItem
Me.mnuFileSourceCodeSave = New System.Windows.Forms.MenuItem
Me.MenuItem5 = New System.Windows.Forms.MenuItem
Me.mnuFileSourceCodeClear = New System.Windows.Forms.MenuItem
Me.MenuItem7 = New System.Windows.Forms.MenuItem
Me.mnuFileExitNoSave = New System.Windows.Forms.MenuItem
Me.mnuFileExit = New System.Windows.Forms.MenuItem
Me.mnuTools = New System.Windows.Forms.MenuItem
Me.mnuToolsScan = New System.Windows.Forms.MenuItem
Me.mnuToolsScanForced = New System.Windows.Forms.MenuItem
Me.mnuToolsCompile = New System.Windows.Forms.MenuItem
Me.mnuToolsCompileForced = New System.Windows.Forms.MenuItem
Me.mnuToolsAssemble = New System.Windows.Forms.MenuItem
Me.mnuToolsAssembleForced = New System.Windows.Forms.MenuItem
Me.mnuToolsMSILrun = New System.Windows.Forms.MenuItem
Me.mnuToolsSep3 = New System.Windows.Forms.MenuItem
Me.mnuToolsRun = New System.Windows.Forms.MenuItem
Me.mnuToolsSep1 = New System.Windows.Forms.MenuItem
Me.mnuToolsSep2 = New System.Windows.Forms.MenuItem
Me.mnuToolsForm2Registry = New System.Windows.Forms.MenuItem
Me.mnuToolsRegistry2Form = New System.Windows.Forms.MenuItem
Me.mnuRegistryClear = New System.Windows.Forms.MenuItem
Me.MenuItem1 = New System.Windows.Forms.MenuItem
Me.mnuToolsOptions = New System.Windows.Forms.MenuItem
Me.mnuToolsOptionsCloneTest = New System.Windows.Forms.MenuItem
Me.mnuToolsViewEventLog = New System.Windows.Forms.MenuItem
Me.mnuToolsDisplayFormControls = New System.Windows.Forms.MenuItem
Me.mnuHelp = New System.Windows.Forms.MenuItem
Me.mnuHelpAbout = New System.Windows.Forms.MenuItem
Me.ofd = New System.Windows.Forms.OpenFileDialog
Me.sfd = New System.Windows.Forms.SaveFileDialog
Me.lblStatusLbl = New System.Windows.Forms.Label
Me.cmdStatusZoom = New System.Windows.Forms.Button
Me.rtbScreen = New System.Windows.Forms.RichTextBox
Me.gbxCE.SuspendLayout()
Me.SuspendLayout()
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(355, 201)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(105, 32)
Me.cmdClose.TabIndex = 1
Me.cmdClose.Text = "Close"
'
'cmdRun
'
Me.cmdRun.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdRun.Location = New System.Drawing.Point(100, 201)
Me.cmdRun.Name = "cmdRun"
Me.cmdRun.Size = New System.Drawing.Size(53, 32)
Me.cmdRun.TabIndex = 3
Me.cmdRun.Text = "Run"
'
'cmdEvaluate
'
Me.cmdEvaluate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdEvaluate.Location = New System.Drawing.Point(7, 201)
Me.cmdEvaluate.Name = "cmdEvaluate"
Me.cmdEvaluate.Size = New System.Drawing.Size(86, 32)
Me.cmdEvaluate.TabIndex = 4
Me.cmdEvaluate.Text = "Evaluate"
'
'gbxCE
'
Me.gbxCE.Controls.Add(Me.chkCEtestEventLog)
Me.gbxCE.Controls.Add(Me.cmdCEtest)
Me.gbxCE.Controls.Add(Me.cmdCEinspector)
Me.gbxCE.Controls.Add(Me.cmdCEobject2XML)
Me.gbxCE.Controls.Add(Me.txtCE)
Me.gbxCE.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxCE.Location = New System.Drawing.Point(467, 7)
Me.gbxCE.Name = "gbxCE"
Me.gbxCE.Size = New System.Drawing.Size(386, 340)
Me.gbxCE.TabIndex = 5
Me.gbxCE.TabStop = False
Me.gbxCE.Text = "Customer Engineering Zone"
Me.gbxCE.Visible = False
'
'chkCEtestEventLog
'
Me.chkCEtestEventLog.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkCEtestEventLog.Location = New System.Drawing.Point(260, 305)
Me.chkCEtestEventLog.Name = "chkCEtestEventLog"
Me.chkCEtestEventLog.Size = New System.Drawing.Size(113, 14)
Me.chkCEtestEventLog.TabIndex = 7
Me.chkCEtestEventLog.Text = "Test event log"
'
'cmdCEtest
'
Me.cmdCEtest.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCEtest.Location = New System.Drawing.Point(180, 305)
Me.cmdCEtest.Name = "cmdCEtest"
Me.cmdCEtest.Size = New System.Drawing.Size(73, 28)
Me.cmdCEtest.TabIndex = 6
Me.cmdCEtest.Text = "Test"
'
'cmdCEinspector
'
Me.cmdCEinspector.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCEinspector.Location = New System.Drawing.Point(93, 305)
Me.cmdCEinspector.Name = "cmdCEinspector"
Me.cmdCEinspector.Size = New System.Drawing.Size(74, 28)
Me.cmdCEinspector.TabIndex = 5
Me.cmdCEinspector.Text = "Inspect"
'
'cmdCEobject2XML
'
Me.cmdCEobject2XML.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCEobject2XML.Location = New System.Drawing.Point(7, 305)
Me.cmdCEobject2XML.Name = "cmdCEobject2XML"
Me.cmdCEobject2XML.Size = New System.Drawing.Size(73, 28)
Me.cmdCEobject2XML.TabIndex = 4
Me.cmdCEobject2XML.Text = "XML"
'
'txtCE
'
Me.txtCE.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.txtCE.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtCE.ForeColor = System.Drawing.Color.Black
Me.txtCE.Location = New System.Drawing.Point(7, 21)
Me.txtCE.Multiline = True
Me.txtCE.Name = "txtCE"
Me.txtCE.ScrollBars = System.Windows.Forms.ScrollBars.Both
Me.txtCE.Size = New System.Drawing.Size(373, 277)
Me.txtCE.TabIndex = 3
Me.txtCE.Text = ""
Me.txtCE.WordWrap = False
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 15
Me.lstStatus.Location = New System.Drawing.Point(7, 257)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(453, 64)
Me.lstStatus.TabIndex = 6
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress.Location = New System.Drawing.Point(7, 340)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(453, 8)
Me.lblProgress.TabIndex = 7
Me.lblProgress.Visible = False
'
'chkViewOutput
'
Me.chkViewOutput.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkViewOutput.Location = New System.Drawing.Point(160, 201)
Me.chkViewOutput.Name = "chkViewOutput"
Me.chkViewOutput.Size = New System.Drawing.Size(64, 32)
Me.chkViewOutput.TabIndex = 8
Me.chkViewOutput.Text = "View output"
'
'mnu
'
Me.mnu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuTools, Me.mnuHelp})
'
'mnuFile
'
Me.mnuFile.Index = 0
Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileSourceCodeLoad, Me.mnuFileSourceCodeSave, Me.MenuItem5, Me.mnuFileSourceCodeClear, Me.MenuItem7, Me.mnuFileExitNoSave, Me.mnuFileExit})
Me.mnuFile.Text = "File"
'
'mnuFileSourceCodeLoad
'
Me.mnuFileSourceCodeLoad.Index = 0
Me.mnuFileSourceCodeLoad.Text = "Load source code..."
'
'mnuFileSourceCodeSave
'
Me.mnuFileSourceCodeSave.Index = 1
Me.mnuFileSourceCodeSave.Text = "Save source code..."
'
'MenuItem5
'
Me.MenuItem5.Index = 2
Me.MenuItem5.Text = "-"
'
'mnuFileSourceCodeClear
'
Me.mnuFileSourceCodeClear.Index = 3
Me.mnuFileSourceCodeClear.Text = "Clear source code"
'
'MenuItem7
'
Me.MenuItem7.Index = 4
Me.MenuItem7.Text = "-"
'
'mnuFileExitNoSave
'
Me.mnuFileExitNoSave.Index = 5
Me.mnuFileExitNoSave.Text = "Exit (don't save any settings)"
'
'mnuFileExit
'
Me.mnuFileExit.Index = 6
Me.mnuFileExit.Text = "Exit"
'
'mnuTools
'
Me.mnuTools.Index = 1
Me.mnuTools.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuToolsScan, Me.mnuToolsScanForced, Me.mnuToolsCompile, Me.mnuToolsCompileForced, Me.mnuToolsAssemble, Me.mnuToolsAssembleForced, Me.mnuToolsMSILrun, Me.mnuToolsSep3, Me.mnuToolsRun, Me.mnuToolsSep1, Me.mnuToolsSep2, Me.mnuToolsForm2Registry, Me.mnuToolsRegistry2Form, Me.mnuRegistryClear, Me.MenuItem1, Me.mnuToolsOptions, Me.mnuToolsOptionsCloneTest, Me.mnuToolsViewEventLog, Me.mnuToolsDisplayFormControls})
Me.mnuTools.Text = "Tools"
'
'mnuToolsScan
'
Me.mnuToolsScan.Index = 0
Me.mnuToolsScan.Text = "Scan"
'
'mnuToolsScanForced
'
Me.mnuToolsScanForced.Index = 1
Me.mnuToolsScanForced.Text = "Scan (forced)"
'
'mnuToolsCompile
'
Me.mnuToolsCompile.Index = 2
Me.mnuToolsCompile.Text = "Compile"
'
'mnuToolsCompileForced
'
Me.mnuToolsCompileForced.Index = 3
Me.mnuToolsCompileForced.Text = "Compile (forced)"
'
'mnuToolsAssemble
'
Me.mnuToolsAssemble.Index = 4
Me.mnuToolsAssemble.Text = "Assemble"
'
'mnuToolsAssembleForced
'
Me.mnuToolsAssembleForced.Index = 5
Me.mnuToolsAssembleForced.Text = "Assemble (forced)"
'
'mnuToolsMSILrun
'
Me.mnuToolsMSILrun.Index = 6
Me.mnuToolsMSILrun.Text = "Run the MSIL code"
'
'mnuToolsSep3
'
Me.mnuToolsSep3.Index = 7
Me.mnuToolsSep3.Text = "-"
'
'mnuToolsRun
'
Me.mnuToolsRun.Index = 8
Me.mnuToolsRun.Text = "Run"
'
'mnuToolsSep1
'
Me.mnuToolsSep1.Index = 9
Me.mnuToolsSep1.Text = "-"
'
'mnuToolsSep2
'
Me.mnuToolsSep2.Index = 10
Me.mnuToolsSep2.Text = "-"
'
'mnuToolsForm2Registry
'
Me.mnuToolsForm2Registry.Index = 11
Me.mnuToolsForm2Registry.Text = "Save Settings"
'
'mnuToolsRegistry2Form
'
Me.mnuToolsRegistry2Form.Index = 12
Me.mnuToolsRegistry2Form.Text = "Restore Settings"
'
'mnuRegistryClear
'
Me.mnuRegistryClear.Index = 13
Me.mnuRegistryClear.Text = "Clear settings"
'
'MenuItem1
'
Me.MenuItem1.Index = 14
Me.MenuItem1.Text = "-"
'
'mnuToolsOptions
'
Me.mnuToolsOptions.Index = 15
Me.mnuToolsOptions.Text = "Options"
'
'mnuToolsOptionsCloneTest
'
Me.mnuToolsOptionsCloneTest.Index = 16
Me.mnuToolsOptionsCloneTest.Text = "Clone Test"
'
'mnuToolsViewEventLog
'
Me.mnuToolsViewEventLog.Index = 17
Me.mnuToolsViewEventLog.Text = "View event log"
'
'mnuToolsDisplayFormControls
'
Me.mnuToolsDisplayFormControls.Index = 18
Me.mnuToolsDisplayFormControls.Text = "Display form's controls"
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
'lblStatusLbl
'
Me.lblStatusLbl.BackColor = System.Drawing.Color.Black
Me.lblStatusLbl.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatusLbl.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(255, Byte), CType(128, Byte))
Me.lblStatusLbl.Location = New System.Drawing.Point(7, 236)
Me.lblStatusLbl.Name = "lblStatusLbl"
Me.lblStatusLbl.Size = New System.Drawing.Size(451, 21)
Me.lblStatusLbl.TabIndex = 9
Me.lblStatusLbl.Text = "Status"
Me.lblStatusLbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdStatusZoom
'
Me.cmdStatusZoom.Location = New System.Drawing.Point(413, 236)
Me.cmdStatusZoom.Name = "cmdStatusZoom"
Me.cmdStatusZoom.Size = New System.Drawing.Size(48, 21)
Me.cmdStatusZoom.TabIndex = 10
Me.cmdStatusZoom.Text = "Zoom"
'
'rtbScreen
'
Me.rtbScreen.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.rtbScreen.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.rtbScreen.Location = New System.Drawing.Point(8, 8)
Me.rtbScreen.Name = "rtbScreen"
Me.rtbScreen.Size = New System.Drawing.Size(448, 184)
Me.rtbScreen.TabIndex = 11
Me.rtbScreen.Text = ""
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(854, 351)
Me.Controls.Add(Me.rtbScreen)
Me.Controls.Add(Me.cmdStatusZoom)
Me.Controls.Add(Me.lblStatusLbl)
Me.Controls.Add(Me.chkViewOutput)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.gbxCE)
Me.Controls.Add(Me.cmdEvaluate)
Me.Controls.Add(Me.cmdRun)
Me.Controls.Add(Me.cmdClose)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Menu = Me.mnu
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Edward Nilges' Version of Quick Basic, a Microsoft Product"
Me.gbxCE.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    ' ***** Quick Basic Engine *****
    Private WithEvents OBJqbe As quickBasicEngine.qbQuickBasicEngine

    ' ***** Tool tips *****
    Private OBJtoolTips As ToolTip

    ' ***** Trace controls *****
    Private WithEvents CMDmore As Button
    Private WithEvents CTLstatusDisplay As statusDisplay

    ' ***** Shared *****
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities

    ' ***** Collection tools *****
    Private OBJcollectionUtilities As collectionUtilities.collectionUtilities

    ' ***** Constants *****
    Private Const MORE_CAPTION As String = "More"
    Private Const LESS_CAPTION As String = "Less"
    Private Const MAX_TRACE As Integer = 100000
    Private Const SCREEN_HEIGHT_TOLERANCE As Single = 0.9
    Private Const SCREEN_WIDTH_TOLERANCE As Single = 0.98
    Private Const RESIZEHANDLER_MAXCACHE As Integer = 100

    ' ***** Options form *****
    Private FRMoptions As options

    ' ***** Extra debug screen *****
    Private TXTsourceDebug As TextBox

    ' ***** Run form *****
    Private FRMrunForm As frmRun

    ' ***** XML and parse stack *****
    Private Class parseStackEntry
        Public GrammarCategory As String            ' Grammar category
        Public StartIndex As Integer                ' Start index in XML display
    End Class
    Private OBJparseStack As Stack                  ' Of parseStackEntry

#End Region ' Form data

#Region " Form events "

    Private Sub chkViewOutput_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkViewOutput.CheckedChanged
        changeView(chkViewOutput.Checked)
    End Sub

    Private Sub cmdCEinspector_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCEinspector.Click
        Dim strReport As String
        If OBJqbe.inspect(strReport) Then
            MsgBox("The quickBasicEngine passes its inspection")
        Else
            MsgBox("The quickBasicEngine has failed its inspection")
        End If
        txtCE.Text = strReport
    End Sub

    Private Sub cmdCEobject2XML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCEobject2XML.Click
        txtCE.Text = OBJqbe.object2XML(booAboutComment:=True)
    End Sub

    Private Sub cmdCEtest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCEtest.Click
        Dim strReport As String
        If FRMoptions.StopButton Then showRunForm()
        Dim objThread As Threading.Thread = New Threading.Thread(AddressOf testInterface)
        objThread.Start()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer()
    End Sub

    Private Sub cmdEvaluate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEvaluate.Click
        If FRMoptions.StopButton Then showRunForm()
        Dim objThread As Threading.Thread = New Threading.Thread(AddressOf evaluateInterface)
        objThread.Start()
    End Sub

    Private Sub cmdMore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        With CMDmore
            If .Text = "More" Then
                showMore()
            ElseIf .Text = "Less" Then
                showLess()
            Else
                MsgBox("Programming error: unexpected caption") : dispose() : End
            End If
        End With
    End Sub

    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click
        run()
    End Sub

    Private Sub cmdStatusZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStatusZoom.Click
        statusZoom()
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        closer()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim booExtension As Boolean
        Dim booPopCheck As Boolean
#If QBGUI_EXTENSIONS Then
                booExtension = True
#End If
#If QBGUI_POPCHECK Then
                booPopCheck = True
#End If
        If ExtensionAvailable <> booExtension _
            OrElse _
            PopCheckAvailable <> booPopCheck Then
            MsgBox("Note that the compiler build is in error. " & _
                   "The settings of the flags QBGUI_EXTENSION and " & _
                   "QUICKBASICENGINE_EXTENSION and/or the settings of " & _
                   "QBGUI_POPCHECK and QUICKBASICENGINE_POPCHECK are " & _
                   "different. The compiler execution has been canceled. " & _
                   "Recompile with matched settings for qbGUI " & _
                   "and quickBasicEngine to fix this problem.")
            End
        End If
        Dim booExecutionRecord As Boolean
        Try
            booExecutionRecord = _
                CBool(GetSetting(Application.ProductName, _
                        "ExecutionRecord", _
                        "ExecutionRecord", _
                        "False"))
        Catch : End Try
        If Not booExecutionRecord Then
            showAbout()
            Try
                SaveSetting(Application.ProductName, _
                            "ExecutionRecord", _
                            "ExecutionRecord", _
                            "True")
                SaveSetting(Application.ProductName, _
                            "ExecutionRecord", _
                            "FirstUsedOn", _
                            CStr(Now))
            Catch : End Try
        End If
        mkScreenTag()
        Height = cmdClose.Top + cmdClose.Height + 48
        Width = cmdClose.Left + cmdClose.Width + 16
        CenterToScreen()
        setToolTips()
        mkMoreButton()
        Try
            FRMoptions = New options
        Catch objException As Exception
            errorHandler("Cannot create the options form: " & _
                         Err.Number & " " & Err.Description & _
                         vbNewLine & vbNewLine & _
                         objException.ToString)
        End Try
        Try
            OBJqbe = New quickBasicEngine.qbQuickBasicEngine
            addLoopEventHandler()
        Catch
            _OBJutilities.errorHandler("Cannot create my Quick Basic engine: " & _
                                       Err.Number & " " & Err.Description, _
                                       Name, "Form1_Load", _
                                       "Cannot proceed")
            Return
        End Try
        registry2Form()
        If Not booExecutionRecord Then
            rtbScreen.Text = "Print ""Hello world"""
        End If
        Try
            OBJcollectionUtilities = New collectionUtilities.collectionUtilities
        Catch
            MsgBox("Cannot create collection utilities: " & _
                   Err.Number & " " & Err.Description)
            End
        End Try
    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        closer()
    End Sub

    Private Sub mnuFileExitNoSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExitNoSave.Click
        closer(False)
    End Sub

    Private Sub mnuFileSourceCodeClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSourceCodeClear.Click
        clearSourceCode()
    End Sub

    Private Sub mnuFileSourceCodeLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSourceCodeLoad.Click
        changeView(False)
        sourceCodeLoad()
    End Sub

    Private Sub mnuFileSourceCodeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSourceCodeSave.Click
        sourceCodeSave()
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        showAbout()
    End Sub

    Private Sub mnuRegistryClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRegistryClear.Click
        registryClear()
    End Sub

    Private Sub mnuToolsAssemble_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsAssemble.Click
        assembleInterface(False)
    End Sub

    Private Sub mnuToolsAssembleForced_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsAssembleForced.Click
        assembleInterface(True)
    End Sub

    Private Sub mnuToolsCompile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsCompile.Click
        compileInterface(False)
    End Sub

    Private Sub mnuToolsCompileForced_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsCompileForced.Click
        compileInterface(True)
    End Sub

    Private Sub mnuToolsDisplayFormControls_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsDisplayFormControls.Click
        displayFormControls()
    End Sub

    Private Sub mnuToolsForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsForm2Registry.Click
        form2Registry()
    End Sub

    Private Sub mnuToolsMSILrun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMSILrun.Click
        MSILrunInterface()
    End Sub

    Private Sub mnuToolsOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsOptions.Click
        showOptions()
    End Sub

    Private Sub mnuToolsOptionsCloneTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsOptionsCloneTest.Click
        cloneTest()
    End Sub

    Private Sub mnuToolsRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsRegistry2Form.Click
        registry2Form()
    End Sub

    Private Sub mnuToolsRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsRun.Click
        run()
    End Sub

    Private Sub mnuToolsScan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsScan.Click
        scanInterface(False)
    End Sub

    Private Sub mnuToolsScanForced_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsScanForced.Click
        scanInterface(True)
    End Sub

    Private Sub mnuToolsViewEventLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsViewEventLog.Click
        viewEventLog()
    End Sub

#End Region ' Form events 

#Region " General procedures "

    ' ----------------------------------------------------------------
    ' Add the loop event handler
    '
    '
    Private Sub addLoopEventHandler()
        AddHandler OBJqbe.loopEvent, AddressOf loopEventHandler
    End Sub

    ' ----------------------------------------------------------------
    ' Add IO timings to Run tag
    '
    '
    Private Sub adjustRunTime(ByVal datStart As Date)
        Try
            With cmdRun
                .Tag = CInt(.Tag) + DateDiff("s", datStart, Now)
            End With
        Catch : End Try
    End Sub

    ' ----------------------------------------------------------------
    ' Interface to assembler
    '
    '
    Private Sub assembleInterface(ByVal booForceOp As Boolean)
        qbeInterface("ASSEMBLE", booForceOp)
    End Sub

    ' ----------------------------------------------------------------
    ' Change view (between source code and output)
    '
    '
    Private Sub changeView(ByVal booOutput As Boolean)
        With rtbScreen
            Dim colTag As Collection = CType(.Tag, Collection)
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            If booOutput Then
                .BackColor = Color.Black
                .ForeColor = Color.LightGreen
                intIndex1 = 2
                intIndex2 = 3
            Else
                .BackColor = CType(colTag.Item(4), Color)
                .ForeColor = CType(colTag.Item(5), Color)
                intIndex1 = 3
                intIndex2 = 2
            End If
            With colTag
                .Remove(intIndex1)
                .Add(rtbScreen.Text, , intIndex1)
            End With
            .Text = CStr(colTag.Item(intIndex2))
        End With
        mnuFileSourceCodeClear.Enabled = Not booOutput
        mnuFileSourceCodeLoad.Enabled = Not booOutput
        mnuFileSourceCodeSave.Enabled = Not booOutput
    End Sub

    ' -----------------------------------------------------------------
    ' Clear source code and associated tags
    '
    '
    Private Sub clearSourceCode()
        With rtbScreen
            .Clear()
            OBJcollectionUtilities.collectionClear(CType(.Tag, Collection))
            mkScreenTag()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Test cloning
    '
    '
    Private Sub cloneTest()
        With OBJqbe
            Dim objClone As quickBasicEngine.qbQuickBasicEngine
            Try
                objClone = OBJqbe.clone
                MsgBox("Clone " & _
                        _OBJutilities.enquote(objClone.Name) & _
                        "of " & _
                        _OBJutilities.enquote(.Name) & " " & _
                        CStr(IIf(.compareTo(objClone) AndAlso objClone.inspect(), _
                                "matches the original object and passes an inspection", _
                                "does not match the original object or fails an inspection")))
            Catch
                errorHandler("Error in clone and/or compare: " & _
                             Err.Number & " " & Err.Description)
            End Try
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Close and end
    '
    '
    Private Overloads Sub closer()
        closer(True)
    End Sub
    Private Overloads Sub closer(ByVal booSave As Boolean)
        If booSave Then form2Registry()
        If Not (OBJqbe Is Nothing) Then
            OBJqbe.dispose() : OBJqbe = Nothing
        End If
        If Not (CMDmore Is Nothing) Then
            CMDmore.dispose() : CMDmore = Nothing
        End If
        If Not (CTLstatusDisplay Is Nothing) Then
            CTLstatusDisplay.dispose() : CTLstatusDisplay = Nothing
        End If
        If Not (OBJcollectionUtilities Is Nothing) Then
            OBJcollectionUtilities.dispose() : OBJcollectionUtilities = Nothing
        End If
        If Not (FRMoptions Is Nothing) Then
            FRMoptions.disposer() : FRMoptions = Nothing
        End If
        If Not (TXTsourceDebug Is Nothing) Then
            disposeSourceDebugScreen()
        End If
        dispose()
        End
    End Sub

    ' ----------------------------------------------------------------
    ' Close run form
    '
    ' 
    Private Sub closeRunForm()
        If (FRMrunForm Is Nothing) Then Return
        With FRMrunForm
            .Hide()
            With CType(.Tag, Collection)
                Dim intIndex1 As Integer
                Dim booVisible As Boolean = CBool(.Item(1))
                Dim booEnabled As Boolean = CBool(.Item(2))
                For intIndex1 = 3 To .Count
                    With CType(.Item(intIndex1), Control)
                        .Visible = booVisible : .Enabled = booEnabled
                    End With
                Next intIndex1
            End With
        End With
        With OBJqbe
            If .getThreadStatus <> "Ready" Then
                .resumeQBE() : .reset()
            End If
        End With
        If screenIsMore() Then
            Dim intOldHeight As Integer = Height
            Height = CInt(_OBJwindowsUtilities.screenHeight * SCREEN_HEIGHT_TOLERANCE)
            resizeHandler(Width, intOldHeight)
        End If
        CenterToScreen()
    End Sub

    ' ----------------------------------------------------------------
    ' Interface to compile
    '
    '
    Private Sub compileInterface(ByVal booForceOp As Boolean)
        qbeInterface("COMPILE", booForceOp)
    End Sub

    ' ----------------------------------------------------------------
    ' Create control view
    '
    ' This method rearranges the compiler views. Its objView parameter
    ' should be a ParamArray. 
    '
    ' The ParamArray should be a vector of elements in the following
    ' form
    '
    '      listBox,label,zoom,status 
    '   
    ' listBox should be one of the display list boxes, label should 
    ' be its associated label, and zoom should be its associated zoom
    ' button. 
    '
    ' Status should be False to make the controls invisible or else
    ' a Double precision value between 0 and 1: in the latter case this
    ' is the percent of the available width the control must occupy.
    '
    ' The controls should be listed in the left-to-right order in which
    ' they need to appear.
    '
    ' This method returns a new Collection that contains
    ' a small number of 4 item subcollections: each subcollection 
    ' contains:
    '
    '
    '      *  The handle of a modified control in Item(1)
    '      *  Its previous Left property in Item(2)
    '      *  Its previous Width property in Item(3) 
    '      *  Its previous Visible property in Item(4) 
    '
    '
    Private Function createControlView(ByVal ParamArray objFormat() As Object) _
            As Collection
        Dim colRestore As Collection
        Try
            colRestore = New Collection
        Catch
            errorHandler("Cannot create restore collection: " & _
                         Err.Number & " " & Err.Description)
            Return (Nothing)
        End Try
        Dim cmdButton As Button
        Dim colEntry As Collection
        Dim dblRatio As Double
        Dim intIndex1 As Integer
        Dim intLeft As Integer = _OBJwindowsUtilities.Grid
        Dim intWidth As Integer = _
            CTLstatusDisplay.GroupBoxHandle.Width _
            - _
            _OBJwindowsUtilities.Grid * 2
        Dim lblLabel As Label
        Dim lstBox As ListBox
        For intIndex1 = 0 To UBound(objFormat) Step 4
            lstBox = CType(objFormat(intIndex1), ListBox)
            lblLabel = CType(objFormat(intIndex1 + 1), Label)
            cmdButton = CType(objFormat(intIndex1 + 2), Button)
            Try
                colEntry = createControlView_mkRestoreEntry(lstBox)
                If (colEntry Is Nothing) Then Return (Nothing)
                colRestore.Add(colEntry)
                colEntry = createControlView_mkRestoreEntry(lblLabel)
                If (colEntry Is Nothing) Then Return (Nothing)
                colRestore.Add(colEntry)
                colEntry = createControlView_mkRestoreEntry(cmdButton)
                If (colEntry Is Nothing) Then Return (Nothing)
                colRestore.Add(colEntry)
            Catch
                errorHandler("Cannot extend restore collection: " & _
                            Err.Number & " " & Err.Description)
                Return (Nothing)
            End Try
            If (objFormat(intIndex1 + 3).GetType.ToString) <> "SYSTEM.BOOLEAN" Then
                lstBox.Width = CInt(intWidth * CDbl(objFormat(intIndex1 + 3)))
                dblRatio = Math.Min(0.5, cmdButton.Width / lblLabel.Width)
                lblLabel.Width = lstBox.Width
                cmdButton.Width = CInt(lblLabel.Width * dblRatio)
                lstBox.Left = intLeft : intLeft += lstBox.Width
                lblLabel.Left = lstBox.Left
                cmdButton.Left = lblLabel.Right - cmdButton.Width
                lstBox.Visible = True
                lblLabel.Visible = True
                cmdButton.Visible = True
                cmdButton.BringToFront()
                lstBox.Refresh() : lblLabel.Refresh() : cmdButton.Refresh()
                CTLstatusDisplay.Refresh()
            Else
                lstBox.Visible = False
                lblLabel.Visible = False
                cmdButton.Visible = False
            End If
        Next intIndex1
        Return (colRestore)
    End Function

    ' ----------------------------------------------------------------
    ' Make the restore entry on behalf of createControlView
    '
    '
    Private Function createControlView_mkRestoreEntry(ByVal ctlControl As Control) _
            As Collection
        Dim colEntry As Collection
        Try
            colEntry = New Collection
            With colEntry
                .Add(ctlControl)
                .Add(ctlControl.Left)
                .Add(ctlControl.Width)
                .Add(ctlControl.Visible)
            End With
            Return (colEntry)
        Catch
            errorHandler("Cannot make restore entry collection: " & _
                         Err.Number & " " & Err.Description)
            Return (Nothing)
        End Try
    End Function

    ' ----------------------------------------------------------------
    ' Create the parse view consisting of a small scan view and a
    ' larger parse tree view
    '
    '
    Private Function createParseView() As Collection
        With CTLstatusDisplay
            Return (createControlView(.ScanListBox, _
                                      .ScanListBoxLabel, _
                                      .ScanZoomButton, _
                                      0.2, _
                                      .ParseListBox, _
                                      .ParseListBoxLabel, _
                                      .ParseZoomButton, _
                                      0.6, _
                                      .RPNlistBox, _
                                      .RPNlistBoxLabel, _
                                      .RPNzoomButton, _
                                      0.2, _
                                      .StackListBox, _
                                      .StackListBoxLabel, _
                                      .StackZoomButton, _
                                      False, _
                                      .StorageListBox, _
                                      .StorageListBoxLabel, _
                                      .StorageZoomButton, _
                                      False))
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Create the runtime view of compiler data structures
    '
    '
    Private Function createRunView() As Collection
        With CTLstatusDisplay
            Return (createControlView(.ScanListBox, _
                                      .ScanListBoxLabel, _
                                      .ScanZoomButton, _
                                      False, _
                                      .ParseListBox, _
                                      .ParseListBoxLabel, _
                                      .ParseZoomButton, _
                                      False, _
                                      .RPNlistBox, _
                                      .RPNlistBoxLabel, _
                                      .RPNzoomButton, _
                                      0.33, _
                                      .StackListBox, _
                                      .StackListBoxLabel, _
                                      .StackZoomButton, _
                                      0.33, _
                                      .StorageListBox, _
                                      .StorageListBoxLabel, _
                                      .StorageZoomButton, _
                                      0.33))
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Create the scan view 
    '
    '
    Private Function createScanView() As Collection
        With CTLstatusDisplay
            Return (createControlView(.ScanListBox, _
                                      .ScanListBoxLabel, _
                                      .ScanZoomButton, _
                                      1, _
                                      .ParseListBox, _
                                      .ParseListBoxLabel, _
                                      .ParseZoomButton, _
                                      False, _
                                      .RPNlistBox, _
                                      .RPNlistBoxLabel, _
                                      .RPNzoomButton, _
                                      False, _
                                      .StackListBox, _
                                      .StackListBoxLabel, _
                                      .StackZoomButton, _
                                      False, _
                                      .StorageListBox, _
                                      .StorageListBoxLabel, _
                                      .StorageZoomButton, _
                                      False))
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Create the standard view of compiler data structures
    '
    '
    Private Function createStdView() As Collection
        With CTLstatusDisplay
            Return (createControlView(.ScanListBox, _
                                      .ScanListBoxLabel, _
                                      .ScanZoomButton, _
                                      0.2, _
                                      .ParseListBox, _
                                      .ParseListBoxLabel, _
                                      .ParseZoomButton, _
                                      0.3, _
                                      .RPNlistBox, _
                                      .RPNlistBoxLabel, _
                                      .RPNzoomButton, _
                                      0.2, _
                                      .StackListBox, _
                                      .StackListBoxLabel, _
                                      .StackZoomButton, _
                                      0.15, _
                                      .StorageListBox, _
                                      .StorageListBoxLabel, _
                                      .StorageZoomButton, _
                                      0.15))
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Dumps the form's controls
    '
    '
    Private Sub displayFormControls()
        txtCE.Text = "Building form control display..."
        Dim strDisplay As String
        displayFormControls_(Me, strDisplay, "")
        txtCE.Text = strDisplay
    End Sub

    ' ----------------------------------------------------------------
    ' Dumps the form's controls (recursion)
    '
    '
    Private Sub displayFormControls_(ByVal ctlParent As Control, _
                                     ByRef strDisplay As String, _
                                     ByVal strIndent As String)
        Dim intIndex1 As Integer
        With ctlParent.Controls
            For intIndex1 = 0 To .Count - 1
                With .Item(intIndex1)
                    _OBJutilities.append(strDisplay, _
                                         vbNewLine, _
                                         strIndent & .GetType.ToString & " " & .Name)
                End With
                displayFormControls_(.Item(intIndex1), strDisplay, strIndent & " ")
            Next intIndex1
        End With
        txtCE.Text = strDisplay
    End Sub

    ' ----------------------------------------------------------------
    ' Display output results
    '
    '
    ' --- Prefixes output with two newlines
    Private Overloads Sub displayOutput(ByVal strText As String)
        displayOutput(strText, vbNewLine & vbNewLine)
    End Sub
    Private Overloads Sub displayOutput(ByVal strText As String, _
                                        ByVal strPrefix As String, _
                                        Optional ByVal booCLS As Boolean = False)
        With rtbScreen
            chkViewOutput.Checked = True
            setScreenModeOutput()
            If booCLS Then
                .Text = strPrefix & strText
            Else
                .Text &= strPrefix & strText
            End If
            .SelectionStart += Len(.Text)
            .SelectionLength = 0
            .ScrollToCaret()
            .Refresh()
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Dispose the source debug screen
    '
    '
    Private Sub disposeSourceDebugScreen()
        TXTsourceDebug.dispose() : TXTsourceDebug = Nothing
    End Sub

    ' ----------------------------------------------------------------
    ' Return the error color
    '
    '
    Private Function errorColor() As Color
        Return Color.Red
    End Function

    ' ----------------------------------------------------------------
    ' Format error information
    '
    '
    Private Function editErrorInfo(ByVal strErrorInfo As String) As String
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim strDelimiter1 As String = vbNewLine & "Error at "
        Dim strDelimiter2 As String = vbNewLine & "Informational message at "
        Dim strOutstring As String
        Do While intIndex1 <= Len(strErrorInfo)
            intIndex2 = Math.Min(InStr(intIndex1, _
                                       strErrorInfo & strDelimiter1, _
                                       strDelimiter1), _
                                 InStr(intIndex1, _
                                       strErrorInfo & strDelimiter2, _
                                       strDelimiter2))
            _OBJutilities.append(strOutstring, _
                                 vbNewLine & vbNewLine, _
                                 Mid(strErrorInfo, _
                                     intIndex1, _
                                     intIndex2 - intIndex1))
            intIndex1 = intIndex2 + Len(vbNewLine)
        Loop
        Return strOutstring
    End Function

    ' ----------------------------------------------------------------
    ' Error handler
    '
    '
    Private Sub errorHandler(ByVal strMessage As String)
        Dim strMessageWork As String = _
            "Error at " & Now & vbNewLine & vbNewLine & strMessage
        Dim booOK As Boolean
        If Len(strMessageWork) > 4096 Then
            Dim frmZoom As zoom.zoom
            Try
                frmZoom = New zoom.zoom
                With frmZoom
                    .ZoomTextBox.Text = strMessageWork
                    .showZoom()
                End With
                Return
            Catch : End Try
        End If
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 Replace(strMessageWork, vbNewLine, ":"))
    End Sub

    ' ----------------------------------------------------------------
    ' Evaluate interface
    '
    '
    Private Sub evaluateInterface()
        With rtbScreen
            Dim strText As String = Trim(.Text)
            Dim intIndex1 As Integer = InStr(strText, vbNewLine)
            Dim strExpression As String = _
                _OBJutilities.item(strText, 1, vbNewLine, False)
            If intIndex1 <> 0 _
               AndAlso _
               intIndex1 <> Len(strText) - Len(vbNewLine) + 1 Then
                strExpression = InputBox("Enter expression", _
                                         "Evaluate expression", _
                                         strExpression)
                If strExpression = "" Then Return
            End If
            Dim booError As Boolean
            displayOutput(OBJqbe.eval(strExpression, booError))
            closeRunForm()
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Find previous selection's color and font
    '
    '
    Private Sub findOldSelectionStyle(ByVal intStartIndex As Integer, _
                                      ByVal intLength As Integer, _
                                      ByRef objColor As Color, _
                                      ByRef objFont As Font)
        With rtbScreen
            objColor = .SelectionColor : objFont = .SelectionFont
        End With
        With CType(CType(rtbScreen.Tag, Collection).Item(6), Collection)
            Dim colEntry As Collection
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                With CType(.Item(intIndex1), Collection)
                    If CInt(.Item(1)) = intStartIndex _
                       AndAlso _
                       CInt(.Item(2)) = intLength Then
                        objColor = CType(.Item(3), Color)
                        objFont = CType(.Item(4), Font)
                        Exit For
                    End If
                End With
            Next intIndex1
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Save form settings in the Registry
    '
    '
    Private Sub form2Registry()
        Try
            With chkCEtestEventLog
                SaveSetting(Application.ProductName, _
                            Me.Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            Dim booReplay As Boolean
            If Not (CTLstatusDisplay Is Nothing) Then
                booReplay = CTLstatusDisplay.Replay
            End If
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "ctlStatusDisplay_Replay", _
                        CStr(booReplay))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "screenStatus", _
                        CStr(screenIsMore()))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "constantFolding", _
                        CStr(FRMoptions.ConstantFolding))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "degenerateOpRemoval", _
                        CStr(FRMoptions.DegenerateOpRemoval))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "assemblyRemovesCode", _
                        CStr(FRMoptions.AssemblyRemovesCode))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "objectTrace", _
                        CStr(FRMoptions.ObjectTrace))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "parseTrace", _
                        CStr(FRMoptions.ParseTrace))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "sourceTrace", _
                        CStr(FRMoptions.SourceTrace))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "inspection", _
                        CStr(FRMoptions.Inspection))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "eventLogging", _
                        CStr(FRMoptions.EventLogging))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "InspectCompilerObjects", _
                        CStr(FRMoptions.InspectCompilerObjects))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "StopButton", _
                        CStr(FRMoptions.StopButton))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "parseDisplay", _
                        CStr(FRMoptions.getParseDisplay.ToString))
            Dim strSourceCode As String
            If Not (OBJqbe Is Nothing) Then
                strSourceCode = OBJqbe.SourceCode
                If strSourceCode = "" AndAlso screenModeSource() Then
                    strSourceCode = rtbScreen.Text
                End If
            End If
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "sourceCode", _
                        strSourceCode)
#If QBGUI_EXTENSIONS Then
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "arrayValueDisplay", _
                        CStr(FRMoptions.ArrayValueDisplay))
            SaveSetting(Application.ProductName, _
                        Me.Name, _
                        "generateNOPs", _
                        CStr(FRMoptions.GenerateNOPs))
#End If
        Catch
            MsgBox("Could not save Registry setting: " & _
                   Err.Number & " " & Err.Description)
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Highlight code
    '
    '
    Private Overloads Sub highlightSourceCode(ByVal intIndex As Integer, _
                                                ByVal intLength As Integer, _
                                                ByVal booChangeColor As Boolean, _
                                                ByVal objColor As Color, _
                                                ByVal booChangeFont As Boolean, _
                                                ByVal objFont As Font)
        Dim objOldColor As Color
        Dim objOldFont As Font
        highlightSourceCode(intIndex, _
                            intLength, _
                            booChangeColor, objColor, _
                            booChangeFont, objFont, _
                            objOldColor, objOldFont)
    End Sub
    Private Overloads Sub highlightSourceCode(ByVal intIndex As Integer, _
                                                ByVal intLength As Integer, _
                                                ByVal booChangeColor As Boolean, _
                                                ByVal objColor As Color, _
                                                ByVal booChangeFont As Boolean, _
                                                ByVal objFont As Font, _
                                                ByRef objOldColor As Color, _
                                                ByRef objOldFont As Font)
        With rtbScreen
            Dim intSelectionStart As Integer = Math.Max(0, intIndex - 1)
            .SelectionStart = intSelectionStart
            .SelectionLength = intLength
            objOldColor = .SelectionColor
            objOldFont = .SelectionFont
            If booChangeFont Then
                .SelectionFont = objFont
            End If
            If booChangeColor Then
                .SelectionColor = objColor
            End If
            .ScrollToCaret()
            .Focus()
            .Refresh()
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Make the More button
    '
    '
    Private Sub mkMoreButton()
        Try
            CMDmore = New Button
            OBJtoolTips.SetToolTip(CMDmore, _
                                    "Shows a more complete view including " & _
                                    "compile, test and inspect information")
            Controls.Add(CMDmore)
            AddHandler CMDmore.Click, AddressOf cmdMore_Click
            With CMDmore
                .Top = cmdClose.Top
                .Width = cmdClose.Width
                .Font = cmdClose.Font
                .Left = cmdClose.Left - .Width - _OBJwindowsUtilities.Grid
                .Height = cmdClose.Height
                .Text = "More"
                .Name = "cmdMore"
            End With
        Catch
            _OBJutilities.errorHandler("Cannot create a More button: " & _
                                       Err.Number & " " & Err.Description)
        End Try
    End Sub

    ' ----------------------------------------------------------------
    ' Change a control for the parse view: record status in collection 
    '
    '
    Private Function mkParse_changeCtl(ByVal colSave As Collection, _
                                       ByVal ctlHandle As Control, _
                                       ByVal intLeft As Integer, _
                                       ByVal intWidth As Integer, _
                                       ByVal booVisible As Boolean) As Boolean
        ' --- Update status collection
        Dim colSubcollection As Collection
        Try
            colSubcollection = New Collection
            With colSubcollection
                .Add(ctlHandle)
                .Add(ctlHandle.Left)
                .Add(ctlHandle.Width)
                .Add(ctlHandle.Visible)
            End With
            colSave.Add(colSubcollection)
        Catch
            errorHandler("Cannot create or populate subcollection, " & _
                         "or add it to save collection " & _
                         Err.Number & " " & Err.Description)
            OBJcollectionUtilities.collectionClear(colSave)
            Return (False)
        End Try
        ' --- Modify the control
        With ctlHandle
            .Left = intLeft
            .Width = intWidth
            .Visible = booVisible
            .Refresh()
        End With
        Return (True)
    End Function

    ' ----------------------------------------------------------------
    ' Create the screen tag
    '
    '
    Private Function mkScreenTag() As Boolean
        With rtbScreen
            Try
                Dim colScreenInfo As New Collection
                With colScreenInfo
                    .Add(True)
                    .Add("")
                    .Add("")
                    .Add(rtbScreen.BackColor)
                    .Add(rtbScreen.ForeColor)
                    Dim colErrorInfo As New Collection
                    .Add(colErrorInfo)
                End With
                .Tag = colScreenInfo
            Catch
                errorHandler("Cannot create screen tag: " & Err.Number & Err.Description)
                End
            End Try
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Make the status display
    '
    '
    Private Function mkStatusDisplay() As Boolean
        Dim intGrid As Integer = _OBJwindowsUtilities.Grid * 3
        Dim intProgressBottom As Integer = _OBJwindowsUtilities.controlBottom(lblProgress)
        CTLstatusDisplay = New statusDisplay(rtbScreen, _
                                                Me, _
                                                0, _
                                                intProgressBottom, _
                                                Width - intGrid, _
                                                ClientSize.Height _
                                                - _
                                                intProgressBottom _
                                                - _
                                                _OBJwindowsUtilities.Grid * 3, _
                                                rtbScreen.BackColor, _
                                                rtbScreen.ForeColor)
        createStdView()
        With CTLstatusDisplay
            .ToolTipObject = OBJtoolTips
            .RPNlistBox.Width -= 2
            AddHandler .displayModifyStorageEvent, _
                       AddressOf statusDisplayModifyStorageEventHandler
        End With
        Return (True)
    End Function

    ' ----------------------------------------------------------------
    ' MSIL run interface
    '
    '
    Private Sub MSILrunInterface()
        Try
            MsgBox("After transliteration to Intermediate Language, this code " & _
                   "produces " & _
                   CStr(OBJqbe.msilRun()))
        Catch
            MsgBox("Not able to translate this code to Intermediate Language: " & _
                   Err.Number & " " & Err.Description)
        End Try
    End Sub

    ' ----------------------------------------------------------------
    ' Pops the XML parse stack, returning the parseStack instance or
    ' Nothing
    '
    '
    ' --- Pop until grammar category is found or the stack is empty
    Private Function popParseStack(ByVal strGC As String) As parseStackEntry
        Dim objNext As parseStackEntry
        Dim strGCwork As String = UCase(Trim(strGC))
        Do While OBJparseStack.Count > 0 _
                 AndAlso _
                 UCase(Trim(CType(OBJparseStack.Peek, parseStackEntry).GrammarCategory)) _
                 <> _
                 strGCwork
            objNext = popParseStack()
        Loop
        Return objNext
    End Function
    ' --- Pop one entry only
    Private Function popParseStack() As parseStackEntry
        If (OBJparseStack Is Nothing) OrElse OBJparseStack.Count = 0 Then
            errorHandler("Internal programming error: parse stack underflow")
            Return Nothing
        End If
        Try
            Dim objPop As parseStackEntry = CType(OBJparseStack.Pop(), parseStackEntry)
            Return objPop
        Catch
            errorHandler("Cannot pop parse stack: " & _
                         Err.Number & " " & Err.Description)
            Return Nothing
        End Try
    End Function

    ' ----------------------------------------------------------------
    ' Extends the XML parse stack
    '
    '
    Private Function pushParseStack(ByVal strGC As String, _
                                    ByVal intIndex As Integer) As Boolean
        If (OBJparseStack Is Nothing) Then
            Try
                OBJparseStack = New Stack
            Catch
                errorHandler("Cannot create parse stack: " & _
                             Err.Number & " " & Err.Description)
                Return False
            End Try
        End If
        Try
            Dim objEntry As New parseStackEntry
            With objEntry
                .GrammarCategory = strGC : .StartIndex = intIndex
            End With
            OBJparseStack.Push(objEntry)
        Catch
            errorHandler("Cannot push XML stack: " & _
                         Err.Number & " " & Err.Description)
            Return False
        End Try
    End Function

    ' ----------------------------------------------------------------
    ' Interface to qbe
    '
    '
    Private Enum ENUqbeInterfaceResult
        failed
        succeeded
        skipped
    End Enum
    Private Overloads Sub qbeInterface(ByVal strGoal As String, _
                                       ByVal booForceOp As Boolean)
        Dim colSave As Collection
        Dim enuResult As ENUqbeInterfaceResult
        Dim strGoalWork As String = UCase(strGoal)
        Dim strReport As String
        Try
            With OBJqbe
                enuResult = ENUqbeInterfaceResult.skipped
                If Not chkViewOutput.Checked _
                   AndAlso _
                   .SourceCode <> rtbScreen.Text _
                   OrElse _
                   .SourceCode <> CStr(screenSourceCode()) Then
                    .SourceCode = rtbScreen.Text
                End If
                Dim intGoal As Integer
                Select Case strGoalWork
                    Case "SCAN" : intGoal = 1
                    Case "COMPILE" : intGoal = 2
                    Case "ASSEMBLE" : intGoal = 3
                    Case Else : intGoal = 4
                End Select
                If booForceOp OrElse Not .scanned Then
                    If Not (CTLstatusDisplay Is Nothing) Then
                        CTLstatusDisplay.ScanListBox.Items.Clear()
                        colSave = createScanView()
                    End If
                    enuResult = CType(IIf(.scan, _
                                          ENUqbeInterfaceResult.succeeded, _
                                          ENUqbeInterfaceResult.failed), _
                                      ENUqbeInterfaceResult)
                    If enuResult = ENUqbeInterfaceResult.succeeded Then
                        If FRMoptions.Inspection Then
                           If Not OBJqbe.inspect(strReport) Then
                               errorHandler("Inspection of quickBasicEngine has failed " & _
                                            "following scan" & _
                                            vbNewLine & vbNewLine & _
                                            strReport)
                               enuResult = ENUqbeInterfaceResult.failed
                           End If
                        End If
                    End If
                    If Not colSave Is Nothing Then
                        restoreStatusDisplay(colSave)
                    End If
                End If
                Dim booRunFormEnabled As Boolean
                If intGoal >= 2 _
                   AndAlso _
                   enuResult <> ENUqbeInterfaceResult.failed _
                   AndAlso _
                   (booForceOp OrElse Not .compiled) Then
                    If Not (FRMrunForm Is Nothing) Then
                        booRunFormEnabled = FRMrunForm.Enabled
                        FRMrunForm.Enabled = False
                    End If
                    If Not (CTLstatusDisplay Is Nothing) Then
                        CTLstatusDisplay.ParseListBox.Items.Clear()
                        CTLstatusDisplay.RPNlistBox.Items.Clear()
                    End If
                    enuResult = ENUqbeInterfaceResult.failed
                    If Not (CTLstatusDisplay Is Nothing) Then
                        colSave = createParseView()
                        If (colSave Is Nothing) Then Return
                    End If
                    enuResult = CType(IIf(.compile, _
                                          ENUqbeInterfaceResult.succeeded, _
                                          ENUqbeInterfaceResult.failed), _
                                      ENUqbeInterfaceResult)
                    If enuResult = ENUqbeInterfaceResult.succeeded _
                       AndAlso _
                       FRMoptions.Inspection Then
                        If Not OBJqbe.inspect(strReport) Then
                            errorHandler("Inspection of quickBasicEngine has failed " & _
                                         "following compile" & _
                                         vbNewLine & vbNewLine & _
                                         strReport)
                            enuResult = ENUqbeInterfaceResult.failed
                        End If
                    End If
                    If Not (colSave Is Nothing) Then
                        restoreStatusDisplay(colSave)
                    End If
                    If Not (FRMrunForm Is Nothing) Then
                        FRMrunForm.Enabled = booRunFormEnabled
                    End If
                End If
                If intGoal >= 3 _
                   AndAlso _
                   enuResult <> ENUqbeInterfaceResult.failed _
                   AndAlso _
                   (booForceOp OrElse Not .assembled) Then
                    enuResult = CType(IIf(.assemble, _
                                          ENUqbeInterfaceResult.succeeded, _
                                          ENUqbeInterfaceResult.failed), _
                                      ENUqbeInterfaceResult)
                    If enuResult = ENUqbeInterfaceResult.succeeded _
                       AndAlso _
                       FRMoptions.Inspection Then
                        If Not OBJqbe.inspect(strReport) Then
                            errorHandler("Inspection of quickBasicEngine has failed " & _
                                         "following assembly" & _
                                         vbNewLine & vbNewLine & _
                                         strReport)
                            enuResult = ENUqbeInterfaceResult.failed
                        End If
                    End If
                End If
                If intGoal >= 4 _
                   AndAlso _
                   enuResult <> ENUqbeInterfaceResult.failed Then
                    If Not (CTLstatusDisplay Is Nothing) Then
                        colSave = createRunView()
                    End If
                    enuResult = CType(IIf(.run, _
                                          ENUqbeInterfaceResult.succeeded, _
                                          ENUqbeInterfaceResult.failed), _
                                      ENUqbeInterfaceResult)
                    If enuResult = ENUqbeInterfaceResult.succeeded _
                       AndAlso _
                       FRMoptions.Inspection Then
                        If Not OBJqbe.inspect(strReport) Then
                            errorHandler("Inspection of quickBasicEngine has failed " & _
                                         "following run" & _
                                         vbNewLine & vbNewLine & _
                                         strReport)
                            enuResult = ENUqbeInterfaceResult.failed
                        End If
                    End If
                    If Not (colSave Is Nothing) Then
                        restoreStatusDisplay(colSave)
                    End If
                End If
                If Not (FRMrunForm Is Nothing) Then closeRunForm()
            End With
        Catch
            With _OBJutilities
                displayOutput _
                (.string2Box(.soft2HardParagraph(editErrorInfo(Err.Number & _
                                                                 " " & _
                                                                 Err.Description), _
                                                 40)))
            End With
        End Try
        If enuResult = ENUqbeInterfaceResult.succeeded _
           AndAlso _
           (strGoalWork = "ASSEMBLE" OrElse strGoalWork = "RUN") _
           AndAlso _
           Not (CTLstatusDisplay Is Nothing) Then
            ' --- Assembled code to list box
            With CTLstatusDisplay
                .RPNlistBox.Items.Clear()
                Dim intIndex1 As Integer
                Dim intCount As Integer
                Try
                    ' If a stop was executed this collection may not exist
                    intCount = OBJqbe.PolishCollection.Count
                Catch : End Try
                For intIndex1 = 1 To intCount
                    .codeUpdate(CType(OBJqbe.PolishCollection.Item(intIndex1), _
                                      qbPolish.qbPolish))
                Next intIndex1
            End With
        End If
        Dim datStart As Date = Now
        MsgBox(_OBJutilities.properCase(strGoal) & " result: " & _
               CStr(IIf(enuResult = ENUqbeInterfaceResult.skipped, _
                        "skipped because source not changed and " & _
                        "step was already performed", _
                        enuResult.ToString)))
        adjustRunTime(datStart)
        If Not (CTLstatusDisplay Is Nothing) Then
            With CTLstatusDisplay.ScanListBox()
                .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
                .Refresh()
            End With
        End If
    End Sub

    ' ----------------------------------------------------------------
    ' Restore form settings from the Registry
    '
    '
    Private Sub registry2Form()
        Try
            With chkCEtestEventLog
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            If CBool(GetSetting(Application.ProductName, _
                                Me.Name, _
                                "ctlStatusDisplay_Replay", _
                                "False")) Then
                showMore()
                CTLstatusDisplay.Replay = True
            End If
            If CBool(GetSetting(Application.ProductName, _
                                Me.Name, _
                                "screenStatus", _
                                "False")) Then
                showMore()
            Else
                showLess()
            End If
            With FRMoptions
#If QBGUI_EXTENSIONS Then
                .ArrayValueDisplay = CBool(GetSetting(Application.ProductName, _
                                                        Me.Name, _
                                                        "ArrayValueDisplay", _
                                                        CStr(.ArrayValueDisplay)))
                .GenerateNOPs = CBool(GetSetting(Application.ProductName, _
                                                        Me.Name, _
                                                        "GenerateNOPs", _
                                                        CStr(OBJqbe.GenerateNOPs)))
#End If
                .AssemblyRemovesCode = CBool(GetSetting(Application.ProductName, _
                                             Me.Name, _
                                             "AssemblyRemovesCode", _
                                             CStr(OBJqbe.AssemblyRemovesCode)))
                .ConstantFolding = CBool(GetSetting(Application.ProductName, _
                                             Me.Name, _
                                             "ConstantFolding", _
                                             CStr(OBJqbe.ConstantFolding)))
                .DegenerateOpRemoval = CBool(GetSetting(Application.ProductName, _
                                             Me.Name, _
                                             "DegenerateOpRemoval", _
                                             CStr(OBJqbe.DegenerateOpRemoval)))
                setQBEoptions()
                .Inspection = CBool(GetSetting(Application.ProductName, _
                                                Me.Name, _
                                                "Inspection", _
                                                CStr(.Inspection)))
                .ObjectTrace = CBool(GetSetting(Application.ProductName, _
                                                Me.Name, _
                                                "ObjectTrace", _
                                                CStr(.ObjectTrace)))
                .ParseTrace = CBool(GetSetting(Application.ProductName, _
                                                Me.Name, _
                                                "ParseTrace", _
                                                CStr(.ParseTrace)))
                .SourceTrace = CBool(GetSetting(Application.ProductName, _
                                                Me.Name, _
                                                "SourceTrace", _
                                                CStr(.SourceTrace)))
                .EventLogging = CBool(GetSetting(Application.ProductName, _
                                                 Me.Name, _
                                                 "EventLogging", _
                                                 CStr(.EventLogging)))
                .InspectCompilerObjects = CBool(GetSetting(Application.ProductName, _
                                                            Me.Name, _
                                                            "InspectCompilerObjects", _
                                                            CStr(.InspectCompilerObjects)))
                .StopButton = CBool(GetSetting(Application.ProductName, _
                                               Me.Name, _
                                               "StopButton", _
                                               CStr(.StopButton)))
                Try
                    .setParseDisplay(GetSetting(Application.ProductName, _
                                                Me.Name, _
                                                "parseDisplay", _
                                                CStr(.getParseDisplay)))
                Catch : End Try
                setQBEoptions()
                splitDebugScreenAdjust()
            End With
            If screenModeSource() Then
                With rtbScreen
                    .Text = GetSetting(Application.ProductName, _
                                        Me.Name, _
                                        "sourceCode", _
                                        .Text)
                End With
            End If
        Catch
            MsgBox("Could not obtain Registry setting: " & _
                   Err.Number & " " & Err.Description)
        End Try
    End Sub

    ' ----------------------------------------------------------------
    ' Registry clear
    '
    '
    Private Function registryClear() As Boolean
        DeleteSetting(Application.ProductName)
        Return (True)
    End Function

    ' ----------------------------------------------------------------
    ' Remove the loop event handler
    '
    '
    Private Sub removeLoopEventHandler()
        RemoveHandler OBJqbe.loopEvent, AddressOf loopEventHandler
    End Sub

    ' ----------------------------------------------------------------
    ' Handle resize
    '
    '
    ' This method caches the geometry of constituent controls for each 
    ' dimension, where a "dimension" consists of width and height and
    ' "geometry" includes Top and Left position in addition to dimension.
    '
    ' The colCache is keyed on _width.height, where width and height are 
    ' control dimensions as set previously. Each cache item is an unkeyed 
    ' sub-Collection.
    '
    ' Each item in the sub-Collection is a sub-sub entry containing 5 items: 
    ' item(1) is the "index path" of the control (see below), item(2) is the 
    ' former control width, item(3) is its former height, item(4) is its 
    ' former leftmost position and item(5) is the former top of the control.
    '
    ' The "index path" of the control is the series of indexes of the control
    ' in the Controls collection of the form and then each container control.
    '
    '
    Private Sub resizeHandler(ByVal intOldWidth As Integer, _
                              ByVal intOldHeight As Integer)
        Static booCache As Boolean = True
        Static colCache As Collection
        If booCache AndAlso (colCache Is Nothing) Then
            Try
                colCache = New Collection
            Catch
                _OBJutilities.errorHandler("Cannot create cache: " & _
                                        Err.Number & " " & Err.Description, _
                                        Me.Name, _
                                        "resizeHandler", _
                                        "Continuing: control dimensions may not " & _
                                        "be correct")
                booCache = False
            End Try
        End If
        If booCache Then
            Dim colCacheEntry As Collection
            If Not (colCache Is Nothing) Then
                Try
                    colCacheEntry = CType(colCache.Item(resizeHandler_mkCacheKey_(Width, Height)), _
                                            Collection)
                Catch : End Try
                If Not (colCacheEntry Is Nothing) Then
                    If resizeHandler_resizeToPreviousValues(colCacheEntry) Then
                        Return
                    End If
                End If
            End If
        End If
        If intOldWidth <> 0 AndAlso intOldHeight <> 0 Then
            _OBJwindowsUtilities.resizeConstituentControls(Me, intOldWidth, intOldHeight, True)
            chkCEtestEventLog.Height = chkCEtestEventLog.Height + 4
        End If
        intOldWidth = Width : intOldHeight = Height
        If booCache Then
            resizeHandler_cacheUpdate(colCache, Width, Height)
        End If
    End Sub

    ' ----------------------------------------------------------------
    ' Update the resize cache (removing old values) on behalf of
    ' resizeHandler
    '
    '
    Private Function resizeHandler_cacheUpdate(ByRef colCache As Collection, _
                                               ByVal intWidth As Integer, _
                                               ByVal intHeight As Integer) _
            As Boolean
        With colCache
            Dim intExtra As Integer = .Count - RESIZEHANDLER_MAXCACHE + 1
            For intExtra = intExtra To 0 Step -1
                Try
                    .Remove(1)
                Catch
                    _OBJutilities.errorHandler("Can't remove cache item: " & _
                                               Err.Number & " " & Err.Description, _
                                               Name, _
                                               "resizeHandler_cacheUpdate", _
                                               "Returning False")
                    Return False
                End Try
            Next intExtra
            Dim colNewEntry As Collection
            Try
                colNewEntry = New Collection
            Catch
                _OBJutilities.errorHandler("Can't create entry collection: " & _
                                            Err.Number & " " & Err.Description, _
                                            Name, _
                                            "resizeHandler_cacheUpdate_mkEntry", _
                                            "Returning False")
                Return False
            End Try
            If Not resizeHandler_cacheUpdate_extendEntry(colNewEntry, Me, "") Then
                Return False
            End If
            Dim strCacheKey As String = resizeHandler_mkCacheKey_(intWidth, intHeight)
            Try
                colCache.Remove(strCacheKey)
            Catch : End Try
            Try
                colCache.Add(colNewEntry, strCacheKey)
            Catch
                _OBJutilities.errorHandler("Can't extend the cache: " & _
                                            Err.Number & " " & Err.Description, _
                                            Name, _
                                            "resizeHandler_cacheUpdate", _
                                            "Returning False")
                Return False
            End Try
            Return True
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Extends one cache entry on behalf of resizeHandler_cacheUpdate,
    ' with information on one container
    '
    '
    Private Function resizeHandler_cacheUpdate_extendEntry(ByRef colNewEntry As Collection, _
                                                           ByVal ctlContainer As Control, _
                                                           ByVal strIndexPath As String) _
            As Boolean
        With ctlContainer
            Dim colNewSubEntry As Collection
            Dim intIndex1 As Integer
            Dim strIndexNextpath As String
            Dim strIndexSubpath As String = strIndexPath & _
                                            CStr(IIf(strIndexPath = "", "", "."))
            For intIndex1 = 0 To .Controls.Count - 1
                With .Controls.Item(intIndex1)
                    Try
                        colNewSubEntry = New Collection
                        strIndexNextpath = strIndexSubpath & CStr(intIndex1)
                        colNewSubEntry.Add(strIndexNextpath)
                        colNewSubEntry.Add(.Width)
                        colNewSubEntry.Add(.Height)
                        colNewSubEntry.Add(.Left)
                        colNewSubEntry.Add(.Top)
                    Catch
                        _OBJutilities.errorHandler("Can't create subentry collection: " & _
                                                    Err.Number & " " & Err.Description, _
                                                    Name, _
                                                    "resizeHandler_cacheUpdate_mkEntry", _
                                                    "Returning False")
                        Return False
                    End Try
                    Try
                        colNewEntry.Add(colNewSubEntry)
                    Catch
                        _OBJutilities.errorHandler("Can't extend entry: " & _
                                                    Err.Number & " " & Err.Description, _
                                                    Name, _
                                                    "resizeHandler_cacheUpdate_mkEntry", _
                                                    "Returning False")
                        Return False
                    End Try
                    If .Controls.Count <> 0 Then
                        If Not resizeHandler_cacheUpdate_extendEntry(colNewEntry, _
                                                                     ctlContainer.Controls.Item(intIndex1), _
                                                                     strIndexNextpath) Then
                            Return False
                        End If
                    End If
                End With
            Next intIndex1
            Return True
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Resolve index path of control on behalf of resize handler
    '
    '
    Private Function resizeHandler_index2Control(ByVal strIndexes As String) _
            As Control
        Dim strSplit() As String
        Try
            strSplit = Split(strIndexes, ".")
        Catch
            _OBJutilities.errorHandler("Cannot split: " & Err.Number & " " & Err.Description, _
                                       Name, "resizeHandler_index2Control", _
                                       "Returning nothing")
            Return Nothing
        End Try
        Dim ctlControl As Control = Me
        Dim intIndex1 As Integer
        For intIndex1 = 0 To UBound(strSplit)
            ctlControl = ctlControl.Controls.Item(CInt(strSplit(intIndex1)))
        Next intIndex1
        Return ctlControl
    End Function

    ' ----------------------------------------------------------------
    ' Make resize cache key (width.height) on behalf of
    ' resizeHandler
    '
    '
    Private Function resizeHandler_mkCacheKey_(ByVal intWidth As Integer, _
                                               ByVal intHeight As Integer) _
            As String
        Return intWidth & "." & intHeight
    End Function

    ' ----------------------------------------------------------------
    ' Resize to previous values on behalf of resizeHandler
    '
    '
    Private Function resizeHandler_resizeToPreviousValues(ByVal colEntry As Collection) _
            As Boolean
        With colEntry
            Dim ctlNext As Control
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                With CType(.Item(intIndex1), Collection)
                    ctlNext = resizeHandler_index2Control(CStr(.Item(1)))
                    If (ctlNext Is Nothing) Then Return False
                    ctlNext.Width = CInt(.Item(2))
                    ctlNext.Height = CInt(.Item(3))
                    ctlNext.Left = CInt(.Item(4))
                    ctlNext.Top = CInt(.Item(5))
                    ctlNext.Refresh()
                End With
            Next intIndex1
            Return True
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Restore the status disply from a colSave collection as created
    ' by mkParseDisplay
    '
    '
    Private Function restoreStatusDisplay(ByVal colSave As Collection) As Boolean
        With colSave
            Dim colHandle As Collection
            Dim ctlHandle As Control
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
                colHandle = CType(.Item(intIndex1), Collection)
                With colHandle
                    ctlHandle = CType(.Item(1), Control)
                    ctlHandle.Left = CType(.Item(2), Integer)
                    ctlHandle.Width = CType(.Item(3), Integer)
                    ctlHandle.Visible = CType(.Item(4), Boolean)
                End With
            Next intIndex1
            CTLstatusDisplay.Refresh()
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Interface to run
    '
    '
    Private Sub run()
        If FRMoptions.StopButton Then showRunForm()
        Dim objThread As Threading.Thread = New Threading.Thread(AddressOf runInterface)
        objThread.Start()
    End Sub

    ' ----------------------------------------------------------------
    ' Interface to run
    '
    '
    Private Sub runInterface()
        Dim datStart As Date = Now
        cmdRun.Tag = 0
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 0, _
                                                 "Starting running a program at " & CStr(datStart), _
                                                 1, _
                                                 strIndent:="    ")
        qbeInterface("RUN", False)
        Dim strAdjustedTime As String
        Dim lngUnadjustedTime As Long = DateDiff("s", datStart, Now)
        Try
            strAdjustedTime = ": this run took about " & _
                              Math.Max(lngUnadjustedTime - CLng(cmdRun.Tag), 0) & " " & _
                              "second(s) NOT including input and output time"
        Catch : End Try
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 -1, _
                                                 "Completed running program at " & CStr(Now) & ": " & _
                                                 "this run took about " & _
                                                 lngUnadjustedTime & " " & _
                                                 "second(s) including input and output time" & _
                                                 strAdjustedTime, _
                                                 strIndent:="    ")
    End Sub

    ' ----------------------------------------------------------------
    ' Return the colors used to highlight scanned code
    '
    '
    Private Function scanColor(ByVal enuType As qbTokenType.qbTokenType.ENUtokenType) As Color
        Select Case enuType
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersand
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeAmpersandSuffix
                Return Color.Black
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeApostrophe
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeColon
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeCurrency
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeExclamation
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier
                Return Color.Black
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeInvalid
                Return Color.Red
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeNewline
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeNull
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeOperator
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeParenthesis
                Return Color.Purple
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypePercent
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypePeriod
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypePound
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypePound
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeSemicolon
                Return Color.Blue
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeString
                Return Color.DarkOrange
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedInteger
                Return Color.DarkOrange
            Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeUnsignedRealNumber
                Return Color.DarkOrange
            Case Else
                Return Color.Blue
        End Select
    End Function

    ' ----------------------------------------------------------------
    ' Interface to scanner
    '
    '
    Private Sub scanInterface(ByVal booForceOp As Boolean)
       qbeInterface("SCAN", booForceOp)
    End Sub

    ' ----------------------------------------------------------------
    ' Return status of screen (True: screen is More: False: screen is
    ' less)
    '
    '
    Private Function screenIsMore() As Boolean
        Select Case CMDmore.Text
            Case MORE_CAPTION : Return (False)
            Case LESS_CAPTION : Return (True)
        End Select
    End Function

    ' ----------------------------------------------------------------
    ' Return the screen mode
    '
    '
    Private Function screenMode() As Boolean
        Return CBool(CType(rtbScreen.Tag, Collection).Item(1))
    End Function

    ' ----------------------------------------------------------------
    ' Return True if the screen mode is "output"
    '
    '
    Private Function screenModeOutput() As Boolean
        Return Not screenMode()
    End Function

    ' ----------------------------------------------------------------
    ' Return True if the screen mode is "source code"
    '
    '
    Private Function screenModeSource() As Boolean
        Return screenMode()
    End Function

    ' ----------------------------------------------------------------
    ' Return screen source code
    '
    '
    Private Function screenSourceCode() As String
        With rtbScreen
            If screenModeSource() Then Return .Text
            Return CStr(CType(.Tag, Collection).Item(2))
        End With
    End Function

    ' ----------------------------------------------------------------
    ' Set Quick Basic Engine options
    '
    '
    Private Sub setQBEoptions()
        With FRMoptions
            OBJqbe.AssemblyRemovesCode = .AssemblyRemovesCode
            OBJqbe.ConstantFolding = .ConstantFolding
            OBJqbe.DegenerateOpRemoval = .DegenerateOpRemoval
            OBJqbe.EventLogging = .EventLogging
            OBJqbe.InspectCompilerObjects = .InspectCompilerObjects
#If QBGUI_EXTENSIONS Then
            OBJqbe.GenerateNOPs = .GenerateNOPs
#End If
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Set the screen mode to output
    '
    '
    Private Sub setScreenModeOutput()
        With CType(rtbScreen.Tag, Collection)
            .Remove(1)
            .Add(False, , 1)
        End With
    End Sub

    ' ----------------------------------------------------------------
    ' Set tool tips
    '
    '
    Private Function setToolTips() As Boolean
        Try
            OBJtoolTips = New ToolTip
            With OBJtoolTips
                .SetToolTip(cmdCEinspector, _
                            "Carries out an internal inspection of the " & _
                            "Quick Basic object")
                .SetToolTip(cmdCEobject2XML, _
                            "Converts the state of the QB object to XML")
                .SetToolTip(cmdCEtest, _
                            "Carries out an internal test of the " & _
                            "Quick Basic object")
                .SetToolTip(cmdClose, _
                            "Saves form selections and closes this form")
                .SetToolTip(cmdEvaluate, _
                            "Evaluates a Quick Basic expression and displays " & _
                            "the result")
                .SetToolTip(cmdRun, _
                            "Runs complete Quick Basic programs with output " & _
                            "to a green/black screen: evaluates expressions and " & _
                            "leaves their result in the stack")
                .SetToolTip(cmdStatusZoom, _
                            "Shows status reports in a scrollable text box")
                .SetToolTip(chkViewOutput, _
                            "Toggles between the source code and the output")
                .SetToolTip(lstStatus, _
                            "Shows status and progress")
                .SetToolTip(txtCE, _
                            "Output from XML, test and inspect functions")
                .SetToolTip(rtbScreen, _
                            "Shows source code alternately with output")
            End With
        Catch

        End Try
    End Function

    ' ----------------------------------------------------------------
    ' Show the About screen
    '
    '
    Private Sub showAbout()
        MsgBox("Edward G. Nilges' simulation of Microsoft Quick Basic" & _
               vbNewLine & vbNewLine & vbNewLine & _
               "This application and form is an interface to the quickBasicEngine class" & _
               vbNewLine & vbNewLine & vbNewLine & _
               OBJqbe.EasterEgg)
    End Sub

    ' ----------------------------------------------------------------
    ' Show simple screen
    '
    '
    Private Sub showLess()
        Dim booSave As Boolean = Visible
        Visible = False
        If Not (CTLstatusDisplay Is Nothing) Then
            CTLstatusDisplay.Visible = False
        End If
        CMDmore.Text = MORE_CAPTION
        OBJtoolTips.SetToolTip(CMDmore, _
                               "Shows extended information about the " & _
                               "scanning, compilation and intepretation " & _
                               "process")
        gbxCE.Visible = False
        ClientSize = New Size(_OBJwindowsUtilities.controlRight(cmdClose) _
                              + _
                              _OBJwindowsUtilities.Grid, _
                              lblProgress.Top _
                              + _
                              lblProgress.Height _
                              + _
                              _OBJwindowsUtilities.Grid)
        Me.CenterToParent()
        With lstStatus
            .Height = lblProgress.Top - .Top
        End With
        Visible = booSave
    End Sub

    ' ----------------------------------------------------------------
    ' Show complex screen
    '
    '
    '      "Less is more" - Mies van der Rohe
    '      "Less is less" - Robert Venturi
    '
    '
    Private Sub showMore()
        Dim booSave As Boolean = Visible
        Visible = False
        ClientSize = New Size(CInt(_OBJwindowsUtilities.screenWidth _
                                    * _
                                    SCREEN_WIDTH_TOLERANCE _
                                    - _
                                    (Size.Width - ClientSize.Width)), _
                              CInt(_OBJwindowsUtilities.screenHeight _
                                    * _
                                    SCREEN_HEIGHT_TOLERANCE _
                                    - _
                                    (Size.Height - ClientSize.Height)))
        If (CTLstatusDisplay Is Nothing) Then
           If Not mkStatusDisplay() Then Return
        End If
        CMDmore.Text = LESS_CAPTION
        OBJtoolTips.SetToolTip(CMDmore, _
                               "Shows a reduced amount of information")
        gbxCE.Visible = True
        CTLstatusDisplay.Visible = True
        CTLstatusDisplay.Refresh()
        Me.CenterToParent()
        With CTLstatusDisplay
            Dim intOldStatusWidth As Integer = .Width
            Dim intOldstatusHeight As Integer = .Height
            .Width = gbxCE.Right
            .GroupBoxHandle.Width = .Width - _OBJwindowsUtilities.Grid
            _OBJwindowsUtilities.resizeConstituentControls(.GroupBoxHandle, _
                                                           intOldStatusWidth, _
                                                           intOldstatusHeight)
            .StorageListBoxLabel.Width = .GroupBoxHandle.Width _
                                         - _
                                         _OBJwindowsUtilities.Grid _
                                         - _
                                         .StorageListBoxLabel.Left
            .StorageZoomButton.Left = .StorageListBoxLabel.Right _
                                      - _
                                      .StorageZoomButton.Width
            .StorageListBox.Width = .StorageListBoxLabel.Width
        End With
        resizeHandler(Math.Max(rtbScreen.Width + gbxCE.Width + _OBJwindowsUtilities.Grid, _
                               CTLstatusDisplay.Width) _
                      + _
                      _OBJwindowsUtilities.Grid * 3, _
                      CTLstatusDisplay.Bottom + _OBJwindowsUtilities.Grid * 7)
        With lstStatus
            .Height = lblProgress.Top - .Top
        End With
        showMore_ensureStatusVisibility()
        Visible = booSave
    End Sub

    ' -----------------------------------------------------------------
    ' Make sure all members of the status control are visible
    '
    '
    Private Sub showMore_ensureStatusVisibility()
        With CTLstatusDisplay.GroupBoxHandle
            Dim intIndex1 As Integer
            For intIndex1 = 0 To .Controls.Count - 1
                .Controls(intIndex1).Visible = True
            Next intIndex1
            CTLstatusDisplay.ReplayCheckbox.Checked = True
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Show the options form
    '
    '
    ' To add an option to this procedure, add a Boolean flag as seen 
    ' below to save and restore the initial value of the option.
    '
    '
    Private Sub showOptions()
        With FRMoptions
#If QBGUI_EXTENSIONS Then
            Dim booArrayValueDisplay As Boolean = .ArrayValueDisplay
            Dim booGenerateNOPs As Boolean = .GenerateNOPs
#End If
            Dim booObjectTrace As Boolean = .ObjectTrace
            Dim booParseTrace As Boolean = .ParseTrace
            Dim booSourceTrace As Boolean = .SourceTrace
            Dim booEventLogging As Boolean = .EventLogging
            Dim booInspection As Boolean = .Inspection
            Dim booInspectCompilerObjects As Boolean = .InspectCompilerObjects
            Dim booStopButton As Boolean = .StopButton
            .ShowDialog()
            If .WasClosed Then
                setQBEoptions()
                splitDebugScreenAdjust()
#If QBGUI_EXTENSIONS Then
                booArrayValueDisplay = .ArrayValueDisplay
                booGenerateNOPs = .GenerateNOPs
#End If
            Else
#If QBGUI_EXTENSIONS Then
                .ArrayValueDisplay = booArrayValueDisplay
                .GenerateNOPs = booGenerateNOPs
#End If
                .AssemblyRemovesCode = OBJqbe.AssemblyRemovesCode
                .ConstantFolding = OBJqbe.ConstantFolding
                .DegenerateOpRemoval = OBJqbe.DegenerateOpRemoval
                .EventLogging = OBJqbe.EventLogging
                .InspectCompilerObjects = OBJqbe.InspectCompilerObjects
                .ObjectTrace = booObjectTrace
                .ParseTrace = booObjectTrace
                .SourceTrace = booSourceTrace
                .Inspection = booInspection
                .StopButton = booStopButton
            End If
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Show the run form
    '
    '
    Private Overloads Sub showRunForm()
        Dim intTopSave As Integer = Top
        Dim intWidthSave As Integer
        If Not (FRMrunForm Is Nothing) Then
            FRMrunForm.dispose()
            FRMrunForm = Nothing
        End If
        ' Show flush with the bottom of the main form
        With _OBJwindowsUtilities
            Dim intScreenHeight As Integer = CInt(.screenHeight * SCREEN_HEIGHT_TOLERANCE)
            Dim intRunFormHeight As Integer = _
                Math.Max(.Grid * 8, intScreenHeight - (Height + .Grid))
            If intRunFormHeight + Height + .Grid > intScreenHeight Then
                Dim intOldHeight As Integer = Height
                Height = intScreenHeight - intRunFormHeight + .Grid
                resizeHandler(Width, intOldHeight)
            End If
            Try
                FRMrunForm = New frmRun(OBJqbe, _
                                        1, _
                                        Width, _
                                        intRunFormHeight, _
                                        Left, _
                                        Bottom)
            Catch
                errorHandler("Cannot recreate re-sized run form: " & _
                                Err.Number & " " & Err.Description)
                Return
            End Try
        End With
        With FRMrunForm
            .Show()
            .BringToFront()
            .Tag = _OBJwindowsUtilities.controlChange(cmdClose, _
                                                      True, _
                                                      False, _
                                                      True)
        End With
        Return
    End Sub

    ' -----------------------------------------------------------------
    ' Load source code
    '
    '
    Private Function sourceCodeLoad() As Boolean
        With ofd
            .ShowDialog()
            Try
                rtbScreen.Text = _OBJutilities.file2String(.FileName)
            Catch : End Try
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Save source code
    '
    '
    Private Function sourceCodeSave() As Boolean
        With sfd
            .ShowDialog()
            _OBJutilities.string2File(rtbScreen.Text, .FileName)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Adjusts the debug screens
    '
    '
    Private Sub splitDebugScreenAdjust()
        If FRMoptions.ObjectTrace _
           AndAlso _
           FRMoptions.SourceTrace _
           AndAlso _
           Not splitDebugScreenExists() Then
            splitDebugScreenMake()
        ElseIf splitDebugScreenExists() Then
            splitDebugScreenDestroy()
        End If
    End Sub

    ' -----------------------------------------------------------------
    ' Tell caller if split screen exists
    '
    '
    Private Function splitDebugScreenExists() As Boolean
        Return (Not (TXTsourceDebug Is Nothing))
    End Function

    ' -----------------------------------------------------------------
    ' Destroy the split screen
    '
    '
    Private Sub splitDebugScreenDestroy()
        disposeSourceDebugScreen()
        txtCE.Height *= 2
    End Sub

    ' -----------------------------------------------------------------
    ' Create split debug screen
    '
    '
    Private Sub splitDebugScreenMake()
        txtCE.Height \= 2
        Try
            TXTsourceDebug = New TextBox
            With TXTsourceDebug
                .Top = txtCE.Bottom
                .Left = txtCE.Left
                .Width = txtCE.Width
                .Height = txtCE.Height
                .BorderStyle = txtCE.BorderStyle
                .BackColor = txtCE.BackColor
                .Font = New Font(txtCE.Font, txtCE.Font.Style)
                .ReadOnly = True
                .Multiline = txtCE.Multiline
                .ScrollBars = txtCE.ScrollBars
                .Refresh()
            End With
            txtCE.Parent.Controls.Add(TXTsourceDebug)
        Catch objException As Exception
            errorHandler("Unable to create extra debug screen: " & _
                         Err.Number & " " & Err.Description & _
                         vbNewLine & vbNewLine & _
                         objException.ToString)
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Return True if a status display is active
    '
    '
    Private Function statusDisplayInEffect() As Boolean
        If (CTLstatusDisplay Is Nothing) Then Return (False)
        Return (CTLstatusDisplay.Visible)
    End Function

    ' -----------------------------------------------------------------
    ' Test interface
    '
    '
    Private Sub testInterface()
        With txtCE
            Dim strReport As String
            .Text = "Testing..."
            .Refresh()
            OBJqbe.test(strReport, chkCEtestEventLog.Checked)
            .Text = strReport
            closeRunForm()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Zoom the status information
    '
    '
    Private Sub statusZoom()
        Dim frmZoom As zoom.zoom
        Try
            frmZoom = New zoom.zoom(lstStatus, CDbl(2), CDbl(3))
            frmZoom.showZoom()
            Return
        Catch : End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Update progress report
    '
    '
    Private Sub updateProgress(ByVal strActivity As String, _
                               ByVal strEntity As String, _
                               ByVal intEntityNumber As Integer, _
                               ByVal intEntityCount As Integer)
        Dim strPercentComplete As String
        If intEntityCount <> 0 Then
            strPercentComplete = ": " & _
                                 CStr(Math.Round(intEntityNumber / intEntityCount * 100, 2)) & _
                                 "% complete"
        End If
        Dim strReport As String = strActivity & " at " & strEntity & " " & _
                                    intEntityNumber & " of " & intEntityCount & _
                                    strPercentComplete
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, strReport)
        With lblProgress
            If intEntityNumber >= intEntityCount Then
                .Visible = False
            Else
                .Width = CInt(_OBJutilities.histogram(intEntityNumber, _
                                                    dblRangeMax:=lstStatus.Width, _
                                                    dblValueMax:=intEntityCount))
                .Visible = True
            End If
            .Refresh()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' View event log
    '
    '
    Private Sub viewEventLog()
        If (OBJqbe.EventLog Is Nothing) Then
            MsgBox("No event log exists") : Return
        End If
        Dim frmEventLogFormat As eventLogFormat
        Try
            frmEventLogFormat = New eventLogFormat
        Catch
            errorHandler("Cannot make event log display: " & _
                         Err.Number & " " & Err.Description)
            Return
        End Try
        removeLoopEventHandler()
        With frmEventLogFormat
            .QuickBasicEngine = OBJqbe
            .ShowDialog()
        End With
        addLoopEventHandler()
    End Sub

#End Region ' General procedures 

#Region " Event handlers "

    ' -------------------------------------------------------------------------
    ' codeChange
    '
    '
    Private Sub codeChangeEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                       ByVal objPolish As qbPolish.qbPolish, _
                                       ByVal intIndex As Integer) _
            Handles OBJqbe.codeChangeEvent
        If (CTLstatusDisplay Is Nothing) Then Return
        Dim intIndex1 As Integer = intIndex - 1
        With CTLstatusDisplay.RPNlistBox
            .Items.RemoveAt(intIndex1)
            .Items.Insert(intIndex1, intIndex1 + 1 & " " & objPolish.ToString)
        End With
    End Sub

    ' -------------------------------------------------------------------------
    ' codeGen
    '
    '
    Private Sub codeGenEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                    ByVal objPolish As qbPolish.qbPolish) Handles OBJqbe.codeGenEvent
        If (CTLstatusDisplay Is Nothing) Then Return
        CTLstatusDisplay.codeUpdate(objPolish)
    End Sub

    ' -------------------------------------------------------------------------
    ' codeHighlight
    '
    '
    Private Sub codeHighlightEventHandler() Handles CTLstatusDisplay.codeHighlightEvent
        chkViewOutput.Checked = False
    End Sub

    ' -------------------------------------------------------------------------
    ' codeRemove
    '
    '
    Private Sub codeRemoveEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                       ByVal intIndex As Integer) Handles OBJqbe.codeRemoveEvent
        If (CTLstatusDisplay Is Nothing) Then Return
        With CTLstatusDisplay
            With .RPNlistBox
                Dim intIndex1 As Integer = intIndex - 1
                .Items.RemoveAt(intIndex1)
                If .Items.Count >= 1 Then
                    .SelectedIndex = intIndex1
                End If
                .Refresh()
            End With
            .Refresh()
        End With
    End Sub

    ' -------------------------------------------------------------------------
    ' compilerError
    '
    '
    Private Sub compilerErrorEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                          ByVal strMessage As String, _
                                          ByVal intIndex As Integer, _
                                          ByVal intContextLength As Integer, _
                                          ByVal intLinenumber As Integer, _
                                          ByVal strHelp As String, _
                                          ByVal strCode As String) Handles OBJqbe.compileErrorEvent
        errorHandler("Error at line " & intLinenumber & _
                     "(character " & intIndex & "): " & _
                     strMessage & _
                     vbNewLine & vbNewLine & _
                     strHelp & _
                     "Code: " & _
                     _OBJutilities.enquote(_OBJutilities.ellipsis(strCode, 64)))
        Dim objOldColor As Color
        Dim objOldFont As Font
        highlightSourceCode(intIndex, _
                            intContextLength, _
                            True, errorColor, _
                            True, New Font(rtbScreen.Font, FontStyle.Strikeout), _
                            objOldColor, objOldFont)
        Try
            Dim objEntry As Collection = New Collection
            With objEntry
                .Add(intIndex)
                .Add(intContextLength)
                .Add(objOldColor)
                .Add(objOldFont)
            End With
            CType(CType(rtbScreen.Tag, Collection).Item(6), Collection).Add(objEntry)
        Catch
            errorHandler("Cannot update error tag: " & _
                         Err.Number & " " & Err.Description)
        End Try
    End Sub

    ' -------------------------------------------------------------------------
    ' compilerError
    '
    '
    Private Sub interpretErrorEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                            ByVal strMessage As String, _
                                            ByVal intIndex As Integer, _
                                            ByVal strHelp As String) _
                Handles OBJqbe.interpretErrorEvent
        Dim strFullMessage As String = "Runtime error at opcode " & intIndex & ": " & _
                                        strMessage & _
                                        vbNewLine & vbNewLine & _
                                        strHelp
        errorHandler(strFullMessage)
    End Sub

    ' -------------------------------------------------------------------------
    ' interpretCls
    '
    '
    Private Sub interpretClsEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine) _
                Handles OBJqbe.interpretClsEvent
        Dim datStart As Date = Now
        displayOutput("", "", booCLS:=True)
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' interpretInput
    '
    '
    Private Sub interpretInputEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                           ByRef strChars As String) _
                Handles OBJqbe.interpretInputEvent
        strChars = InputBox("Interpreter input event", _
                            "Interpreter input event", _
                            "")
    End Sub

    ' -------------------------------------------------------------------------
    ' interpretPrint
    '
    '
    Private Sub interpretPrintEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                           ByVal strOutstring As String) _
                Handles OBJqbe.interpretPrintEvent
        Dim datStart As Date = Now
        displayOutput(strOutstring, "")
        adjustRunTime(datStart)
    End Sub

#If QBGUI_POPCHECK Then
    ' -------------------------------------------------------------------------
    ' interpretTrace (popCheck in effect)
    '
    '
    Private Sub interpretTraceEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                           ByVal intIndex As Integer, _
                                           ByVal objStack As Stack, _
                                           ByVal colStorage As Collection) _
                Handles OBJqbe.interpretTraceEvent
        Dim datStart As Date = Now
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Running code at IP " & intIndex)
        If (CTLstatusDisplay Is Nothing) _
           AndAlso _
           Not FRMoptions.SourceTrace _
           AndAlso _
           Not FRMoptions.ObjectTrace Then
            Return
        End If
        Dim strStack As String = OBJqbe.stack2String(objStack)
        Dim booIncludeArrayValues As Boolean
#If QBGUI_EXTENSIONS Then
        booIncludeArrayValues = FRMoptions.ArrayValueDisplay
#End If
        Dim strStorage As String = OBJqbe.storage2String(colStorage, booIncludeArrayValues)
        If Not (CTLstatusDisplay Is Nothing) Then
            CTLstatusDisplay.interpreterDisplayUpdate(intIndex, _
                                                      strStack, _
                                                      strStorage)
        End If
        If FRMoptions.ObjectTrace Then
            interpretTraceEventHandler_objectTrace(intIndex, strStack, strStorage)
        End If
        If FRMoptions.SourceTrace Then
            Dim txtDisplay As TextBox
            If Not FRMoptions.ObjectTrace Then
                txtDisplay = txtCE
            Else
                txtDisplay = TXTsourceDebug
            End If
            interpretTraceEventHandler_sourceTrace(intIndex, txtDisplay)
        End If
        adjustRunTime(datStart)
    End Sub
#End If

#If Not QBGUI_POPCHECK Then
    ' -------------------------------------------------------------------------
    ' interpretTrace (popCheck not in effect)
    '
    '
    Private Sub interpretTraceEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                           ByVal intIndex As Integer, _
                                           ByVal objStack() As qbVariable.qbVariable, _
                                           ByVal intStackTop As Integer, _
                                           ByVal colStorage As Collection) _
                Handles OBJqbe.interpretTraceFastEvent
        Dim datStart As Date = Now
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Running code at IP " & intIndex)
        If (CTLstatusDisplay Is Nothing) _
           AndAlso _
           Not FRMoptions.SourceTrace _
           AndAlso _
           Not FRMoptions.ObjectTrace Then
            Return
        End If
        Dim strStack As String = OBJqbe.stack2StringFast(objStack, intStackTop)
        Dim booIncludeArrayValues As Boolean
#If QBGUI_EXTENSIONS Then
        booIncludeArrayValues = FRMoptions.ArrayValueDisplay
#End If
        Dim strStorage As String = OBJqbe.storage2String(colStorage, booIncludeArrayValues)
        If Not (CTLstatusDisplay Is Nothing) Then
            CTLstatusDisplay.interpreterDisplayUpdate(intIndex, _
                                                      strStack, _
                                                      strStorage)
        End If
        If FRMoptions.ObjectTrace Then
            interpretTraceEventHandler_objectTrace(intIndex, strStack, strStorage)
        End If
        If FRMoptions.SourceTrace Then
            Dim txtDisplay As TextBox
            If Not FRMoptions.ObjectTrace Then
                txtDisplay = txtCE
            Else
                txtDisplay = TXTsourceDebug
            End If
            interpretTraceEventHandler_sourceTrace(intIndex, txtDisplay)
        End If
        adjustRunTime(datStart)
    End Sub
#End If

    ' -------------------------------------------------------------------------
    ' Produce object code trace on behalf of interpretTraceEventHandler
    '
    '
    Private Sub interpretTraceEventHandler_objectTrace(ByVal intIndex As Integer, _
                                                       ByVal strStack As String, _
                                                       ByVal strStorage As String)
        Dim datStart As Date = Now
        With txtCE
            If strStack <> "" Then
                strStack = _OBJutilities.string2Box(strStack, "Stack")
            End If
            If strStorage <> "" Then
                strStorage = _OBJutilities.string2Box _
                                (_OBJutilities.soft2HardParagraph(_OBJutilities.breakLongWords(strStorage, 50), _
                                                                  bytLineWidth:=50), _
                                "Storage")
            End If
            Dim objPolishHandle As qbPolish.qbPolish = _
                CType(OBJqbe.PolishCollection.Item(intIndex), _
                        qbPolish.qbPolish)
            Dim strDisplay As String
            With objPolishHandle
                strDisplay = _
                    _OBJutilities.soft2HardParagraph _
                        (_OBJutilities.breakLongWords("Opcode: " & objPolishHandle.ToString, 50), _
                         50)
                strDisplay &= _
                    vbNewLine & vbNewLine & _
                    _OBJutilities.string2Box _
                    (_OBJutilities.soft2HardParagraph(.opcodeToDescription, 50))
                If objPolishHandle.TokenStartIndex <= OBJqbe.Scanner.TokenCount Then
                    strDisplay &= _
                        vbNewLine & vbNewLine & _
                        _OBJutilities.string2Box _
                        (OBJqbe.Scanner.Line _
                         (OBJqbe.Scanner.tokenLineNumber _
                          (Math.Max(.TokenStartIndex - 1, 1))), _
                        "Source Code")
                End If
            End With
            If strStorage <> "" Then
                strDisplay &= vbNewLine & vbNewLine & strStorage
            End If
            If strStack <> "" Then
                strDisplay &= vbNewLine & vbNewLine & strStack
            End If
            strDisplay = _OBJutilities.string2Box(strDisplay, _
                                                  "IP: " & intIndex)
            If Len(.Text) > MAX_TRACE Then
                .Text = "This trace is incomplete: older lines have been removed" & _
                        vbNewLine
            End If
            .Text &= vbNewLine & vbNewLine & strDisplay
            .SelectionStart = Len(.Text)
            .SelectionLength = 0
            .ScrollToCaret()
            .Refresh()
        End With
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' Produce source code trace on behalf of interpretTraceEventHandler
    '
    '
    Private Sub interpretTraceEventHandler_sourceTrace(ByVal intIndex As Integer, _
                                                       ByVal txtDisplay As TextBox)
        Dim datStart As Date = Now
        With OBJqbe.Scanner
            txtDisplay.SelectionStart = .LineStartIndex(.tokenLineNumber(intIndex)) - 1
            txtDisplay.SelectionLength = .LineLength(.tokenLineNumber(intIndex))
        End With
        With txtDisplay
            .ScrollToCaret()
            .Focus()
            .Refresh()
            .Parent.Focus()
            .Parent.Refresh()
        End With
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' loopEvent
    '
    '
    ' Note: no compile-time handles clause is included in the loopEvent handler
    ' since we want to replace its normal handler by a new handler when loop
    ' events are handled in subsidiary forms.
    '
    ' See the Load procedure for the handler wiring.
    '
    '
    Private Sub loopEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                 ByVal strActivity As String, _
                                 ByVal strEntity As String, _
                                 ByVal intNumber As Integer, _
                                 ByVal intCount As Integer, _
                                 ByVal intLevel As Integer, _
                                 ByVal strComment As String)
        Dim datStart As Date = Now
        updateProgress(strActivity, strEntity, intNumber, intCount)
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' msgEvent
    '
    '
    Public Sub msgEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                               ByVal strMessage As String) _
               Handles OBJqbe.msgEvent
        Dim datStart As Date = Now
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, strMessage)
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' parseEvent
    '
    '
    Public Sub parseEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                 ByVal strGrammarCategory As String, _
                                 ByVal booTerminal As Boolean, _
                                 ByVal intSrcStartIndex As Integer, _
                                 ByVal intSrcLength As Integer, _
                                 ByVal intTokStartIndex As Integer, _
                                 ByVal intTokLength As Integer, _
                                 ByVal intObjStartIndex As Integer, _
                                 ByVal intObjLength As Integer, _
                                 ByVal strComment As String, _
                                 ByVal intLevel As Integer) _
               Handles OBJqbe.parseEvent
        Dim datStart As Date = Now
        Dim strGrammarCategoryDisplay As String = strGrammarCategory
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "Have parsed grammar category " & _
                                                 strGrammarCategory & ": " & _
                                                 strComment)
        If booTerminal Then
            strGrammarCategoryDisplay = _
                _OBJutilities.enquote(_OBJutilities.string2Display(strGrammarCategoryDisplay))
        End If
        If FRMoptions.ParseTrace Then
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                        -1, _
                                                        "Parsing at token " & intTokStartIndex & ": " & _
                                                        "have parsed " & _
                                                        _OBJutilities.string2Display(strGrammarCategoryDisplay), _
                                                        strIndent:=" ")
        End If
        If UCase(Trim(strGrammarCategory)) = "COMMENT" Then
            highlightSourceCode(intSrcStartIndex, _
                                intSrcLength, _
                                True, Color.Gray, _
                                True, New Font(rtbScreen.Font, FontStyle.Bold))
        Else
            Dim objOldColor As Color
            Dim objOldFont As Font
            highlightSourceCode(intSrcStartIndex, _
                                intSrcLength, _
                                False, rtbScreen.SelectionColor, _
                                True, New Font(rtbScreen.Font, FontStyle.Bold), _
                                objOldColor, objOldFont)
            If objOldColor.A = errorColor().A _
               AndAlso _
               objOldColor.R = errorColor().R _
               AndAlso _
               objOldColor.G = errorColor().G _
               AndAlso _
               objOldColor.B = errorColor().B Then
                ' --- GS was flagged as an error by a previous pass: set its style to the
                ' --- style in effect before the error, and, then, to the gs style
                findOldSelectionStyle(intSrcStartIndex, intSrcLength, _
                                      objOldColor, objOldFont)
                highlightSourceCode(intSrcStartIndex, _
                                    intSrcLength, _
                                    True, objOldColor, _
                                    True, objOldFont)
                highlightSourceCode(intSrcStartIndex, _
                                    intSrcLength, _
                                    False, rtbScreen.SelectionColor, _
                                    True, New Font(rtbScreen.Font, FontStyle.Bold))
            End If
        End If
        If booTerminal Then
            CTLstatusDisplay.ScanListBox.SelectedIndex = intTokStartIndex - 1
        End If
        If FRMoptions.getParseDisplay = options.ENUparseDisplay.XML _
           AndAlso _
           Not (CTLstatusDisplay Is Nothing) Then
            Dim objParseStackEntry As parseStackEntry = popParseStack(strGrammarCategory)
            If Not (objParseStackEntry Is Nothing) Then
                With CTLstatusDisplay.ParseListBox()
                    .Items.Insert(objParseStackEntry.StartIndex, _
                                  _OBJutilities.mkXMLTag(strGrammarCategory))
                    If booTerminal Then
                        .Items.Add(_OBJutilities.xmlMeta2Name(OBJqbe.Scanner.sourceMid(intTokStartIndex, _
                                                                                       intTokLength)))
                    End If
                    .Items.Add(_OBJutilities.mkXMLTag(strGrammarCategory, True))
                    Dim booSave As Boolean = CTLstatusDisplay.ParseListBoxActive
                    CTLstatusDisplay.ParseListBoxActive = False
                    .SelectedIndex = .Items.Count - 1
                    CTLstatusDisplay.ParseListBoxActive = booSave
                    .Refresh()
                End With
            End If
        ElseIf FRMoptions.getParseDisplay = options.ENUparseDisplay.Outline Then
            CTLstatusDisplay.parseUpdate(intLevel, _
                                         strGrammarCategoryDisplay, _
                                         intSrcStartIndex, _
                                         intSrcLength, _
                                         intObjStartIndex, _
                                         intObjLength, _
                                         strComment)
        End If
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' Parse failure event handler
    '
    '
    Private Sub parseFailEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                      ByVal strGC As String) _
            Handles OBJqbe.parseFailEvent
        If FRMoptions.ParseTrace Then
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     -1, _
                                                     "Parse of grammar category " & strGC & " " & _
                                                     "has failed", _
                                                     strIndent:=" ")
        End If
        If (CTLstatusDisplay Is Nothing) Then Return
        popParseStack(strGC)
        Dim objParseStackEntry As parseStackEntry = popParseStack()
        If objParseStackEntry Is Nothing Then Return
        With CTLstatusDisplay.ParseListBox
            Dim intIndex1 As Integer
            For intIndex1 = .Items.Count - 1 To 0 Step -1
                If intIndex1 < objParseStackEntry.StartIndex Then Exit For
                .Items.RemoveAt(intIndex1)
                .Refresh()
            Next intIndex1
        End With
    End Sub

    ' -------------------------------------------------------------------------
    ' Parse start event handler
    '
    '
    Private Sub parseStartEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                       ByVal strGC As String) _
            Handles OBJqbe.parseStartEvent
        If FRMoptions.ParseTrace Then
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     "Parse attempt for the grammar category " & _
                                                     strGC & " " & "is starting", _
                                                     1, _
                                                     strIndent:=" ")

        End If
        If Not (CTLstatusDisplay Is Nothing) Then
            Try
                pushParseStack(strGC, CTLstatusDisplay.ParseListBox.Items.Count)
            Catch
                errorHandler("Cannot extend parse stack: " & _
                            Err.Number & " " & Err.Description)
            End Try
        End If
    End Sub

    ' -------------------------------------------------------------------------
    ' scanEvent
    '
    '
    Public Sub scanEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                ByVal objQBToken As qbToken.qbToken) _
           Handles OBJqbe.scanEvent
        Dim datStart As Date = Now
        With objQBToken
            highlightSourceCode(.StartIndex, _
                                .Length, _
                                True, scanColor(objQBToken.TokenType), _
                                True, New Font(rtbScreen.Font, FontStyle.Regular))
        End With
        With objQBToken
            CTLstatusDisplay.scanUpdate(.StartIndex, _
                                        .Length, _
                                        .TokenType.ToString, _
                                        .LineNumber)
        End With
        adjustRunTime(datStart)
    End Sub

    ' -------------------------------------------------------------------------
    ' Display variable: acquire changes (if any)
    '
    '
    Public Sub statusDisplayModifyStorageEventHandler(ByVal intVariableIndex As Integer)
        Dim frmModify As modifyVariable
        Try
            frmModify = New modifyVariable
        Catch
            errorHandler("Cannot create zoomer form: " & _
                         Err.Number & " " & Err.Description)
        End Try
        With frmModify
            Dim strStorage As String = CStr(CTLstatusDisplay.StorageListBox.Items.Item(intVariableIndex - 1))
            Dim intLocation As Integer = CInt(_OBJutilities.word(strStorage, 1))
            .Variable = OBJqbe.Variable(_OBJutilities.word(strStorage, 2))
            .VariableLocation = intLocation
            .ShowDialog()
            If .Change Then
                OBJqbe.Variable(intVariableIndex) = .Variable
                With CTLstatusDisplay.StorageListBox.Items
                    Dim intListIndex As Integer = intVariableIndex - 1
                    .RemoveAt(intListIndex)
                    .Insert(intListIndex, _
                            CStr(intLocation) & " " & _
                            frmModify.Variable.VariableName & " " & _
                            frmModify.Variable.ToString)
                End With
            End If
            .dispose()
        End With
    End Sub

    ' -------------------------------------------------------------------------
    ' userErrorEvent
    '
    '
    Public Sub userErrorEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                     ByVal strDescription As String, _
                                     ByVal strHelp As String) _
               Handles OBJqbe.userErrorEvent
        errorHandler(strDescription & vbNewLine & vbNewLine & _
                     strHelp)
    End Sub

#End Region

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
    ' *      *  A list box shows the interpretation storage area      *
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
    ' * This control exposes and overrides the following properties,  *
    ' * methods, and events in addition to the base properties,       *
    ' * methods and events of controls.                               *
    ' *                                                               *
    ' *                                                               *
    ' *      *  clear: this method resets the control                 *
    ' *                                                               *
    ' *      *  codeChanged: this method should be executed when the  *
    ' *         code already on display is no longer current. When    *
    ' *         this method is executed, the background and the       *
    ' *         foreground colors of the code display control are     *
    ' *         changed to match the corresponding colors of its      *
    ' *         label.                                                *
    ' *                                                               *
    ' *         Use of the clear method changes the state of the code *
    ' *         display back to normal.                               *
    ' *                                                               *
    ' *      *  codeHighlightEvent: this event is Raised when code is *
    ' *         about to be highlighted in the source control speci-  *
    ' *         fied. It's a reminder that this control is about to be*
    ' *         used.                                                 * 
    ' *                                                               *
    ' *         The control has to be declared WithEvents for this    *
    ' *         event to be sensed.                                   *
    ' *                                                               *
    ' *      *  codeUpdate(op, operand, comment): this method adds a  *
    ' *         line of compiled reverse-Polish code to the end of the*
    ' *         code display. This method requires two parameters.    *
    ' *                                                               *
    ' *         + op: the string form of an RPN operator              *
    ' *         + operand: the string form of an RPN operand          *
    ' *         + comment: comment                                    *
    ' *                                                               *
    ' *         An overload of codeUpdate, codeUpdate(i), will REMOVE *
    ' *         the ith line of code from the display, where i is its *
    ' *         index from 0. Note: this overload is responsible for  *
    ' *         maintaining a map so that clicks to other controls    *
    ' *         highlight the right code.                             *
    ' *                                                               *
    ' *      *  displayModifyStorageEvent(string): this event is      *
    ' *         passed an expression in the string parameter, in this *
    ' *         format:                                               *
    ' *                                                               *
    ' *         <location> <name> <fromString>                        *
    ' *                                                               *
    ' *         The event is expected to display the storage item in  *
    ' *         such a way it can be modified, and to change the      *
    ' *         variable if necessary.                                *
    ' *                                                               *
    ' *      *  dispose: this method gets rid of reference objects    *
    ' *         associated with the statusDisplay; for best results   *
    ' *         use this method when you no longer need this control. *
    ' *                                                               *
    ' *      *  Expression: this write-only property can be set to    *
    ' *         the expression being evaluated.                       *
    ' *                                                               *
    ' *      *  groupBoxHandle: this read-only property returns the   *
    ' *         handle of the group box used to display the parse     *
    ' *         info.                                                 *
    ' *                                                               *
    ' *      *  interpreterDisplayUpdate(ip,stack): this method       *
    ' *         updates the interpreter display.                      *
    ' *                                                               *
    ' *         + ip should be the instruction pointer                *
    ' *                                                               *
    ' *         + stack should be the serialized interpreter stack as *
    ' *           a string                                            *
    ' *                                                               *
    ' *      *  new(source,[parent[,left[,top[,width[,height          *
    ' *         [,backColor,foreColor]]]]]]): the constructor must    *
    ' *         specify the control in the parent containing source   *
    ' *         code. It can specify the parent control and the       *
    ' *         control geometry as the left side, top, width and     *
    ' *         height of the statusDisplay.                          *
    ' *                                                               *
    ' *         Optionally the constructor can also specify a control *
    ' *         color scheme and the background and foreground colors *
    ' *         of labels.                                            *
    ' *                                                               *
    ' *         As shown above, all parameters may be omitted.  The   *
    ' *         parent defaults to Nothing.  The default geometry uses*
    ' *         the size of a default group box and a Location of     *
    ' *         0, 0.                                                 *
    ' *                                                               *
    ' *         To obtain the corresponding default value of any      *
    ' *         geometry parameter, use a value of -1.                *
    ' *                                                               *
    ' *      *  ParseListBox: this read-only property returns a handle*
    ' *         to the list box that displays the parse.              *
    ' *                                                               *
    ' *      *  ParseListBoxActive: this read-write property returns  *
    ' *         and may be set to True to make the parse list box     *
    ' *         respond to changes in its SelectedIndex; note that the*
    ' *         default value of this property is True.               *
    ' *                                                               *
    ' *      *  ParseListBoxLabel: this read-only property returns a  *
    ' *         handle to the label of the list box that displays the *
    ' *         parse.                                                *
    ' *                                                               *
    ' *      *  parseUpdate(srcIndex,srcLength,objIndex,objLength,    *
    ' *         name,level): this method updates the parse display.   *
    ' *         The parse display consists of a visual parse outline: *
    ' *         clicking on this outline highlights text belonging to *
    ' *         grammar symbols as well as the object code produced   *
    ' *         (if any) for grammar symbols.                         *
    ' *                                                               *
    ' *         + srcIndex: the index of the first character of the   *
    ' *           source code that corresponds to the grammar symbol. *
    ' *                                                               *
    ' *         + srcLength: the length of the source code that       *
    ' *           corresponds to the grammar symbol.                  *
    ' *                                                               *
    ' *         + objIndex: the first operation compiled if any for   *
    ' *           the grammar symbol or 0                             *
    ' *                                                               *
    ' *         + objLength: the number of operations compiled for the*
    ' *           grammar symbol                                      *
    ' *                                                               *
    ' *         + name: the name of the parsed grammar symbol         *
    ' *                                                               *
    ' *         + level: the depth of the parsed grammar symbol       *
    ' *                                                               *
    ' *      *  ParseZoomButton: this read-only property returns the  *
    ' *         zoom button associated with the parse tree display    *
    ' *                                                               *
    ' *      *  Replay: this read-write property returns True when the*
    ' *         replay feature was selected, False otherwise. It can  *
    ' *         also change the status of the replay.                 *  
    ' *                                                               *
    ' *      *  ReplayCheckBox: this read-only property returns the   *
    ' *         handle of the Replay check box.                       *
    ' *                                                               *
    ' *      *  RPNlistBox: this read-only property returns a handle  *
    ' *         to the list box that displays the Reverse Polish      *
    ' *         Notation code.                                        *
    ' *                                                               *
    ' *      *  RPNlistBoxLabel: this read-only property returns a    *
    ' *         handle to the label of the list box that displays the *
    ' *         Reverse Polish code.                                  *
    ' *                                                               *
    ' *      *  RPNzoomButton: this read-only property returns the    *
    ' *         zoom button associated with the RPN display           *
    ' *                                                               *
    ' *      *  ScanlistBox: this read-only property returns a handle *
    ' *         to the list box that displays the scan output.        *
    ' *                                                               *
    ' *      *  ScanlistBoxLabel: this read-only property returns a   *
    ' *         handle to the label of the list box that displays the *
    ' *         scan output.                                          *
    ' *                                                               *
    ' *      *  scanUpdate(index,length,type,lineNumber): this method *
    ' *         updates the scan display.  The scan display consists  *
    ' *         of highlights to the expression and a listbox of scan *
    ' *         information.                                          *
    ' *                                                               *
    ' *         + index: the scan index                               *
    ' *                                                               *
    ' *         + length: the length of the token being scanned       *
    ' *                                                               *
    ' *         + type: an enumerator of type ENUscannedType and the  *
    ' *           type of the token being scanned                     *
    ' *                                                               *
    ' *         + lineNumber: the token's line number                 *  
    ' *                                                               *
    ' *      *  ScanZoomButton: this read-only property returns the   *
    ' *         zoom button associated with the scanner display       *
    ' *                                                               *
    ' *      *  StackListBox: this read-only property returns a handle*
    ' *         to the list box that displays the stack.              *
    ' *                                                               *
    ' *      *  StackListBoxLabel: this read-only property returns a  *
    ' *         handle to the label of the list box that displays the *
    ' *         stack.                                                *
    ' *                                                               *
    ' *      *  StackZoomButton: this read-only property returns the  *
    ' *         zoom button associated with the stack display         *
    ' *                                                               *
    ' *      *  StorageListBox: this read-only property returns a     *
    ' *         handle to the list box that displays the stack.       *
    ' *                                                               *
    ' *      *  StorageListBoxLabel: this read-only property returns a*
    ' *         handle to the label of the list box that displays the *
    ' *         storage.                                              *
    ' *                                                               *
    ' *      *  StorageZoomButton: this read-only property returns the*
    ' *         zoom button associated with the storage display       *
    ' *                                                               *
    ' *      *  ToolTipObject: this write-only property should be     *
    ' *         assigned a ToolTip object: setting it will cause tool *
    ' *         tips to be created for each constituent control in    *
    ' *         the status display. For best results set this property*
    ' *         right after creating the status display.              *
    ' *                                                               *
    ' *                                                               *
    ' * C H A N G E   R E C O R D ----------------------------------- *
    ' *   DATE     PROGRAMMER     DESCRIPTION OF CHANGE               *
    ' * --------   ----------     ----------------------------------- *
    ' * 01 27 03   Nilges         Version 1.0                         *
    ' * 11 22 03   Nilges         Added RPNlistBox                    *
    ' * 11 23 03   Nilges         Changed to add line number to scan  *
    ' * 11 23 03   Nilges         Added ScanListBox, ParseListBox,    *
    ' *                           RPNlistBox and StackListBox and     *
    ' *                           associated properties to get ctls   *
    ' * 11 29 03   Nilges         Added storage display               *
    ' * 11 23 03   Nilges         Added ScanZoomButton,               *
    ' *                           ParseZoomButton, RPNzoomButton,     *
    ' *                           StackZoomButton and                 *
    ' *                           StorageZoomButton properties        *
    ' * 03 06 04   Nilges         Added ParseListBoxActive property   *
    ' * 03 17 04   Nilges         Added ReplayCheckBox property       *
    ' * 04 27 04   Nilges         Added popup menu to storage box, for*
    ' *                           zooming contents                    *
    ' *                                                               *
    ' *                                                               *
    ' *****************************************************************
    Private Class statusDisplay
        Inherits Control

        ' ***** Shared *****
        Private Shared _INTsequenceNumber As Integer
        Private Shared _OBJutilities As utilities.utilities
        Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities

        ' ***** Collection tools *****
        Private OBJcollectionUtilities As collectionUtilities.collectionUtilities

        ' ***** Object state *****
        ' --- The state's structure
        Private Structure TYPstate
            Dim booUsable As Boolean                ' True: object is usable
            Dim gbxContainer As GroupBox            ' Container
            Dim rtbSource As RichTextBox            ' Source code text box
            Dim lblParseOutline As Label            ' The parse outline
            Dim lstParseOutline As ListBox
            Dim lblScanned As Label                 ' Scanned code     
            Dim lstScanned As ListBox
            Dim lblRPN As Label                     ' RPN code
            Dim lstRPN As ListBox
            Dim lblStack As Label                   ' Stack
            Dim lstStack As ListBox
            Dim lblStorage As Label                 ' Storage
            Dim lstStorage As ListBox
            Dim chkReplay As CheckBox               ' Instant replay
            Dim cmdReplay As Button                 ' Instant replay
            Dim cmdReset As Button                  ' Instant replay reset
            Dim cmdStep As Button                   ' Instant replay step
            Dim cmdScanZoom As Button               ' Zoom buttons
            Dim cmdParseZoom As Button
            Dim cmdRPNZoom As Button
            Dim cmdStackZoom As Button
            Dim cmdStorageZoom As Button
            Dim booParseListBoxActive As Boolean
            Dim mnuCurrentPopup As ContextMenu
            Dim ctlCurrentPopup As Control
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
        Private WithEvents CMDstorageZoom As Button
        Private WithEvents CHKreplay As CheckBox
        Private WithEvents CMDreplay As Button
        Private WithEvents CMDreset As Button
        Private WithEvents CMDstep As Button

        ' ***** Constants *****
        ' --- The following constants should sum to 1
        Private Const LISTBOX_PROPORTION_SCANNED As Single = 0.25
        Private Const LISTBOX_PROPORTION_PARSE As Single = 0.3
        Private Const LISTBOX_PROPORTION_RPN As Single = 0.17
        Private Const LISTBOX_PROPORTION_STACK As Single = 0.14
        Private Const LISTBOX_PROPORTION_STORAGE As Single = 0.14

        ' ***** Events *****
        Public Event codeHighlightEvent()
        Public Event displayModifyStorageEvent(ByVal intIndex As Integer)

        ' ***** Public procedures **************************************  

        ' --------------------------------------------------------------
        ' Reset the control
        '
        '
        Public Function clear() As Boolean
            If Not checkUsable_("clear") Then Return (False)
            With USRstate
                clearListBoxes_()
                replayErase_()
                changeCodeDisplay_(True)
                Return (True)
            End With
        End Function

        ' --------------------------------------------------------------
        ' Indicate code may no longer correspond to source
        '
        '
        Public Function codeChanged() As Boolean
            If Not checkUsable_("codeChanged") Then Return (False)
            changeCodeDisplay_(False)
            Return (True)
        End Function

        ' --------------------------------------------------------------
        ' Update the RPN code display
        '
        '
        ' --- Insert one line of code at the bottom using a Polish object
        Public Function codeUpdate(ByVal objPolish As qbPolish.qbPolish) As Boolean
            If Not checkUsable_("codeUpdate") Then Return (False)
            With USRstate.lstRPN
                .Items.Add(.Items.Count + 1 & " " & objPolish.ToString)
                .SelectedIndex = .Items.Count - 1
                .Refresh()
            End With
            With objPolish
                If USRstate.chkReplay.Checked _
                   AndAlso _
                   Not replayStore_("codeUpdate", _
                                    .opcodeToString, _
                                    _OBJutilities.object2String(.Operand), _
                                    .Comment) Then
                    Return (False)
                End If
            End With
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
        ' Return group box handle
        '
        '
        Public ReadOnly Property GroupBoxHandle() As GroupBox
            Get
                If Not checkUsable_("GroupBoxHandle Get") Then Return (Nothing)
                Return (USRstate.gbxContainer)
            End Get
        End Property

        ' --------------------------------------------------------------
        ' Interpreter display update
        '
        '
        Public Function interpreterDisplayUpdate(ByVal intIP As Integer, _
                                                 ByVal strStack As String, _
                                                 ByVal strStorage As String) As Boolean
            If Not checkUsable_("interpreterDisplay") Then Return (False)
            With USRstate
                .gbxContainer.Visible = False
                ' --- Check the instruction pointer
                If intIP < 1 OrElse intIP > .lstRPN.Items.Count Then
                    .gbxContainer.Visible = True
                    errorHandler_("Invalid instruction pointer " & intIP, _
                                  "interpreterDisplay", _
                                  "No change is being made to the display")
                    Return (False)
                End If
                ' --- Highlight current op
                If CBool(.lstRPN.Tag) Then
                    .lstRPN.SelectedIndex = intIP - 1
                    .lstRPN.Refresh()
                End If
                ' --- Display the stack
                _OBJwindowsUtilities.string2Listbox(strStack, .lstStack)
                ' --- Display the storage
                _OBJwindowsUtilities.string2Listbox(strStorage, .lstStorage)
                ' --- Save the event information
                If .chkReplay.Checked Then replayStore_("interpreterDisplayUpdate", _
                                                        CStr(intIP), _
                                                        strStack, _
                                                        strStorage)
                ' --- Open the kimono
                .gbxContainer.Visible = True
                .gbxContainer.Refresh()
                Me.Refresh()
                Return (True)
            End With
        End Function

        ' --------------------------------------------------------------
        ' Object constructor
        '
        '
        Public Sub New(ByVal rtbSource As RichTextBox)
            new_(rtbSource, Nothing, -1, -1, -1, -1, Color.Black, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control)
            new_(rtbSource, ctlParent, -1, -1, -1, -1, Color.Black, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control, _
                       ByVal intLeft As Integer)
            new_(rtbSource, ctlParent, intLeft, -1, -1, -1, Color.Black, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer)
            new_(rtbSource, ctlParent, intLeft, intTop, -1, -1, Color.Black, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer)
            new_(rtbSource, ctlParent, intLeft, intTop, intWidth, -1, Color.Black, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer, _
                       ByVal intHeight As Integer)
            new_(rtbSource, ctlParent, intLeft, intTop, intWidth, intHeight, Color.Black, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal objBackColor As Color)
            new_(rtbSource, ctlParent, intLeft, intTop, intWidth, intHeight, objBackColor, Color.OrangeRed)
        End Sub
        Public Sub New(ByVal rtbSource As RichTextBox, _
                       ByVal ctlParent As Control, _
                       ByVal intLeft As Integer, _
                       ByVal intTop As Integer, _
                       ByVal intWidth As Integer, _
                       ByVal intHeight As Integer, _
                       ByVal objBackColor As Color, _
                       ByVal objForeColor As Color)
            new_(rtbSource, ctlParent, intLeft, intTop, intWidth, intHeight, objBackColor, objForeColor)
        End Sub
        Private Sub new_(ByVal rtbSource As RichTextBox, _
                         ByVal ctlParent As Control, _
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
                    ' --- Create collection utilities
                    Try
                        OBJcollectionUtilities = New collectionUtilities.collectionUtilities
                    Catch
                        errorHandler_("Cannot create collection utilities", _
                                      "New", _
                                      "Object won't be usable")
                        Return
                    End Try
                    ' --- Save the text control
                    .rtbSource = rtbSource
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
                    ' --- Create scan display below expression and almost flush with left              
                    Dim intTopForm As Integer = _
                        _OBJwindowsUtilities.Grid * 2
                    Dim intWidthUsable As Integer = _
                        USRstate.gbxContainer.Width _
                        - _
                        _OBJwindowsUtilities.Grid * 2
                    ' Create label   
                    .lblScanned = New Label
                    .gbxContainer.Controls.Add(.lblScanned)
                    With .lblScanned
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.Grid
                        .Top = intTopForm
                        .Width = CInt(intWidthUsable * LISTBOX_PROPORTION_SCANNED)
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Scanned Tokens"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box with events            
                    LSTscanned = New ListBox
                    .lstScanned = LSTscanned
                    .gbxContainer.Controls.Add(.lstScanned)
                    With .lstScanned
                        .Left = USRstate.lblScanned.Left
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblScanned)
                        .Width = USRstate.lblScanned.Width
                        .Height = USRstate.gbxContainer.Height _
                                  - _
                                  new__buttonRowHeight_() _
                                  - _
                                  .Top
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                        AddHandler .SelectedIndexChanged, AddressOf lstScanned_SelectedIndexChanged_
                        AddHandler .Click, AddressOf listbox_Click_
                        AddHandler .MouseUp, AddressOf listbox_MouseUp_
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
                        .Top = intTopForm
                        .Width = CInt(intWidthUsable * LISTBOX_PROPORTION_PARSE)
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Parse Outline"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box with events
                    LSTparseOutline = New ListBox
                    .lstParseOutline = LSTparseOutline
                    .gbxContainer.Controls.Add(.lstParseOutline)
                    With LSTparseOutline
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblParseOutline)
                        .Left = USRstate.lblParseOutline.Left
                        .Width = USRstate.lblParseOutline.Width
                        .Height = USRstate.lstScanned.Height
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                        AddHandler .SelectedIndexChanged, _
                                   AddressOf lstParseOutline_SelectedIndexChanged_
                        AddHandler .Click, AddressOf listbox_Click_
                        AddHandler .MouseUp, AddressOf listbox_MouseUp_
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
                        .Width = CInt(intWidthUsable * LISTBOX_PROPORTION_RPN)
                        .Top = intTopForm
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
                        AddHandler .Click, AddressOf listbox_Click_
                        AddHandler .MouseUp, AddressOf listbox_MouseUp_
                    End With
                    ' Create zoomer
                    .cmdRPNZoom = new__createZoom_(.lblRPN, _
                                                   CMDrpnZoom, _
                                                   .gbxContainer, _
                                                   .lstRPN)
                    ' Make sure it has default "current" status
                    changeCodeDisplay_(True)
                    ' --- Create stack display below expression and to right of object code display    
                    ' Create label
                    .lblStack = New Label
                    .gbxContainer.Controls.Add(.lblStack)
                    With .lblStack
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.lstRPN)
                        .Width = CInt(intWidthUsable * LISTBOX_PROPORTION_STACK)
                        .Top = intTopForm
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
                        AddHandler .Click, AddressOf listbox_Click_
                        AddHandler .MouseUp, AddressOf listbox_MouseUp_
                    End With
                    ' Create zoomer
                    .cmdStackZoom = new__createZoom_(.lblStack, _
                                                     CMDstackZoom, _
                                                     .gbxContainer, _
                                                     .lstStack)
                    ' --- Create storage display below expression and to right of stack display    
                    ' Create label
                    .lblStorage = New Label
                    .gbxContainer.Controls.Add(.lblStorage)
                    With .lblStorage
                        .BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
                        .Left = _OBJwindowsUtilities.controlRight(USRstate.lstStack)
                        .Width = CInt(intWidthUsable * LISTBOX_PROPORTION_STORAGE)
                        .Top = intTopForm
                        .BackColor = objBackColor : .ForeColor = objForeColor
                        .Text = "Storage"
                        .TextAlign = ContentAlignment.MiddleLeft
                        .Font = New Font(.Font, FontStyle.Bold)
                    End With
                    ' Create list box          
                    .lstStorage = New ListBox
                    .gbxContainer.Controls.Add(.lstStorage)
                    With .lstStorage
                        .Top = _OBJwindowsUtilities.controlBottom(USRstate.lblStorage)
                        .Left = USRstate.lblStorage.Left
                        .Width = USRstate.lblStorage.Width
                        .Height = USRstate.lstScanned.Height
                        .BackColor = Color.LightGray
                        .Font = New Font("Courier New", 8, FontStyle.Regular)
                        AddHandler .Click, AddressOf listbox_Click_
                        AddHandler .MouseUp, AddressOf listbox_MouseUp_
                    End With
                    ' Create zoomer
                    .cmdStorageZoom = new__createZoom_(.lblStorage, _
                                                       CMDstorageZoom, _
                                                       .gbxContainer, _
                                                       .lstStorage)
                    ' --- Create replay commands under other controls
                    Dim intBottom As Integer = LSTscanned.Bottom
                    ' Replay check box                                                                 
                    CHKreplay = New CheckBox
                    AddHandler CHKreplay.CheckedChanged, AddressOf chkReplay_CheckedChanged_
                    .chkReplay = CHKreplay
                    .gbxContainer.Controls.Add(.chkReplay)
                    With .chkReplay
                        .Left = USRstate.lblScanned.Left
                        .Top = intBottom
                        .Text = "Replay"
                    End With
                    ' Replay command button                                                                 
                    CMDreplay = New Button
                    CMDreplay.Height = CInt(CMDreplay.Height * 1.1)
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
                    CMDreset.Height = CMDreplay.Height
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
                    CMDstep.Height = CMDreplay.Height
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
                    .booParseListBoxActive = True
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
            chkBox.dispose() : cmdButton.dispose()
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
                    .Width = CInt(.Width * 0.75)
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
        ' Return parse list box
        '
        '
        Public ReadOnly Property ParseListBox() As ListBox
            Get
                If Not checkUsable_("ParseListBox Get") Then Return (Nothing)
                Return USRstate.lstParseOutline
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return and set parse list box activity
        '
        '
        Public Property ParseListBoxActive() As Boolean
            Get
                If Not checkUsable_("ParseListBoxActive Get") Then Return False
                Return USRstate.booParseListBoxActive
            End Get
            Set(ByVal booNewValue As Boolean)
                If Not checkUsable_("ParseListBoxActive Set") Then Return
                USRstate.booParseListBoxActive = booNewValue
            End Set
        End Property

        ' ---------------------------------------------------------------
        ' Return parse list box label
        '
        '
        Public ReadOnly Property ParseListBoxLabel() As Label
            Get
                If Not checkUsable_("ParseListBoxLabel Get") Then Return (Nothing)
                Return USRstate.lblParseOutline
            End Get
        End Property

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
        Public Function parseUpdate(ByVal intLevel As Integer, _
                                    ByVal strName As String, _
                                    ByVal intSrcIndex As Integer, _
                                    ByVal intSrcLength As Integer, _
                                    ByVal intObjIndex As Integer, _
                                    ByVal intObjLength As Integer, _
                                    ByVal strComments As String) As Boolean
            If Not checkUsable_("parseUpdate") Then Return (False)
            If intSrcIndex < 1 _
               OrElse _
               intSrcLength < 0 _
               OrElse _
               intObjIndex < 0 Then
                errorHandler_("Invalid parameters", _
                              "parseUpdate", _
                              "Returning False")
                Return (False)
            End If
            With USRstate.lstParseOutline
                Dim strDisplay As String = parseDisplayMk_(intLevel, _
                                                            strName, _
                                                            intSrcIndex, intSrcLength, _
                                                            intObjIndex, Math.Max(0, intObjLength), _
                                                            strComments)
                Try
                    .Items.Insert(0, strDisplay)
                Catch
                    errorHandler_("Cannot extend the parse outline: " & _
                                  Err.Number & " " & Err.Description, _
                                  "parseUpdate", _
                                  "Making object unusable")
                    USRstate.booUsable = False
                    Return (False)
                End Try
                .Refresh()
                parseDisplay2Selection_(strDisplay)
            End With
            If USRstate.chkReplay.Checked Then replayStore_("parseUpdate", _
                                                            CStr(intLevel), _
                                                            strName, _
                                                            CStr(intSrcIndex), _
                                                            CStr(intSrcLength), _
                                                            CStr(intObjIndex), _
                                                            CStr(intObjLength), _
                                                            strComments)
            Return (True)
        End Function

        ' ---------------------------------------------------------------
        ' Return parse zoom button
        '
        '
        Public ReadOnly Property ParseZoomButton() As Button
            Get
                If Not checkUsable_("ParseZoomButton Get") Then Return (Nothing)
                Return USRstate.cmdParseZoom
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return Replay status
        '
        '
        Public Property Replay() As Boolean
            Get
                If Not checkUsable_("Replay Get") Then Return (False)
                Return (USRstate.chkReplay.Checked)
            End Get
            Set(ByVal booNewValue As Boolean)
                If Not checkUsable_("Replay Get") Then Return
                USRstate.chkReplay.Checked = booNewValue
            End Set
        End Property

        ' ---------------------------------------------------------------
        ' Return Replay check box
        '
        '
        Public ReadOnly Property ReplayCheckbox() As CheckBox
            Get
                If Not checkUsable_("ReplayCheckbox Get") Then Return (Nothing)
                Return (USRstate.chkReplay)
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return RPN list box
        '
        '
        Public ReadOnly Property RPNlistBox() As ListBox
            Get
                If Not checkUsable_("RPNlistBox Get") Then Return (Nothing)
                Return USRstate.lstRPN
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return RPN list box label
        '
        '
        Public ReadOnly Property RPNlistBoxLabel() As Label
            Get
                If Not checkUsable_("RPNlistBoxLabel Get") Then Return (Nothing)
                Return USRstate.lblRPN
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return RPN zoom button
        '
        '
        Public ReadOnly Property RPNzoomButton() As Button
            Get
                If Not checkUsable_("RPNzoomButton Get") Then Return (Nothing)
                Return USRstate.cmdRPNZoom
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return scan list box
        '
        '
        Public ReadOnly Property ScanListBox() As ListBox
            Get
                If Not checkUsable_("ScanListBox Get") Then Return (Nothing)
                Return USRstate.lstScanned
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return scan list box label
        '
        '
        Public ReadOnly Property ScanListBoxLabel() As Label
            Get
                If Not checkUsable_("ScanListBoxLabel Get") Then Return (Nothing)
                Return USRstate.lblScanned
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Update the scan display and highlight scanned token
        '
        '
        ' This method adds a line in the form
        '
        '
        '      tokenType at start-end: value
        '
        '
        Public Function scanUpdate(ByVal intIndex As Integer, _
                                   ByVal intLength As Integer, _
                                   ByVal strTokenType As String, _
                                   ByVal intLineNumber As Integer) As Boolean
            If Not checkUsable_("scanUpdate") Then Return (False)
            If intIndex < 1 OrElse intLength < 0 Then
                errorHandler_("Invalid parameters", _
                              "scanUpdate", _
                              "Returning False")
                Return (False)
            End If
            With USRstate.lstScanned
                Try
                    .Items.Add(strTokenType & " " & _
                               "on line " & intLineNumber & " " & _
                               "at " & _
                               intIndex & " to " & intIndex + intLength - 1)
                Catch
                    errorHandler_("Cannot extend the parse outline: " & _
                                  Err.Number & " " & Err.Description, _
                                  "parseUpdate", _
                                  "Making object unusable")
                    USRstate.booUsable = False
                    Return (False)
                End Try
                .SelectedIndex = .Items.Count - 1
                .Refresh()
            End With
            If USRstate.chkReplay.Checked Then replayStore_("scanUpdate", _
                                                            CStr(intIndex), _
                                                            CStr(intLength), _
                                                            strTokenType, _
                                                            CStr(intLineNumber))
            Return (True)
        End Function

        ' ---------------------------------------------------------------
        ' Return scan zoom button
        '
        '
        Public ReadOnly Property ScanZoomButton() As Button
            Get
                If Not checkUsable_("ScanZoomButton Get") Then Return (Nothing)
                Return USRstate.cmdScanZoom
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return stack list box
        '
        '
        Public ReadOnly Property StackListBox() As ListBox
            Get
                If Not checkUsable_("StackListBox Get") Then Return (Nothing)
                Return USRstate.lstStack
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return stack list box label
        '
        '
        Public ReadOnly Property StackListBoxLabel() As Label
            Get
                If Not checkUsable_("StackListBoxLabel Get") Then Return (Nothing)
                Return USRstate.lblStack
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return stack zoom button
        '
        '
        Public ReadOnly Property StackZoomButton() As Button
            Get
                If Not checkUsable_("StackZoomButton Get") Then Return (Nothing)
                Return USRstate.cmdStackZoom
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return storage list box
        '
        '
        Public ReadOnly Property StorageListBox() As ListBox
            Get
                If Not checkUsable_("StorageListBox Get") Then Return (Nothing)
                Return USRstate.lstStorage
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return storage list box label
        '
        '
        Public ReadOnly Property StorageListBoxLabel() As Label
            Get
                If Not checkUsable_("StorageListBoxLabel Get") Then Return (Nothing)
                Return USRstate.lblStorage
            End Get
        End Property

        ' --------------------------------------------------------------
        ' Update the storage display
        '
        '
        ' --- Insert one cell at the bottom using a Variable object
        Public Overloads Function storageUpdate(ByVal objQBvariable As qbVariable.qbVariable) _
               As Boolean
            With objQBvariable
                If .Dope.isArray Then
                    Return (storageUpdate(objQBvariable.Dope.ToString))
                Else
                    Return (storageUpdate(objQBvariable.ToString))
                End If
            End With
        End Function
        ' --- Insert one cell at the bottom using a string
        Private Overloads Function storageUpdate(ByVal strStorage As String) _
                As Boolean
            If Not checkUsable_("storageUpdate") Then Return (False)
            With USRstate.lstStorage
                .Items.Add(strStorage)
                .Refresh()
            End With
            If USRstate.chkReplay.Checked _
                AndAlso _
                Not replayStore_("storageUpdate", strStorage) Then
                Return (False)
            End If
            Return (True)
        End Function

        ' ---------------------------------------------------------------
        ' Return storage zoom button
        '
        '
        Public ReadOnly Property StorageZoomButton() As Button
            Get
                If Not checkUsable_("StorageZoomButton Get") Then Return (Nothing)
                Return USRstate.cmdStorageZoom
            End Get
        End Property

        ' ---------------------------------------------------------------
        ' Return storage zoom button
        '
        '
        Public WriteOnly Property ToolTipObject() As ToolTip
            Set(ByVal objNewValue As ToolTip)
                If Not checkUsable_("ToolTipObject Set") Then Return
                setToolTips_(objNewValue)
            End Set
        End Property

        ' ***** Private procedures **************************************

        ' ---------------------------------------------------------------
        ' Change the code display
        '
        '
        Private Function changeCodeDisplay_(ByVal booEnable As Boolean) _
                As Boolean
            With USRstate.lstRPN
                Static objBackColor As Color = .BackColor
                Static objForeColor As Color = .ForeColor
                If booEnable Then
                    .BackColor = objBackColor : .ForeColor = objForeColor
                    .Tag = True
                Else
                    .BackColor = USRstate.lblRPN.BackColor
                    .ForeColor = USRstate.lblRPN.ForeColor
                    .Tag = False
                End If
            End With
        End Function

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
            With USRstate
                Dim intStartIndex As Integer
                Dim intLength As Integer
                If parseDisplayDecode_(strDisplay, 2, intStartIndex, intLength) Then
                    With USRstate.rtbSource
                        If intStartIndex > 0 Then
                            If .SelectionStart <> intStartIndex OrElse .SelectionLength <> intLength Then
                                .Select(intStartIndex - 1, intLength)
                                .ScrollToCaret()
                                .Focus()
                                .Refresh()
                            End If
                        End If
                    End With
                    RaiseEvent codeHighlightEvent()
                End If
                Return True
            End With
        End Function

        ' ---------------------------------------------------------------
        ' Obtain source and object locations from parse display
        '
        '
        ' Expects a string in the form 
        '
        '      <dewey> <grammar symbol>: source code from <start> to <end>:
        '      object code from token <start> to <end>: comments
        '
        '
        ' in strParseDisplay.
        '
        '
        Private Function parseDisplayDecode_(ByVal strParseDisplay As String, _
                                                ByVal intIndex As Integer, _
                                                ByRef intStartIndex As Integer, _
                                                ByRef intLength As Integer) As Boolean
            With _OBJutilities
                Try
                    Dim strItems() As String = Split(strParseDisplay, _
                                                     ": source code from ")
                    If UBound(strItems) < 1 Then Return False
                    Dim strWork As String = strItems(1)
                    strItems = Split(strWork, ":")
                    strItems = Split(strItems(0), " ")
                    If UBound(strItems) <> 3 Then Return False
                    If strItems(1) <> "to" Then Return False
                    intStartIndex = CInt(strItems(0))
                    intLength = CInt(strItems(2)) _
                                - _
                                intStartIndex _
                                + _
                                1
                Catch
                    Return False
                End Try
                Return True
            End With
        End Function

        ' ---------------------------------------------------------------
        ' Make parse display
        '
        '
        ' Returns a string in the form 
        '
        '      <dewey> <grammar symbol>: source code from <start> to <end>:
        '      object code from token <start> to <end>: comments
        '
        '
        Private Function parseDisplayMk_(ByVal intLevel As Integer, _
                                         ByVal strGrammarSymbol As String, _
                                         ByVal intSrcStartIndex As Integer, _
                                         ByVal intSrcLength As Integer, _
                                         ByVal intObjStartIndex As Integer, _
                                         ByVal intObjLength As Integer, _
                                         ByVal strComments As String) _
                As String
            Return (_OBJutilities.copies(" ", intLevel - 1) & " " & _
                    strGrammarSymbol & ": " & _
                    "source code from " & _
                    intSrcStartIndex & " to " & intSrcStartIndex + intSrcLength - 1 & ": " & _
                    "object code from " & _
                    intObjStartIndex & " to " & intObjStartIndex + intObjLength - 1 & _
                    CStr(IIf(strComments = "", "", ": ")) & _
                    strComments)
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
                            Dim objPolish As New qbPolish.qbPolish
                            With objPolish
                                .opcodeFromString(strSplit(1))
                                .Operand = strSplit(2)
                                .Comment = strSplit(3)
                                Me.codeUpdate(objPolish)
                                .dispose() : objPolish = Nothing
                            End With
                        Case "INTERPRETERDISPLAYUPDATE"
                            Me.interpreterDisplayUpdate(CInt(strSplit(1)), _
                                                        strSplit(2), _
                                                        strSplit(3))
                        Case "PARSEUPDATE"
                            Me.parseUpdate(CInt(strSplit(1)), _
                                           strSplit(2), _
                                           CInt(strSplit(3)), _
                                           CInt(strSplit(4)), _
                                           CInt(strSplit(5)), _
                                           CInt(strSplit(6)), _
                                           strSplit(5))
                        Case "SCANUPDATE"
                            Me.scanUpdate(CInt(strSplit(1)), _
                                          CInt(strSplit(2)), _
                                          strSplit(3), _
                                          CInt(strSplit(2)))
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
                Return (True)
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
            If Not booValue Then replayErase_()
            INTmacroIndex = 0
            Return (True)
        End Function

        ' -----------------------------------------------------------------
        ' Set tool tips
        '
        '
        Private Sub setToolTips_(ByVal objToolTipObject As ToolTip)
            With USRstate
                objToolTipObject.SetToolTip(.chkReplay, _
                                            "Check this to obtain controls that replay the " & _
                                            "compilation and interpretation process")
                objToolTipObject.SetToolTip(.cmdParseZoom, _
                                            "Click this button to see the parse tree in a " & _
                                            "larger read-only text box")
                objToolTipObject.SetToolTip(.cmdReplay, _
                                            "Click this button to replay all parsing events " & _
                                            "since you turned on replay, or you reset replay: " & _
                                            "see also the Step button")
                objToolTipObject.SetToolTip(.cmdReset, _
                                            "Click this button to reset and to clear the " & _
                                            "storage of events for replay")
                objToolTipObject.SetToolTip(.cmdRPNZoom, _
                                            "Click this button to see the Reverse Polish " & _
                                            "notation object code in a larger read-only text box")
                objToolTipObject.SetToolTip(.cmdScanZoom, _
                                            "Click this button to see the parse tree in a " & _
                                            "larger read-only text box")
                objToolTipObject.SetToolTip(.cmdStackZoom, _
                                            "Click this button to see the stack in a " & _
                                            "larger read-only text box")
                objToolTipObject.SetToolTip(.cmdStep, _
                                            "Click this button to step through the replay " & _
                                            "slowly")
                objToolTipObject.SetToolTip(.cmdStorageZoom, _
                                            "Click this button to see storage in a " & _
                                            "larger read-only text box")
                objToolTipObject.SetToolTip(.lstRPN, _
                                            "Shows the emitted Reverse Polish Notation code")
                objToolTipObject.SetToolTip(.lstParseOutline, _
                                            "Shows the parse tree")
                objToolTipObject.SetToolTip(.lstScanned, _
                                            "Shows the scanned tokens")
                objToolTipObject.SetToolTip(.lstStack, _
                                            "Shows the interpreter's stack containing " & _
                                            "serialized variables")
#If QBGUI_EXTENSIONS Then
                objToolTipObject.SetToolTip(.lstStorage, _
                                            "Shows the interpreter's storage containing " & _
                                            "serialized variables and types: " & _
                                            "note that array values won't be shown " & _
                                            "by default, but may be obtained using the " & _
                                            "options screen")
#Else
                objToolTipObject.SetToolTip(.lstStorage, _
                                            "Shows the interpreter's storage containing " & _
                                            "serialized variables and types: " & _
                                            "note that array values won't be shown")
#End If
            End With
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
                MsgBox("Nothing is available for replay: " & _
                       "use Evaluate or Run first")
                Return
            End If
            Try
                replay_()
            Catch : End Try
        End Sub

        ' -----------------------------------------------------------------
        ' Reset command button clicked
        '
        '
        Private Sub cmdReset_Click_(ByVal objSender As Object, _
                                    ByVal objEventArgs As System.EventArgs)
            INTmacroIndex = 0
            If (COLmacro Is Nothing) Then Return
            INTmacroIndex = Math.Min(1, COLmacro.Count)
            clearListBoxes_()
        End Sub

        ' -----------------------------------------------------------------
        ' Step command button clicked
        '
        '
        Private Sub cmdStep_Click_(ByVal objSender As Object, _
                                   ByVal objEventArgs As System.EventArgs)
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
                ctlZoom = New zoom.zoom(ctlHandle, CDbl(3.75), CDbl(3))
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
        ' Storage list box clicked
        '
        '
        Private Sub listbox_Click_(ByVal objSender As Object, _
                                   ByVal objEventArgs As System.EventArgs)
            With CType(objSender, ListBox)
                .Tag = CStr(.Items(.SelectedIndex))
            End With
        End Sub

        ' -----------------------------------------------------------------
        ' Storage list box clicked
        '
        '
        Private Sub listbox_MouseUp_(ByVal objSender As Object, _
                                     ByVal objEventArgs As System.Windows.Forms.MouseEventArgs)
            If objEventArgs.Button <> MouseButtons.Right Then Exit Sub
            Dim lstBox As ListBox = CType(objSender, ListBox)
            Dim strTag As String = Trim(CStr(lstBox.Tag))
            If strTag = "" Then Return
            Try
                With USRstate
                    If (.mnuCurrentPopup Is Nothing) Then
                        .mnuCurrentPopup = New ContextMenu
                    End If
                    With .mnuCurrentPopup
                        .MenuItems.Clear()
                        .MenuItems.Add(strTag)
                        AddHandler .MenuItems.Item(0).Click, AddressOf mnuPopup_click_
                        .Show(lstBox, New Point(objEventArgs.X, objEventArgs.Y))
                    End With
                    .ctlCurrentPopup = CType(objSender, Control)
                End With
            Catch
                errorHandler_("Cannot create popup menu: " & _
                              Err.Number & " " & Err.Description, _
                              "listbox_MouseUp_", _
                              "Continuing: no popups will occur")
                Return
            End Try
        End Sub

        ' -----------------------------------------------------------------
        ' Parse outline clicked
        '
        '
        Private Sub lstParseOutline_SelectedIndexChanged_(ByVal objSender As Object, _
                                                          ByVal objEventArgs As System.EventArgs)
            If Not Me.ParseListBoxActive Then Return
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
                Dim intLength As Integer
                Dim intStartIndex As Integer
                If parseDisplayDecode_(CStr(.Items(.SelectedIndex)), 1, intStartIndex, intLength) Then
                    USRstate.rtbSource.Select(intStartIndex, intLength)
                End If
            End With
        End Sub

        ' ------------------------------------------------------------------
        ' Respond to popup menu click
        '
        '
        Private Sub mnuPopup_click_(ByVal objSender As Object, _
                                    ByVal objEventArgs As System.EventArgs)
            With USRstate
                Dim lstCurrentPopup As ListBox
                Try
                    lstCurrentPopup = CType(.ctlCurrentPopup, ListBox)
                Catch
                    Return
                End Try
                If (.ctlCurrentPopup Is .lstStorage) Then
                    RaiseEvent displayModifyStorageEvent(lstCurrentPopup.SelectedIndex + 1)
                End If
            End With
        End Sub

    End Class

#End Region

End Class

