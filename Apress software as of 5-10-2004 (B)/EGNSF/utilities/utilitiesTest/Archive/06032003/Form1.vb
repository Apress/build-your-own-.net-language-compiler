Option Strict

' *********************************************************************
' *                                                                   *
' * utilitiesTest                                                     *
' *                                                                   *
' * Tag usage: the Tag of cmdSourceLoad contains its initial caption. *
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
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblUtilities As System.Windows.Forms.Label
    Friend WithEvents lstUtilities As System.Windows.Forms.ListBox
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents cmdStatusZoom As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    Friend WithEvents cmdUtilitiesZoom As System.Windows.Forms.Button
    Friend WithEvents gbxSource As System.Windows.Forms.GroupBox
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents txtSourceOtherFileid As System.Windows.Forms.TextBox
    Friend WithEvents gbxSourceOther As System.Windows.Forms.GroupBox
    Friend WithEvents radSourceOther As System.Windows.Forms.RadioButton
    Friend WithEvents ofdSource As System.Windows.Forms.OpenFileDialog
    Friend WithEvents radSourceWindowsUtilitiesNet As System.Windows.Forms.RadioButton
    Friend WithEvents radSourceUtilitiesNet As System.Windows.Forms.RadioButton
    Friend WithEvents radSourceUtilitiesCOM As System.Windows.Forms.RadioButton
    Friend WithEvents chkSourceLoadOnStartup As System.Windows.Forms.CheckBox
    Friend WithEvents cmdSourceOtherFindFile As System.Windows.Forms.Button
    Friend WithEvents cmdSourceLoad As System.Windows.Forms.Button
    Friend WithEvents cmdUsageAnalysis As System.Windows.Forms.Button
    Friend WithEvents lblProgress2 As System.Windows.Forms.Label
    Friend WithEvents lblProgress3 As System.Windows.Forms.Label
    Friend WithEvents cmdProcedureDump As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents nudStatusDetail As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblStatusDetail As System.Windows.Forms.Label
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.cmdClose = New System.Windows.Forms.Button
Me.lblUtilities = New System.Windows.Forms.Label
Me.lstUtilities = New System.Windows.Forms.ListBox
Me.lstStatus = New System.Windows.Forms.ListBox
Me.lblStatus = New System.Windows.Forms.Label
Me.cmdStatusZoom = New System.Windows.Forms.Button
Me.lblProgress = New System.Windows.Forms.Label
Me.cmdAbout = New System.Windows.Forms.Button
Me.cmdUtilitiesZoom = New System.Windows.Forms.Button
Me.gbxSource = New System.Windows.Forms.GroupBox
Me.cmdSourceLoad = New System.Windows.Forms.Button
Me.chkSourceLoadOnStartup = New System.Windows.Forms.CheckBox
Me.radSourceUtilitiesCOM = New System.Windows.Forms.RadioButton
Me.radSourceOther = New System.Windows.Forms.RadioButton
Me.gbxSourceOther = New System.Windows.Forms.GroupBox
Me.txtSourceOtherFileid = New System.Windows.Forms.TextBox
Me.cmdSourceOtherFindFile = New System.Windows.Forms.Button
Me.radSourceWindowsUtilitiesNet = New System.Windows.Forms.RadioButton
Me.radSourceUtilitiesNet = New System.Windows.Forms.RadioButton
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.ofdSource = New System.Windows.Forms.OpenFileDialog
Me.cmdUsageAnalysis = New System.Windows.Forms.Button
Me.lblProgress2 = New System.Windows.Forms.Label
Me.lblProgress3 = New System.Windows.Forms.Label
Me.cmdProcedureDump = New System.Windows.Forms.Button
Me.nudStatusDetail = New System.Windows.Forms.NumericUpDown
Me.lblStatusDetail = New System.Windows.Forms.Label
Me.cmdTest = New System.Windows.Forms.Button
Me.gbxSource.SuspendLayout()
Me.gbxSourceOther.SuspendLayout()
CType(Me.nudStatusDetail, System.ComponentModel.ISupportInitialize).BeginInit()
Me.SuspendLayout()
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(528, 488)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(96, 88)
Me.cmdClose.TabIndex = 0
Me.cmdClose.Text = "Close"
'
'lblUtilities
'
Me.lblUtilities.BackColor = System.Drawing.Color.Khaki
Me.lblUtilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblUtilities.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblUtilities.Location = New System.Drawing.Point(8, 8)
Me.lblUtilities.Name = "lblUtilities"
Me.lblUtilities.Size = New System.Drawing.Size(616, 20)
Me.lblUtilities.TabIndex = 1
Me.lblUtilities.Text = "Utilities Available: double-click to see information about any utility"
Me.lblUtilities.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lstUtilities
'
Me.lstUtilities.BackColor = System.Drawing.SystemColors.Control
Me.lstUtilities.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstUtilities.ItemHeight = 14
Me.lstUtilities.Location = New System.Drawing.Point(8, 28)
Me.lstUtilities.Name = "lstUtilities"
Me.lstUtilities.Size = New System.Drawing.Size(616, 228)
Me.lstUtilities.TabIndex = 2
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.SystemColors.Control
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 14
Me.lstStatus.Location = New System.Drawing.Point(7, 280)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(616, 130)
Me.lstStatus.TabIndex = 4
'
'lblStatus
'
Me.lblStatus.BackColor = System.Drawing.Color.Khaki
Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatus.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblStatus.Location = New System.Drawing.Point(7, 260)
Me.lblStatus.Name = "lblStatus"
Me.lblStatus.Size = New System.Drawing.Size(401, 20)
Me.lblStatus.TabIndex = 3
Me.lblStatus.Text = "Status"
Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdStatusZoom
'
Me.cmdStatusZoom.Location = New System.Drawing.Point(560, 260)
Me.cmdStatusZoom.Name = "cmdStatusZoom"
Me.cmdStatusZoom.Size = New System.Drawing.Size(64, 20)
Me.cmdStatusZoom.TabIndex = 5
Me.cmdStatusZoom.Text = "Zoom"
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.Khaki
Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblProgress.Location = New System.Drawing.Point(7, 416)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(616, 32)
Me.lblProgress.TabIndex = 6
Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblProgress.Visible = False
'
'cmdAbout
'
Me.cmdAbout.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdAbout.Location = New System.Drawing.Point(8, 488)
Me.cmdAbout.Name = "cmdAbout"
Me.cmdAbout.Size = New System.Drawing.Size(96, 24)
Me.cmdAbout.TabIndex = 7
Me.cmdAbout.Text = "About"
'
'cmdUtilitiesZoom
'
Me.cmdUtilitiesZoom.Location = New System.Drawing.Point(560, 8)
Me.cmdUtilitiesZoom.Name = "cmdUtilitiesZoom"
Me.cmdUtilitiesZoom.Size = New System.Drawing.Size(64, 20)
Me.cmdUtilitiesZoom.TabIndex = 8
Me.cmdUtilitiesZoom.Text = "Zoom"
'
'gbxSource
'
Me.gbxSource.Controls.Add(Me.cmdSourceLoad)
Me.gbxSource.Controls.Add(Me.chkSourceLoadOnStartup)
Me.gbxSource.Controls.Add(Me.radSourceUtilitiesCOM)
Me.gbxSource.Controls.Add(Me.radSourceOther)
Me.gbxSource.Controls.Add(Me.gbxSourceOther)
Me.gbxSource.Controls.Add(Me.radSourceWindowsUtilitiesNet)
Me.gbxSource.Controls.Add(Me.radSourceUtilitiesNet)
Me.gbxSource.Location = New System.Drawing.Point(112, 488)
Me.gbxSource.Name = "gbxSource"
Me.gbxSource.Size = New System.Drawing.Size(408, 152)
Me.gbxSource.TabIndex = 9
Me.gbxSource.TabStop = False
Me.gbxSource.Text = "Source"
'
'cmdSourceLoad
'
Me.cmdSourceLoad.Location = New System.Drawing.Point(184, 112)
Me.cmdSourceLoad.Name = "cmdSourceLoad"
Me.cmdSourceLoad.Size = New System.Drawing.Size(216, 32)
Me.cmdSourceLoad.TabIndex = 11
Me.cmdSourceLoad.Text = "Load %FILEID"
'
'chkSourceLoadOnStartup
'
Me.chkSourceLoadOnStartup.Checked = True
Me.chkSourceLoadOnStartup.CheckState = System.Windows.Forms.CheckState.Checked
Me.chkSourceLoadOnStartup.Location = New System.Drawing.Point(8, 112)
Me.chkSourceLoadOnStartup.Name = "chkSourceLoadOnStartup"
Me.chkSourceLoadOnStartup.Size = New System.Drawing.Size(176, 16)
Me.chkSourceLoadOnStartup.TabIndex = 10
Me.chkSourceLoadOnStartup.Text = "Load source file upon start up"
'
'radSourceUtilitiesCOM
'
Me.radSourceUtilitiesCOM.Location = New System.Drawing.Point(8, 40)
Me.radSourceUtilitiesCOM.Name = "radSourceUtilitiesCOM"
Me.radSourceUtilitiesCOM.Size = New System.Drawing.Size(216, 16)
Me.radSourceUtilitiesCOM.TabIndex = 9
Me.radSourceUtilitiesCOM.Text = "Visual Basic 6 utilities"
'
'radSourceOther
'
Me.radSourceOther.Location = New System.Drawing.Point(240, 40)
Me.radSourceOther.Name = "radSourceOther"
Me.radSourceOther.Size = New System.Drawing.Size(120, 16)
Me.radSourceOther.TabIndex = 8
Me.radSourceOther.Text = "Other"
'
'gbxSourceOther
'
Me.gbxSourceOther.Controls.Add(Me.txtSourceOtherFileid)
Me.gbxSourceOther.Controls.Add(Me.cmdSourceOtherFindFile)
Me.gbxSourceOther.Location = New System.Drawing.Point(8, 56)
Me.gbxSourceOther.Name = "gbxSourceOther"
Me.gbxSourceOther.Size = New System.Drawing.Size(392, 48)
Me.gbxSourceOther.TabIndex = 2
Me.gbxSourceOther.TabStop = False
'
'txtSourceOtherFileid
'
Me.txtSourceOtherFileid.Location = New System.Drawing.Point(88, 16)
Me.txtSourceOtherFileid.Name = "txtSourceOtherFileid"
Me.txtSourceOtherFileid.Size = New System.Drawing.Size(296, 20)
Me.txtSourceOtherFileid.TabIndex = 6
Me.txtSourceOtherFileid.Text = "c:\egnsf\utilities\utilities.VB"
'
'cmdSourceOtherFindFile
'
Me.cmdSourceOtherFindFile.Location = New System.Drawing.Point(8, 16)
Me.cmdSourceOtherFindFile.Name = "cmdSourceOtherFindFile"
Me.cmdSourceOtherFindFile.Size = New System.Drawing.Size(72, 24)
Me.cmdSourceOtherFindFile.TabIndex = 5
Me.cmdSourceOtherFindFile.Text = "Find file..."
'
'radSourceWindowsUtilitiesNet
'
Me.radSourceWindowsUtilitiesNet.Location = New System.Drawing.Point(240, 16)
Me.radSourceWindowsUtilitiesNet.Name = "radSourceWindowsUtilitiesNet"
Me.radSourceWindowsUtilitiesNet.Size = New System.Drawing.Size(144, 16)
Me.radSourceWindowsUtilitiesNet.TabIndex = 1
Me.radSourceWindowsUtilitiesNet.Text = "Windows utilities (.Net)"
'
'radSourceUtilitiesNet
'
Me.radSourceUtilitiesNet.Checked = True
Me.radSourceUtilitiesNet.Location = New System.Drawing.Point(8, 16)
Me.radSourceUtilitiesNet.Name = "radSourceUtilitiesNet"
Me.radSourceUtilitiesNet.Size = New System.Drawing.Size(216, 16)
Me.radSourceUtilitiesNet.TabIndex = 0
Me.radSourceUtilitiesNet.TabStop = True
Me.radSourceUtilitiesNet.Text = "Utilities for Web and Windows (.Net)"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(528, 584)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(96, 24)
Me.cmdForm2Registry.TabIndex = 10
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(528, 616)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(96, 24)
Me.cmdRegistry2Form.TabIndex = 11
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdUsageAnalysis
'
Me.cmdUsageAnalysis.Location = New System.Drawing.Point(8, 520)
Me.cmdUsageAnalysis.Name = "cmdUsageAnalysis"
Me.cmdUsageAnalysis.Size = New System.Drawing.Size(96, 24)
Me.cmdUsageAnalysis.TabIndex = 12
Me.cmdUsageAnalysis.Text = "Usage Analysis"
'
'lblProgress2
'
Me.lblProgress2.BackColor = System.Drawing.Color.Khaki
Me.lblProgress2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblProgress2.Location = New System.Drawing.Point(7, 448)
Me.lblProgress2.Name = "lblProgress2"
Me.lblProgress2.Size = New System.Drawing.Size(616, 16)
Me.lblProgress2.TabIndex = 13
Me.lblProgress2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblProgress2.Visible = False
'
'lblProgress3
'
Me.lblProgress3.BackColor = System.Drawing.Color.Khaki
Me.lblProgress3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblProgress3.Location = New System.Drawing.Point(7, 464)
Me.lblProgress3.Name = "lblProgress3"
Me.lblProgress3.Size = New System.Drawing.Size(616, 16)
Me.lblProgress3.TabIndex = 14
Me.lblProgress3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
Me.lblProgress3.Visible = False
'
'cmdProcedureDump
'
Me.cmdProcedureDump.Location = New System.Drawing.Point(8, 552)
Me.cmdProcedureDump.Name = "cmdProcedureDump"
Me.cmdProcedureDump.Size = New System.Drawing.Size(96, 32)
Me.cmdProcedureDump.TabIndex = 15
Me.cmdProcedureDump.Text = "Procedure Dump"
'
'nudStatusDetail
'
Me.nudStatusDetail.BackColor = System.Drawing.SystemColors.Control
Me.nudStatusDetail.Location = New System.Drawing.Point(512, 260)
Me.nudStatusDetail.Name = "nudStatusDetail"
Me.nudStatusDetail.Size = New System.Drawing.Size(48, 20)
Me.nudStatusDetail.TabIndex = 16
Me.nudStatusDetail.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
Me.nudStatusDetail.Value = New Decimal(New Integer() {5, 0, 0, 0})
'
'lblStatusDetail
'
Me.lblStatusDetail.BackColor = System.Drawing.Color.Khaki
Me.lblStatusDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatusDetail.Font = New System.Drawing.Font("Georgia", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblStatusDetail.Location = New System.Drawing.Point(408, 260)
Me.lblStatusDetail.Name = "lblStatusDetail"
Me.lblStatusDetail.Size = New System.Drawing.Size(104, 20)
Me.lblStatusDetail.TabIndex = 17
Me.lblStatusDetail.Text = "Status Detail"
Me.lblStatusDetail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(8, 616)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.Size = New System.Drawing.Size(96, 24)
Me.cmdTest.TabIndex = 18
Me.cmdTest.Text = "Test"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(630, 651)
Me.ControlBox = False
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.lblStatusDetail)
Me.Controls.Add(Me.nudStatusDetail)
Me.Controls.Add(Me.cmdProcedureDump)
Me.Controls.Add(Me.lblProgress3)
Me.Controls.Add(Me.lblProgress2)
Me.Controls.Add(Me.cmdUsageAnalysis)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.gbxSource)
Me.Controls.Add(Me.cmdUtilitiesZoom)
Me.Controls.Add(Me.cmdAbout)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.cmdStatusZoom)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.lblStatus)
Me.Controls.Add(Me.lstUtilities)
Me.Controls.Add(Me.lblUtilities)
Me.Controls.Add(Me.cmdClose)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "utilitiesTest"
Me.gbxSource.ResumeLayout(False)
Me.gbxSourceOther.ResumeLayout(False)
CType(Me.nudStatusDetail, System.ComponentModel.ISupportInitialize).EndInit()
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    ' ***** Constants *****
    ' --- About info  
    Private Const ABOUTINFO As String = _
    "utilitiesTest" & vbNewline & vbNewline & vbNewline & _
    "This form and application tests the utilities and the windowsUtilities libraries." & vbnewline & vbnewline & _
    "The utilities library is a collection of more than fifty general-purpose methods " & vbnewline & _
    "for Web based and Windows parsing, collection management, simple IO and other " & vbNewline & _    
    "common tasks." & vbNewline & vbNewline & _
    "The windowsUtilities library is a smaller collection of tools for classic Windows " & vbnewline & _
    "applications." & vbNewline & vbNewline & _
    "This application was developed by: " & vbnewline & vbNewline & _
    "Edward G. Nilges" & vbNewLine & _
    "Sarita Flats" & vbNewLine & _
    "39 Gordon St." & vbNewLine & _
    "PO Box 16334" & vbNewLine & _
    "Suva, Fiji" & vbNewLine & _
    "Intl dialing: 679 330 0084" & vbNewLine & _
    "spinoza1111@yahoo.COM" & vbNewLine & _
    "http://members.screenz.com/edNilges"
    ' --- File identifiers
    Private Const UTILITIESFILEIDNET As String = "..\..\..\utilities\utilities.VB"
    Private Const UTILITIESFILEIDCOM As String = "..\..\..\vb6\utilities\utilities.CLS"
    Private Const WINDOWSUTILITIESFILEIDNET As String = "..\..\..\windowsUtilities\windowsUtilities.VB"

    ' ***** Shared utilities *****
    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities
    Private Shared _OBJutilitiesCOM As prjUtilities.clsUtilities

    ' ***** Procedure information *****
    ' --- Procedure type
    Private Enum ENUprocedureType
        subroutineProcedure
        functionProcedure
        propertyProcedure                
        invalidProcedure
    End Enum		
    ' --- Procedure information
    Private Structure TYPoverload
        Dim intStartIndex As Integer        ' Start index from one
        Dim intLength As Integer            ' Length
        Dim enuType As ENUprocedureType     ' Sub, func, property or not valid
        Dim colParameters As Collection     ' Parameters: no key
                                            ' Item(1): return type
                                            ' Items 2..n: subcollection
                                            '      Item(1): parameter name
                                            '      Item(2): True (ByVal) or False (ByRef)
                                            '      Item(3): dimensions or 0 for a scalar
                                            '      Item(4): type
    End Structure
    Private Structure TYPprocedure
        Dim strName As String		        ' Function, subroutine or property name
        Dim strText As String               ' Procedure's text
        Dim usrOverload() As TYPoverload    ' One or more noncontiguous overloads
        Dim strAbstract As String           ' From comment header
        Dim colUses As Collection           ' List of procedures used by this procedure
        Dim colUsedBy As Collection           ' List of procedures using this procedure
    End Structure
    Private Structure TYPindexedProcedures
        Dim booNet As Boolean               ' True: .Net: False: VB-6
        Dim colIndex As Collection          ' Index: key is name: data is index in table
        Dim usrProcedures() As TYPprocedure 
    End Structure    
    Private USRutilitiesProcedures As TYPindexedProcedures

#End Region ' " Form data "

#Region " Form events "

    Private Sub cmdAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbout.Click
        showAbout
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer
        End
    End Sub
    
    Private Sub cmdProcedureDump_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProcedureDump.Click
        Msgbox(dumpProcedures(USRutilitiesProcedures))
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form
    End Sub

    Private Sub cmdSourceLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceLoad.Click
        getProcedures(sourceFileid, USRutilitiesProcedures)
    End Sub

    Private Sub cmdSourceOtherLoadFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSourceOtherFindFile.Click
        findFile       
        adjustSourceDisplay 
    End Sub

    Private Sub cmdStatusZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStatusZoom.Click
        zoomInterface(lstStatus)
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click
        testInterface
    End Sub

    Private Sub cmdUsageAnalysis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUsageAnalysis.Click
        usageAnalysis(USRutilitiesProcedures)
    End Sub

    Private Sub cmdUtilitiesZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUtilitiesZoom.Click
        zoomInterface(lstUtilities)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ofdSource.InitialDirectory = Application.StartupPath
        Try
            If Not CBool(GetSetting(Application.ProductName, Name, "firstUse", "False")) Then
                showAbout
                SaveSetting(Application.ProductName, Name, "firstUse", "True")
            End If            
        Catch: End Try        
        Show
        Opacity = .5
        Refresh
        updateStatusListBox("Loading", 1)
        registry2Form
        Opacity = .75
        Refresh
        If chkSourceLoadOnStartup.Checked _
           AndAlso _
           Not getProcedures(sourceFileid, USRutilitiesProcedures) Then closer
        Opacity = 1
        Refresh
        Try
            _OBJutilitiesCOM = New prjUtilities.clsUtilities
        Catch ex As Exception
            errorHandler("Can't create COM utilities: " & Err.Number & " " & Err.Description, _
                         "form1_Load", _
                         "Won't be able to test")
        End Try        
        updateStatusListBox(-1, "Load complete")
    End Sub

    Private Sub lstUtilities_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstUtilities.DoubleClick
        displayInfoForm(_OBJutilities.item(CStr(lstUtilities.Items(lstUtilities.SelectedIndex)), _
                                           1, _
                                           ":", _
                                           False), _
                        USRutilitiesProcedures) 
    End Sub

    Private Sub radSourceOther_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radSourceOther.Click
        adjustSourceDisplay
    End Sub

    Private Sub radSourceUtilitiesCOM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radSourceUtilitiesCOM.Click
        adjustSourceDisplay
    End Sub

    Private Sub radSourceUtilitiesNet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radSourceUtilitiesNet.Click
        adjustSourceDisplay
    End Sub

    Private Sub radSourceWindowsUtilitiesNet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radSourceWindowsUtilitiesNet.Click
        adjustSourceDisplay
    End Sub

#End Region ' " Form events "

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Get one-line abstract
    '
    '
    Private Function abstract2Oneline(ByVal strAbstract As String) As String
        Dim strSplit() As String  
        Try
            strSplit = split(strAbstract, vbNewline)
        Catch ex As Exception
            errorHandler("Can't split string: " & Err.Number & " " & Err.Description, _
                         "", _
                         "Returning a null string" & vbNewline & vbNewline & _
                         ex.ToString)
            Return("")                         
        End Try       
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim strNext As String
        For intIndex1 = 0 To UBound(strSplit)
            intCount = 0
            For intIndex2 = 1 To Len(strSplit(intIndex1))
                strNext = Mid(strSplit(intIndex1), intIndex2, 1)
                If strNext >= "A" AndAlso strNext <= "Z" _
                   OrElse _
                   strNext >= "a" AndAlso strNext <= "z" Then
                    intCount += 1
                End If                    
            Next intIndex2   
            If intCount >= Len(strSplit(intIndex1)) * .75 Then 
                Return(strSplit(intIndex1))
            End If                     
        Next intIndex1         
    End Function    
    
    ' -----------------------------------------------------------------
    ' Adjust source display to selections
    '
    '
    Private Sub adjustSourceDisplay
        txtSourceOtherFileid.Text = sourceFileid
        gbxSourceOther.Enabled = radSourceOther.Checked
        With cmdSourceLoad
            If (.Tag Is Nothing) Then .Tag = .Text
            .Text = Replace(CStr(.Tag), "%FILEID", sourceFileid)
        End With        
        cmdTest.Enabled = Not (radSourceOther.Checked)
    End Sub    

    ' -----------------------------------------------------------------
    ' Initialize the procedure table
    '
    '
    Private Function allocateProcinfo(ByRef usrProcedures As TYPindexedProcedures) As Boolean
        With usrProcedures
            Try
                Redim .usrProcedures(0)
                .colIndex = New Collection
            Catch ex As Exception
                errorHandler("Cannot allocate procedure array or index: " & _
                            Err.Number & " " & Err.Description, _
                            "allocateProcinfo", _
                            ex.toString)
                Return(False)
            End Try
            Return(True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Close the application and return to Windows
    '
    '
    Private Sub closer 
        form2Registry 
        destroyProcInfo(USRutilitiesProcedures)
        Dispose
        End
    End Sub

    ' -----------------------------------------------------------------
    ' Decomment source code
    '
    '
    Private Function decomment(ByVal strSource As String) As String 
        Dim objRE As System.Text.RegularExpressions.Regex
        Try
            objRE = New System.Text.RegularExpressions.Regex _
                    (_OBJutilities.commonRegularExpressions("vbCommentLine"))
        Catch ex As Exception
            errorHandler("Cannot create regular expression: " & _
                         Err.Number & " " & Err.Description, _
                         "decomment", _
                         "Returning unchanged source code")  
            Return(strSource)                                       
        End Try        
        Dim intIndex1 As Integer
        Dim objMatch As System.Text.RegularExpressions.Match
        Dim strOutstring As String
        Do While intIndex1 < Len(strSource)
            Try
                objMatch = objRE.Match(strSource, intIndex1)
            Catch ex As Exception
                errorHandler("Cannot apply regular expression: " & _
                             Err.Number & " " & Err.Description, _
                             "decomment", _
                             "Returning unchanged source code")   
                Return(strSource)                                          
            End Try        
            With objMatch
                If .Length = 0 Then 
                    strOutstring &= Mid(strSource, intIndex1 + 1)
                    Exit Do
                End If        
                strOutstring &= Mid(strSource, intIndex1 + 1, .Index - intIndex1) 
                If .Index >= Len(vbNewline) _
                   AndAlso _
                   Mid(strSource, .Index + 1 - Len(vbNewline), Len(vbNewline)) _
                   <> _
                   vbNewline Then
                    strOutstring &= vbNewline
                End If                     
                intIndex1 = .Index + .Length
            End With            
        Loop       
        Return(strOutstring)
    End Function
    
    ' -----------------------------------------------------------------
    ' Dispose of the procedure info
    '
    '
    Private Sub destroyProcInfo(ByRef usrProcedureInfo As TYPindexedProcedures)
        With usrProcedureInfo
            If Not procInfoExists(usrProcedureInfo) Then Return
            _OBJutilities.collectionClear(.colIndex): .colIndex = Nothing
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            For intIndex1 = 1 To UBound(.usrProcedures)
                With .usrProcedures(intIndex1)
                    .colUsedBy = Nothing: .colUses = Nothing
                    For intIndex2 = 1 To UBound(.usrOverload)
                        With .usrOverload(intIndex2)
                            .colParameters = Nothing 
                        End With                        
                    Next intIndex2                 
                End With              
            Next intIndex1            
        End With        
    End Sub    

    ' -----------------------------------------------------------------
    ' Determine procedure type
    '
    '
    Private Function determineProcedureType(ByVal strHeader As String) As ENUprocedureType
        If Instr(strHeader & " ", "SUB ", CompareMethod.Text) <> 0 Then
            Return(ENUprocedureType.subroutineProcedure)
        ElseIf Instr(strHeader & " ", "FUNCTION ", CompareMethod.Text) <> 0 Then
            Return(ENUprocedureType.functionProcedure)
        ElseIf Instr(strHeader & " ", "PROPERTY ", CompareMethod.Text) <> 0 Then
            Return(ENUprocedureType.propertyProcedure)
        End If
        errorHandler("Cannot determine procedure type in " & _
                     _OBJutilities.enquote(_OBJutilities.string2Display(strHeader)), _
                     "determineProcedureType", _
                     "Returning invalid")
        Return(ENUprocedureType.invalidProcedure)
    End Function
    
    ' -----------------------------------------------------------------
    ' Display the test form
    '
    '
    Private Sub displayInfoForm(ByVal strUtility As String, _
                                ByVal usrIndexedProcedures As TYPindexedProcedures)
        Dim intIndex1 As Integer
        Try
            intIndex1 = CInt(usrIndexedProcedures.colIndex(strUtility))
        Catch ex As Exception
            errorHandler("Can't create utilities info form: " & _
                        Err.Number & " " & Err.Description, _
                        "displayInfoForm", _
                        "Returning to caller" & vbNewline & vbNewline & _
                        ex.ToString)
            Return                         
        End Try     
        zoomInterface(usrIndexedProcedures.usrProcedures(intIndex1).strAbstract & _
                            vbNewline & vbNewline & _
                            displayInfoForm_usageAnalysis _
                            (usrIndexedProcedures.usrProcedures(intIndex1)) & _
                            vbNewline & vbNewline & _
                            "S O U R C E   C O D E" & _
                            vbNewline & vbNewline & _
                            usrIndexedProcedures.usrProcedures(intIndex1).strText, _
                        dblHeight:=1)
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Format the usage analysis on behalf of displayInfoForm
    '
    '
    Private Function displayInfoForm_usageAnalysis _
            (ByVal usrProcedure As TYPprocedure) As String
        Dim strUsageAnalysis As String = "Used/uses analysis is not available: click the Usage " & _
                                         "Analysis button on the main form to obtain this information"            
        With usrProcedure
            If (.colUsedBy Is Nothing) AndAlso (.colUses Is Nothing) Then
                Return(strUsageAnalysis)
            ElseIf (Not .colUsedBy Is Nothing) AndAlso (Not .colUses Is Nothing) Then
                Dim intIndex1 As Integer
                strUsageAnalysis = ""
                Dim strUsed As String                                         
                With .colUsedBy
                    For intIndex1 = 1 To .Count
                        _OBJutilities.append(strUsed, ",", CStr(.Item(intIndex1)))
                    Next intIndex1
                End With                
                Dim strUses As String                                         
                With .colUses
                    For intIndex1 = 1 To .Count
                         _OBJutilities.append(strUses, ",", CStr(.Item(intIndex1)))
                    Next intIndex1
                End With      
                strUsageAnalysis &= _
                    CStr(Iif(strUses <> "", _
                             "This procedure uses the following procedures: " & strUses, _
                             "This procedure does not use any other procedure")) & _
                    vbNewline & vbNewline & _          
                    CStr(Iif(strUsed <> "", _
                             "This procedure is used by the following procedures: " & strUsed, _
                             "This procedure isn't used by any other procedure"))  
            Else
                errorHandler("Internal programming error: usage analysis state", _
                             "", _
                             "Returning not available message")
                Return(strUsageAnalysis)
            End If            
            Return(strUsageAnalysis)
        End With        
    End Function    
    
    ' -----------------------------------------------------------------
    ' Dump the indexed procedure table
    '
    '
    Private Function dumpProcedures(ByVal usrIndexedProcedures As TYPindexedProcedures) _
            As String
        Dim strOutstring As String = "No procedure information exists"
        If procInfoExists(usrIndexedProcedures) Then            
            With usrIndexedProcedures
                strOutstring = "Index: " & _OBJutilities.collection2String(.colIndex) & _
                               vbNewline & vbNewline & vbNewline
                Dim intIndex1 As Integer
                For intIndex1 = 1 To UBound(.usrProcedures)
                    With .usrProcedures(intIndex1)
                        _OBJutilities.append(strOutstring, _
                                             vbNewline & vbNewline, _
                                             .strName & ": " & abstract2OneLine(.strAbstract) & vbNewline & _
                                             "Uses: " & _OBJutilities.collection2String(.colUses) & vbNewline & _
                                             "Used by: " & _OBJutilities.collection2String(.colUsedBy))
                    End With                    
                Next intIndex1            
            End With      
        End If            
        Return(strOutstring)      
    End Function            

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
        If Not booInfo Then updateStatusListBox
        updateStatusListBox(_OBJutilities.string2Box(_OBJutilities.soft2HardParagraph(strFullMessage, _
                                                                                      72), _
                                                     CStr(Iif(booInfo, "I N F O", "E R R O R"))), _
                            booDateStamp:=False)
    End Sub

    ' -----------------------------------------------------------------
    ' Expand the procedure table
    '
    '
    Private Function expandProctable(ByRef usrIndexedProcedures As TYPindexedProcedures, _
                                     ByVal strName As String, _
                                     ByVal strText As String, _
                                     ByVal intStartIndex As Integer, _
                                     ByVal intLength As Integer, _
                                     ByVal enuType As ENUprocedureType, _
                                     ByVal strAbstract As String, _
                                     ByVal colParameters As Collection) As Boolean
        With usrIndexedProcedures
            Dim intIndex1 As Integer
            Try
                intIndex1 = CInt(.colIndex(strName))
            Catch: End Try            
            If intIndex1 = 0 Then
                ' New procedure
                Try
                    Redim Preserve .usrProcedures(UBound(.usrProcedures) + 1)
                    With .usrProcedures(UBound(.usrProcedures))
                        Redim .usrOverload(0)
                    End With
                Catch ex As Exception
                    errorHandler("Cannot expand procedure table: " & _
                                Err.Number & " " & Err.Description, _
                                "expandProctable", _
                                "Returning False" & _
                                vbNewline & vbNewline & _
                                ex.toString)
                    Return(False)
                End Try
                With .usrProcedures(UBound(.usrProcedures))
                    .strName = strName
                    .strAbstract = strAbstract
                    .strText = decomment(strText)
                End With
                intIndex1 = UBound(.usrProcedures)
                Try
                    .colIndex.Add(intIndex1, strName)
                Catch ex As Exception
                    errorHandler("Cannot expand procedure table's index: " & _
                                Err.Number & " " & Err.Description, _
                                "expandProctable", _
                                "Returning False" & _
                                vbNewline & vbNewline & _
                                ex.toString)
                    Return(False)
                End Try                
            Else
                .usrProcedures(intIndex1).strText &= vbNewline & strText                
            End If
            With .usrProcedures(intIndex1)
                Try
                    Redim Preserve .usrOverload(UBound(.usrOverload) + 1)
                Catch ex As Exception
                    errorHandler("Cannot expand procedure overload table: " & _
                                Err.Number & " " & Err.Description, _
                                "expandProctable_overload_", _
                                "Returning False" & _
                                vbNewline & vbNewline & _
                                ex.toString)
                    Return(False)
                End Try     
                With .usrOverload(UBound(.usrOverload))   
                    .intStartIndex = intStartIndex
                    .intLength = intLength
                    .enuType = .enuType
                    .colParameters = colParameters
                End With        
            End With                                
            Return(True)
        End With                                            
    End Function
    
    ' -----------------------------------------------------------------
    ' Find file with dialog
    '
    '
    Private Function findFile As String
        With ofdSource
            .InitialDirectory = getPath(txtSourceOtherFileid.Text)
            .ShowDialog
            If .FileName = "" Then Return("")
            txtSourceOtherFileid.Text = .FileName
            Return(.FileName)
        End With        
    End Function 
       
    ' -----------------------------------------------------------------
    ' Form to registry
    '
    '
    Private Sub form2Registry
        Try
            With radSourceUtilitiesNet
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With            
            With radSourceUtilitiesCOM
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With            
            With radSourceWindowsUtilitiesNet
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With            
            With radSourceOther
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With            
            With txtSourceOtherFileid
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            .Text)
            End With            
            With chkSourceLoadOnStartup
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With            
            With nudStatusDetail
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Value))
            End With            
        Catch ex As Exception
            errorHandler("Cannot save settings: " & Err.Number & " " & Err.Description, _
                         "form2Registry", _
                         "Continuing" & vbNewline & vbNewline & _
                         ex.ToString)
        End Try        
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Get directory from file id
    '
    '
    Private Function getFileDirectory(ByVal strFileid As String) As String
        Dim strFileDirectory As String
        Dim strTitle As String
        parseFileid(strFileid, strFileDirectory, strTitle)
        Return(strFileDirectory)
    End Function

    ' -----------------------------------------------------------------
    ' Get name from file id
    '
    '
    Private Function getFileName(ByVal strFileid As String) As String
        Dim strDirectory As String
        Dim strTitle As String
        parseFileid(strFileid, strDirectory, strTitle)
        Dim intIndex1 As Integer = InstrRev(strTitle, ".")
        If intIndex1 = 0 Then Return(strTitle)
        Return(Mid(strTitle, 1, intIndex1 - 1)) 
    End Function

    ' -----------------------------------------------------------------
    ' Get title from file id
    '
    '
    Private Function getFileTitle(ByVal strFileid As String) As String
        Dim strDirectory As String
        Dim strTitle As String
        parseFileid(strFileid, strDirectory, strTitle)
        Return(strTitle)
    End Function
    
    ' -----------------------------------------------------------------
    ' Extract the path
    '
    '
    Private Function getPath(ByVal strFileid As String) As String
        Dim intIndex1 As Integer = InstrRev(strFileid, "\")
        Dim intIndex2 As Integer = InstrRev(strFileid, ":")
        If intIndex1 > intIndex2 Then
            Return(Mid(strFileid, 1, intIndex1 - 1))
        ElseIf intIndex2 > intIndex1 Then
            Return(Mid(strFileid, 1, intIndex2))
        Else
            Return("")                        
        End If        
    End Function     

    ' -----------------------------------------------------------------
    ' Get all procedures  
    '
    '
    Private Function getProcedures(ByVal strFileId As String, _
                                   ByRef usrProcedureInfo As TYPindexedProcedures) As Boolean  
        Dim booExists As Boolean
        Dim strDir As String
        Dim strInfile As String = ""
        Try
            strDir = CurDir
            ChDir(Application.StartupPath)
            booExists = _OBJutilities.fileExists(strFileid) 
            If booExists Then 
                strInfile = getProcedures_normalizeNewline(_OBJutilities.file2String(strFileId))
            End If                
            ChDir(strDir)
        Catch: End Try
        destroyProcInfo(usrProcedureInfo) 
        If Not allocateProcinfo(usrProcedureInfo) Then Return(False)
        If Not (getProcedures_file2Procedures(strInfile, _
                                                usrProcedureInfo, _
                                                getFileName(strFileId))) Then
            Return(False)
        End If
        getProcedures_procedures2Listbox(usrProcedureInfo.usrProcedures, lstUtilities) 
        Return(True)
    End Function

    ' -----------------------------------------------------------------
    ' Get procedures from file on behalf of getProcedures
    '
    '
    Private Function getProcedures_file2Procedures(ByVal strInfile As String, _
                                                   ByRef usrIndexedProcedures As TYPindexedProcedures, _
                                                   ByVal strFileName As String) As Boolean
        Dim strActivity As String = "Obtaining procedures from file"
        updateStatusListBox(strActivity, 1)
        With usrIndexedProcedures
            .booNet = True
            Select Case _OBJutilities.vbCodeType(strInfile)
                Case "VBNET":
                Case "VB6": .booNet = False
                Case "":
                    MsgBox("Cannot determine VB code type, assuming .Net")
                Case Else:
                    errorHandler("Internal programming error: " & _
                                "unexpected return value from vbCodeType", _
                                "getProcedures_file2Procedures", _
                                "Assuming VB Net code")
            End Select    
        End With    
        Dim objRegex(1) As System.Text.RegularExpressions.Regex
        Try
            objRegex(0) = New System.Text.RegularExpressions.Regex _
                              (_OBJutilities.commonRegularExpressions("vbCommentBlock") & _ 
                               getProcedures_removeScopes _
                               (_OBJutilities.commonRegularExpressions _
                                (CStr(Iif(usrIndexedProcedures.booNet, _
                                          "vbProcedureHeaderNET", _
                                          "vbProcedureHeaderCOM"))))) 
        Catch ex As Exception
            errorHandler("Cannot create regular expression: " & Err.Number & " " & Err.Description, _
                         "instrument_getFiles_getProcedures", _
                         ex.ToString)
            Return(False)
        End Try
        Dim booFoundEnd As Boolean
        Dim colParameters As Collection
        Dim enuType As ENUprocedureType
        Dim intParameterLength As Integer
        Dim intIndex1 As Integer = 1
        Dim intStartIndex As Integer
        Dim intEndIndex As Integer
        Dim objMatch As System.Text.RegularExpressions.Match
        Dim strName As String
        Dim strText As String
        Dim strType As String
        Dim strWork As String
        Do While intIndex1 <= Len(strInfile)
            updateProgress(lblProgress, _
                           strActivity, _
                           "character", _
                           intIndex1, _
                           Len(strInfile))
            Try
                objMatch = objRegex(0).Match(strInfile, _
                                             getProcedures_findProbableStart(strInfile, _
                                                                             intIndex1) - 1)
            Catch ex As Exception
                errorHandler("Regex match for start of procedure failed: " & _
                             Err.Number & " " & Err.Description, _
                             "instrument_getFiles_getProcedures", _
                             "Terminating search" & vbNewline & vbNewline & ex.ToString)
                Exit Do
            End Try
            With objMatch
                If .Length <> 0 Then
                    intStartIndex = .Index + 1
                    strName = _OBJutilities.word(.Value, _OBJutilities.words(.Value))
                Else
                    intIndex1 = Len(strInfile) + 1: Exit Do
                End If 
                enuType = determineProcedureType(.Value)
                booFoundEnd = True
                strType = procedureType2Keyword(enuType)
                Try
                    objRegex(1) = Nothing
                    objRegex(1) = New System.Text.RegularExpressions.Regex _
                                      (_OBJutilities.commonRegularExpressions("newlineWebWindows") & _
                                       "[ ]*End " & strType)
                    objMatch = objRegex(1).Match(strInfile, intStartIndex) 
                Catch ex As Exception
                    errorHandler("Regex match for end of procedure failed: " & _
                                 Err.Number & " " & Err.Description, _
                                 "instrument_getFiles_getProcedures", _
                                 "Ending procedure at eof" & vbNewline & vbNewline & ex.ToString)
                    booFoundEnd = False                                 
                End Try       
                With objMatch
                    booFoundEnd = booFoundEnd AndAlso .Length <> 0
                    If booFoundEnd Then 
                        intEndIndex = .Index + .Length
                    Else
                        intEndIndex = Len(strInfile) + 1
                    End If                        
                    strText = Mid(strInfile, intStartIndex, intEndIndex - intStartIndex + 1)
                End With  
                colParameters = getProcedures_parseParameters(Mid(strInfile, .Index + 1 + .Length, 8192), _
                                                              intParameterLength)                      
                If Not expandProctable(usrIndexedProcedures, _ 
                                       strName, _
                                       strText, _
                                       intStartIndex, _
                                       Len(strText), _
                                       enuType, _
                                       getProcedures_parseAbstract(Mid(strInfile, .Index + 1, 8192), _
                                                                   usrIndexedProcedures.booNet), _ 
                                       colParameters) Then 
                    Exit Do
                End If
                updateStatusListBox("Found " & strType & " " & strName)
                intIndex1 = intEndIndex + 1
            End With
        Loop
        updateStatusListBox(-1, strActivity & " complete")
        lblProgress.Visible = False
        Return(intIndex1 > Len(strInfile))
    End Function
    
    ' -----------------------------------------------------------------
    ' Find probable start of procedure based on structure of the actual
    ' utilities file
    '
    '
    ' We locate the comment bar (consisting of apostrophe and dashes
    ' as above) that precedes the string "\n    Public". This assists
    ' the regular expression scan in finding the next procedure quickly.
    '
    ' intIndex is returned without change if the probable start cannot be
    ' located.
    '
    ' Note that this method will skip over Private and Friend procedures as 
    ' intended.
    '
    '
    Private Function getProcedures_findProbableStart(ByVal strInfile As String, _
                                                     ByVal intIndex As Integer) _
                     As Integer
        Dim intIndex1 As Integer = Instr(intIndex, strInfile, vbNewline & "    Public") 
        If intIndex1 = 0 Then Return(intIndex)
        Dim intIndex2 As Integer = InstrRev(strInfile, vbNewline & "    ' ----------", intIndex1)
        If intIndex2 > intIndex Then Return(intIndex2)
        Return(intIndex1) 
    End Function                     

    ' -----------------------------------------------------------------
    ' Normalize the newline on behalf of getProcedures
    '
    '
    Private Function getProcedures_normalizeNewline(ByVal strInfile As String) As String
        updateStatusListBox("Normalizing the new line string", 1)
        Dim strNewline As String = _OBJutilities.determineNewline(strInfile, intMaxLength:=1024)
        If strNewline = "" Then
            errorHandler("Input file does not appear to have a consistent newline", _
                         "getProcedures_normalizeNewline", _
                         "Using the standard value")
            Return(strInfile)                         
        End If        
        updateStatusListBox(-1, "The newline string is " & _OBJutilities.string2Display(strNewline))
        Return(Replace(strInfile, strNewline, vbNewline))
    End Function
        
    ' -----------------------------------------------------------------
    ' Parse the abstract on behalf of getProcedures
    '
    '
    Private Function getProcedures_parseAbstract(ByVal strSource As String, _
                                                 ByVal booNet As Boolean) As String
        ' --- Get the abstract
        Dim intIndex1 As Integer  
        Dim strSplit() As String 
        Try
            strSplit = Split(Mid(strSource, intIndex1 + Len(vbNewline)), vbNewline)
        Catch ex As Exception
            errorHandler("Cannot split string: " & Err.Number & " " & Err.Description, _
                         "getProcedures_parseAbstract", _
                         "Returning a null abstract")
            Return("")
        End Try
        Dim strOutstring As String
        Dim strText As String
        For intIndex1 = 0 To UBound(strSplit)
            If Not getProcedures_parseComment(strSplit(intIndex1), strText) Then 
                If booNet OrElse intIndex1 > 1 Then Exit For
            End If                
            _OBJutilities.append(strOutstring, vbNewline, strText)
        Next intIndex1
        Return(strOutstring)
    End Function

    ' -----------------------------------------------------------------
    ' Parse comment on behalf of getProcedures
    '
    '
    Private Function getProcedures_parseComment(ByVal strSourceLine As String, _
                                                ByRef strText As String) As Boolean
        Dim strTextTrim As String = Trim(strSourceLine)
        If Mid(strTextTrim, 1, 1) <> "'" Then Return(False)
        strText = Mid(strTextTrim, 2) 
        Return(True)
    End Function
        
    ' -----------------------------------------------------------------
    ' Parse the parameters and the As clause on behalf of getProcedures
    '
    '
    ' This method returns a new collection. Item(1) will be Nothing or
    ' the As clause's type (Object when the variable type is not 
    ' specified). Items(2)..n will be subcollections:
    '
    '
    '      *  Item(1) of each subcollection will be the variable name
    '
    '      *  Item(2) will be True (item is passed ByVal) or False 
    '         (item is passed ByRef)
    '
    '      *  Item(3) will be its number of dimensions (0 for nonarrays)
    '
    '      *  Item(4) will be its variable type (Object when the variable type
    '         is not specified, even for subroutines)
    '
    '
    Private Function getProcedures_parseParameters(ByVal strSource As String, _
                                                   ByRef intParameterLength As Integer) As Collection
        Dim colNew As Collection
        Try
            colNew = New Collection
        Catch ex As Exception
            errorHandler("Can't create collection: " & Err.Number & " " & Err.Description, _
                         "", _
                         "Returning Nothing")
            Return(Nothing)                         
        End Try        
        ' --- Assemble possibly continued statement
        Dim intIndex1 As Integer = 1
        Dim intIndex2 As Integer
        Dim intIndex3 As Integer
        Dim strParameterList As String
        Do While intIndex1 <= Len(strSource)
            intIndex2 = Instr(intIndex1, strSource & " _" & vbNewline, " _" & vbNewline)
            intIndex3 = Instr(intIndex1, strSource & vbNewline, vbNewline)
            strParameterList &= Mid(strSource, intIndex1, Math.Min(intIndex2, intIndex3) - intIndex1)
            If intIndex2 > intIndex3 Then Exit Do
            intIndex1 = intIndex2 + Len(" _" & vbNewline)
        Loop     
        intParameterLength = Len(strParameterList)   
        ' --- Parse the parameters
        Dim booByVal As Boolean
        Dim colSubentry as Collection
        Dim intDimensions As Integer
        Dim objRE As New System.Text.RegularExpressions.Regex _
            (_OBJutilities.commonRegularExpressions("vbParameterDefNet"))
        Dim objMatch As System.Text.RegularExpressions.Match  
        Dim strName As String
        Dim strSplit() As String
        Dim strType As String
        intIndex1 = 1
        Do While intIndex1 <= Len(strSource)
            Try
                objMatch = objRE.Match(strParameterList, intIndex1)
            Catch ex As Exception
                errorHandler("Can't apply regular expression: " & _
                             Err.Number & " " & Err.Description, _
                             "getProcedures_parseParameters", _
                             "Returning Nothing")
                Return(Nothing)                         
            End Try   
            With objMatch
                If .Length = 0 Then Exit Do
                Try
                    colSubentry = New Collection
                Catch ex As Exception
                    errorHandler("Can't create subcollection: " & _
                                Err.Number & " " & Err.Description, _
                                "getProcedures_parseParameters", _
                                "Returning Nothing")
                    Return(Nothing)                         
                End Try                
                booByVal = True
                intIndex2 = 1
                Select Case UCase(_OBJutilities.word(.Value, intIndex2))
                    Case "BYVAL": intIndex2 += 1
                    Case "BYREF": booByVal = False: intIndex2 += 1
                    Case Else:
                End Select
                strName = _OBJutilities.word(.Value, intIndex2)
                intDimensions = 0
                Try
                    strSplit = split(strName, "(")
                Catch ex As Exception
                    errorHandler("Can't split: " & _
                                Err.Number & " " & Err.Description, _
                                "getProcedures_parseParameters", _
                                "Returning Nothing")
                    Return(Nothing)                         
                End Try                    
                strName = strSplit(0)
                intDimensions = 0
                If UBound(strSplit) > 0 Then
                    Try
                        strSplit = split(strSplit(1), ",")
                    Catch ex As Exception
                        errorHandler("Can't split: " & _
                                    Err.Number & " " & Err.Description, _
                                    "getProcedures_parseParameters", _
                                    "Returning Nothing")
                        Return(Nothing)                         
                    End Try                    
                    intDimensions = UBound(strSplit) + 1
                End If
                intIndex2 += 1      
                strType = "Object"
                If intIndex2 < _OBJutilities.words(.Value) _
                   AndAlso _
                   UCase(_OBJutilities.word(.Value, intIndex2)) = "AS" Then
                    strType = _OBJutilities.word(.Value, intIndex2 + 1)
                End If                
                With colSubentry
                    Try
                        .Add(strName) 
                        .Add(booByVal)
                        .Add(intDimensions)
                        .Add(strType)
                        colNew.Add(colSubentry)
                    Catch ex As Exception
                        errorHandler("Can't make subentry and/or add it to collection: " & _
                                    Err.Number & " " & Err.Description, _
                                    "getProcedures_parseParameters", _
                                    "Nothing")
                        Return(Nothing)                         
                    End Try                    
                End With
                intIndex1 = .Index + 1 + .Length
                If Mid(strParameterList, intIndex1, 1) <> "," Then Exit Do                            
            End With                 
        Loop            
        strType = "Object"
        intIndex1 = Instr(intIndex1, strSource, " As ")
        If intIndex1 <> 0 Then strType = _OBJutilities.item(Mid(strSource, intIndex1), _
                                                            2, _
                                                            _OBJutilities.range2String(0, 32), _ 
                                                            True)
        Try
            colNew.Add(strType, , 1)
        Catch ex As Exception
            errorHandler("Can't add return type to collection: " & _
                        Err.Number & " " & Err.Description, _
                        "getProcedures_parseParameters", _
                        "Nothing")
            Return(Nothing)                         
        End Try    
        Return(colNew)    
    End Function

    ' -----------------------------------------------------------------
    ' Transfer procedures to List box on behalf of getProcedures
    '
    '
    Private Sub getProcedures_procedures2Listbox(ByRef usrProcedureTable() As TYPprocedure, _
                                                 ByRef lstBox As ListBox)  
        With lstBox
            .Items.Clear: .Refresh
            Dim intIndex1 As Integer
            For intIndex1 = 1 To UBound(usrProcedureTable)
                With usrProcedureTable(intIndex1)
                    lstBox.Items.Add(.strName & ": " & abstract2OneLine(.strAbstract))
                End With
                .SelectedIndex = .Items.Count - 1
                .Refresh
            Next intIndex1
            .SelectedIndex = CInt(Iif(.Items.Count = 0, -1, 0))
            .Refresh
        End With
    End Sub
    
    ' -----------------------------------------------------------------
    ' Remove Private and Friend scopes from the procedures re
    '
    '
    Private Function getProcedures_removeScopes(ByVal strRE As String) As String
        Dim strReplace As String = "(Public )|(Private )|(Friend )"
        Dim intIndex1 As Integer = Instr(strRE, strReplace)
        If intIndex1 = 0 Then
            errorHandler("Cannot remove scopes from procedure regular expression: " & _
                         "expected contents not found", _
                         "", _
                         "Continuing with unmodified regular expression")
            Return(strRE)                         
        End If        
        Return(Replace(strRE, strReplace, "(Public )"))
    End Function    

    ' -----------------------------------------------------------------
    ' Parse path into directory and file title
    '
    '
    Private Sub parseFileid(ByVal strPath As String, _
                            ByRef strDirectory As String, _
                            ByRef strFileTitle As String)
        Dim intIndex1 As Integer = InstrRev(strPath, "\")
        strDirectory = "": strFileTitle = strPath
        If intIndex1 = 0 Then Return
        strDirectory = Mid(strPath, 1, intIndex1 - 1)
        strFileTitle = Mid(strPath, intIndex1 + 1)
    End Sub
    
    ' -----------------------------------------------------------------
    ' Return True (procedure exists) or False
    '
    '
    Private Function procedureExists(ByVal usrIndexedProcedures As TYPindexedProcedures, _
                                     ByVal strName As String) As Boolean
        Return(procedureName2Index(usrIndexedProcedures, strName) <> 0)
    End Function                                     
    
    ' -----------------------------------------------------------------
    ' Return procedure's index or 0
    '
    '
    Private Function procedureName2Index(ByVal usrIndexedProcedures As TYPindexedProcedures, _
                                         ByVal strName As String) As Integer
        Dim intIndex1 As Integer                                         
        Try
            intIndex1 = CInt(usrIndexedProcedures.colIndex(strName))
        Catch: End Try        
        Return(intIndex1)
    End Function                                     
    
    ' -----------------------------------------------------------------
    ' Convert procedure type to keyword
    '
    '
    Private Function procedureType2Keyword(ByVal enuType As ENUprocedureType) As String
        Select Case enuType
            Case ENUprocedureType.functionProcedure: Return("Function")
            Case ENUprocedureType.invalidProcedure: Return("")
            Case ENUprocedureType.propertyProcedure: Return("Property")
            Case ENUprocedureType.subroutineProcedure: Return("Sub")
            Case Else
                errorHandler("Programming error: unexpected procedure type " & _
                             enuType.ToString, _
                             "procedureType2Keyword", _
                             "Returning a null string")
                Return("")                             
        End Select
    End Function    

    ' -----------------------------------------------------------------
    ' Convert procedure type enumerator to VB-compatible name
    '
    '
    Private Function procInfoExists(ByVal usrIndexedProcedures As TYPindexedProcedures) As Boolean
        Return(Not (usrIndexedProcedures.colIndex Is Nothing))
    End Function

    ' -----------------------------------------------------------------
    ' Convert procedure type enumerator to VB-compatible name
    '
    '
    Private Function procType2String(ByVal enuType As ENUprocedureType) As String
        Select Case enuType
            Case ENUprocedureType.functionProcedure: Return("Function")
            Case ENUprocedureType.propertyProcedure: Return("Property")
            Case ENUprocedureType.subroutineProcedure: Return("Sub")
            Case Else: Return("Invalid")
        End Select
    End Function
    
    ' -----------------------------------------------------------------
    ' Registry to form
    '
    '
    Private Sub registry2Form
        Try
            With radSourceUtilitiesNet
                .Checked = CBool(GetSetting(Application.ProductName, _
                                    Name, _
                                    .Name, _
                                    CStr(.Checked)))
            End With            
            With radSourceUtilitiesCOM
                .Checked = CBool(GetSetting(Application.ProductName, _
                                    Name, _
                                    .Name, _
                                    CStr(.Checked)))
            End With            
            With radSourceWindowsUtilitiesNet
                .Checked = CBool(GetSetting(Application.ProductName, _
                                    Name, _
                                    .Name, _
                                    CStr(.Checked)))
            End With            
            With radSourceOther
                .Checked = CBool(GetSetting(Application.ProductName, _
                                    Name, _
                                    .Name, _
                                    CStr(.Checked)))
            End With            
            With chkSourceLoadOnStartup
                .Checked = CBool(GetSetting(Application.ProductName, _
                                    Name, _
                                    .Name, _
                                    CStr(.Checked)))
            End With            
            With txtSourceOtherFileid
                .Text = GetSetting(Application.ProductName, _
                                   Name, _
                                   .Name, _
                                   .Text) 
            End With   
            adjustSourceDisplay         
            With nudStatusDetail
                .Value = CInt(GetSetting(Application.ProductName, _
                                         Name, _
                                         .Name, _
                                         CStr(.Value)))
            End With            
        Catch ex As Exception
            errorHandler("Cannot obtain settings: " & Err.Number & " " & Err.Description, _
                         "registry2Form", _
                         "Continuing" & vbNewline & vbNewline & _
                         ex.ToString)
        End Try        
    End Sub    

    ' -----------------------------------------------------------------
    ' Show the Easter Egg
    '
    '
    Private Sub showAbout 
        MsgBox(ABOUTINFO)
    End Sub

    ' -----------------------------------------------------------------
    ' Adjust form boundaries, around bottom and right side of control
    '
    '
    Private Sub shrinkWrap(ByVal ctlControl As Control)
        Width =  _OBJwindowsUtilities.controlRight(ctlControl) _
                 + _
                 _OBJwindowsUtilities.Grid * 2
        Height = _OBJwindowsUtilities.controlBottom(ctlControl) _
                 + _
                 _OBJwindowsUtilities.Grid * 4
        CenterToParent
        Refresh
    End Sub
    
    ' -----------------------------------------------------------------
    ' Radio button setting to source file id
    '
    '
    Private Function sourceFileid As String
        If radSourceUtilitiesNet.Checked Then Return(UTILITIESFILEIDNET)
        If radSourceWindowsUtilitiesNet.Checked Then Return(WINDOWSUTILITIESFILEIDNET)
        If radSourceUtilitiesCOM.Checked Then Return(UTILITIESFILEIDCOM)
        If radSourceOther.Checked Then Return(txtSourceOtherFileid.Text)
        radSourceUtilitiesNet.Checked = True
        Return(sourceFileid)
    End Function    
    
    ' -----------------------------------------------------------------
    ' Test the library
    '
    '
    Private Sub testInterface
        Dim booOK As Boolean
        Dim strReport As String
        If radSourceUtilitiesNet.Checked Then
            booOK = _OBJutilities.test(strReport)
        ElseIf radSourceWindowsUtilitiesNet.Checked Then
            booOK = _OBJwindowsUtilities.test(strReport)
        ElseIf radSourceUtilitiesCOM.Checked Then
            If (_OBJutilitiesCOM Is Nothing) Then 
                MsgBox("COM utilities unavailable"): Return
            End If            
            booOK = _OBJutilitiesCOM.test(strReport)
        Else
            MsgBox("Nothing to test"): Return             
        End If        
        If Msgbox("Test has " & _
                  CStr(Iif(booOK, "succeeded", "failed")) & ": " & _
                  "click Yes to see report: click No to return to the main form") _
           = _
           MsgBoxResult.No Then Return
        zoomInterface(strReport, dblHeight:=1)                                                 
    End Sub    
    
    ' -----------------------------------------------------------------
    ' Update progress report
    '
    '
    Private Sub updateProgress(ByVal lblProgress As Label, _
                               ByVal strActivity As String, _
                               ByVal strEntity As String, _
                               ByVal intEntityNumber As Integer, _
                               ByVal intEntityCount As Integer)
        Dim strPercentComplete As String
        If intEntityCount <> 0 Then
            strPercentComplete = ": " & _
                                 CStr(Math.Round(intEntityNumber/intEntityCount * 100, 2)) & _
                                 "% complete"
        End If
        Dim strReport As String = strActivity & " at " & strEntity & " " & _
                                    intEntityNumber & " of " & intEntityCount & _
                                    strPercentComplete
        updateStatusListBox(strReport)
        With lblProgress
            .Width = CInt(_OBJutilities.histogram(intEntityNumber, _
                                             dblRangeMax:=lstStatus.Width, _
                                             dblValueMax:=intEntityCount))
            .Text = strReport                                             
            .Visible = True
            .Refresh
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Update the status list box
    '
    '
    ' --- No level change
    Private Overloads Sub updateStatusListBox(ByVal strMessage As String, _
                                              Optional ByVal booDateStamp As Boolean = True)
        updateStatusListBox(0, strMessage, 0)
    End Sub
    ' --- Change level before displaying the message
    Private Overloads Sub updateStatusListBox(ByVal intLevelChangeBefore As Integer, _
                                              ByVal strMessage As String, _
                                              Optional ByVal booDateStamp As Boolean = True)
        updateStatusListBox(intLevelChangeBefore, strMessage, 0)
    End Sub
    ' --- Change level after displaying the message
    Private Overloads Sub updateStatusListBox(ByVal strMessage As String, _
                                              ByVal intLevelChangeAfter As Integer, _
                                              Optional ByVal booDateStamp As Boolean = True)
        updateStatusListBox(0, strMessage, intLevelChangeAfter)
    End Sub
    ' --- Reset level with no display
    Private Overloads Sub updateStatusListBox()
        lstStatus.Tag = 0
    End Sub
    ' --- Change level before and after displaying the message
    Private Overloads Sub updateStatusListBox(ByVal intLevelChangeBefore As Integer, _
                                              ByVal strMessage As String, _
                                              ByVal intLevelChangeAfter As Integer, _
                                              Optional ByVal booDateStamp As Boolean = True)
        Dim intLevel As Integer
        Try
            intLevel = CInt(lstStatus.Tag)
        Catch: End Try
        intLevel += intLevelChangeBefore
        If intLevel < 0 Then intLevel = 0
        If strMessage <> "" AndAlso intLevel <= nudStatusDetail.Value Then
            With lstStatus
                Dim strSplit() As String 
                Try
                    strSplit = split(strMessage, vbNewline)
                Catch ex As Exception
                    MsgBox("Error in " & _
                           Me.Name & ".updateStatusListBox" & ": " & _
                           "cannot split: " & _
                           Err.Number & " " & Err.Description)
                    Return
                End Try
                Dim strDateStamp As String
                If booDateStamp Then strDateStamp = Trim(CStr(Now))
                Dim strIndent As String = _OBJutilities.copies(" ", intLevel * 5)
                Dim intIndex1 As Integer
                If UBound(strSplit) = 0 Then
                    .Items.Add(strIndent & strDateStamp & " " & _
                               Replace(_OBJutilities.soft2HardParagraph(strSplit(0), bytLineWidth:=50), _
                                       vbNewline, _
                                       vbNewline & _
                                       strIndent & _
                                       _OBJutilities.copies(" ", Len(strDateStamp) + 1)))
                Else
                    For intIndex1 = 0 To UBound(strSplit)
                        .Items.Add(strIndent & _
                                   strDateStamp & " " & _
                                   strSplit(intIndex1))
                    Next intIndex1
                End If
                .SelectedIndex = .Items.Count - 1
                .Focus
                '.Refresh
            End With
        End If
        intLevel += intLevelChangeAfter
        lstStatus.Tag = intLevel
    End Sub
    
    ' -----------------------------------------------------------------
    ' Usage analysis
    '
    '
    Private Sub usageAnalysis(ByVal usrIndexedProcedures As TYPindexedProcedures)
        Dim booError As Boolean
        Dim intIndirectRefCount As Integer
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim intIndex2 As Integer
        Dim intIndex3 As Integer
        Dim intIndex4 As Integer
        Dim strActivity As String  
        Dim strActivity2 As String  
        Dim strActivity3 As String  
        Dim strActivity4 As String  
        Dim strCalledProcedure As String
        Dim strIndirectRefCounts As String  
        Dim strNext As String
        Dim strNext2 As String
        If Not procInfoExists(usrIndexedProcedures) Then
            MsgBox("No procedures have been loaded"): Return
        End If        
        updateStatusListBox("Usage Analysis", 1)
        ' --- Pass one: get direct usage
        strActivity = "Usage Analysis (pass 1: get direct usage)"
        updateStatusListBox(strActivity, 1)
        With usrIndexedProcedures
            Dim objRE As System.Text.RegularExpressions.Regex
            Dim objMatch As System.Text.RegularExpressions.Match
            Try
                objRE = New System.Text.RegularExpressions.Regex _
                        (_OBJutilities.commonRegularExpressions(CStr(Iif(.booNet, _
                                                                         "compoundIdentifierNet", _
                                                                         "compoundIdentifierCOM"))))
            Catch ex As Exception
                errorHandler("Cannot create regular expression: " & _
                             Err.Number & " " & Err.Description, _
                             "", _
                             "Returning with no usage analysis")
                Return                             
            End Try         
            For intIndex1 = 1 To UBound(.usrProcedures)
                With .usrProcedures(intIndex1)
                    updateProgress(lblProgress, _
                                   strActivity, _
                                   "procedure", _
                                   intIndex1, _
                                   UBound(usrIndexedProcedures.usrProcedures))                
                    updateStatusListBox(1, "")                                   
                    Try
                        .colUsedBy = Nothing: .colUses = Nothing
                        .colUsedBy = New Collection
                        .colUses = New Collection
                    Catch ex As Exception
                        errorHandler("Cannot create used/uses collections: " & _
                                    Err.Number & " " & Err.Description, _
                                    "", _
                                    "Returning with partial or empty usage analysis")
                        Return                             
                    End Try 
                    intIndex2 = 0
                    strActivity2 = "Analyzing procedure " & .strName
                    updateStatusListBox(strActivity2, 1)
                    intCount = 0
                    Do While intIndex2 < Len(.strText)
                        intCount += 1
                        If intCount Mod 32 = 0 Then
                            updateProgress(lblProgress2, _
                                           strActivity2, _
                                           "character", _
                                           intIndex2, _
                                           Len(.strText))                
                        End If                                       
                        Try
                            objMatch = objRE.Match(.strText, intIndex2)
                        Catch ex As Exception
                            errorHandler("Cannot create used/uses collections: " & _
                                        Err.Number & " " & Err.Description, _
                                        "", _
                                        "Returning with partial or empty usage analysis")
                            Return                             
                        End Try  
                        With objMatch
                            If .Length = 0 Then Exit Do
                            If UCase(.Value) _
                               <> _
                               UCase(usrIndexedProcedures.usrProcedures(intIndex1).strName) _
                               AndAlso _
                               procedureExists(usrIndexedProcedures, .Value) Then
                                Try
                                    usrIndexedProcedures.usrProcedures(intIndex1).colUses.Add _
                                    (.Value, .Value)
                                Catch: End Try                                
                            End If                            
                            intIndex2 = .Index + .Length                                       
                        End With       
                    Loop    
                    updateStatusListBox(-1, "")
                    lblProgress2.Visible = False
                    updateStatusListBox(-1, "Completed: " & strActivity2)                   
                End With                                                   
            Next intIndex1
        End With         
        lblProgress.Visible = False                           
        updateStatusListBox(-1, "Completed: " & strActivity)
        ' --- Pass two: get indirect usage
        strActivity = "Usage Analysis (pass 2: get indirect usage)"
        updateStatusListBox(strActivity, 1)
        Do 
            strActivity2 = "Iterating through all procedures to add indirect usage" & _
                           strIndirectRefCounts
            updateStatusListBox(strActivity2, 1)
            intIndirectRefCount = 0
            With usrIndexedProcedures
                For intIndex1 = 1 To UBound(.usrProcedures)
                    updateProgress(lblProgress, _
                                   strActivity2, _
                                   "procedure", _
                                   intIndex1, _
                                   UBound(.usrProcedures)) 
                    With .usrProcedures(intIndex1)
                        strActivity3 = "Expanding direct to indirect usage for " & .strName
                        updateStatusListBox(strActivity2, 1)                                   
                        For intIndex2 = 1 To .colUses.Count
                            updateProgress(lblProgress2, _
                                           strActivity3, _
                                           "directly called procedure", _
                                           intIndex2, _
                                           UBound(usrIndexedProcedures.usrProcedures))    
                            strCalledProcedure = Cstr(.colUses.Item(intIndex2))                                       
                            intIndex3 = procedureName2Index(usrIndexedProcedures, strCalledProcedure) 
                            If intIndex3 <> 0 Then
                                strActivity4 = "Expanding called procedure " & strCalledProcedure
                                updateStatusListBox(1, strActivity3)
                                With usrIndexedProcedures.usrProcedures(intIndex3).colUses
                                    For intIndex4 = 1 To .Count
                                        If intIndex4 Mod 10 = 0 Then
                                            updateProgress(lblProgress3, _
                                                           strActivity4, _
                                                           "procedure", _
                                                           intIndex1, _
                                                           UBound(usrIndexedProcedures.usrProcedures))
                                        End If                                                               
                                        strNext = CStr(.Item(intIndex4))
                                        booError = False
                                        Try
                                            usrIndexedProcedures.usrProcedures(intIndex1).colUses.Add _
                                                (strNext, strNext)
                                        Catch
                                            booError = True
                                        End Try      
                                        If Not booError Then intIndirectRefCount += 1                                  
                                    Next intIndex4        
                                End With       
                                lblProgress3.Visible = False             
                                updateStatusListBox(-1, "Completed: " & strActivity4)                                   
                            End If                                               
                        Next intIndex2 
                        lblProgress2.Visible = False                   
                        updateStatusListBox(-1, "Completed: " & strActivity3)                                   
                    End With                                                            
                Next intIndex1
                updateStatusListBox(-1, "Completed: " & strActivity2) 
                lblProgress.Visible = False            
            End With       
            If strIndirectRefCounts = "" Then 
                strIndirectRefCounts = ": Indirect references discovered in these iterations: " & _
                                       CStr(intIndirectRefCount)
            Else
                strIndirectRefCounts &= "," & CStr(intIndirectRefCount)
            End If                                       
        Loop Until intIndirectRefCount = 0            
        updateStatusListBox(-1, "Completed " & strActivity)
        ' --- Pass three: get used by information
        strActivity = "Usage Analysis (pass 3: get used by information)"
        updateStatusListBox(strActivity, 1)
        With usrIndexedProcedures
            For intIndex1 = 1 To UBound(.usrProcedures)
                updateProgress(lblProgress, _
                               strActivity, _
                               "procedure", _
                               intIndex1, _
                               UBound(.usrProcedures)) 
                updateStatusListBox(1, "")                               
                With .usrProcedures(intIndex1)
                    strNext = .strName
                    strActivity2 = "Checking procedure calls from " & .strName
                    For intIndex2 = 1 To UBound(usrIndexedProcedures.usrProcedures)
                        updateProgress(lblProgress2, _
                                       strActivity2, _
                                       "called procedure", _
                                       intIndex2, _
                                       UBound(usrIndexedProcedures.usrProcedures)) 
                        If .strName <> usrIndexedProcedures.usrProcedures(intIndex2).strName Then                                      
                            strNext2 = ""
                            Try
                                strNext2 = CStr(usrIndexedProcedures.usrProcedures(intIndex2).colUses(strNext))
                            Catch: End Try   
                            If strNext2 <> "" Then
                                Try
                                    .colUsedBy.Add(usrIndexedProcedures.usrProcedures(intIndex2).strName)
                                Catch: End Try
                            End If      
                        End If                                        
                    Next intIndex2    
                End With            
                lblProgress2.Visible = False
                updateStatusListBox(-1, "Completed: " & strActivity2)                               
            Next intIndex1 
            lblProgress.Visible = False           
        End With        
        updateStatusListBox(-1, "Completed: " & strActivity)
        ' --- Miller time
        updateStatusListBox(-1, "Usage Analysis complete")
    End Sub    

    ' -----------------------------------------------------------------
    ' Zoom control
    '
    '
    Private Sub zoomInterface(ByVal objZoomed As Object, _
                              Optional ByVal dblWidth As Double = 1.5, _
                              Optional ByVal dblHeight As Double = 3)
        Dim ctlZoomedHandle As Control                              
        If (TypeOf objZoomed Is System.String) Then                              
            Dim txtBox As TextBox                                
            Try
                txtBox = New TextBox
                With txtBox
                    .Height = Height: .Width = Width: .Multiline = True
                    .Text = CStr(objZoomed)
                End With            
            Catch ex As Exception
                errorHandler("Can't create textbox: " & _
                            Err.Number & " " & Err.Description, _
                            "displayInfoForm", _
                            "Returning to caller" & vbNewline & vbNewline & _
                            ex.ToString)
                Return                         
            End Try    
            ctlZoomedHandle = txtBox
        ElseIf (TypeOf objZoomed Is Control) 
            ctlZoomedHandle = CType(objZoomed, Control)            
        End If             
        Dim ctlZoom As zoom.zoom
        Try
            ctlZoom = New zoom.zoom(ctlZoomedHandle)
        Catch ex As Exception
            errorHandler("Cannot create zoom control: " & _
                         Err.Number & " " & Err.Description, _
                         "zoomInterface", _
                         ex.ToString)
            Return
        End Try
        With ctlZoom
            .setSize(dblWidth, dblHeight)
            .ShowZoom
            .Dispose
        End With
    End Sub

#End Region ' " General procedures "

End Class
