Option Strict

' *********************************************************************
' *                                                                   *
' * Variable test form                                                *
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
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdInspect As System.Windows.Forms.Button
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents cmdObject2XML As System.Windows.Forms.Button
    Friend WithEvents lblVariables As System.Windows.Forms.Label
    Friend WithEvents lstVariables As System.Windows.Forms.ListBox
    Friend WithEvents cmdRandomize As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents cmdStatusZoom As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents cmdClearSettings As System.Windows.Forms.Button
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    Friend WithEvents cmdToDescription As System.Windows.Forms.Button
    Friend WithEvents lblVariableExpression As System.Windows.Forms.Label
    Friend WithEvents cmdCreateVariable As System.Windows.Forms.Button
    Friend WithEvents cmdCloneVariable As System.Windows.Forms.Button
    Friend WithEvents cmdDisposeVariable As System.Windows.Forms.Button
    Friend WithEvents cmdVariablesCreate As System.Windows.Forms.Button
    Friend WithEvents cmdVariableClear As System.Windows.Forms.Button
    Friend WithEvents cmdMkRandomVariable As System.Windows.Forms.Button
    Friend WithEvents txtVariableExpression As System.Windows.Forms.TextBox
    Friend WithEvents cmdEmpiricalDope As System.Windows.Forms.Button
    Friend WithEvents cmdToString As System.Windows.Forms.Button
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents cmdReset As System.Windows.Forms.Button
    Friend WithEvents cmdCloseNosave As System.Windows.Forms.Button
    Friend WithEvents cmdValue As System.Windows.Forms.Button
    Friend WithEvents cmdValueAssignment As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblVariableExpression = New System.Windows.Forms.Label
Me.txtVariableExpression = New System.Windows.Forms.TextBox
Me.cmdCreateVariable = New System.Windows.Forms.Button
Me.cmdCloneVariable = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.lblVariables = New System.Windows.Forms.Label
Me.lstVariables = New System.Windows.Forms.ListBox
Me.cmdInspect = New System.Windows.Forms.Button
Me.cmdTest = New System.Windows.Forms.Button
Me.cmdObject2XML = New System.Windows.Forms.Button
Me.cmdDisposeVariable = New System.Windows.Forms.Button
Me.cmdVariablesCreate = New System.Windows.Forms.Button
Me.cmdVariableClear = New System.Windows.Forms.Button
Me.cmdMkRandomVariable = New System.Windows.Forms.Button
Me.cmdRandomize = New System.Windows.Forms.Button
Me.lblStatus = New System.Windows.Forms.Label
Me.lstStatus = New System.Windows.Forms.ListBox
Me.cmdStatusZoom = New System.Windows.Forms.Button
Me.lblProgress = New System.Windows.Forms.Label
Me.cmdClearSettings = New System.Windows.Forms.Button
Me.cmdAbout = New System.Windows.Forms.Button
Me.cmdToDescription = New System.Windows.Forms.Button
Me.cmdEmpiricalDope = New System.Windows.Forms.Button
Me.cmdToString = New System.Windows.Forms.Button
Me.cmdClear = New System.Windows.Forms.Button
Me.cmdReset = New System.Windows.Forms.Button
Me.cmdCloseNosave = New System.Windows.Forms.Button
Me.cmdValue = New System.Windows.Forms.Button
Me.cmdValueAssignment = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblVariableExpression
'
Me.lblVariableExpression.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblVariableExpression.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariableExpression.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblVariableExpression.Location = New System.Drawing.Point(7, 8)
Me.lblVariableExpression.Name = "lblVariableExpression"
Me.lblVariableExpression.Size = New System.Drawing.Size(693, 16)
Me.lblVariableExpression.TabIndex = 0
Me.lblVariableExpression.Text = "Variable Expression: boldface if variable exists: normal face otherwise"
Me.lblVariableExpression.TextAlign = System.Drawing.ContentAlignment.TopCenter
'
'txtVariableExpression
'
Me.txtVariableExpression.Font = New System.Drawing.Font("Courier New", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtVariableExpression.Location = New System.Drawing.Point(7, 24)
Me.txtVariableExpression.Name = "txtVariableExpression"
Me.txtVariableExpression.Size = New System.Drawing.Size(693, 23)
Me.txtVariableExpression.TabIndex = 1
Me.txtVariableExpression.Text = "Unknown"
'
'cmdCreateVariable
'
Me.cmdCreateVariable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCreateVariable.Location = New System.Drawing.Point(8, 56)
Me.cmdCreateVariable.Name = "cmdCreateVariable"
Me.cmdCreateVariable.TabIndex = 2
Me.cmdCreateVariable.Text = "Create"
'
'cmdCloneVariable
'
Me.cmdCloneVariable.Enabled = False
Me.cmdCloneVariable.Location = New System.Drawing.Point(168, 56)
Me.cmdCloneVariable.Name = "cmdCloneVariable"
Me.cmdCloneVariable.TabIndex = 3
Me.cmdCloneVariable.Text = "Clone"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(8, 440)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(104, 24)
Me.cmdForm2Registry.TabIndex = 4
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(120, 440)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(104, 24)
Me.cmdRegistry2Form.TabIndex = 5
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(624, 56)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(75, 56)
Me.cmdClose.TabIndex = 6
Me.cmdClose.Text = "Close"
'
'lblVariables
'
Me.lblVariables.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblVariables.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariables.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblVariables.Location = New System.Drawing.Point(7, 120)
Me.lblVariables.Name = "lblVariables"
Me.lblVariables.Size = New System.Drawing.Size(686, 20)
Me.lblVariables.TabIndex = 7
Me.lblVariables.Text = "Variables: click to select: double click to set test object"
Me.lblVariables.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'lstVariables
'
Me.lstVariables.BackColor = System.Drawing.SystemColors.Control
Me.lstVariables.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstVariables.ItemHeight = 16
Me.lstVariables.Location = New System.Drawing.Point(8, 144)
Me.lstVariables.Name = "lstVariables"
Me.lstVariables.Size = New System.Drawing.Size(696, 116)
Me.lstVariables.TabIndex = 8
'
'cmdInspect
'
Me.cmdInspect.Enabled = False
Me.cmdInspect.Location = New System.Drawing.Point(88, 56)
Me.cmdInspect.Name = "cmdInspect"
Me.cmdInspect.TabIndex = 11
Me.cmdInspect.Text = "Inspect"
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(7, 88)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.TabIndex = 12
Me.cmdTest.Text = "Stress Test"
'
'cmdObject2XML
'
Me.cmdObject2XML.Location = New System.Drawing.Point(168, 88)
Me.cmdObject2XML.Name = "cmdObject2XML"
Me.cmdObject2XML.TabIndex = 13
Me.cmdObject2XML.Text = "XML"
'
'cmdDisposeVariable
'
Me.cmdDisposeVariable.Enabled = False
Me.cmdDisposeVariable.Location = New System.Drawing.Point(408, 56)
Me.cmdDisposeVariable.Name = "cmdDisposeVariable"
Me.cmdDisposeVariable.TabIndex = 14
Me.cmdDisposeVariable.Text = "Dispose "
'
'cmdVariablesCreate
'
Me.cmdVariablesCreate.Location = New System.Drawing.Point(560, 120)
Me.cmdVariablesCreate.Name = "cmdVariablesCreate"
Me.cmdVariablesCreate.Size = New System.Drawing.Size(144, 21)
Me.cmdVariablesCreate.TabIndex = 16
Me.cmdVariablesCreate.Text = "Create Random Variables"
'
'cmdVariableClear
'
Me.cmdVariableClear.Location = New System.Drawing.Point(488, 120)
Me.cmdVariableClear.Name = "cmdVariableClear"
Me.cmdVariableClear.Size = New System.Drawing.Size(72, 21)
Me.cmdVariableClear.TabIndex = 18
Me.cmdVariableClear.Text = "Clear (all)"
'
'cmdMkRandomVariable
'
Me.cmdMkRandomVariable.Location = New System.Drawing.Point(88, 88)
Me.cmdMkRandomVariable.Name = "cmdMkRandomVariable"
Me.cmdMkRandomVariable.TabIndex = 19
Me.cmdMkRandomVariable.Text = "Random"
'
'cmdRandomize
'
Me.cmdRandomize.Location = New System.Drawing.Point(344, 440)
Me.cmdRandomize.Name = "cmdRandomize"
Me.cmdRandomize.Size = New System.Drawing.Size(104, 24)
Me.cmdRandomize.TabIndex = 20
Me.cmdRandomize.Text = "Randomize"
'
'lblStatus
'
Me.lblStatus.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblStatus.Location = New System.Drawing.Point(8, 264)
Me.lblStatus.Name = "lblStatus"
Me.lblStatus.Size = New System.Drawing.Size(689, 20)
Me.lblStatus.TabIndex = 21
Me.lblStatus.Text = "Status"
Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.SystemColors.Control
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 16
Me.lstStatus.Location = New System.Drawing.Point(8, 288)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(697, 132)
Me.lstStatus.TabIndex = 22
'
'cmdStatusZoom
'
Me.cmdStatusZoom.Location = New System.Drawing.Point(624, 264)
Me.cmdStatusZoom.Name = "cmdStatusZoom"
Me.cmdStatusZoom.Size = New System.Drawing.Size(80, 21)
Me.cmdStatusZoom.TabIndex = 23
Me.cmdStatusZoom.Text = "Zoom"
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblProgress.Location = New System.Drawing.Point(7, 420)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(697, 14)
Me.lblProgress.TabIndex = 24
Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.TopCenter
Me.lblProgress.Visible = False
'
'cmdClearSettings
'
Me.cmdClearSettings.Location = New System.Drawing.Point(232, 440)
Me.cmdClearSettings.Name = "cmdClearSettings"
Me.cmdClearSettings.Size = New System.Drawing.Size(104, 24)
Me.cmdClearSettings.TabIndex = 26
Me.cmdClearSettings.Text = "Clear Settings"
'
'cmdAbout
'
Me.cmdAbout.Location = New System.Drawing.Point(456, 440)
Me.cmdAbout.Name = "cmdAbout"
Me.cmdAbout.Size = New System.Drawing.Size(88, 24)
Me.cmdAbout.TabIndex = 27
Me.cmdAbout.Text = "About"
'
'cmdToDescription
'
Me.cmdToDescription.Enabled = False
Me.cmdToDescription.Location = New System.Drawing.Point(248, 88)
Me.cmdToDescription.Name = "cmdToDescription"
Me.cmdToDescription.TabIndex = 28
Me.cmdToDescription.Text = "Describe"
'
'cmdEmpiricalDope
'
Me.cmdEmpiricalDope.Enabled = False
Me.cmdEmpiricalDope.Location = New System.Drawing.Point(328, 88)
Me.cmdEmpiricalDope.Name = "cmdEmpiricalDope"
Me.cmdEmpiricalDope.Size = New System.Drawing.Size(72, 23)
Me.cmdEmpiricalDope.TabIndex = 29
Me.cmdEmpiricalDope.Text = "Dope"
'
'cmdToString
'
Me.cmdToString.Enabled = False
Me.cmdToString.Location = New System.Drawing.Point(408, 88)
Me.cmdToString.Name = "cmdToString"
Me.cmdToString.TabIndex = 30
Me.cmdToString.Text = "toString"
'
'cmdClear
'
Me.cmdClear.Enabled = False
Me.cmdClear.Location = New System.Drawing.Point(248, 56)
Me.cmdClear.Name = "cmdClear"
Me.cmdClear.TabIndex = 31
Me.cmdClear.Text = "Clear"
'
'cmdReset
'
Me.cmdReset.Enabled = False
Me.cmdReset.Location = New System.Drawing.Point(328, 56)
Me.cmdReset.Name = "cmdReset"
Me.cmdReset.TabIndex = 32
Me.cmdReset.Text = "Reset"
'
'cmdCloseNosave
'
Me.cmdCloseNosave.Location = New System.Drawing.Point(552, 440)
Me.cmdCloseNosave.Name = "cmdCloseNosave"
Me.cmdCloseNosave.Size = New System.Drawing.Size(152, 25)
Me.cmdCloseNosave.TabIndex = 33
Me.cmdCloseNosave.Text = "Close-don't save settings"
'
'cmdValue
'
Me.cmdValue.Enabled = False
Me.cmdValue.Location = New System.Drawing.Point(488, 56)
Me.cmdValue.Name = "cmdValue"
Me.cmdValue.TabIndex = 34
Me.cmdValue.Text = "Get Value"
'
'cmdValueAssignment
'
Me.cmdValueAssignment.Enabled = False
Me.cmdValueAssignment.Location = New System.Drawing.Point(488, 88)
Me.cmdValueAssignment.Name = "cmdValueAssignment"
Me.cmdValueAssignment.TabIndex = 35
Me.cmdValueAssignment.Text = "Set Value"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(710, 467)
Me.Controls.Add(Me.cmdValueAssignment)
Me.Controls.Add(Me.cmdValue)
Me.Controls.Add(Me.cmdCloseNosave)
Me.Controls.Add(Me.cmdReset)
Me.Controls.Add(Me.cmdClear)
Me.Controls.Add(Me.cmdToString)
Me.Controls.Add(Me.cmdEmpiricalDope)
Me.Controls.Add(Me.cmdToDescription)
Me.Controls.Add(Me.cmdAbout)
Me.Controls.Add(Me.cmdClearSettings)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.cmdStatusZoom)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.lblStatus)
Me.Controls.Add(Me.cmdRandomize)
Me.Controls.Add(Me.cmdMkRandomVariable)
Me.Controls.Add(Me.cmdVariableClear)
Me.Controls.Add(Me.cmdVariablesCreate)
Me.Controls.Add(Me.cmdDisposeVariable)
Me.Controls.Add(Me.cmdObject2XML)
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.cmdInspect)
Me.Controls.Add(Me.lstVariables)
Me.Controls.Add(Me.lblVariables)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.cmdCloneVariable)
Me.Controls.Add(Me.cmdCreateVariable)
Me.Controls.Add(Me.txtVariableExpression)
Me.Controls.Add(Me.lblVariableExpression)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "qbVariable"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities
    Private Shared _OBJqbVariableType As qbVariableType.qbVariableType
    Private WithEvents OBJqbVariable As qbVariable.qbVariable
    Private COLtestObjects As Collection             ' Of qbVariable 
    Private OBJtoolTip As ToolTip
    Private FRMmodifyVariable As modifyVariable

    ' --- About information Easter egg
    Private Const ABOUTINFO As String = _
        "qbVariableTest" & _
        vbNewLine & vbNewLine & vbNewLine & _
        "This form and application tests the qbVariable object, designed as the OO " & _
        "model for Quick Basic variables." & vbNewLine & vbNewLine & _
        "This form and application was developed commencing on 4/5/2003 by" & vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & vbNewLine & _
        "http://members.screenz.com/edNilges" & vbNewLine & vbNewLine & _
        "For more information, click the Zoom button on the right side of the main form " & _
        "above the status box, and scroll to see a boxed set of instructions."
    Private Const ABOUTINFO_EXTENSION As String = _
        "To try this form and application out, type a simple variable type/value such as " & _
        """integer:10"" in the text box at the top of the form. Then, click " & _
        "the button labeled ""Create"" near the upper left hand " & _
        "corner of the form to see that an object is created...and inspected for validity!" & _
        vbNewLine & vbNewLine & _
        "Then, click the Inspect and Object to XML buttons to see how the object " & _
        "self-inspects for its own sanity (and yours) and converts its state to " & _
        "eXtended Markup Language." & vbNewLine & vbNewLine & _
        "Then, click the Dispose button to get rid of the object, and " & _
        "enter a more complex variable type/value such as ""Variant,String:""Test"""" or " & _
        """Array,Integer,1,10:32767(10)"" (ten integers with the value 32767)." & vbNewLine & vbNewLine & _
        "Inspect these examples using the Inspect button, " & _
        "use the Object to XML button to see their state, " & _
        "and dispose each one in turn." & vbNewLine & vbNewLine & _
        "The Describe button provides the English description of a data type underlying " & _
        "a value. " & _
        "Dispose of the existing data type, enter ""Array,Long,1,2,1,3"", " & _
        "click the Create variable type button, and click Describe to try this out." & vbNewLine & vbNewLine & _
        "Finally, the ""Stress Test"" button will stress-test this object. Click Test, " & _
        "sit back, and watch the progress of testing in the status box. " & _
        "This will provide some assurance if you've changed the source code."

    Private BOOnoEvents As Boolean

#End Region

#Region " Form events "

    Private Sub cmdAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbout.Click
        about(False)
    End Sub

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        With OBJqbVariable
            If .clearVariable Then
                msgboxInterface("Cleared the qbVariable to " & _OBJutilities.ellipsis(.ToString, 64))
            End If
        End With
    End Sub

    Private Sub cmdClearSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearSettings.Click
        Select Case msgboxInterface("Do you want to clear the Registry settings (defaults) " & _
                                    "of this form and application?" & vbNewLine & vbNewLine & _
                                    "Note that this will restore default settings.", _
                                    MsgBoxStyle.YesNo)
            Case MsgBoxResult.Yes
                If clearRegistry() Then registry2Form()
            Case MsgBoxResult.No
            Case Else
                _OBJutilities.errorHandler("Unexpected reply from MsgBox", _
                                           Me.Name, "cmdClearSettings_Click", _
                                           "Continuing despite this serious error")
        End Select
    End Sub

    Private Sub cmdCloneVariable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCloneVariable.Click
        cloneVariable()
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer(False)
    End Sub

    Private Sub cmdCloseNosave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCloseNosave.Click
        closer(True)
    End Sub

    Private Sub cmdCreateVariable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateVariable.Click
        createVariable()
    End Sub

    Private Sub cmdDisposeVariable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDisposeVariable.Click
        disposeVariable(OBJqbVariable)
        changeVTcontrols(False)
    End Sub

    Private Sub cmdEmpiricalDope_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmpiricalDope.Click
        msgboxInterface(OBJqbVariable.empiricalDope)
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry()
    End Sub

    Private Sub cmdInspect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInspect.Click
        inspectInterface()
    End Sub

    Private Sub cmdMkRandomVariable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMkRandomVariable.Click
        mkRandomType()
    End Sub

    Private Sub cmdObject2XML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdObject2XML.Click
        Dim strXML As String
        If (OBJqbVariable Is Nothing) OrElse Not OBJqbVariable.Usable Then
            strXML = OBJqbVariable.class2XML
        Else
            strXML = OBJqbVariable.object2XML
        End If
        zoomInterface(strXML)
    End Sub

    Private Sub cmdRandomize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRandomize.Click
        Randomize()
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form()
    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
        With OBJqbVariable
            If .resetVariable Then
                msgboxInterface("Reset the qbVariable to " & _OBJutilities.ellipsis(.ToString, 64))
            End If
        End With
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click
        testInterface()
    End Sub

    Private Sub cmdToDescription_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToDescription.Click
        msgboxInterface("The variableType inside " & Me.Name & " has this description: " & _
                        vbNewLine & vbNewLine & _
                        OBJqbVariable.Dope.toDescription)
    End Sub

    Private Sub cmdToString_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToString.Click
        msgboxInterface(OBJqbVariable.ToString)
    End Sub

    Private Sub cmdValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValue.Click
        showValue()
    End Sub

    Private Sub cmdValueAssignment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValueAssignment.Click
        valueAssignment()
    End Sub

    Private Sub cmdVariableTypesClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVariableClear.Click
        clearVT()
    End Sub

    Private Sub cmdVariableTypesCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVariablesCreate.Click
        mkRandomTypes()
    End Sub

    Private Sub cmdStatusZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStatusZoom.Click
        zoomInterface(lstStatus, 1.25, 3)
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        closer(False)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CenterToScreen()
        Opacity = 0.5
        Show()
        Refresh()
        Dim booNotFirstTime As Boolean
        Try
            booNotFirstTime = CBool(GetSetting(Application.ProductName, Me.Name, "notFirstTime"))
        Catch : End Try
        If Not booNotFirstTime Then
            If Not about(True) Then End
            Try
                SaveSetting(Application.ProductName, Me.Name, "notFirstTime", CStr(True))
            Catch : End Try
        End If
        Try
            FRMmodifyVariable = New modifyVariable
        Catch
            msgboxInterface("Can't create modifyVariable form: " & _
                            Err.Number & " " & Err.Description)
            closer(True)
        End Try
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Loading", 1)
        registry2Form()
        Try
            COLtestObjects = New Collection
        Catch
            _OBJutilities.errorHandler("Cannot initialize the test object collection: " & _
                                      Err.Number & " " & Err.Description, _
                                      Me.Name, _
                                      "Form1_Load", _
                                      "Terminating this application")
            closer(True)
        End Try
        setToolTips()
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Loading complete")
        Opacity = 1
        Refresh()
   End Sub

   Private Sub lstVariableTypes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstVariables.DoubleClick
       assignSelectedObject(lstVariables.SelectedIndex)
   End Sub

   Private Sub lstVariableTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstVariables.Click
       If (OBJqbVariable Is Nothing) Then
           cmdDisposeVariable.Visible = True
       End If
   End Sub

    Private Sub txtVariableExpression_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVariableExpression.Leave
        disposeVariable(OBJqbVariable)
        changeVTcontrols(False)
    End Sub

#End Region

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Show the application's Easter Egg
    '
    '
    Private Function about(ByVal booAllowCancel As Boolean) As Boolean
        Dim objMsgBoxResponse As MsgBoxResult = _
            msgboxInterface(ABOUTINFO & _
                            vbNewLine & vbNewLine & vbNewLine & _
                            "Information about the qbVariable class" & _
                            vbNewLine & vbNewLine & _
                            OBJqbVariable.About, _
                            CType(IIf(booAllowCancel, MsgBoxStyle.OKCancel, MsgBoxStyle.OKOnly), _
                                  MsgBoxStyle))
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "More information About this form", _
                                                 1)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 _OBJutilities.string2Box _
                                                 (_OBJutilities.soft2HardParagraph(ABOUTINFO_EXTENSION)), _
                                                 booSplitLines:=True, _
                                                 booIncludeDate:=False)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 -1, _
                                                 "End of information About this form")
        If Not booAllowCancel Then Return (True)
        Return (objMsgBoxResponse = MsgBoxResult.OK)
    End Function

    ' -----------------------------------------------------------------
    ' Add a test object to collection and screen
    '
    '
    Private Function addTestObject(ByVal objNew As qbVariable.qbVariable) _
            As Boolean
        Try
            COLtestObjects.Add(objNew)
        Catch
            _OBJutilities.errorHandler("Can't add test object to collection: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "addTestObject", _
                                       "Continuing without saving the test object")
            Return (False)
        End Try
        With lstVariables
            .Items.Add(objNew.Name & " " & _
                       "(" & _
                       _OBJutilities.enquote(_OBJutilities.ellipsis(objNew.ToString, 32)) & " " & _
                       ")")
            .SelectedIndex = .Items.Count - 1
            .Focus()
            .Refresh()
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Move the selected object to the test object
    '
    '
    Private Sub assignSelectedObject(ByVal intIndex As Integer)
        Dim booClone As Boolean
        Select Case msgboxInterface("Should the selected object be cloned? " & _
                                    vbNewLine & vbNewLine & _
                                    "Click Yes to make a clone. " & _
                                    vbNewLine & vbNewLine & _
                                    "Click No to assign the selected object to the main test object " & _
                                    "without making a clone. " & _
                                    vbNewLine & vbNewLine & _
                                    "Click Cancel to cancel this operation", _
                                    MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Yes : booClone = True
            Case MsgBoxResult.No : booClone = False
            Case MsgBoxResult.Cancel : Return
            Case Else
                _OBJutilities.errorHandler("Unexpected result from MsgBox", _
                                           Me.Name, _
                                           "assignSelectedObject", _
                                           "Returning to caller. This is an apparent Visual Studio error.")
                Return
        End Select
        If booClone Then
            OBJqbVariable = CType(COLtestObjects.Item(intIndex + 1), qbVariable.qbVariable)
            changeVTcontrols(True)
        Else
            createVariable(False)
        End If
        txtVariableExpression.Text = OBJqbVariable.ToString
    End Sub

    ' -----------------------------------------------------------------
    ' Change the enabled status of the variable type controls
    '
    '
    Private Sub changeVTcontrols(ByVal booEnabled As Boolean)
        cmdCloneVariable.Enabled = booEnabled
        cmdCreateVariable.Enabled = Not booEnabled
        cmdDisposeVariable.Enabled = booEnabled
        cmdInspect.Enabled = booEnabled
        cmdMkRandomVariable.Enabled = Not booEnabled
        cmdToDescription.Enabled = booEnabled
        cmdEmpiricalDope.Enabled = booEnabled
        cmdToString.Enabled = booEnabled
        cmdClear.Enabled = booEnabled
        cmdReset.Enabled = booEnabled
        cmdValue.Enabled = booEnabled
        cmdValueAssignment.Enabled = booEnabled
        txtVariableExpression.Font = _
            New Font(txtVariableExpression.Font, _
                     CType(IIf(booEnabled, FontStyle.Bold, FontStyle.Regular), _
                           FontStyle))
    End Sub

    ' -----------------------------------------------------------------
    ' Clear the Registry settings associated with this form
    '
    '
    Private Function clearRegistry() As Boolean
        Try
            DeleteSetting(Application.ProductName)
            registry2Form()
        Catch
            _OBJutilities.errorHandler("Cannot clear Registry settings for this " & _
                                       "application: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, "clearRegistry", _
                                       "Returning False but continuing")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Clear variable type collection on screen and in storage
    '
    '
    Private Function clearVT() As Boolean
        If Not clearVTcollection() Then Return (False)
        With lstVariables
            .Items.Clear()
            .Refresh()
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Clear the variable type collection in storage
    '
    '
    Private Function clearVTcollection() As Boolean
        If (COLtestObjects Is Nothing) Then Return (True)
       With COLtestObjects
            Dim intCount As Integer = .Count
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            Dim objHandle As qbVariable.qbVariable
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Clearing the variable type collection", 1)
            intIndex2 = 1
            For intIndex1 = .Count To 1 Step -1
                updateProgress("Clearing the variable type collection", _
                               "test object", _
                               intIndex2, _
                               intCount)
                _OBJwindowsUtilities.updateStatusListBox(lstStatus, "", 1)
                objHandle = CType(.Item(intIndex1), qbVariable.qbVariable)
                If Not disposeVariable(objHandle) Then Return (False)
                objHandle = Nothing
                .Remove(intIndex1)
                _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "")
                intIndex2 += 1
            Next intIndex1
            lblProgress.Visible = False
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Finished clearing the variable type collection")
            Return (True)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Clone the variable  
    '
    '
    Private Function cloneVariable() As Boolean
        Dim objClone As qbVariable.qbVariable
        Try
            objClone = OBJqbVariable.clone
        Catch
            _OBJutilities.errorHandler("Cannot create clone: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "cloneVariableType", _
                                       "Terminating this application")
            closer(True)
        End Try
        Return (addTestObject(objClone))
    End Function

    ' -----------------------------------------------------------------
    ' Terminate the application
    '
    '
    Private Sub closer(ByVal booFast As Boolean)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Closing", 1)
        objectCleanup()
        If Not booFast Then form2Registry()
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Close complete")
        End
    End Sub

    ' -----------------------------------------------------------------
    ' Create the variable, disposing of the old object as needed:
    ' change display
    '
    '
    Private Overloads Function createVariable() As Boolean
        Return createVariable(True)
    End Function
    Private Overloads Function createVariable(ByVal booInspect As Boolean) As Boolean
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Creation of the variable type " & _
                            txtVariableExpression.Text, _
                            1)
        If Not (OBJqbVariable Is Nothing) Then destroyVariable()
        Try
            OBJqbVariable = New qbVariable.qbVariable(txtVariableExpression.Text)
        Catch
            msgboxInterface("Cannot create variable: " & Err.Number & " " & Err.Description)
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Creation of vt failed")
            Return False
        End Try
        If (OBJqbVariable Is Nothing) OrElse Not OBJqbVariable.Usable Then
            With _OBJutilities
                msgboxInterface("Did not create variable: " & _
                        "there is probably a syntax error in the fromString " & _
                        "expression " & _
                        .enquote(.ellipsis(txtVariableExpression.Text, 32)) & ": " & _
                        "variable is nothing or not usable")
            End With
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Creation of vt failed")
            Return False
        End If
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Creation of vt complete")
        changeVTcontrols(True)
        txtVariableExpression.Text = OBJqbVariable.ToString
        If booInspect Then inspectInterface()
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Destroy the vt object and change the display
    '
    '
    Private Sub destroyVariable()
        disposeVariable(OBJqbVariable)
        OBJqbVariable = Nothing
        changeVTcontrols(False)
    End Sub

    ' -----------------------------------------------------------------
    ' Dispose the variable  
    '
    '
    Private Function disposeVariable(ByVal objVariable As qbVariable.qbVariable) _
            As Boolean
        If (objVariable Is Nothing) OrElse Not objVariable.Usable Then Return False
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Disposing the variable " & _
                            _OBJutilities.enquote(objVariable.Name) & " " & _
                            "(" & toStringInterface(objVariable) & ")", _
                            1)
        Dim booSave As Boolean = BOOnoEvents : BOOnoEvents = True
        objVariable.disposeInspect()
        BOOnoEvents = booSave
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Disposal complete")
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Save form settings
    '
    '
    Private Sub form2Registry()
        Try
            With txtVariableExpression
                SaveSetting(Application.ProductName, Me.Name, .Name, .Text)
            End With
        Catch
            _OBJutilities.errorHandler("Cannot save form selections: " & _
                                      Err.Number & " " & Err.Description, _
                                      Me.Name, _
                                      "form2Registry", _
                                      "Continuing execution")
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Prompt for containment test operands
    '
    '
    ' Note that the Tag of each object is set to True when the object
    ' was created by this routine, False otherwise.
    '
    '
    Private Function getContainmentTesterOperands _
                     (ByRef objOperand1 As qbVariable.qbVariable, _
                      ByRef objOperand2 As qbVariable.qbVariable) _
        As Boolean
        Dim strOperand1 As String
        Dim strOperand2 As String
        objOperand1 = Nothing : objOperand2 = Nothing
        If Not (OBJqbVariable Is Nothing) AndAlso OBJqbVariable.Usable Then
            objOperand1 = OBJqbVariable
            strOperand1 = OBJqbVariable.ToString
        Else
            strOperand1 = OBJqbVariable.mkRandomVariable
        End If
        Dim booHaveOperand2 As Boolean
        Dim intIndex1 As Integer = lstVariables.SelectedIndex + 1
        If intIndex1 > 0 Then
            objOperand2 = CType(COLtestObjects.Item(intIndex1), _
                                qbVariable.qbVariable)
            booHaveOperand2 = objOperand2.Usable
        Else
            strOperand2 = OBJqbVariable.Dope.mkRandomType
        End If
        Dim strNext As String = strOperand1 & ":" & strOperand2
        Dim strOperands() As String
        Do
            strNext = InputBox("Enter containment test operands separated by colon", _
                               "Containment Test", _
                               strNext)
            If Trim(strNext) = "" Then Return (False)
            strOperands = Split(strNext, ":")
        Loop Until UBound(strOperands) = 1
        If UCase(Trim(strOperands(0))) <> UCase(Trim(strOperand1)) OrElse (objOperand1 Is Nothing) Then
            Try
                objOperand1 = New qbVariable.qbVariable
                objOperand1.fromString(strOperands(0))
                objOperand1.Tag = True
            Catch
                _OBJutilities.errorHandler("Can't create qbVariableType 1: " & _
                                           Err.Number & " " & Err.Description, _
                                           Me.Name, _
                                           "", _
                                           "Returning False to cancel test")
                Return (False)
            End Try
        End If
        If UCase(Trim(strOperands(1))) <> UCase(Trim(strOperand2)) OrElse Not booHaveOperand2 Then
            Try
                objOperand2 = New qbVariable.qbVariable
                objOperand2.fromString(strOperands(1))
                objOperand2.Tag = True
            Catch
                _OBJutilities.errorHandler("Can't create qbVariableType 2: " & _
                                           Err.Number & " " & Err.Description, _
                                           Me.Name, _
                                           "", _
                                           "Returning False to cancel test")
                Return (False)
            End Try
        End If
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Inspect the test object
    '
    '
    Private Sub inspectInterface()
        Dim strReport As String
        Select Case msgboxInterface("Inspection of object has " & _
                           CStr(IIf(OBJqbVariable.inspect(strReport), _
                                    "succeeded", "failed")) & ": " & _
                           "Click OK to view report: click Cancel to return to main form", _
                           MsgBoxStyle.OKCancel)
            Case MsgBoxResult.OK : zoomInterface(strReport)
            Case MsgBoxResult.Cancel
            Case Else
                _OBJutilities.errorHandler("Unexpected Select Case", _
                                          Me.Name, _
                                          "inspectInterface", _
                                          "Terminating")
                closer(True)
        End Select
    End Sub

    ' -----------------------------------------------------------------
    ' Make one random type
    '
    '
    Private Sub mkRandomType()
        txtVariableExpression.Text = _OBJqbVariableType.mkRandomType
    End Sub

    ' -----------------------------------------------------------------
    ' Make several random types
    '
    '
    ' --- Prompts for the count
    Private Overloads Sub mkRandomTypes()
        Dim booOK As Boolean
        Dim intCount As Integer = 10
        Dim strCount As String
        Do
            strCount = Trim(InputBox("Enter the number of objects to be created", _
                            "mkrandomTypes", _
                            CStr(intCount)))
            If strCount = "" Then Return
            Try
                intCount = CInt(strCount)
                booOK = True
            Catch : End Try
        Loop Until booOK
        If booOK Then mkRandomTypes(intCount)
    End Sub
    ' --- Accepts the count
    Private Overloads Sub mkRandomTypes(ByVal intCount As Integer)
        Dim booOK As Boolean
        Dim intIndex1 As Integer
        Dim intErr As Integer
        Dim objNew As qbVariable.qbVariable
        Dim strErr As String
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Creating random types", 1)
        For intIndex1 = 1 To intCount
            updateProgress("Creating random types", _
                           "type", _
                           intIndex1, _
                           intCount)
            booOK = False
            Try
                objNew = New qbVariable.qbVariable
                booOK = objNew.fromString(_OBJqbVariableType.mkRandomType)
            Catch
                intErr = Err.Number : strErr = Err.Description
            End Try
            If Not booOK Then
                _OBJutilities.errorHandler("Cannot create test object " & _
                                           intIndex1 & " of " & intCount & _
                                           CStr(IIf(strErr <> "", _
                                                    ": " & intErr & " " & strErr, _
                                                    "")), _
                                           "Terminating test object creation")
                Return
            End If
            addTestObject(objNew)
        Next intIndex1
        lblProgress.Visible = False
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Created test objects")
        With lstVariables
            .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
            .SelectedIndex = -1
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Display and log all messages
    '
    '
    ' --- Default style
    Private Overloads Function msgboxInterface(ByVal strMessage As String) _
            As Microsoft.VisualBasic.MsgBoxResult
        Return msgboxInterface(strMessage, _
                               Microsoft.VisualBasic.MsgBoxStyle.Information)
    End Function
    ' --- Default style
    Private Overloads Function msgboxInterface(ByVal strMessage As String, _
                                               ByVal objStyle _
                                               As Microsoft.VisualBasic.MsgBoxStyle) _
            As Microsoft.VisualBasic.MsgBoxResult
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, strMessage)
        Return MsgBox(strMessage, objStyle)
    End Function

    ' -----------------------------------------------------------------
    ' Dispose of all objects
    '
    '
    Private Function objectCleanup() As Boolean
        If Not (OBJqbVariable Is Nothing) Then
            OBJqbVariable.disposeInspect()
        End If
        Return (clearVTcollection())
    End Function

    ' -----------------------------------------------------------------
    ' Restore form settings
    '
    '
    Private Sub registry2Form()
        Try
            With txtVariableExpression
                .Text = GetSetting(Application.ProductName, Me.Name, .Name, .Text)
            End With
        Catch
            _OBJutilities.errorHandler("Cannot get Registry settings: " & _
                                      Err.Number & " " & Err.Description, _
                                      Me.Name, _
                                      "registry2Form", _
                                      "Continuing")
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Create tool tips
    '
    '
    Private Function setToolTips() As Boolean
        Try
            OBJtoolTip = New ToolTip
            With OBJtoolTip
                .SetToolTip(cmdAbout, _
                            "Display information about this form and the " & _
                            "qbVariable class")
                .SetToolTip(cmdClear, _
                            "Clears the variable or array to default values")
                .SetToolTip(cmdClearSettings, _
                            "Clears form selections as saved in the Registry " & _
                            "to default values")
                .SetToolTip(cmdCloneVariable, _
                            "Makes a copy of the object")
                .SetToolTip(cmdClose, _
                            "Saves form selections in the Registry and " & _
                            "dismisses the application")
                .SetToolTip(cmdCloseNosave, _
                            "Dismisses the application: doesn't save selections")
                .SetToolTip(cmdCreateVariable, _
                            "Creates the qbVariable object based on the " & _
                            "variable expression in the black and white text " & _
                            "box at the top of the screen")
                .SetToolTip(cmdDisposeVariable, _
                            "Disposes of the qbVariable object: will inspect " & _
                            "for correctness")
                .SetToolTip(cmdEmpiricalDope, _
                            "This is not one of the Three Stooges: instead, it is " & _
                            "the result of examining the value of the " & _
                            "value and determining its best representation from " & _
                            "this value, alone.")
                .SetToolTip(cmdForm2Registry, _
                            "Save form selections in the Registry such that " & _
                            "they can be used in the next activation of this " & _
                            "application")
                .SetToolTip(cmdInspect, _
                            "Inspects the state of the test qbVariable for internal " & _
                            "errors, and displays a report in a scrollable, read-only " & _
                            "text box")
                .SetToolTip(cmdMkRandomVariable, _
                            "Creates a random variable expression for your testing")
                .SetToolTip(cmdObject2XML, _
                            "Converts the object state to an eXtended Markup Language tag")
                .SetToolTip(cmdRandomize, _
                            "Randomizes such that the next use of the Test button, " & _
                            "to stress-test the qbVariable class, will use " & _
                            "unpredictable values")
                .SetToolTip(cmdRegistry2Form, _
                            "Resets form selections to values saved in the Registry by " & _
                            "the ""Save Settings"" button or Closing the form")
                .SetToolTip(cmdReset, _
                            "Resets the qbVariable to the Unknown variable")
                .SetToolTip(cmdStatusZoom, _
                            "Shows the status and progress information in the List box " & _
                            "below in a read-only, scrollable text box that can be copied " & _
                            "to Notepad or Word")
                .SetToolTip(cmdTest, _
                            "Runs a comprehensive stress test of the qbVariable object. " & _
                            "Note that the test cases will be random but determined " & _
                            "unless you have clicked the Randomize button; in this " & _
                            "case, the test cases will be fully random.")
                .SetToolTip(cmdToDescription, _
                            "Describes the underlying variable type of the variable")
                .SetToolTip(cmdToString, _
                            "Converts the variable object back to its fromString expression")
                .SetToolTip(cmdValue, _
                            "Shows the scalar value of this variable")
                .SetToolTip(cmdVariableClear, _
                            "Clears the randomly created variables in the list box " & _
                            "below this button")
                .SetToolTip(cmdVariablesCreate, _
                            "Creates 0..n random variables: you are prompted for the exact number")
                .SetToolTip(lstStatus, _
                            "Shows status and progress information")
                .SetToolTip(lstVariables, _
                            "Identifies random variables created using the Create Random Variable " & _
                            "button that is above this list box")
                .SetToolTip(txtVariableExpression, _
                            "Enter the fromString variable expression in the overall form " & _
                            "type:value, where ""type"" is a qbVariableType fromString " & _
                            "expression, and ""value"" is a single numeric or string value, " & _
                            "a ""decorated"" value such as integer(0), or a list of such " & _
                            "values when ""type"" is an array or a User Data Type." & _
                            vbNewLine & vbNewLine & _
                            "For example, enter Array,Byte,1,2:0,255 to define the two-element " & _
                            "one-dimensional array of Bytes containing 0 and 255.")
            End With
        Catch
            Return False
        End Try
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Show the scalar value of the test variable
    '
    '
    Private Sub showValue()
        With OBJqbVariable
            msgboxInterface("The .Net value of the variable with object name " & _
                            "[" & .Name & "] " & _
                            "and variable name " & _
                            _OBJutilities.enquote(.VariableName) & " is " & _
                            _OBJutilities.object2String(.value, True))
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Test interface
    '
    '
    Private Sub testInterface()
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "Testing the qbVariable class", _
                                                 1)
        Dim strReport As String
        If (OBJqbVariable Is Nothing) Then
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     "Creating the variable object for the test", 1)
            If Not createVariable(False) Then
                msgboxInterface("Not able to create test variable: the fromString for this variable, " & _
                       _OBJutilities.enquote(txtVariableExpression.Text) & ", " & _
                       "may have a syntax error. Retrying with the Unknown variable: " & _
                       "note that the fromString expression in the text box has been erased.")
                txtVariableExpression.Text = "Unknown"
                If Not createVariable(False) Then
                    msgboxInterface("Can't create the Unknown variable. Contact support.")
                    Return
                End If
                _OBJwindowsUtilities.updateStatusListBox(lstStatus, "", -1)
                Return
            End If
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     -1, "Created the variable object for the test")
        End If
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "Commencing the test method", 1)
        Dim booOK As Boolean = OBJqbVariable.test(strReport, False)
        Dim strMessage As String = "Test of object has " & _
                                    CStr(IIf(booOK, "succeeded", "failed"))
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, strMessage)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -2, "Test method complete")
        Select Case msgboxInterface(strMessage & ": " & _
                                    "Click OK to view report: click Cancel to return to main form", _
                                    MsgBoxStyle.OKCancel)
            Case MsgBoxResult.OK : zoomInterface(strReport)
            Case MsgBoxResult.Cancel
            Case Else
                _OBJutilities.errorHandler("Unexpected Select Case", _
                                          Me.Name, _
                                          "inspectInterface", _
                                          "Terminating")
                closer(True)
        End Select
        lblProgress.Visible = False
    End Sub

    ' -----------------------------------------------------------------
    ' Interface to toString where object may not be usable
    '
    '
    Private Function toStringInterface(ByVal objHandle As qbVariable.qbVariable) _
            As String
        With objHandle
            If Not .Usable Then Return ("unusable")
            Return .ToString
        End With
    End Function

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

    ' -------------------------------------------------------------------
    ' Assign a scalar to the value
    '
    '
    Private Sub valueAssignment()
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "Obtaining and assigning a scalar value " & _
                                                 "to the qbVariable", _
                                                 1)
        With FRMmodifyVariable
            .Variable = OBJqbVariable
            Try
                .ShowDialog()
            Catch
                MsgBox("Cannot assign value: " & _
                       Err.Number & " " & Err.Description)
                Return
            End Try
            If Not .Change Then Return
            txtVariableExpression.Text = OBJqbVariable.ToString
            txtVariableExpression.Refresh()
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     -1, _
                                                     "Assignment succeeded: value is " & _
                                                     txtVariableExpression.Text)
        End With
    End Sub

    ' -------------------------------------------------------------------
    ' Interface to the zoom object
    '
    '
    Private Overloads Sub zoomInterface(ByVal objZoomed As Object)
        zoomInterface(objZoomed, CDbl(3.25), CDbl(2))
    End Sub
    Private Overloads Sub zoomInterface(ByVal objZoomed As Object, _
                                        ByVal dblWidthChange As Double, _
                                        ByVal dblHeightChange As Double)
        Dim objZoom As zoom.zoom
        Try
            objZoom = New zoom.zoom
        Catch
        End Try
        With objZoom
            .ZoomTextBox.Font = New Font(.ZoomTextBox.Font, FontStyle.Bold)
            .setSize(CDbl(dblWidthChange), CDbl(dblHeightChange))
            If (TypeOf objZoomed Is String) Then
                .ZoomTextBox.Text = CStr(objZoomed)
            ElseIf (TypeOf objZoomed Is Control) Then
                .Control = CType(objZoomed, Control)
            Else
                _OBJutilities.errorHandler("objZoomed has unsupported type", _
                                          Me.Name, "zoomInterface", _
                                          "Terminating application")
                closer(True)
            End If
            .showZoom()
        End With
    End Sub

#End Region

#Region " Event handlers "

    Private Sub testEventHandler(ByVal strDesc As String, _
                                 ByVal intLevelChange As Integer) Handles OBJqbVariable.msgEvent
        If BOOnoEvents Then Return
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, strDesc, intLevelChange)
    End Sub

    Private Sub testProgressHandler(ByVal strActivity As String, _
                                    ByVal strEntity As String, _
                                    ByVal intEntityNumber As Integer, _
                                    ByVal intEntityCount As Integer, _
                                    ByVal intLevel As Integer, _
                                    ByVal strComments As String) Handles OBJqbVariable.progressEvent
        If BOOnoEvents Then Return
        updateProgress(strActivity, strEntity, intEntityNumber, intEntityCount)
    End Sub

    Private Sub testProgressSharedHandler(ByVal strActivity As String, _
                                            ByVal strEntity As String, _
                                            ByVal intEntityNumber As Integer, _
                                            ByVal intEntityCount As Integer, _
                                            ByVal intLevel As Integer, _
                                            ByVal strComments As String) Handles OBJqbVariable.progressEventShared
        If BOOnoEvents Then Return
        updateProgress(strActivity, strEntity, intEntityNumber, intEntityCount)
    End Sub

    Private Sub testProgressHandlerShared(ByVal strActivity As String, _
                                            ByVal strEntity As String, _
                                            ByVal intEntityNumber As Integer, _
                                            ByVal intEntityCount As Integer, _
                                            ByVal intLevel As Integer, _
                                            ByVal strComments As String) Handles OBJqbVariable.progressEventShared
        If BOOnoEvents Then Return
        updateProgress(strActivity, strEntity, intEntityNumber, intEntityCount)
    End Sub

#End Region ' Event handlers 

End Class
