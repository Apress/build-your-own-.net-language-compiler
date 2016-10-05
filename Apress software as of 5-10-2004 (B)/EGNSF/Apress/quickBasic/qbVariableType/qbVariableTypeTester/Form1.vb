Option Strict

' *********************************************************************
' *                                                                   *
' * Variable type test form                                           *
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
    Friend WithEvents lblVariableTypeExpression As System.Windows.Forms.Label
    Friend WithEvents txtVariableTypeExpression As System.Windows.Forms.TextBox
    Friend WithEvents cmdCreateVariableType As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdInspect As System.Windows.Forms.Button
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents cmdCloneVariableType As System.Windows.Forms.Button
    Friend WithEvents cmdObject2XML As System.Windows.Forms.Button
    Friend WithEvents cmdDisposeVariableType As System.Windows.Forms.Button
    Friend WithEvents lblVariableTypes As System.Windows.Forms.Label
    Friend WithEvents lstVariableTypes As System.Windows.Forms.ListBox
    Friend WithEvents cmdCompareToType As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteVariableType As System.Windows.Forms.Button
    Friend WithEvents cmdVariableTypesCreate As System.Windows.Forms.Button
    Friend WithEvents cmdVariableTypeContainment As System.Windows.Forms.Button
    Friend WithEvents chkContainedTypeWithState As System.Windows.Forms.CheckBox
    Friend WithEvents cmdVariableTypesClear As System.Windows.Forms.Button
    Friend WithEvents cmdMkRandomType As System.Windows.Forms.Button
    Friend WithEvents cmdRandomize As System.Windows.Forms.Button
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents cmdStatusZoom As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents cmdContainmentTester As System.Windows.Forms.Button
    Friend WithEvents cmdClearSettings As System.Windows.Forms.Button
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    Friend WithEvents cmdToDescription As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents cmdToString As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblVariableTypeExpression = New System.Windows.Forms.Label
Me.txtVariableTypeExpression = New System.Windows.Forms.TextBox
Me.cmdCreateVariableType = New System.Windows.Forms.Button
Me.cmdCloneVariableType = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.lblVariableTypes = New System.Windows.Forms.Label
Me.lstVariableTypes = New System.Windows.Forms.ListBox
Me.cmdCompareToType = New System.Windows.Forms.Button
Me.cmdDeleteVariableType = New System.Windows.Forms.Button
Me.cmdInspect = New System.Windows.Forms.Button
Me.cmdTest = New System.Windows.Forms.Button
Me.cmdObject2XML = New System.Windows.Forms.Button
Me.cmdDisposeVariableType = New System.Windows.Forms.Button
Me.cmdVariableTypeContainment = New System.Windows.Forms.Button
Me.cmdVariableTypesCreate = New System.Windows.Forms.Button
Me.chkContainedTypeWithState = New System.Windows.Forms.CheckBox
Me.cmdVariableTypesClear = New System.Windows.Forms.Button
Me.cmdMkRandomType = New System.Windows.Forms.Button
Me.cmdRandomize = New System.Windows.Forms.Button
Me.lblStatus = New System.Windows.Forms.Label
Me.lstStatus = New System.Windows.Forms.ListBox
Me.cmdStatusZoom = New System.Windows.Forms.Button
Me.lblProgress = New System.Windows.Forms.Label
Me.cmdContainmentTester = New System.Windows.Forms.Button
Me.cmdClearSettings = New System.Windows.Forms.Button
Me.cmdAbout = New System.Windows.Forms.Button
Me.cmdToDescription = New System.Windows.Forms.Button
Me.Button2 = New System.Windows.Forms.Button
Me.Button3 = New System.Windows.Forms.Button
Me.cmdToString = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblVariableTypeExpression
'
Me.lblVariableTypeExpression.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblVariableTypeExpression.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariableTypeExpression.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblVariableTypeExpression.Location = New System.Drawing.Point(8, 9)
Me.lblVariableTypeExpression.Name = "lblVariableTypeExpression"
Me.lblVariableTypeExpression.Size = New System.Drawing.Size(808, 19)
Me.lblVariableTypeExpression.TabIndex = 0
Me.lblVariableTypeExpression.Text = "Variable Type Expression: boldface if variable type exists: normal face otherwise" & _
""
Me.lblVariableTypeExpression.TextAlign = System.Drawing.ContentAlignment.TopCenter
'
'txtVariableTypeExpression
'
Me.txtVariableTypeExpression.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtVariableTypeExpression.Location = New System.Drawing.Point(8, 28)
Me.txtVariableTypeExpression.Name = "txtVariableTypeExpression"
Me.txtVariableTypeExpression.Size = New System.Drawing.Size(808, 24)
Me.txtVariableTypeExpression.TabIndex = 1
Me.txtVariableTypeExpression.Text = "Unknown"
'
'cmdCreateVariableType
'
Me.cmdCreateVariableType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCreateVariableType.Location = New System.Drawing.Point(8, 55)
Me.cmdCreateVariableType.Name = "cmdCreateVariableType"
Me.cmdCreateVariableType.Size = New System.Drawing.Size(152, 25)
Me.cmdCreateVariableType.TabIndex = 2
Me.cmdCreateVariableType.Text = "Create Variable Type"
'
'cmdCloneVariableType
'
Me.cmdCloneVariableType.Enabled = False
Me.cmdCloneVariableType.Location = New System.Drawing.Point(168, 55)
Me.cmdCloneVariableType.Name = "cmdCloneVariableType"
Me.cmdCloneVariableType.Size = New System.Drawing.Size(152, 25)
Me.cmdCloneVariableType.TabIndex = 3
Me.cmdCloneVariableType.Text = "Clone Variable Type"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(8, 632)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(125, 23)
Me.cmdForm2Registry.TabIndex = 4
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(144, 632)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(125, 23)
Me.cmdRegistry2Form.TabIndex = 5
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(664, 120)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(152, 108)
Me.cmdClose.TabIndex = 6
Me.cmdClose.Text = "Close"
'
'lblVariableTypes
'
Me.lblVariableTypes.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblVariableTypes.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariableTypes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblVariableTypes.Location = New System.Drawing.Point(8, 240)
Me.lblVariableTypes.Name = "lblVariableTypes"
Me.lblVariableTypes.Size = New System.Drawing.Size(808, 23)
Me.lblVariableTypes.TabIndex = 7
Me.lblVariableTypes.Text = "Variable types: click to select: double click to set test object"
Me.lblVariableTypes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
'
'lstVariableTypes
'
Me.lstVariableTypes.BackColor = System.Drawing.SystemColors.Control
Me.lstVariableTypes.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstVariableTypes.ItemHeight = 17
Me.lstVariableTypes.Location = New System.Drawing.Point(8, 264)
Me.lstVariableTypes.Name = "lstVariableTypes"
Me.lstVariableTypes.Size = New System.Drawing.Size(808, 174)
Me.lstVariableTypes.TabIndex = 8
'
'cmdCompareToType
'
Me.cmdCompareToType.Location = New System.Drawing.Point(8, 120)
Me.cmdCompareToType.Name = "cmdCompareToType"
Me.cmdCompareToType.Size = New System.Drawing.Size(316, 48)
Me.cmdCompareToType.TabIndex = 9
Me.cmdCompareToType.Text = "Compare to variable type %VTNAME"
Me.cmdCompareToType.Visible = False
'
'cmdDeleteVariableType
'
Me.cmdDeleteVariableType.Location = New System.Drawing.Point(335, 120)
Me.cmdDeleteVariableType.Name = "cmdDeleteVariableType"
Me.cmdDeleteVariableType.Size = New System.Drawing.Size(321, 48)
Me.cmdDeleteVariableType.TabIndex = 10
Me.cmdDeleteVariableType.Text = "Delete variable type %VTNAME"
Me.cmdDeleteVariableType.Visible = False
'
'cmdInspect
'
Me.cmdInspect.Enabled = False
Me.cmdInspect.Location = New System.Drawing.Point(8, 88)
Me.cmdInspect.Name = "cmdInspect"
Me.cmdInspect.Size = New System.Drawing.Size(152, 24)
Me.cmdInspect.TabIndex = 11
Me.cmdInspect.Text = "Inspect"
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(168, 88)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.Size = New System.Drawing.Size(152, 24)
Me.cmdTest.TabIndex = 12
Me.cmdTest.Text = "Stress Test"
'
'cmdObject2XML
'
Me.cmdObject2XML.Enabled = False
Me.cmdObject2XML.Location = New System.Drawing.Point(328, 88)
Me.cmdObject2XML.Name = "cmdObject2XML"
Me.cmdObject2XML.Size = New System.Drawing.Size(152, 24)
Me.cmdObject2XML.TabIndex = 13
Me.cmdObject2XML.Text = "Object to XML"
'
'cmdDisposeVariableType
'
Me.cmdDisposeVariableType.Enabled = False
Me.cmdDisposeVariableType.Location = New System.Drawing.Point(326, 55)
Me.cmdDisposeVariableType.Name = "cmdDisposeVariableType"
Me.cmdDisposeVariableType.Size = New System.Drawing.Size(153, 25)
Me.cmdDisposeVariableType.TabIndex = 14
Me.cmdDisposeVariableType.Text = "Dispose "
'
'cmdVariableTypeContainment
'
Me.cmdVariableTypeContainment.Location = New System.Drawing.Point(6, 175)
Me.cmdVariableTypeContainment.Name = "cmdVariableTypeContainment"
Me.cmdVariableTypeContainment.Size = New System.Drawing.Size(319, 48)
Me.cmdVariableTypeContainment.TabIndex = 15
Me.cmdVariableTypeContainment.Text = "Test for containment in %VTNAME"
Me.cmdVariableTypeContainment.Visible = False
'
'cmdVariableTypesCreate
'
Me.cmdVariableTypesCreate.Location = New System.Drawing.Point(648, 240)
Me.cmdVariableTypesCreate.Name = "cmdVariableTypesCreate"
Me.cmdVariableTypesCreate.Size = New System.Drawing.Size(168, 24)
Me.cmdVariableTypesCreate.TabIndex = 16
Me.cmdVariableTypesCreate.Text = "Create Random Types"
'
'chkContainedTypeWithState
'
Me.chkContainedTypeWithState.Location = New System.Drawing.Point(336, 177)
Me.chkContainedTypeWithState.Name = "chkContainedTypeWithState"
Me.chkContainedTypeWithState.Size = New System.Drawing.Size(312, 47)
Me.chkContainedTypeWithState.TabIndex = 17
Me.chkContainedTypeWithState.Text = "containedTypeWithState"
Me.chkContainedTypeWithState.Visible = False
'
'cmdVariableTypesClear
'
Me.cmdVariableTypesClear.Location = New System.Drawing.Point(552, 240)
Me.cmdVariableTypesClear.Name = "cmdVariableTypesClear"
Me.cmdVariableTypesClear.Size = New System.Drawing.Size(96, 24)
Me.cmdVariableTypesClear.TabIndex = 18
Me.cmdVariableTypesClear.Text = "Clear (all)"
'
'cmdMkRandomType
'
Me.cmdMkRandomType.Location = New System.Drawing.Point(662, 55)
Me.cmdMkRandomType.Name = "cmdMkRandomType"
Me.cmdMkRandomType.Size = New System.Drawing.Size(153, 23)
Me.cmdMkRandomType.TabIndex = 19
Me.cmdMkRandomType.Text = "Random Type"
'
'cmdRandomize
'
Me.cmdRandomize.Location = New System.Drawing.Point(664, 88)
Me.cmdRandomize.Name = "cmdRandomize"
Me.cmdRandomize.Size = New System.Drawing.Size(152, 24)
Me.cmdRandomize.TabIndex = 20
Me.cmdRandomize.Text = "Randomize"
'
'lblStatus
'
Me.lblStatus.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblStatus.Location = New System.Drawing.Point(8, 440)
Me.lblStatus.Name = "lblStatus"
Me.lblStatus.Size = New System.Drawing.Size(808, 23)
Me.lblStatus.TabIndex = 21
Me.lblStatus.Text = "Status"
Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.SystemColors.Control
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 17
Me.lstStatus.Location = New System.Drawing.Point(8, 464)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(808, 140)
Me.lstStatus.TabIndex = 22
'
'cmdStatusZoom
'
Me.cmdStatusZoom.Location = New System.Drawing.Point(720, 440)
Me.cmdStatusZoom.Name = "cmdStatusZoom"
Me.cmdStatusZoom.Size = New System.Drawing.Size(96, 24)
Me.cmdStatusZoom.TabIndex = 23
Me.cmdStatusZoom.Text = "Zoom"
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblProgress.Location = New System.Drawing.Point(8, 608)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(808, 16)
Me.lblProgress.TabIndex = 24
Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.TopCenter
Me.lblProgress.Visible = False
'
'cmdContainmentTester
'
Me.cmdContainmentTester.Location = New System.Drawing.Point(664, 632)
Me.cmdContainmentTester.Name = "cmdContainmentTester"
Me.cmdContainmentTester.Size = New System.Drawing.Size(144, 23)
Me.cmdContainmentTester.TabIndex = 25
Me.cmdContainmentTester.Text = "Containment tester"
'
'cmdClearSettings
'
Me.cmdClearSettings.Location = New System.Drawing.Point(280, 632)
Me.cmdClearSettings.Name = "cmdClearSettings"
Me.cmdClearSettings.Size = New System.Drawing.Size(124, 23)
Me.cmdClearSettings.TabIndex = 26
Me.cmdClearSettings.Text = "Clear Settings"
'
'cmdAbout
'
Me.cmdAbout.Location = New System.Drawing.Point(432, 632)
Me.cmdAbout.Name = "cmdAbout"
Me.cmdAbout.Size = New System.Drawing.Size(125, 23)
Me.cmdAbout.TabIndex = 27
Me.cmdAbout.Text = "About"
'
'cmdToDescription
'
Me.cmdToDescription.Enabled = False
Me.cmdToDescription.Location = New System.Drawing.Point(487, 55)
Me.cmdToDescription.Name = "cmdToDescription"
Me.cmdToDescription.Size = New System.Drawing.Size(151, 23)
Me.cmdToDescription.TabIndex = 28
Me.cmdToDescription.Text = "Describe"
'
'Button2
'
Me.Button2.Location = New System.Drawing.Point(552, 240)
Me.Button2.Name = "Button2"
Me.Button2.Size = New System.Drawing.Size(96, 24)
Me.Button2.TabIndex = 18
Me.Button2.Text = "Clear (all)"
'
'Button3
'
Me.Button3.Location = New System.Drawing.Point(648, 240)
Me.Button3.Name = "Button3"
Me.Button3.Size = New System.Drawing.Size(168, 24)
Me.Button3.TabIndex = 16
Me.Button3.Text = "Create Random Types"
'
'cmdToString
'
Me.cmdToString.Enabled = False
Me.cmdToString.Location = New System.Drawing.Point(488, 88)
Me.cmdToString.Name = "cmdToString"
Me.cmdToString.Size = New System.Drawing.Size(152, 24)
Me.cmdToString.TabIndex = 29
Me.cmdToString.Text = "Convert to String"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
Me.ClientSize = New System.Drawing.Size(820, 662)
Me.Controls.Add(Me.cmdToString)
Me.Controls.Add(Me.cmdToDescription)
Me.Controls.Add(Me.cmdAbout)
Me.Controls.Add(Me.cmdClearSettings)
Me.Controls.Add(Me.cmdContainmentTester)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.cmdStatusZoom)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.lblStatus)
Me.Controls.Add(Me.cmdRandomize)
Me.Controls.Add(Me.cmdMkRandomType)
Me.Controls.Add(Me.cmdVariableTypesClear)
Me.Controls.Add(Me.chkContainedTypeWithState)
Me.Controls.Add(Me.cmdVariableTypesCreate)
Me.Controls.Add(Me.cmdVariableTypeContainment)
Me.Controls.Add(Me.cmdDisposeVariableType)
Me.Controls.Add(Me.cmdObject2XML)
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.cmdInspect)
Me.Controls.Add(Me.cmdDeleteVariableType)
Me.Controls.Add(Me.cmdCompareToType)
Me.Controls.Add(Me.lstVariableTypes)
Me.Controls.Add(Me.lblVariableTypes)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.cmdCloneVariableType)
Me.Controls.Add(Me.cmdCreateVariableType)
Me.Controls.Add(Me.txtVariableTypeExpression)
Me.Controls.Add(Me.lblVariableTypeExpression)
Me.Controls.Add(Me.Button2)
Me.Controls.Add(Me.Button3)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "qbVariableType"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private OBJtoolTip As ToolTip

    Private Shared _OBJutilities As utilities.utilities
    Private Shared _OBJwindowsUtilities As windowsUtilities.windowsUtilities
    Private WithEvents OBJqbVariableType As qbVariableType.qbVariableType
    Private COLtestObjects As Collection             ' Of qbVariableType

    ' --- About information Easter egg
    Private Const ABOUTINFO As String = _
        "qbVariableTypeTester" & _
        vbNewLine & vbNewLine & vbNewLine & _
        "This form and application tests the qbVariableType object, designed as the OO " & _
        "model for Quick Basic variable types." & vbNewLine & vbNewLine & _
        "This form and application was developed commencing on 4/5/2003 by" & vbNewLine & vbNewLine & _
        "Edward G. Nilges" & vbNewLine & _
        "spinoza1111@yahoo.COM" & vbNewLine & _
        "http://members.screenz.com/edNilges"
    Private Const ABOUTINFO_EXTENSION As String = _
        "To try this form and application out, type a simple variable type name such as " & _
        """integer"" in the text box at the top of the form. Then, click " & _
        "the button labeled ""Create variable type"" near the upper left hand " & _
        "corner of the form to see that an object is created." & vbNewLine & vbNewLine & _
        "Then, click the Inspect and Object to XML buttons to see how the object " & _
        "self-inspects for its own sanity (and yours) and converts its state to " & _
        "eXtended Markup Language." & vbNewLine & vbNewLine & _
        "Then, click the Dispose button to get rid of the object, and " & _
        "enter a more complex variable type such as Variant,String or " & _
        "Array,Integer,1,10, or even Variant,(Array,Integer,1,10)." & vbNewLine & vbNewLine & _
        "Inspect these examples, use the Object to XML button to see their state, " & _
        "and dispose each one in turn." & vbNewLine & vbNewLine & _
        "The Describe button provides the English description of a data type. " & _
        "Dispose of the existing data type, enter ""Array,Long,1,2,1,3"", " & _
        "click the Create variable type button, and click Describe to try this out." & vbNewLine & vbNewLine & _
        "Finally, the ""Stress Test"" button will stress-test this object. Click Test, " & _
        "sit back, and watch the progress of testing in the status box. " & _
        "This will provide some assurance if you've changed the source code."

    ' ***** Constants *****
    Private SCREEN_WIDTH_TOLERANCE As Single = 0.9
    Private SCREEN_HEIGHT_TOLERANCE As Single = 0.9

#End Region

#Region " Form events "

    Private Sub cmdAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbout.Click
        about(False)
    End Sub

    Private Sub cmdClearSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearSettings.Click
        Select Case MsgBox("Do you want to clear the Registry settings (defaults) " & _
                           "of this form and application?" & vbNewline & vbNewline & _
                           "Note that this will restore default settings.", _
                           MsgBoxStyle.YesNo)
            Case MsgBoxResult.Yes: 
                If clearRegistry Then registry2Form
            Case MsgBoxResult.No: 
            Case Else:
                _OBJutilities.errorHandler("Unexpected reply from MsgBox", _
                                           Me.Name, "cmdClearSettings_Click", _
                                           "Continuing despite this serious error")
        End Select
    End Sub

    Private Sub cmdCloneVariableType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCloneVariableType.Click
        cloneVariableType
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        closer(False)
    End Sub

    Private Sub cmdCompareToType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCompareToType.Click
        compareToInterface(OBJqbVariableType, _
                           CType(COLtestObjects.Item(lstVariableTypes.SelectedIndex + 1), _
                                 qbVariableType.qbVariableType))
    End Sub

    Private Sub cmdContainmentTester_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdContainmentTester.Click
        Dim objTest(1) As qbVariableType.qbVariableType
        If Not getContainmentTesterOperands(objTest(0), objTest(1)) Then Return
        testContainment(objtest(0), objTest(1), True)
        If CBool(objTest(0).Tag) Then
            objTest(0).dispose: objTest(0) = Nothing
        End If
        If CBool(objTest(1).Tag) Then
            objTest(1).dispose: objTest(1) = Nothing
        End If
    End Sub

    Private Sub cmdCreateVariableType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateVariableType.Click
        createVariableType
    End Sub

    Private Sub cmdDeleteVariableType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteVariableType.Click
        deleteVariableType(lstVariableTypes.SelectedIndex)
    End Sub

    Private Sub cmdDisposeVariableType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDisposeVariableType.Click
        destroyVariableType
        changeVTcontrols(False)
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForm2Registry.Click
        form2Registry
    End Sub

    Private Sub cmdInspect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInspect.Click
        inspectInterface
    End Sub

    Private Sub cmdMkRandomType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMkRandomType.Click
        mkRandomType
    End Sub

    Private Sub cmdObject2XML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdObject2XML.Click
        zoomInterface(OBJqbVariableType.object2XML, 3, 2.35)
    End Sub

    Private Sub cmdRandomize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRandomize.Click
        Randomize
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegistry2Form.Click
        registry2Form
    End Sub

    Private Sub cmdTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTest.Click
        testInterface
    End Sub

    Private Sub cmdToDescription_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToDescription.Click
        MsgBox("The variableType " & Me.Name & " has this description: " & _
               vbNewline & vbNewline & _
               OBJqbVariableType.toDescription)
    End Sub

    Private Sub cmdToString_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToString.Click
        _OBJutilities.errorHandler("The toString value is " & _
                                   vbNewline & vbNewline & _
                                   objQBvariableType.toString, _
                                   Me.Name, _
                                   "cmdToString_Click", _
                                   "This shows all the information about the variable type " & _
                                   "in a serialized form, and can recreate the same type " & _
                                   "when used with the fromString method", _
                                   booInfo:=True, _
                                   booMsgbox:=True)
    End Sub

    Private Sub cmdVariableTypesClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVariableTypesClear.Click
        clearVT
    End Sub

    Private Sub cmdVariableTypeContainment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVariableTypeContainment.Click
        testContainment
    End Sub

    Private Sub cmdVariableTypesCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVariableTypesCreate.Click
        mkRandomTypes
    End Sub

    Private Sub cmdStatusZoom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStatusZoom.Click
        zoomInterface(lstStatus, 1.1, 5)
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        closer(False)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        Dim intOldWidth As Integer = Size.Width
        Dim intOldHeight As Integer = Size.Height
        Me.Size = New Size(CInt(_OBJwindowsUtilities.screenWidth _
                                * _
                                SCREEN_WIDTH_TOLERANCE), _
                           CInt(_OBJwindowsUtilities.screenHeight _
                                * _
                                SCREEN_HEIGHT_TOLERANCE))
        _OBJwindowsUtilities.resizeConstituentControls(Me, intOldWidth, intOldHeight)
        Me.CenterToScreen()
        Opacity = 0.5
        Show()
        Refresh()
        Try
            OBJtoolTip = New ToolTip
        Catch
            _OBJutilities.errorHandler("Cannot create the ToolTip object: " & _
                                       Err.Number & " " & Err.Description, _
                                       Me.Name, _
                                       "Form1_Load", _
                                       "Continuing: tool tips won't be displayed")
        End Try
        setToolTips()
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
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Loading complete")
        Opacity = 1
        Refresh()
   End Sub

   Private Sub lstVariableTypes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstVariableTypes.DoubleClick
       assignSelectedObject(lstVariableTypes.SelectedIndex)
   End Sub

   Private Sub lstVariableTypes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstVariableTypes.Click
       changeTestObjectControls(True)
       If (OBJqbVariableType Is Nothing) Then
           changeTestObjectControls(False)
           cmdDeleteVariableType.Visible = True
       End If
   End Sub

#End Region

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Show the application's Easter Egg
    '
    '
    Private Function about(ByVal booAllowCancel As Boolean) As Boolean
        Dim strAboutInfo As String = ABOUTINFO & _
                                    vbNewLine & vbNewLine & vbNewLine & _
                                    OBJqbVariableType.About
        Dim objMsgBoxResponse As MsgBoxResult = _
            MsgBox(strAboutInfo & _
                   vbNewLine & vbNewLine & vbNewLine & _
                   "To see further information about using this form, and obtain this documentation " & _
                   "on the Clipboard, simply Zoom the status box on the main form to " & _
                   "see this information (and extended information) " & _
                   "at the bottom of a scrollable list", _
                   CType(IIf(booAllowCancel, MsgBoxStyle.OKCancel, MsgBoxStyle.OKOnly), _
                         MsgBoxStyle))
        strAboutInfo &= vbNewLine & vbNewLine & vbNewLine & _
                            "Trying this Application Out" & _
                            vbNewLine & vbNewLine & _
                            ABOUTINFO_EXTENSION
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "Information about this form and the " & _
                                                 "qbVariableType class", _
                                                 1)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 _OBJutilities.string2Box(_OBJutilities.soft2HardParagraph(strAboutInfo)), _
                                                 booSplitLines:=True, _
                                                 booIncludeDate:=False)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 -1, _
                                                 "End of About information")
        If Not booAllowCancel Then Return (True)
        Return (objMsgBoxResponse = MsgBoxResult.OK)
    End Function

    ' -----------------------------------------------------------------
    ' Add a test object to collection and screen
    '
    '
    Private Function addTestObject(ByVal objNew As qbVariableType.qbVariableType) _
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
        With lstVariableTypes
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
        Select Case MsgBox("Should the selected object be cloned? " & _
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
            OBJqbVariableType = CType(COLtestObjects.Item(intIndex + 1), qbVariableType.qbVariableType)
            changeVTcontrols(True)
        Else
            createVariableType()
        End If
        txtVariableTypeExpression.Text = OBJqbVariableType.ToString
    End Sub

    ' -----------------------------------------------------------------
    ' Change the visibility of the test object controls
    '
    '
    Private Sub changeTestObjectControls(ByVal booVisible As Boolean)
        With cmdCompareToType
            If (.Tag Is Nothing) Then .Tag = .Text
            If lstVariableTypes.SelectedIndex <> -1 Then
                .Text = Replace(CStr(.Tag), "%VTNAME", _
                                CType(COLtestObjects.Item(lstVariableTypes.SelectedIndex + 1), _
                                      qbVariableType.qbVariableType).Name)
            End If
            .Visible = booVisible
        End With
        With cmdDeleteVariableType
            If (.Tag Is Nothing) Then .Tag = .Text
            If lstVariableTypes.SelectedIndex <> -1 Then
                .Text = Replace(CStr(.Tag), "%VTNAME", _
                                CType(COLtestObjects.Item(lstVariableTypes.SelectedIndex + 1), _
                                      qbVariableType.qbVariableType).Name)
            End If
            .Visible = booVisible
        End With
        With cmdVariableTypeContainment
            If (.Tag Is Nothing) Then .Tag = .Text
            If lstVariableTypes.SelectedIndex <> -1 Then
                .Text = Replace(CStr(.Tag), "%VTNAME", _
                                CType(COLtestObjects.Item(lstVariableTypes.SelectedIndex + 1), _
                                      qbVariableType.qbVariableType).Name)
            End If
            .Visible = booVisible
        End With
        chkContainedTypeWithState.Visible = booVisible
    End Sub

    ' -----------------------------------------------------------------
    ' Change the enabled status of the variable type controls
    '
    '
    Private Sub changeVTcontrols(ByVal booEnabled As Boolean)
        cmdCloneVariableType.Enabled = booEnabled
        cmdCreateVariableType.Enabled = Not booEnabled
        cmdDisposeVariableType.Enabled = booEnabled
        cmdInspect.Enabled = booEnabled
        cmdObject2XML.Enabled = booEnabled
        cmdMkRandomType.Enabled = Not booEnabled
        cmdToDescription.Enabled = booEnabled
        cmdToString.Enabled = booEnabled
        txtVariableTypeExpression.Font = _
            New Font(txtVariableTypeExpression.Font, _
                     CType(IIf(booEnabled, FontStyle.Bold, FontStyle.Regular), _
                           FontStyle))
        If Not booEnabled Then
            changeTestObjectControls(False)
        End If
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
        With lstVariableTypes
            .Items.Clear()
            .Refresh()
        End With
        changeTestObjectControls(False)
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
            Dim objHandle As qbVariableType.qbVariableType
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Clearing the variable type collection", 1)
            intIndex2 = 1
            For intIndex1 = .Count To 1 Step -1
                updateProgress("Clearing the variable type collection", _
                               "test object", _
                               intIndex2, _
                               intCount)
                _OBJwindowsUtilities.updateStatusListBox(lstStatus, "", 1)
                objHandle = CType(.Item(intIndex1), qbVariableType.qbVariableType)
                If Not disposeVariableType(objHandle) Then Return (False)
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
    ' Clone the variable type
    '
    '
    Private Function cloneVariableType() As Boolean
        Dim objClone As qbVariableType.qbVariableType
        Try
            objClone = OBJqbVariableType.clone
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
    ' compareTo Interface
    '
    '
    Private Sub compareToInterface(ByVal objQBV1 As qbVariableType.qbVariableType, _
                                   ByVal objQBV2 As qbVariableType.qbVariableType)
        Try
            Dim strExplanation As String
            If objQBV1.compareTo(objQBV2, strExplanation) Then
                MsgBox(_OBJutilities.enquote(objQBV1.Name) & " " & _
                       "is identical to " & _
                       _OBJutilities.enquote(objQBV2.Name) & ": " & _
                       strExplanation)
            Else
                MsgBox(_OBJutilities.enquote(objQBV1.Name) & " " & _
                       "is not identical to " & _
                       _OBJutilities.enquote(objQBV2.Name) & ": " & _
                       strExplanation)
            End If
        Catch
            _OBJutilities.errorHandler("Error occured in compareTo method: " & _
                                      Err.Number & " " & Err.Description, _
                                      Me.Name, "", _
                                      "Continuing")
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Create the variable type, disposing of the old object as needed:
    ' change display
    '
    '
    ' --- Always displays
    Private Function createVariableType() As Boolean
        Return createVariableType(True)
    End Function
    ' --- Option to suppress display
    Private Function createVariableType(ByVal booDisplay As Boolean) As Boolean
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                 "Creation of the variable type " & _
                                                 txtVariableTypeExpression.Text, _
                                                 1)
        If Not (OBJqbVariableType Is Nothing) Then destroyVariableType()
        Try
            OBJqbVariableType = New qbVariableType.qbVariableType(txtVariableTypeExpression.Text)
        Catch
            If booDisplay Then
                MsgBox("Cannot create variableType: " & Err.Number & " " & Err.Description)
            End If
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, "", -1)
            Return False
        End Try
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Creation of vt complete")
        changeVTcontrols(True)
        If lstVariableTypes.SelectedIndex <> -1 Then
            changeTestObjectControls(True)
        End If
        Dim strReport As String
        If Not OBJqbVariableType.inspect(strReport) Then
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     "Failed inspection report", _
                                                     1)
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     strReport, _
                                                     booSplitLines:=True)
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                     -1, _
                                                     "End of failed inspection report")
            If booDisplay Then
                MsgBox("Created object fails its inspection. Zoom the status box " & _
                       "to see the inspection report at the bottom of the display")
            End If
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, "", -1)
            Return False
        End If
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Delete variable type
    '
    '
    Private Sub deleteVariableType(ByVal intIndex As Integer)
        Dim intIndex1 As Integer = intIndex + 1
        Dim objHandle As qbVariableType.qbVariableType = _
            CType(COLtestObjects.Item(intIndex1), qbVariableType.qbVariableType)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Destroying the variable type " & _
                            _OBJutilities.enquote(objHandle.Name) & " " & _
                            "(" & objHandle.ToString & ")", _
                            1)
        disposeVariableType(objHandle)
        changeTestObjectControls(False)
        COLtestObjects.Remove(intIndex1)
        lstVariableTypes.Items.RemoveAt(intIndex)
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Destruction of vt complete")
    End Sub

    ' -----------------------------------------------------------------
    ' Destroy the vt object and change the display
    '
    '
    Private Sub destroyVariableType()
        disposeVariableType(OBJqbVariableType)
        OBJqbVariableType = Nothing
        changeVTcontrols(False)
    End Sub

    ' -----------------------------------------------------------------
    ' Dispose the variable type
    '
    '
    Private Function disposeVariableType(ByVal objType As qbVariableType.qbVariableType) _
            As Boolean
        If Not objType.Usable Then
            _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Cannot dispose of the variable type " & _
                                _OBJutilities.enquote(objType.Name) & " " & _
                                "because it is not usable")
            Exit Function
        End If
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Disposing the variable type " & " " & _
                            "(" & toStringInterface(objType) & ")", _
                            1)
        objType.disposeInspect()
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, -1, "Disposal complete")
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Save form settings
    '
    '
    Private Sub form2Registry()
        Try
            With txtVariableTypeExpression
                SaveSetting(Application.ProductName, Me.Name, .Name, .Text)
            End With
            With chkContainedTypeWithState
                SaveSetting(Application.ProductName, Me.Name, .Name, CStr(.Checked))
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
                     (ByRef objOperand1 As qbVariableType.qbVariableType, _
                      ByRef objOperand2 As qbVariableType.qbVariableType) _
        As Boolean
        Dim strOperand1 As String
        Dim strOperand2 As String
        objOperand1 = Nothing : objOperand2 = Nothing
        If Not (OBJqbVariableType Is Nothing) AndAlso OBJqbVariableType.Usable Then
            objOperand1 = OBJqbVariableType
            strOperand1 = OBJqbVariableType.ToString
        Else
            strOperand1 = OBJqbVariableType.mkRandomType
        End If
        Dim booHaveOperand2 As Boolean
        Dim intIndex1 As Integer = lstVariableTypes.SelectedIndex + 1
        If intIndex1 > 0 Then
            objOperand2 = CType(COLtestObjects.Item(intIndex1), _
                                qbVariableType.qbVariableType)
            booHaveOperand2 = objOperand2.Usable
        Else
            strOperand2 = OBJqbVariableType.mkRandomType
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
                objOperand1 = New qbVariableType.qbVariableType
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
                objOperand2 = New qbVariableType.qbVariableType
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
        Select Case MsgBox("Inspection of object has " & _
                           CStr(IIf(OBJqbVariableType.inspect(strReport), _
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
        txtVariableTypeExpression.Text = _OBJqbVariableType.mkRandomType
    End Sub

    ' -----------------------------------------------------------------
    ' Make several random types
    '
    '
    Private Sub mkRandomTypes()
        Dim booOK As Boolean
        Static intCount As Integer = 10
        Dim intIndex1 As Integer
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
        Dim intErr As Integer
        Dim objNew As qbVariableType.qbVariableType
        Dim strErr As String
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, "Creating random types", 1)
        For intIndex1 = 1 To intCount
            updateProgress("Creating random types", _
                           "type", _
                           intIndex1, _
                           intCount)
            booOK = False
            Try
                objNew = New qbVariableType.qbVariableType
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
        With lstVariableTypes
            .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
            .SelectedIndex = -1
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Dispose of all objects
    '
    '
    Private Function objectCleanup() As Boolean
        If Not (OBJqbVariableType Is Nothing) Then
            With OBJqbVariableType
                If .Usable AndAlso Not .disposeInspect Then Return (False)
            End With
        End If
        Return (clearVTcollection())
    End Function

    ' -----------------------------------------------------------------
    ' Restore form settings
    '
    '
    Private Sub registry2Form()
        Try
            With txtVariableTypeExpression
                .Text = GetSetting(Application.ProductName, Me.Name, .Name, .Text)
            End With
            With chkContainedTypeWithState
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Me.Name, _
                                            .Name, _
                                            CStr(.Checked)))
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
    ' Set tool tips
    '
    '
    Private Function setToolTips() As Boolean
        Try
            With OBJtoolTip
                .SetToolTip(cmdAbout, _
                            "Replays information About this application and " & _
                            "the qbVariableType object")
                .SetToolTip(cmdClearSettings, _
                            "Resets form selections, as saved in the Registry, " & _
                            "to initial values")
                .SetToolTip(cmdCloneVariableType, _
                            "Makes a clone (copy) of the variable type, and places " & _
                            "the name of the clone in the list box provided")
                .SetToolTip(cmdClose, _
                            "Saves all settings and dismisses the form")
                .SetToolTip(cmdCompareToType, _
                            "Compares the test variable type (at the top of the screen) " & _
                            "to the selected variable type in the Variable Types list box")
                .SetToolTip(cmdContainmentTester, _
                            "Allows you to test whether a given variable type logically " & _
                            """contains"" another variable type, in the sense that " & _
                            "all values of the contained type can be values of the container.")
                .SetToolTip(cmdCreateVariableType, _
                            "Creates the object corresponding to the variable type expression " & _
                            "in the black on white test box at the top of this form")
                .SetToolTip(cmdDeleteVariableType, _
                            "Removes the highlighted variable type from the Variable Types " & _
                            "list box")
                .SetToolTip(cmdDisposeVariableType, _
                            "Disposes the test variable type that is identified in the " & _
                            "text box at the top of this form")
                .SetToolTip(cmdForm2Registry, _
                            "Saves form selections in the Registry")
                .SetToolTip(cmdInspect, _
                            "Inspects the test variable type (identified in the text box) " & _
                            "for internal errors, and displays the inspection report")
                .SetToolTip(cmdMkRandomType, _
                            "Creates a random variable type, and identifies it in the " & _
                            "text box")
                .SetToolTip(cmdObject2XML, _
                            "Converts the state of the test object to eXtended Markup Language")
                .SetToolTip(cmdRandomize, _
                            "Creates ""nondeterministic"" tests with nonpredictable results")
                .SetToolTip(cmdRegistry2Form, _
                            "Restores form selections from the Registry")
                .SetToolTip(cmdStatusZoom, _
                            "Displays all status information in a scrollable, " & _
                            "read-only Text box, suitable for copying into the Clipboard, " & _
                            "and from there into Notepad or Word")
                .SetToolTip(cmdTest, _
                            "Conducts a stress test of the qbVariableType class")
                .SetToolTip(cmdToDescription, _
                            "Converts the variable type to documentation")
                .SetToolTip(cmdToString, _
                            "Converts the variable type to its serialized expression form")
                .SetToolTip(cmdVariableTypeContainment, _
                            "Allows you to test variable type containment using type expressions")
                .SetToolTip(cmdVariableTypesClear, _
                            "Clears the list box (below) which displays cloned and randomly-generated " & _
                            "variable type objects")
                .SetToolTip(cmdVariableTypesCreate, _
                            "Creates 0..n random variable types in the list box below after " & _
                            "prompting for the exact count")
                .SetToolTip(lstStatus, _
                            "Provides status information: use the Zoom box to see status " & _
                            "information in full")
                .SetToolTip(lstVariableTypes, _
                            "Identifies cloned and random variable types")
                .SetToolTip(txtVariableTypeExpression, _
                            "Enter a variable type name (such as ""Integer"") " & _
                            "or enter a complex type expression " & _
                            "such as ""Array,String,1,10,0,4""")
            End With
        Catch
            Return False
        End Try
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Test containment
    '
    '
    ' --- Use the test object and the list object
    Private Overloads Sub testContainment()
        testContainment(OBJqbVariableType, _
                        CType(COLtestObjects.Item(lstVariableTypes.SelectedIndex + 1), _
                              qbVariableType.qbVariableType), _
                        chkContainedTypeWithState.Checked)
    End Sub
    ' --- Use parameters
    Private Overloads Sub testContainment(ByVal objTest1 As qbVariableType.qbVariableType, _
                                          ByVal objTest2 As qbVariableType.qbVariableType, _
                                          ByVal booWithState As Boolean)
        With objTest1
            Dim booContained As Boolean
            Dim strExplanation As String = _
                "retry, with the containedTypeWithState check box set, " & _
                "to obtain an explanation"
            If booWithState Then
                booContained = .containedTypeWithState(objTest1, _
                                                       objTest2, _
                                                       strExplanation)
            Else
                booContained = _OBJqbVariableType.containedType(objTest1, objTest2)
            End If
            MsgBox("The test qb variable type object " & _
                   _OBJutilities.enquote(.Name) & " " & _
                   "is " & _
                   CStr(IIf(booContained, "", "not ")) & _
                   "contained in the qb variable type object " & _
                   _OBJutilities.enquote(objTest2.Name) & _
                   vbNewLine & vbNewLine & _
                   strExplanation)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Test interface
    '
    '
    Private Sub testInterface()
        If MsgBox("Note: this button will conduct a stress test of the " & _
                    "qbVariableType class. It will parse several fixed and " & _
                    "random variable type ""fromString"" expressions, " & _
                    "and it will test type containment. It will take between " & _
                    "5 and 10 seconds on an adequate modern Pentium whatever." & _
                    vbNewLine & vbNewLine & _
                    "At this writing, you will be able to conduct at least " & _
                    "ten ""deterministic"" tests if you haven't clicked the " & _
                    "Randomize button on the main form, and they should run to " & _
                    "completion. They will always make the same random selections, " & _
                    "in generating tests, and each should complete successfully." & _
                    vbNewLine & vbNewLine & _
                    "However, certain ""non-deterministic"" tests may fail " & _
                    "if you've clicked Randomize. The failure will be on an " & _
                    "invalid index." & _
                    vbNewLine & vbNewLine & _
                    "This known problem is on the Issues list in the readme.TXT " & _
                    "file for the qbVariableType project for repair." & _
                    vbNewLine & vbNewLine & _
                    "To get started click OK. To cancel the test click Cancel.", _
                    MsgBoxStyle.OKCancel) _
            = _
            MsgBoxResult.Cancel Then
             Return
        End If
        Dim booEnabled As Boolean
        booEnabled = Enabled
        Enabled = False : Refresh()
        If (OBJqbVariableType Is Nothing) Then
            If Not createVariableType(False) Then
                _OBJwindowsUtilities.updateStatusListBox(lstStatus, _
                                                         "Replacing invalid variable type")
                txtVariableTypeExpression.Text = "Unknown"
                txtVariableTypeExpression.Refresh()
                If Not createVariableType(False) Then
                    Enabled = booEnabled : Refresh() : Return
                End If
            End If
        End If
        Dim strReport As String
        Select Case MsgBox("Test of object has " & _
                           CStr(IIf(OBJqbVariableType.test(strReport), _
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
        lblProgress.Visible = False
        Enabled = booEnabled : Refresh()
    End Sub

    ' -----------------------------------------------------------------
    ' Interface to toString where object may not be usable
    '
    '
    Private Function toStringInterface(ByVal objHandle As qbVariableType.qbVariableType) _
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
            .Width = CInt(_OBJutilities.histogram(intEntityNumber, _
                                                  dblRangeMax:=lstStatus.Width, _
                                                  dblValueMax:=intEntityCount))
            .Visible = True
            .Refresh()
        End With
    End Sub

    ' -------------------------------------------------------------------
    ' Interface to the zoom object
    '
    '
    Private Overloads Sub zoomInterface(ByVal objZoomed As Object)
        zoomInterface(objZoomed, 3, 2)
    End Sub
    Private Overloads Sub zoomInterface(ByVal objZoomed As Object, _
                                        ByVal dblWidthMultiple As Double, _
                                        ByVal dblHeightMultiple As Double)

        Dim objZoom As zoom.zoom
        Try
            objZoom = New zoom.zoom
        Catch
        End Try
        With objZoom
            .ZoomTextBox.Font = New Font(.ZoomTextBox.Font, FontStyle.Bold)
            .setSize(dblWidthMultiple, dblHeightMultiple)
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
                                 ByVal intLevelChange As Integer) Handles OBJqbVariableType.testEvent
        _OBJwindowsUtilities.updateStatusListBox(lstStatus, strDesc, intLevelChange)
    End Sub

    Private Sub testProgressHandler(ByVal strDesc As String, _
                                    ByVal strEntity As String, _
                                    ByVal intNumber As Integer, _
                                    ByVal intCount As Integer) Handles OBJqbVariableType.testProgressEvent
        updateProgress(strDesc, strEntity, intNumber, intCount)
    End Sub

#End Region ' Event handlers 

End Class
