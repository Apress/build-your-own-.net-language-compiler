Option Strict On

Imports windowsUtilities.windowsUtilities

Imports utilities

' *********************************************************************
' *                                                                   *
' * ruleEntry                                                         *
' *                                                                   *
' *                                                                   *
' * This form allows the creditEvaluation user to enter rules, and it *
' * exposes the following properties and methods:                     * 
' *                                                                   *
' *                                                                   *
' *      *  DataNames: this read-write property returns and may be    *
' *         set to the available data names                           *
' *                                                                   *
' *      *  DefaultScreen: this read-write property may be set to True*
' *         to display the screen that displays the default action for*
' *         a set of screens.                                         *
' *                                                                   *
' *         This screen makes the Condition section inaccessible.     *
' *                                                                   *
' *      *  ExampleRules: this write-only property should be set to a *
' *         newline-separated list of example rules when the form is  *
' *         displayed so the user may select and modify these rules.  *
' *                                                                   *
' *      *  RuleCondition: this read-write property returns the rule  *
' *         condition as set by the user and may initialize this      *
' *         condition.                                                *
' *                                                                   *
' *      *  RuleExplanation: this read-only property returns the rule *
' *         explanatory comments as set by the user                   *
' *                                                                   *
' *      *  RuleResult: this read-only property returns the rule      *
' *         result as set by the user:                                *
' *                                                                   *
' *         + If the user selects acceptance the RuleResult is the    *
' *           Single precision annual percentage rate.                *
' *                                                                   *
' *         + If the user selects decline the RuleResult is the       *
' *           string "DECLINE"                                        *
' *                                                                   *
' *         + If the user selects nothing the RuleResult is Nothing   *
' *                                                                   *
' *      *  toString: the toString method is overridden and it returns*
' *         the rule in the form condition,action: comments.          * 
' *                                                                   *
' *      *  WasClosed:t this read-only property returns True when the *
' *         form was dismissed using its Close button, False          *
' *         otherwise.                                                *
' *                                                                   *
' *                                                                   *
' *********************************************************************

Public Class ruleEntry
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
    Friend WithEvents gbxCondition As System.Windows.Forms.GroupBox
    Friend WithEvents txtCondition As System.Windows.Forms.TextBox
    Friend WithEvents gbxPolicy As System.Windows.Forms.GroupBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents radPolicyDecline As System.Windows.Forms.RadioButton
    Friend WithEvents nudPolicyAcceptAPR As System.Windows.Forms.NumericUpDown
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents lblCondition As System.Windows.Forms.Label
    Friend WithEvents lblExplanation As System.Windows.Forms.Label
    Friend WithEvents radPolicyAccept As System.Windows.Forms.RadioButton
    Friend WithEvents txtExplanation As System.Windows.Forms.TextBox
    Friend WithEvents lstExampleRules As System.Windows.Forms.ListBox
    Friend WithEvents lblExampleRules As System.Windows.Forms.Label
    Friend WithEvents lstDataNames As System.Windows.Forms.ListBox
    Friend WithEvents lblConditionDataNames As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.gbxCondition = New System.Windows.Forms.GroupBox
Me.lblExplanation = New System.Windows.Forms.Label
Me.txtExplanation = New System.Windows.Forms.TextBox
Me.lblCondition = New System.Windows.Forms.Label
Me.txtCondition = New System.Windows.Forms.TextBox
Me.gbxPolicy = New System.Windows.Forms.GroupBox
Me.radPolicyAccept = New System.Windows.Forms.RadioButton
Me.nudPolicyAcceptAPR = New System.Windows.Forms.NumericUpDown
Me.radPolicyDecline = New System.Windows.Forms.RadioButton
Me.cmdClose = New System.Windows.Forms.Button
Me.cmdCancel = New System.Windows.Forms.Button
Me.lstExampleRules = New System.Windows.Forms.ListBox
Me.lblExampleRules = New System.Windows.Forms.Label
Me.lstDataNames = New System.Windows.Forms.ListBox
Me.lblConditionDataNames = New System.Windows.Forms.Label
Me.gbxCondition.SuspendLayout()
Me.gbxPolicy.SuspendLayout()
CType(Me.nudPolicyAcceptAPR, System.ComponentModel.ISupportInitialize).BeginInit()
Me.SuspendLayout()
'
'gbxCondition
'
Me.gbxCondition.Controls.Add(Me.lblExplanation)
Me.gbxCondition.Controls.Add(Me.txtExplanation)
Me.gbxCondition.Controls.Add(Me.lblCondition)
Me.gbxCondition.Controls.Add(Me.txtCondition)
Me.gbxCondition.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxCondition.Location = New System.Drawing.Point(8, 8)
Me.gbxCondition.Name = "gbxCondition"
Me.gbxCondition.Size = New System.Drawing.Size(408, 104)
Me.gbxCondition.TabIndex = 0
Me.gbxCondition.TabStop = False
Me.gbxCondition.Text = "Condition"
'
'lblExplanation
'
Me.lblExplanation.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
Me.lblExplanation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblExplanation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblExplanation.Location = New System.Drawing.Point(8, 56)
Me.lblExplanation.Name = "lblExplanation"
Me.lblExplanation.Size = New System.Drawing.Size(392, 16)
Me.lblExplanation.TabIndex = 7
Me.lblExplanation.Text = "Explanation"
Me.lblExplanation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'txtExplanation
'
Me.txtExplanation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtExplanation.Location = New System.Drawing.Point(8, 72)
Me.txtExplanation.Name = "txtExplanation"
Me.txtExplanation.Size = New System.Drawing.Size(392, 20)
Me.txtExplanation.TabIndex = 6
Me.txtExplanation.Text = ""
'
'lblCondition
'
Me.lblCondition.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
Me.lblCondition.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblCondition.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblCondition.Location = New System.Drawing.Point(8, 16)
Me.lblCondition.Name = "lblCondition"
Me.lblCondition.Size = New System.Drawing.Size(392, 16)
Me.lblCondition.TabIndex = 5
Me.lblCondition.Text = "Condition"
Me.lblCondition.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'txtCondition
'
Me.txtCondition.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtCondition.Location = New System.Drawing.Point(8, 32)
Me.txtCondition.Name = "txtCondition"
Me.txtCondition.Size = New System.Drawing.Size(392, 20)
Me.txtCondition.TabIndex = 0
Me.txtCondition.Text = ""
'
'gbxPolicy
'
Me.gbxPolicy.Controls.Add(Me.radPolicyAccept)
Me.gbxPolicy.Controls.Add(Me.nudPolicyAcceptAPR)
Me.gbxPolicy.Controls.Add(Me.radPolicyDecline)
Me.gbxPolicy.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxPolicy.Location = New System.Drawing.Point(424, 8)
Me.gbxPolicy.Name = "gbxPolicy"
Me.gbxPolicy.Size = New System.Drawing.Size(224, 72)
Me.gbxPolicy.TabIndex = 1
Me.gbxPolicy.TabStop = False
Me.gbxPolicy.Text = "Policy"
'
'radPolicyAccept
'
Me.radPolicyAccept.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radPolicyAccept.Location = New System.Drawing.Point(8, 40)
Me.radPolicyAccept.Name = "radPolicyAccept"
Me.radPolicyAccept.Size = New System.Drawing.Size(144, 16)
Me.radPolicyAccept.TabIndex = 5
Me.radPolicyAccept.Text = "Accept at this APR:"
'
'nudPolicyAcceptAPR
'
Me.nudPolicyAcceptAPR.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.nudPolicyAcceptAPR.Location = New System.Drawing.Point(160, 40)
Me.nudPolicyAcceptAPR.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
Me.nudPolicyAcceptAPR.Name = "nudPolicyAcceptAPR"
Me.nudPolicyAcceptAPR.Size = New System.Drawing.Size(56, 20)
Me.nudPolicyAcceptAPR.TabIndex = 4
Me.nudPolicyAcceptAPR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'radPolicyDecline
'
Me.radPolicyDecline.Checked = True
Me.radPolicyDecline.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radPolicyDecline.Location = New System.Drawing.Point(8, 16)
Me.radPolicyDecline.Name = "radPolicyDecline"
Me.radPolicyDecline.Size = New System.Drawing.Size(144, 16)
Me.radPolicyDecline.TabIndex = 2
Me.radPolicyDecline.TabStop = True
Me.radPolicyDecline.Text = "Decline"
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(568, 88)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(80, 24)
Me.cmdClose.TabIndex = 2
Me.cmdClose.Text = "Close"
'
'cmdCancel
'
Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCancel.Location = New System.Drawing.Point(424, 88)
Me.cmdCancel.Name = "cmdCancel"
Me.cmdCancel.Size = New System.Drawing.Size(80, 24)
Me.cmdCancel.TabIndex = 3
Me.cmdCancel.Text = "Cancel"
'
'lstExampleRules
'
Me.lstExampleRules.BackColor = System.Drawing.SystemColors.Control
Me.lstExampleRules.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstExampleRules.Location = New System.Drawing.Point(136, 136)
Me.lstExampleRules.Name = "lstExampleRules"
Me.lstExampleRules.Size = New System.Drawing.Size(512, 108)
Me.lstExampleRules.TabIndex = 8
'
'lblExampleRules
'
Me.lblExampleRules.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
Me.lblExampleRules.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblExampleRules.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblExampleRules.Location = New System.Drawing.Point(136, 120)
Me.lblExampleRules.Name = "lblExampleRules"
Me.lblExampleRules.Size = New System.Drawing.Size(512, 16)
Me.lblExampleRules.TabIndex = 7
Me.lblExampleRules.Text = "Example Rules"
Me.lblExampleRules.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lstDataNames
'
Me.lstDataNames.BackColor = System.Drawing.SystemColors.Control
Me.lstDataNames.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstDataNames.Location = New System.Drawing.Point(8, 136)
Me.lstDataNames.Name = "lstDataNames"
Me.lstDataNames.Size = New System.Drawing.Size(120, 108)
Me.lstDataNames.TabIndex = 6
'
'lblConditionDataNames
'
Me.lblConditionDataNames.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(224, Byte), CType(192, Byte))
Me.lblConditionDataNames.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblConditionDataNames.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblConditionDataNames.Location = New System.Drawing.Point(8, 120)
Me.lblConditionDataNames.Name = "lblConditionDataNames"
Me.lblConditionDataNames.Size = New System.Drawing.Size(120, 16)
Me.lblConditionDataNames.TabIndex = 5
Me.lblConditionDataNames.Text = "Data Names"
Me.lblConditionDataNames.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'ruleEntry
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(656, 253)
Me.ControlBox = False
Me.Controls.Add(Me.lstExampleRules)
Me.Controls.Add(Me.lblExampleRules)
Me.Controls.Add(Me.lstDataNames)
Me.Controls.Add(Me.lblConditionDataNames)
Me.Controls.Add(Me.cmdCancel)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.gbxPolicy)
Me.Controls.Add(Me.gbxCondition)
Me.Name = "ruleEntry"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "ruleEntry"
Me.gbxCondition.ResumeLayout(False)
Me.gbxPolicy.ResumeLayout(False)
CType(Me.nudPolicyAcceptAPR, System.ComponentModel.ISupportInitialize).EndInit()
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private BOOclose As Boolean

#End Region ' Form data

#Region " Form properties and methods "

    Public Property DataNames() As String
        Get
            Return listBox2String(lstDataNames, strSep:=" ")
        End Get
        Set(ByVal strNewValue As String)
            string2Listbox(strNewValue, lstDataNames, strSep:=" ")
            With lstDataNames
                .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
            End With
        End Set
    End Property

    Public Property DefaultScreen() As Boolean
        Get
            Return gbxCondition.Enabled()
        End Get
        Set(ByVal booNewValue As Boolean)
            gbxCondition.Enabled = Not booNewValue
        End Set
    End Property

    Public WriteOnly Property ExampleRules() As String
        Set(ByVal strNewValue As String)
            string2Listbox(strNewValue, lstExampleRules)
            With lstExampleRules
                .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
            End With
        End Set
    End Property

    Public Property RuleCondition() As String
        Get
            Return txtCondition.Text
        End Get
        Set(ByVal strNewValue As String)
            txtCondition.Text = strNewValue
        End Set
    End Property

    Public ReadOnly Property RuleExplanation() As String
        Get
            Return txtExplanation.Text
        End Get
    End Property

    Public ReadOnly Property RuleResult() As Object
        Get
            Return getRuleResult()
        End Get
    End Property

    Public Overrides Function toString() As String
        With Me
            Return .RuleCondition & ", " & CStr(.RuleResult) & ": " & .RuleExplanation
        End With
    End Function

    Public ReadOnly Property WasClosed() As Boolean
        Get
            Return BOOclose
        End Get
    End Property

#End Region ' Form properties

#Region " Form events "

    Private Sub cmdCancel_Click(ByVal objSender As System.Object, _
                                ByVal objEvents As System.EventArgs) _
            Handles cmdCancel.Click
        Hide()
    End Sub

    Private Sub cmdClose_Click(ByVal objSender As System.Object, _
                               ByVal objEvents As System.EventArgs) _
            Handles cmdClose.Click
        BOOclose = True
        Hide()
    End Sub

    Private Sub lstDataNames_DoubleClick(ByVal objSender As System.Object, _
                                         ByVal objEvents As System.EventArgs) _
            Handles lstDataNames.DoubleClick
        With lstDataNames
            insertDataName(CStr(.Items(.SelectedIndex)))
        End With
    End Sub

    Private Sub lstExampleRules_DoubleClick(ByVal objSender As System.Object, _
                                            ByVal objEvents As System.EventArgs) _
            Handles lstExampleRules.DoubleClick
        With lstExampleRules
            rule2Screen(CStr(.Items.Item(.SelectedIndex)))
        End With
    End Sub

    Private Sub ruleEntry_Load(ByVal objSender As System.Object, _
                               ByVal objEvents As System.EventArgs) _
            Handles MyBase.Load
    End Sub

#End Region ' Form events

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Set the radio buttons and the nud to the action information
    '
    '
    Private Sub action2Screen(ByVal objAction As Object)
        If (TypeOf objAction Is String) AndAlso UCase(CStr(objAction)) = "DECLINE" Then
            radPolicyDecline.Checked = True
            nudPolicyAcceptAPR.Enabled = False
        ElseIf (TypeOf objAction Is Single) Then
            radPolicyAccept.Checked = True
            nudPolicyAcceptAPR.Value = CInt(CSng(objAction) * 100)
            nudPolicyAcceptAPR.Enabled = True
        Else
            utilities.utilities.errorHandler("Action object " & _
                                             utilities.utilities.object2String(objAction) & " " & _
                                             "has an unexpected format", _
                                             Name, "action2Screen")
        End If
    End Sub

    ' -----------------------------------------------------------------
    ' Get the rule's result as an APR or DECLINE
    '
    '
    Private Function getRuleResult() As Object
        If radPolicyAccept.Checked Then
            Return CSng(nudPolicyAcceptAPR.Value / 100)
        ElseIf radPolicyDecline.Checked Then
            Return "DECLINE"
        Else
            Return Nothing
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Insert the data name
    '
    '
    Private Sub insertDataName(ByVal strName As String)
        With txtCondition
            .Text = Mid(.Text, 1, .SelectionStart) & _
                    strName & _
                    Mid(.Text, .SelectionStart + 1)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Put the rule in the top part of the screen
    '
    '
    Private Sub rule2Screen(ByVal strRule As String)
        Dim strCondition As String
        Dim strComment As String
        Dim objAction As Object
        parseRule.parseRule(strRule, strCondition, objAction, strComment)
        txtCondition.Text = strRule
        txtExplanation.Text = strComment
        action2Screen(objAction)
    End Sub

#End Region ' General procedures

End Class
