Option Strict On

Imports utilities.utilities
Imports windowsUtilities.windowsUtilities

' *********************************************************************
' *                                                                   *
' * creditEvaluation                                                  *
' *                                                                   *
' *                                                                   *
' * This form uses Tags as follows.                                   *
' *                                                                   *
' *                                                                   *
' *      *  The lstRules list box is tagged with Nothing or a Quick   *
' *         Basic engine, and, it is tagged with a QBE when the QBE   *
' *         contains a valid result (and, the rules have not been     *
' *         changed).                                                 *
' *                                                                   *
' *      *  The lstStatus list box is tagged with the status report   *
' *         level by the updateStatusListBox method in                *
' *         windowsUtilities.                                         *
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
    Friend WithEvents lblRules As System.Windows.Forms.Label
    Friend WithEvents cmdRulesDelete As System.Windows.Forms.Button
    Friend WithEvents cmdEvaluateCredit As System.Windows.Forms.Button
    Friend WithEvents cmdRulesAdd As System.Windows.Forms.Button
    Friend WithEvents lstRules As System.Windows.Forms.ListBox
    Friend WithEvents lstStatus As System.Windows.Forms.ListBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents cmdAbout As System.Windows.Forms.Button
    Friend WithEvents cmdRulesEdit As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdClearRegistry As System.Windows.Forms.Button
    Friend WithEvents cmdRulesDefaultPolicy As System.Windows.Forms.Button
    Friend WithEvents cmdRulesClear As System.Windows.Forms.Button
    Friend WithEvents cmdRulesShowCode As System.Windows.Forms.Button
    Friend WithEvents cmdCloseNoSave As System.Windows.Forms.Button
    Friend WithEvents cmdRulesAllCode As System.Windows.Forms.Button
    Friend WithEvents cmdRulesExplain As System.Windows.Forms.Button
    Friend WithEvents cmdRulesExplainRule As System.Windows.Forms.Button
    Friend WithEvents gbxApplicant As System.Windows.Forms.GroupBox
    Friend WithEvents chkBankrupt As System.Windows.Forms.CheckBox
    Friend WithEvents nud60daysPastDue As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbl60DayPastDue As System.Windows.Forms.Label
    Friend WithEvents nud30DaysPastDue As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbl30dayPastDue As System.Windows.Forms.Label
    Friend WithEvents gbxHousing As System.Windows.Forms.GroupBox
    Friend WithEvents txtHousingOther As System.Windows.Forms.TextBox
    Friend WithEvents radHousingOther As System.Windows.Forms.RadioButton
    Friend WithEvents radHousingRents As System.Windows.Forms.RadioButton
    Friend WithEvents radHousingOwns As System.Windows.Forms.RadioButton
    Friend WithEvents nudAnnualIncome As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblAnnualIncome As System.Windows.Forms.Label
    Friend WithEvents txtOutput As System.Windows.Forms.TextBox
    Friend WithEvents chkRulesThorough As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblRules = New System.Windows.Forms.Label
Me.cmdRulesDelete = New System.Windows.Forms.Button
Me.cmdEvaluateCredit = New System.Windows.Forms.Button
Me.cmdRulesAdd = New System.Windows.Forms.Button
Me.lstRules = New System.Windows.Forms.ListBox
Me.lstStatus = New System.Windows.Forms.ListBox
Me.cmdClose = New System.Windows.Forms.Button
Me.lblProgress = New System.Windows.Forms.Label
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdClearRegistry = New System.Windows.Forms.Button
Me.cmdAbout = New System.Windows.Forms.Button
Me.cmdRulesEdit = New System.Windows.Forms.Button
Me.cmdRulesDefaultPolicy = New System.Windows.Forms.Button
Me.cmdRulesClear = New System.Windows.Forms.Button
Me.cmdRulesShowCode = New System.Windows.Forms.Button
Me.cmdCloseNoSave = New System.Windows.Forms.Button
Me.cmdRulesAllCode = New System.Windows.Forms.Button
Me.cmdRulesExplain = New System.Windows.Forms.Button
Me.cmdRulesExplainRule = New System.Windows.Forms.Button
Me.gbxApplicant = New System.Windows.Forms.GroupBox
Me.chkBankrupt = New System.Windows.Forms.CheckBox
Me.nud60daysPastDue = New System.Windows.Forms.NumericUpDown
Me.lbl60DayPastDue = New System.Windows.Forms.Label
Me.nud30DaysPastDue = New System.Windows.Forms.NumericUpDown
Me.lbl30dayPastDue = New System.Windows.Forms.Label
Me.gbxHousing = New System.Windows.Forms.GroupBox
Me.txtHousingOther = New System.Windows.Forms.TextBox
Me.radHousingOther = New System.Windows.Forms.RadioButton
Me.radHousingRents = New System.Windows.Forms.RadioButton
Me.radHousingOwns = New System.Windows.Forms.RadioButton
Me.nudAnnualIncome = New System.Windows.Forms.NumericUpDown
Me.lblAnnualIncome = New System.Windows.Forms.Label
Me.txtOutput = New System.Windows.Forms.TextBox
Me.chkRulesThorough = New System.Windows.Forms.CheckBox
Me.gbxApplicant.SuspendLayout()
CType(Me.nud60daysPastDue, System.ComponentModel.ISupportInitialize).BeginInit()
CType(Me.nud30DaysPastDue, System.ComponentModel.ISupportInitialize).BeginInit()
Me.gbxHousing.SuspendLayout()
CType(Me.nudAnnualIncome, System.ComponentModel.ISupportInitialize).BeginInit()
Me.SuspendLayout()
'
'lblRules
'
Me.lblRules.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblRules.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblRules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblRules.Location = New System.Drawing.Point(10, 166)
Me.lblRules.Name = "lblRules"
Me.lblRules.Size = New System.Drawing.Size(892, 74)
Me.lblRules.TabIndex = 8
Me.lblRules.Text = "Credit scoring rules"
'
'cmdRulesDelete
'
Me.cmdRulesDelete.Location = New System.Drawing.Point(614, 203)
Me.cmdRulesDelete.Name = "cmdRulesDelete"
Me.cmdRulesDelete.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesDelete.TabIndex = 9
Me.cmdRulesDelete.Text = "Delete Rule"
'
'cmdEvaluateCredit
'
Me.cmdEvaluateCredit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdEvaluateCredit.Location = New System.Drawing.Point(10, 9)
Me.cmdEvaluateCredit.Name = "cmdEvaluateCredit"
Me.cmdEvaluateCredit.Size = New System.Drawing.Size(249, 83)
Me.cmdEvaluateCredit.TabIndex = 10
Me.cmdEvaluateCredit.Text = "Evaluate Credit"
'
'cmdRulesAdd
'
Me.cmdRulesAdd.Location = New System.Drawing.Point(182, 175)
Me.cmdRulesAdd.Name = "cmdRulesAdd"
Me.cmdRulesAdd.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesAdd.TabIndex = 11
Me.cmdRulesAdd.Text = "Add Rule"
'
'lstRules
'
Me.lstRules.BackColor = System.Drawing.SystemColors.Control
Me.lstRules.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstRules.ItemHeight = 17
Me.lstRules.Location = New System.Drawing.Point(10, 249)
Me.lstRules.Name = "lstRules"
Me.lstRules.Size = New System.Drawing.Size(892, 106)
Me.lstRules.TabIndex = 12
'
'lstStatus
'
Me.lstStatus.BackColor = System.Drawing.SystemColors.Control
Me.lstStatus.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstStatus.ItemHeight = 17
Me.lstStatus.Location = New System.Drawing.Point(10, 554)
Me.lstStatus.Name = "lstStatus"
Me.lstStatus.Size = New System.Drawing.Size(892, 106)
Me.lstStatus.TabIndex = 15
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(154, 112)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(105, 48)
Me.cmdClose.TabIndex = 16
Me.cmdClose.Text = "Close"
'
'lblProgress
'
Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblProgress.Location = New System.Drawing.Point(10, 664)
Me.lblProgress.Name = "lblProgress"
Me.lblProgress.Size = New System.Drawing.Size(892, 16)
Me.lblProgress.TabIndex = 17
Me.lblProgress.Visible = False
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(10, 692)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(124, 28)
Me.cmdForm2Registry.TabIndex = 18
Me.cmdForm2Registry.Text = "Save Settings"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(144, 692)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(125, 28)
Me.cmdRegistry2Form.TabIndex = 19
Me.cmdRegistry2Form.Text = "Restore Settings"
'
'cmdClearRegistry
'
Me.cmdClearRegistry.Location = New System.Drawing.Point(278, 692)
Me.cmdClearRegistry.Name = "cmdClearRegistry"
Me.cmdClearRegistry.Size = New System.Drawing.Size(125, 28)
Me.cmdClearRegistry.TabIndex = 20
Me.cmdClearRegistry.Text = "Clear Settings"
'
'cmdAbout
'
Me.cmdAbout.Location = New System.Drawing.Point(778, 692)
Me.cmdAbout.Name = "cmdAbout"
Me.cmdAbout.Size = New System.Drawing.Size(124, 28)
Me.cmdAbout.TabIndex = 21
Me.cmdAbout.Text = "About"
'
'cmdRulesEdit
'
Me.cmdRulesEdit.Location = New System.Drawing.Point(614, 175)
Me.cmdRulesEdit.Name = "cmdRulesEdit"
Me.cmdRulesEdit.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesEdit.TabIndex = 22
Me.cmdRulesEdit.Text = "Edit Rule"
'
'cmdRulesDefaultPolicy
'
Me.cmdRulesDefaultPolicy.Location = New System.Drawing.Point(326, 175)
Me.cmdRulesDefaultPolicy.Name = "cmdRulesDefaultPolicy"
Me.cmdRulesDefaultPolicy.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesDefaultPolicy.TabIndex = 23
Me.cmdRulesDefaultPolicy.Text = "Default Policy"
'
'cmdRulesClear
'
Me.cmdRulesClear.Location = New System.Drawing.Point(326, 203)
Me.cmdRulesClear.Name = "cmdRulesClear"
Me.cmdRulesClear.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesClear.TabIndex = 25
Me.cmdRulesClear.Text = "Clear"
'
'cmdRulesShowCode
'
Me.cmdRulesShowCode.Location = New System.Drawing.Point(758, 203)
Me.cmdRulesShowCode.Name = "cmdRulesShowCode"
Me.cmdRulesShowCode.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesShowCode.TabIndex = 26
Me.cmdRulesShowCode.Text = "Show Basic code"
'
'cmdCloseNoSave
'
Me.cmdCloseNoSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCloseNoSave.Location = New System.Drawing.Point(10, 112)
Me.cmdCloseNoSave.Name = "cmdCloseNoSave"
Me.cmdCloseNoSave.Size = New System.Drawing.Size(105, 48)
Me.cmdCloseNoSave.TabIndex = 27
Me.cmdCloseNoSave.Text = "Close (don't save settings)"
'
'cmdRulesAllCode
'
Me.cmdRulesAllCode.Location = New System.Drawing.Point(470, 203)
Me.cmdRulesAllCode.Name = "cmdRulesAllCode"
Me.cmdRulesAllCode.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesAllCode.TabIndex = 28
Me.cmdRulesAllCode.Text = "All Basic code"
'
'cmdRulesExplain
'
Me.cmdRulesExplain.Location = New System.Drawing.Point(470, 175)
Me.cmdRulesExplain.Name = "cmdRulesExplain"
Me.cmdRulesExplain.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesExplain.TabIndex = 29
Me.cmdRulesExplain.Text = "Explain"
'
'cmdRulesExplainRule
'
Me.cmdRulesExplainRule.Location = New System.Drawing.Point(758, 175)
Me.cmdRulesExplainRule.Name = "cmdRulesExplainRule"
Me.cmdRulesExplainRule.Size = New System.Drawing.Size(135, 23)
Me.cmdRulesExplainRule.TabIndex = 30
Me.cmdRulesExplainRule.Text = "Explain this rule"
'
'gbxApplicant
'
Me.gbxApplicant.Controls.Add(Me.chkBankrupt)
Me.gbxApplicant.Controls.Add(Me.nud60daysPastDue)
Me.gbxApplicant.Controls.Add(Me.lbl60DayPastDue)
Me.gbxApplicant.Controls.Add(Me.nud30DaysPastDue)
Me.gbxApplicant.Controls.Add(Me.lbl30dayPastDue)
Me.gbxApplicant.Controls.Add(Me.gbxHousing)
Me.gbxApplicant.Controls.Add(Me.nudAnnualIncome)
Me.gbxApplicant.Controls.Add(Me.lblAnnualIncome)
Me.gbxApplicant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxApplicant.Location = New System.Drawing.Point(269, 9)
Me.gbxApplicant.Name = "gbxApplicant"
Me.gbxApplicant.Size = New System.Drawing.Size(633, 148)
Me.gbxApplicant.TabIndex = 31
Me.gbxApplicant.TabStop = False
Me.gbxApplicant.Text = "Applicant Standing"
'
'chkBankrupt
'
Me.chkBankrupt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkBankrupt.ForeColor = System.Drawing.Color.Firebrick
Me.chkBankrupt.Location = New System.Drawing.Point(134, 111)
Me.chkBankrupt.Name = "chkBankrupt"
Me.chkBankrupt.Size = New System.Drawing.Size(212, 25)
Me.chkBankrupt.TabIndex = 15
Me.chkBankrupt.Text = "Bankruptcy (undischarged)"
'
'nud60daysPastDue
'
Me.nud60daysPastDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.nud60daysPastDue.Location = New System.Drawing.Point(326, 74)
Me.nud60daysPastDue.Name = "nud60daysPastDue"
Me.nud60daysPastDue.Size = New System.Drawing.Size(96, 23)
Me.nud60daysPastDue.TabIndex = 14
Me.nud60daysPastDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'lbl60DayPastDue
'
Me.lbl60DayPastDue.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lbl60DayPastDue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lbl60DayPastDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lbl60DayPastDue.Location = New System.Drawing.Point(10, 74)
Me.lbl60DayPastDue.Name = "lbl60DayPastDue"
Me.lbl60DayPastDue.Size = New System.Drawing.Size(307, 23)
Me.lbl60DayPastDue.TabIndex = 13
Me.lbl60DayPastDue.Text = "Sixty day past due"
Me.lbl60DayPastDue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'nud30DaysPastDue
'
Me.nud30DaysPastDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.nud30DaysPastDue.Location = New System.Drawing.Point(326, 46)
Me.nud30DaysPastDue.Name = "nud30DaysPastDue"
Me.nud30DaysPastDue.Size = New System.Drawing.Size(96, 23)
Me.nud30DaysPastDue.TabIndex = 12
Me.nud30DaysPastDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
'
'lbl30dayPastDue
'
Me.lbl30dayPastDue.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lbl30dayPastDue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lbl30dayPastDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lbl30dayPastDue.Location = New System.Drawing.Point(10, 46)
Me.lbl30dayPastDue.Name = "lbl30dayPastDue"
Me.lbl30dayPastDue.Size = New System.Drawing.Size(307, 23)
Me.lbl30dayPastDue.TabIndex = 11
Me.lbl30dayPastDue.Text = "Thirty day past due"
Me.lbl30dayPastDue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'gbxHousing
'
Me.gbxHousing.Controls.Add(Me.txtHousingOther)
Me.gbxHousing.Controls.Add(Me.radHousingOther)
Me.gbxHousing.Controls.Add(Me.radHousingRents)
Me.gbxHousing.Controls.Add(Me.radHousingOwns)
Me.gbxHousing.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxHousing.Location = New System.Drawing.Point(432, 18)
Me.gbxHousing.Name = "gbxHousing"
Me.gbxHousing.Size = New System.Drawing.Size(192, 120)
Me.gbxHousing.TabIndex = 10
Me.gbxHousing.TabStop = False
Me.gbxHousing.Text = "Housing"
'
'txtHousingOther
'
Me.txtHousingOther.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtHousingOther.Location = New System.Drawing.Point(10, 83)
Me.txtHousingOther.Name = "txtHousingOther"
Me.txtHousingOther.Size = New System.Drawing.Size(172, 23)
Me.txtHousingOther.TabIndex = 3
Me.txtHousingOther.Text = ""
'
'radHousingOther
'
Me.radHousingOther.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radHousingOther.Location = New System.Drawing.Point(10, 56)
Me.radHousingOther.Name = "radHousingOther"
Me.radHousingOther.Size = New System.Drawing.Size(172, 24)
Me.radHousingOther.TabIndex = 2
Me.radHousingOther.Text = "Other (please describe)"
'
'radHousingRents
'
Me.radHousingRents.Checked = True
Me.radHousingRents.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radHousingRents.Location = New System.Drawing.Point(86, 28)
Me.radHousingRents.Name = "radHousingRents"
Me.radHousingRents.Size = New System.Drawing.Size(68, 18)
Me.radHousingRents.TabIndex = 1
Me.radHousingRents.TabStop = True
Me.radHousingRents.Text = "Rents"
'
'radHousingOwns
'
Me.radHousingOwns.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radHousingOwns.Location = New System.Drawing.Point(10, 28)
Me.radHousingOwns.Name = "radHousingOwns"
Me.radHousingOwns.Size = New System.Drawing.Size(76, 18)
Me.radHousingOwns.TabIndex = 0
Me.radHousingOwns.Text = "Owns"
'
'nudAnnualIncome
'
Me.nudAnnualIncome.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.nudAnnualIncome.Location = New System.Drawing.Point(326, 18)
Me.nudAnnualIncome.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
Me.nudAnnualIncome.Name = "nudAnnualIncome"
Me.nudAnnualIncome.Size = New System.Drawing.Size(96, 23)
Me.nudAnnualIncome.TabIndex = 9
Me.nudAnnualIncome.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
Me.nudAnnualIncome.Value = New Decimal(New Integer() {20000, 0, 0, 0})
'
'lblAnnualIncome
'
Me.lblAnnualIncome.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.lblAnnualIncome.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblAnnualIncome.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblAnnualIncome.Location = New System.Drawing.Point(10, 18)
Me.lblAnnualIncome.Name = "lblAnnualIncome"
Me.lblAnnualIncome.Size = New System.Drawing.Size(307, 24)
Me.lblAnnualIncome.TabIndex = 8
Me.lblAnnualIncome.Text = "Annual Income"
Me.lblAnnualIncome.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtOutput
'
Me.txtOutput.BackColor = System.Drawing.SystemColors.Control
Me.txtOutput.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtOutput.Location = New System.Drawing.Point(10, 360)
Me.txtOutput.Multiline = True
Me.txtOutput.Name = "txtOutput"
Me.txtOutput.ReadOnly = True
Me.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
Me.txtOutput.Size = New System.Drawing.Size(892, 185)
Me.txtOutput.TabIndex = 32
Me.txtOutput.Text = ""
'
'chkRulesThorough
'
Me.chkRulesThorough.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(192, Byte))
Me.chkRulesThorough.Location = New System.Drawing.Point(16, 192)
Me.chkRulesThorough.Name = "chkRulesThorough"
Me.chkRulesThorough.Size = New System.Drawing.Size(152, 40)
Me.chkRulesThorough.TabIndex = 33
Me.chkRulesThorough.Text = "Thorough rule application"
'
'Form1
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
Me.ClientSize = New System.Drawing.Size(912, 725)
Me.Controls.Add(Me.chkRulesThorough)
Me.Controls.Add(Me.txtOutput)
Me.Controls.Add(Me.gbxApplicant)
Me.Controls.Add(Me.cmdRulesExplainRule)
Me.Controls.Add(Me.cmdRulesExplain)
Me.Controls.Add(Me.cmdRulesAllCode)
Me.Controls.Add(Me.cmdCloseNoSave)
Me.Controls.Add(Me.cmdRulesShowCode)
Me.Controls.Add(Me.cmdRulesClear)
Me.Controls.Add(Me.cmdRulesDefaultPolicy)
Me.Controls.Add(Me.cmdRulesEdit)
Me.Controls.Add(Me.cmdAbout)
Me.Controls.Add(Me.cmdClearRegistry)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.lblProgress)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.lstStatus)
Me.Controls.Add(Me.lstRules)
Me.Controls.Add(Me.cmdRulesAdd)
Me.Controls.Add(Me.cmdEvaluateCredit)
Me.Controls.Add(Me.cmdRulesDelete)
Me.Controls.Add(Me.lblRules)
Me.Name = "Form1"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
Me.Text = "Credit Evaluation Calculator"
Me.gbxApplicant.ResumeLayout(False)
CType(Me.nud60daysPastDue, System.ComponentModel.ISupportInitialize).EndInit()
CType(Me.nud30DaysPastDue, System.ComponentModel.ISupportInitialize).EndInit()
Me.gbxHousing.ResumeLayout(False)
CType(Me.nudAnnualIncome, System.ComponentModel.ISupportInitialize).EndInit()
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private OBJcollectionUtilities As collectionUtilities.collectionUtilities

    Private Const ABOUT As String = "This form and this application demonstrate " & _
                                    "a practical application of the Quick Basic engine " & _
                                    "to business rules in the form of a credit " & _
                                    "evaluation calculator."

    Private Const DATANAMES As String = "annualIncome thirtyDay sixtyDay bankrupt " & _
                                        "owns rents other otherDescription"

    Private Const EXAMPLERULES As String = "annualIncome < 5000,decline: Insufficent annual income" & vbNewLine & _
                                           "annualIncome >= 5000 And annualIncome <= 15000 " & _
                                           "And Not Bankrupt And ThirtyDay < 2 And SixtyDay = 0,.10: " & _
                                           "In the midrange group we accept most applicants at a very favorable rate" & vbNewLine & _
                                           "annualIncome >= 15000 And annualIncome <= 25000," & _
                                           "decline: This income range is declined by this firm as company policy" & vbNewLine & _
                                           "annualIncome > 25000 And Not Bankrupt,.15: " & _
                                           "The high income client will pay a higher interest rate at our firm" & vbNewLine & _
                                           "defaultPolicy, decline: " & _
                                           "Rejects other applicants"

    Private Const SCREEN_WIDTH_TOLERANCE As Single = 0.9
    Private Const SCREEN_HEIGHT_TOLERANCE As Single = 0.9

    Private OBJevaluation As Object
    Private COLevalRuleNumbers As Collection
    Private Enum ENUcontradictionType
        benign      ' Two or more declines
        APR         ' Two or more acceptances
        fatal       ' Acceptance and decline
    End Enum
    Private COLcontradictions As Collection     ' Collection of subcollections
                                                ' Item(1) of each subcollection: rule index
                                                ' Item(2) of each subcollection: type
                                                ' (benign, APR, fatal) 

    Private OBJtooltip As ToolTip

#End Region ' Form data

#Region " Form events "

    Private Sub chkBankrupt_CheckedChanged(ByVal objSender As System.Object, _
                                           ByVal objEvents As System.EventArgs) _
                Handles chkBankrupt.CheckedChanged
        adjustBKdisplay()
    End Sub

    Private Sub cmdAbout_Click(ByVal objSender As System.Object, _
                               ByVal objEvents As System.EventArgs) _
                Handles cmdAbout.Click
        showAbout()
    End Sub

    Private Sub cmdClearRegistry_Click(ByVal objSender As System.Object, _
                                       ByVal objEvents As System.EventArgs) _
                Handles cmdClearRegistry.Click
        clearRegistry()
    End Sub

    Private Sub cmdClose_Click(ByVal objSender As System.Object, _
                               ByVal objEvents As System.EventArgs) _
                Handles cmdClose.Click
        closer(True)
    End Sub

    Private Sub cmdCloseNoSave_Click(ByVal objSender As System.Object, _
                                     ByVal objEvents As System.EventArgs) _
                Handles cmdCloseNoSave.Click
        closer(False)
    End Sub

    Private Sub cmdEvaluateCredit_Click(ByVal objSender As System.Object, _
                                        ByVal objEvents As System.EventArgs) _
                Handles cmdEvaluateCredit.Click
        creditEvaluation()
    End Sub

    Private Sub cmdForm2Registry_Click(ByVal objSender As System.Object, _
                                       ByVal objEvents As System.EventArgs) _
                Handles cmdForm2Registry.Click
        form2Registry()
    End Sub

    Private Sub cmdRegistry2Form_Click(ByVal objSender As System.Object, _
                                       ByVal objEvents As System.EventArgs) _
                Handles cmdRegistry2Form.Click
        registry2Form()
    End Sub

    Private Sub cmdRulesClear_Click(ByVal objSender As System.Object, _
                                    ByVal objEvents As System.EventArgs) _
                Handles cmdRulesClear.Click
        clearRules()
    End Sub

    Private Sub cmdRulesDefaultPolicy_Click(ByVal objSender As System.Object, _
                                            ByVal objEvents As System.EventArgs) _
                Handles cmdRulesDefaultPolicy.Click
        getDefaultPolicy()
    End Sub

    Private Sub cmdRulesAdd_Click(ByVal objSender As System.Object, _
                                  ByVal objEvents As System.EventArgs) _
                Handles cmdRulesAdd.Click
        addRule(getNewRule)
    End Sub

    Private Sub cmdRulesAllCode_Click(ByVal objSender As System.Object, _
                                      ByVal objEvents As System.EventArgs) _
                Handles cmdRulesAllCode.Click
        showAllRuleCode()
    End Sub

    Private Sub cmdRulesDelete_Click(ByVal objSender As System.Object, _
                                     ByVal objEvents As System.EventArgs) _
                Handles cmdRulesDelete.Click
        deleteRule()
    End Sub

    Private Sub cmdRulesEdit_Click(ByVal objSender As System.Object, _
                                   ByVal objEvents As System.EventArgs) _
                Handles cmdRulesEdit.Click
        editRule()
    End Sub

    Private Sub cmdRulesExplain_Click(ByVal objSender As System.Object, _
                                      ByVal objEvents As System.EventArgs) _
                Handles cmdRulesExplain.Click
        explainRules()
    End Sub

    Private Sub cmdRulesExplainRule_Click(ByVal objSender As System.Object, _
                                          ByVal objEvents As System.EventArgs) _
                Handles cmdRulesExplainRule.Click
        explainRule()
    End Sub

    Private Sub cmdRulesShowCode_Click(ByVal objSender As System.Object, _
                                       ByVal objEvents As System.EventArgs) _
                Handles cmdRulesShowCode.Click
        showRuleCode()
    End Sub

    Private Sub lstRules_SelectedIndexChanged(ByVal objSender As System.Object, _
                                              ByVal objEvents As System.EventArgs) _
                Handles lstRules.SelectedIndexChanged
        adjustRuleButtons()
    End Sub

    Private Sub radHousingOther_CheckedChanged(ByVal objSender As System.Object, _
                                               ByVal objEvents As System.EventArgs) _
                Handles radHousingOther.CheckedChanged
        adjustHousingView()
    End Sub

    Private Sub radHousingOwns_CheckedChanged(ByVal objSender As System.Object, _
                                              ByVal objEvents As System.EventArgs) _
                Handles radHousingOther.CheckedChanged
        adjustHousingView()
    End Sub

    Private Sub radHousingRents_CheckedChanged(ByVal objSender As System.Object, _
                                               ByVal objEvents As System.EventArgs) _
                Handles radHousingRents.CheckedChanged
        adjustHousingView()
    End Sub

    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        closer(True)
    End Sub

    Private Sub Form1_Load(ByVal objSender As System.Object, _
                           ByVal objEvents As System.EventArgs) _
                Handles MyBase.Load
        updateStatusListBox(lstStatus, "Loading", 1)
        Dim booUsed As Boolean
        Try
            booUsed = CBool(GetSetting(Application.ProductName, Name, "used", "False"))
        Catch : End Try
        If Not booUsed Then
            showAbout()
            SaveSetting(Application.ProductName, Name, "used", "True")
        End If
        setToolTips()
        Dim intOldWidth As Integer = Width
        Dim intOldHeight As Integer = Height
        Width = CInt(screenWidth() * SCREEN_WIDTH_TOLERANCE)
        Height = CInt(screenHeight() * SCREEN_HEIGHT_TOLERANCE)
        resizeConstituentControls(Me, intOldWidth, intOldHeight)
        CenterToScreen()
        Try
            OBJcollectionUtilities = New collectionUtilities.collectionUtilities
        Catch
            errorHandler("Cannot create collection utilities: " & _
                         Err.Number & Err.Description, _
                         Name, "Form1_Load")
        End Try
        registry2Form()
        updateStatusListBox(lstStatus, "Load complete")
    End Sub

#End Region ' Form events

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Add one complete rule to the list box
    '
    '
    Private Sub addRule(ByVal strRule As String)
        If (strRule Is Nothing) Then Return
        lstRules.Items.Add(strRule)
        destroyRuleQBE()
        adjustRuleButtons()
    End Sub

    ' -----------------------------------------------------------------
    ' Add rule contradiction
    '
    '
    Private Sub addRuleContradiction(ByVal enuType As ENUcontradictionType, _
                                     ByVal intRuleIndex As Integer)
        If (COLcontradictions Is Nothing) Then
            Try
                COLcontradictions = New Collection
            Catch
                errorHandler("Cannot create COLcontradictions: " & _
                             Err.Number & " " & Err.Description, _
                             Name, _
                             "addRuleContradiction")
                Return
            End Try
        End If
        Try
            Dim colHandle As Collection = New Collection
            With colHandle
                .Add(intRuleIndex) : .Add(enuType)
            End With
            COLcontradictions.Add(colHandle)
        Catch
            errorHandler("Cannot extend COLcontradictions: " & _
                         Err.Number & " " & Err.Description, _
                         Name, _
                         "addRuleContradiction")
            Return
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Adjust bankruptcy display
    '
    '
    ' We expect initially to find the color to be used when the 
    ' applicant is bankrupt, and, we save this color. 
    '
    ' If the applicant is BK we use this color. Otherwise the color
    ' is set to black.
    '
    '
    Private Sub adjustBKdisplay()
        With chkBankrupt
            Static objBKcolor As Color = .ForeColor
            .ForeColor = CType(IIf(.Checked, objBKcolor, Color.Black), Color)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Adjust housing view
    '
    '
    Private Sub adjustHousingView()
        txtHousingOther.Enabled = radHousingOther.Checked
    End Sub

    ' -----------------------------------------------------------------
    ' Adjust rule buttons
    '
    '
    Private Sub adjustRuleButtons()
        With lstRules
            Dim booRuleExists As Boolean = .Items.Count > 0
            Dim booRuleSelected As Boolean = .SelectedIndex >= 0
            cmdRulesAllCode.Enabled = booRuleExists
            cmdRulesClear.Enabled = booRuleExists
            cmdRulesDelete.Enabled = booRuleSelected
            cmdRulesEdit.Enabled = booRuleSelected
            cmdRulesExplain.Enabled = booRuleSelected
            cmdRulesShowCode.Enabled = booRuleSelected
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Preprocess and apply the credit evaluation rules: return
    ' DECLINE or annual percentage rate
    '
    '
    Private Function applyEvaluation() As Object
        ' --- May have an existing result 
        If Not (lstRules.Tag Is Nothing) Then
            Return (CBool(CType(lstRules.Tag, _
                                quickBasicEngine.qbQuickBasicEngine).evaluationValue))
        End If
        ' --- Must process
        Try
            destroyRuleCollections()
            default2Last()
            OBJevaluation = Nothing
            Dim objQBE As quickBasicEngine.qbQuickBasicEngine = New quickBasicEngine.qbQuickBasicEngine
            AddHandler objQBE.scanEvent, AddressOf scanEventHandler
            AddHandler objQBE.parseEvent, AddressOf parseEventHandler
            AddHandler objQBE.interpretPrintEvent, AddressOf interpretPrintEventHandler
            AddHandler objQBE.interpretTraceEvent, AddressOf interpretTraceEventHandler
            With objQBE
                .SourceCode = rules2Basic()
                .run()
            End With
            Return (OBJevaluation)
        Catch
            errorHandler("Cannot create or use Quick Basic engine: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "creditEvaluation")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Apply one rule
    '
    '
    Private Sub applyRule(ByVal strOutstring As String)
        Dim booDecline As Boolean
        Dim objNewEvaluation As Object
        Dim strAction As String = word(strOutstring, 1)
        If UCase(strAction) = "DECLINE" Then
            objNewEvaluation = "DECLINE"
            booDecline = True
        Else
            Try
                objNewEvaluation = CSng(strAction)
            Catch : End Try
        End If
        If (objNewEvaluation Is Nothing) Then
            errorHandler("Programming error: the rules print the unexpected string " & _
                         enquote(strOutstring))
            Return
        End If
        Dim intRuleIndex As Integer
        Try
            intRuleIndex = CInt(word(strOutstring, 2))
        Catch
            errorHandler("Programming error: the rules print the unexpected string " & _
                         enquote(strOutstring))
            Return
        End Try
        If (OBJevaluation Is Nothing) AndAlso booDecline Then
            ' First evaluation is decline
            OBJevaluation = "DECLINE"
        ElseIf (OBJevaluation Is Nothing) AndAlso Not booDecline Then
            ' First evaluation is accept: objNewEvaluation is APR
            OBJevaluation = objNewEvaluation
        ElseIf Not (OBJevaluation Is Nothing) AndAlso booDecline Then
            ' A benign or fatal rule contradiction exists
            If OBJevaluation.GetType.ToString = "System.String" Then
                ' Declined for more than one reason 
                addRuleContradiction(ENUcontradictionType.benign, intRuleIndex)
            Else
                ' Accepted and then declined 
                addRuleContradiction(ENUcontradictionType.fatal, intRuleIndex)
            End If
        Else
            ' APR change contradiction: use better APR
            addRuleContradiction(ENUcontradictionType.APR, intRuleIndex)
            OBJevaluation = CSng(Math.Min(CSng(OBJevaluation), CSng(strAction)))
        End If
        If (COLevalRuleNumbers Is Nothing) Then
            Try
                COLevalRuleNumbers = New Collection
            Catch
                errorHandler("Cannot create evaluation rule number collection: " & _
                             Err.Number & " " & Err.Description, _
                             Name, _
                             "interpretPrintEventHandler", _
                             booMsgBox:=True)
            End Try
        End If
        Try
            COLevalRuleNumbers.Add(intRuleIndex)
        Catch
            errorHandler("Cannot extend evaluation rule number collection: " & _
                         Err.Number & " " & Err.Description, _
                         Name, _
                         "interpretPrintEventHandler", _
                         booMsgBox:=True)
        End Try
    End Sub

    ' -----------------------------------------------------------------
    ' Change selected rule
    '
    '
    Private Sub changeRule(ByVal strNewRule As String, _
                           ByVal intIndex As Integer)
        If lstRules.SelectedIndex < 0 Then
            MsgBox("No rule is selected for editing") : Return
        End If
        Dim strPolicy As String = showPolicyScreen()
        If strPolicy = "" Then Return
        changeRule(strPolicy, lstRules.SelectedIndex)
        destroyRuleQBE()
    End Sub

    ' -----------------------------------------------------------------
    ' Erase Registry settings
    '
    '
    Private Function clearRegistry() As Boolean
        Try
            DeleteSetting(Application.ProductName, Name)
        Catch
            errorHandler("Cannot clear Registry values: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "form2Registry")
        End Try
        Return (True)
    End Function

    ' -----------------------------------------------------------------
    ' Erase Rules
    '
    '
    Private Sub clearRules()
        If MsgBox("Click Yes to clear all the rules. Note that " & _
                  "this operation cannot be reversed.", _
                   MsgBoxStyle.YesNo) _
           = _
           MsgBoxResult.No Then Return
        lstRules.Items.Clear()
        adjustRuleButtons()
    End Sub

    ' -----------------------------------------------------------------
    ' Close form and application
    '
    '
    Private Sub closer(ByVal booSave As Boolean)
        If booSave Then form2Registry()
        destroyRuleCollections()
        Dispose()
        End
    End Sub

    ' -----------------------------------------------------------------
    ' Create Let statement list
    '
    '
    Private Function createAssignments() As String
        Return ("Let annualIncome = " & getValue("annualIncome") & vbNewLine & _
                 "Let thirtyDay = " & getValue("thirtyDay") & vbNewLine & _
                 "Let sixtyDay = " & getValue("sixtyDay") & vbNewLine & _
                 "Let bankrupt = " & getValue("bankrupt") & vbNewLine & _
                 "Let owns = " & getValue("owns") & vbNewLine & _
                 "Let rents = " & getValue("rents") & vbNewLine & _
                 "Let other = " & getValue("other") & vbNewLine & _
                 "Let otherDescription = " & getValue("otherDescription"))
    End Function

    ' -----------------------------------------------------------------
    ' Perform the credit evaluation
    '
    '
    Private Sub creditEvaluation()
        With txtOutput
            updateStatusListBox(lstStatus, 1, "Evaluating applicant")
            .Text = "Evaluating the applicant..."
            .Refresh()
            .Text = mkEvaluationReport(applyEvaluation)
            .Refresh()
            updateStatusListBox(lstStatus, -1, "Evaluation complete")
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Make sure the default policy is last
    '
    '
    Private Sub default2Last()
        With lstRules
            ' --- Create index list
            Dim colIndexes As Collection    ' In most cases will contain one entry
            Try
                colIndexes = New Collection
            Catch
                errorHandler("Cannot create collection: " & _
                             Err.Number & " " & Err.Description, _
                             Name, _
                             "default2Last")
                Return
            End Try
            ' --- Find all default policies
            Dim intIndex1 As Integer
            For intIndex1 = 0 To .Items.Count - 1
                If UCase(item(CStr(.Items.Item(intIndex1)), _
                                   1, ",", False)) _
                   = _
                   "DEFAULTPOLICY" Then
                    Try
                        colIndexes.Add(intIndex1)
                    Catch
                        errorHandler("Cannot extend collection: " & _
                                     Err.Number & " " & Err.Description, _
                                     Name, _
                                     "default2Last")
                        OBJcollectionUtilities.collectionClear(colIndexes)
                        Return
                    End Try
                End If
            Next intIndex1
            With colIndexes
                For intIndex1 = 1 To .Count
                    lstRules.Items.Add(lstRules.Items.Item(CInt(.Item(intIndex1))))
                Next intIndex1
                For intIndex1 = .Count To 1 Step -1
                    lstRules.Items.RemoveAt(CInt(.Item(intIndex1)))
                Next intIndex1
            End With
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Delete selected rule
    '
    '
    Private Sub deleteRule()
        With lstRules
            If .SelectedIndex < 0 Then
                MsgBox("No rule is selected for delete") : Return
            End If
            .Items.RemoveAt(.SelectedIndex)
        End With
        destroyRuleQBE()
        adjustRuleButtons()
    End Sub

    ' -----------------------------------------------------------------
    ' Deletes the rule number collections if they exist 
    '
    '
    Private Sub destroyRuleCollections()
        If Not (COLevalRuleNumbers Is Nothing) Then
            OBJcollectionUtilities.collectionClear(COLevalRuleNumbers)
        End If
        If Not (COLcontradictions Is Nothing) Then
            OBJcollectionUtilities.collectionClear(COLcontradictions)
        End If
    End Sub

    ' -----------------------------------------------------------------
    ' Deletes the rule engine if it exists
    '
    '
    Private Sub destroyRuleQBE()
        With lstRules
            If (.Tag Is Nothing) Then Return
            Dim objQBEhandle As quickBasicEngine.qbQuickBasicEngine
            Try
                objQBEhandle = CType(.Tag, quickBasicEngine.qbQuickBasicEngine)
            Catch
                Return
            End Try
            objQBEhandle.dispose()
            objQBEhandle = Nothing
            .Tag = Nothing
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Edit the rule
    '
    '
    Private Sub editRule()
        If lstRules.SelectedIndex < 0 Then
            MsgBox("No rule is selected for editing") : Return
        End If
        Dim strPolicy As String = showPolicyScreen()
        If strPolicy = "" Then Return
        changeRule(strPolicy, lstRules.SelectedIndex)
    End Sub

    ' -----------------------------------------------------------------
    ' Explain the selected rule
    '
    '
    Private Sub explainRule()
        With lstRules
            If .SelectedIndex < 0 Then
                MsgBox("No rule is selected for explanation") : Return
            End If
            txtOutput.Text = rule2Explanation(CStr(.Items.Item(.SelectedIndex)), True)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Explain one rule contradiction 
    '
    '
    Private Function explainContradiction(ByVal colRuleContradiction As Collection) _
            As String
        With colRuleContradiction
            Dim intRuleIndex As Integer = CInt(.Item(1))
            Select Case CType(.Item(2), ENUcontradictionType)
                Case ENUcontradictionType.APR
                    Return ("Note that rule " & intRuleIndex & " confirms a preceding acceptance: " & _
                            "selecting lowest applicable APR")
                Case ENUcontradictionType.benign
                    Return ("Note that rule " & intRuleIndex & " confirms a preceding decline")
                Case ENUcontradictionType.fatal
                    Return ("Note that rule " & intRuleIndex & " contradicts a preceding rule")
                Case Else
                    errorHandler("Unexpected enumerator has occured", _
                                 Name, _
                                 "explainContradiction")
                    Return ("")
            End Select
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Explain the rule contradictions
    '
    '
    Private Function explainContradictions(ByVal colRuleContradictions As Collection) _
            As String
        If (colRuleContradictions Is Nothing) Then Return ""
        With colRuleContradictions
            Dim intIndex1 As Integer
            Dim strOutstring As String
            For intIndex1 = 1 To .Count
                append(strOutstring, _
                       vbNewLine & vbNewLine, _
                       explainContradiction(CType(.Item(intIndex1), _
                                            Collection)))
            Next intIndex1
            Return (strOutstring)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Explain all rules
    '
    '
    Private Sub explainRules()
        With lstRules
            txtOutput.Text = "Obtaining explanation"
            txtOutput.Refresh()
            txtOutput.Text = rules2Explanation(True)
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Find rule identified only by condition and action
    '
    '
    Private Function findRule(ByVal strCondition As String, _
                              ByVal objAction As Object) As Integer
        With lstRules
            Dim intIndex1 As Integer
            Dim objNextAction As Object
            Dim strConditionUCase As String = UCase(strCondition)
            Dim strNextCondition As String
            For intIndex1 = 0 To .Items.Count - 1
                parseRule.parseRule(CStr(.Items.Item(intIndex1)), _
                                    strNextCondition, _
                                    objNextAction)
                If strConditionUCase = strNextCondition _
                   AndAlso _
                   sameAction(objAction, objNextAction) Then
                    Return intIndex1
                End If
            Next intIndex1
            Return -1
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Save Registry settings
    '
    '
    Private Function form2Registry() As Boolean
        Try
            With nudAnnualIncome
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Value))
            End With
            With nud30DaysPastDue
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Value))
            End With
            With nud60daysPastDue
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Value))
            End With
            With chkBankrupt
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            With radHousingOwns
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            With radHousingRents
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            With radHousingOther
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            With txtHousingOther
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            .Text)
            End With
            With chkRulesThorough
                SaveSetting(Application.ProductName, _
                            Name, _
                            .Name, _
                            CStr(.Checked))
            End With
            listBox2Registry(lstRules, _
                             strApplication:=Application.ProductName, _
                             strSection:=Name)
        Catch
            errorHandler("Cannot save Registry values: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "form2Registry")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Search for a fatal contradiction
    '
    '
    Private Function findFatalContradiction(ByVal colRuleContradictions As Collection) _
            As Boolean
        If (colRuleContradictions Is Nothing) Then Return (False)
        With colRuleContradictions
            Dim intIndex1 As Integer
            For intIndex1 = 1 To .Count
               If CType(CType(.Item(intIndex1), Collection).Item(2), ENUcontradictionType) _
                  = _
                  ENUcontradictionType.fatal Then Return (True)
            Next intIndex1
            Return (False)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Get default policy
    '
    '
    Private Sub getDefaultPolicy()
        Dim strDefaultPolicy As String = showDefaultPolicyScreen()
        If strDefaultPolicy = "" Then Return
        addRule(strDefaultPolicy)
    End Sub

    ' -----------------------------------------------------------------
    ' Get new rule
    '
    '
    Private Function getNewRule() As String
        Dim frmRuleEntry As ruleEntry = mkPolicyForm()
        If (frmRuleEntry Is Nothing) Then Return ""
        With frmRuleEntry
            .ShowDialog()
            If .WasClosed Then Return .ToString
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Get value
    '
    '
    Private Function getValue(ByVal strDataName As String) As String
        Select Case UCase(strDataName)
            Case "ANNUALINCOME" : Return CStr(nudAnnualIncome.Value)
            Case "THIRTYDAY" : Return CStr(nud30DaysPastDue.Value)
            Case "SIXTYDAY" : Return CStr(nud60daysPastDue.Value)
            Case "BANKRUPT" : Return CStr(chkBankrupt.Checked)
            Case "OWNS" : Return CStr(radHousingOwns.Checked)
            Case "RENTS" : Return CStr(radHousingRents.Checked)
            Case "OTHER" : Return CStr(radHousingOther.Checked)
            Case "OTHERDESCRIPTION" : Return enquote(txtHousingOther.Text)
            Case Else
                errorHandler("Invalid data name " & _
                             utilities.utilities.enquote(strDataName), _
                             Name, _
                             "getValue")
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Translate the identifier to an external name
    '
    '
    Private Function id2ExternalName(ByVal strID As String) As String
        Select Case UCase(strID)
            Case "ANNUALINCOME" : Return ("annual income")
            Case "THIRTYDAY" : Return ("30-day overdue reports")
            Case "SIXTYDAY" : Return ("60-day overdue reports")
            Case "BANKRUPT" : Return ("undischarged bankruptcy")
            Case "OWNS" : Return ("owns home")
            Case "RENTS" : Return ("rents home")
            Case "OTHER" : Return ("has other housing arrangements")
            Case "DEFAULTPOLICY" : Return ("no other rules apply")
            Case "AND" : Return ("and")
            Case "OR" : Return ("or")
            Case "NOT" : Return ("not")
            Case Else : Return (strID)
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Handle the qbe print event
    '
    '
    Private Sub interpretPrintEventHandler(ByVal objSender As quickBasicEngine.qbQuickBasicEngine, _
                                           ByVal strOutstring As String)
        applyRule(strOutstring)
    End Sub

    ' -----------------------------------------------------------------
    ' Handle the qbe print event
    '
    '
    Private Sub interpretTraceEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                            ByVal intIndex As Integer, _
                                            ByVal objStack As Stack, _
                                            ByVal colStorage As Collection)
        updateStatusListBox(lstStatus, "Executing compiled code at " & intIndex)
    End Sub

    ' -----------------------------------------------------------------
    ' Make the evaluation report
    '
    '
    Private Function mkEvaluationReport(ByVal objEval As Object) As String
        Dim strReport As String
        If findFatalContradiction(COLcontradictions) Then
            strReport = "The decision cannot be performed because the rules are not consistent"
        Else
            strReport = "The application has been "
            If UCase(objEval.GetType.ToString) = "SYSTEM.STRING" Then
                strReport &= "declined"
            Else
                strReport &= "accepted: the annual percentage rate shall be " & _
                            CSng(objEval)
            End If
        End If
        Return (strReport & _
                vbNewLine & vbNewLine & vbNewLine & _
                rules2Explanation(False) & _
                vbNewLine & vbNewLine & vbNewLine & _
                explainContradictions(COLcontradictions))
    End Function

    ' -----------------------------------------------------------------
    ' Make the policy form
    '
    '
    Private Function mkPolicyForm() As ruleEntry
        Dim frmNew As ruleEntry
        Try
            frmNew = New ruleEntry
            With frmNew
                .DataNames = DATANAMES : .ExampleRules = EXAMPLERULES
            End With
        Catch
            errorHandler("Cannot create policy form:" & _
                         Err.Number & " " & Err.Description, _
                         Name, "mkPolicyForm")
            Return (Nothing)
        End Try
        Return (frmNew)
    End Function

    ' -----------------------------------------------------------------
    ' Make rule in the format condition, action: comment
    '
    '
    Private Function mkRule(ByVal strCondition As String, _
                            ByVal objAction As Object, _
                            ByVal strComment As String) As String
        Return strCondition & ", " & CStr(objAction) & ": " & strComment
    End Function

    ' -----------------------------------------------------------------
    ' Translate opcode to external description
    '
    '
    Private Function op2ExternalName(ByVal strOp As String) As String
        Select Case strOp
            Case "=" : Return ("equals")
            Case "<>" : Return ("does not equal")
            Case "<" : Return ("is less than")
            Case "<=" : Return ("is less than or equal to")
            Case ">=" : Return ("is greater than or equal to")
            Case ">" : Return ("is greater than")
            Case Else : Return (strOp)
        End Select
    End Function

    ' -----------------------------------------------------------------
    ' Show the About information
    '
    '
    Private Sub parseEventHandler(ByVal objQBsender As quickBasicEngine.qbQuickBasicEngine, _
                                    ByVal strGrammarCategory As String, _
                                    ByVal booTerminal As Boolean, _
                                    ByVal intSrcStartIndex As Integer, _
                                    ByVal intSrcLength As Integer, _
                                    ByVal intTokStartIndex As Integer, _
                                    ByVal intTokLength As Integer, _
                                    ByVal intObjStartIndex As Integer, _
                                    ByVal intObjLength As Integer, _
                                    ByVal strComment As String, _
                                    ByVal intLevel As Integer)
        updateStatusListBox(lstStatus, _
                            "Parsing Basic rules at token " & intTokStartIndex)
    End Sub
    ' -----------------------------------------------------------------
    ' Restore Registry settings
    '
    '
    Private Function registry2Form() As Boolean
        Try
            With nudAnnualIncome
                .Value = CInt(GetSetting(Application.ProductName, _
                                         Name, _
                                         .Name, _
                                         CStr(.Value)))
            End With
            With nud30DaysPastDue
                .Value = CInt(GetSetting(Application.ProductName, _
                                         Name, _
                                         .Name, _
                                         CStr(.Value)))
            End With
            With nud60daysPastDue
                .Value = CInt(GetSetting(Application.ProductName, _
                                         Name, _
                                         .Name, _
                                         CStr(.Value)))
            End With
            With chkBankrupt
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            adjustBKdisplay()
            With radHousingOwns
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With radHousingRents
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With radHousingOther
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            With txtHousingOther
                .Text = CStr(GetSetting(Application.ProductName, _
                                            Name, _
                                            .Name, _
                                            CStr(.Text)))
            End With
            With chkRulesThorough
                .Checked = CBool(GetSetting(Application.ProductName, _
                                            Name, _
                                            .Name, _
                                            CStr(.Checked)))
            End With
            registry2ListBox(lstRules, _
                             strApplication:=Application.ProductName, _
                             strSection:=Name)
            If lstRules.Items.Count = 0 Then
                If MsgBox("No rules are available. Click Yes to obtain a set of " & _
                          "example rules. Click No for an empty set of rules.", _
                          MsgBoxStyle.YesNo) _
                   = _
                   MsgBoxResult.Yes Then
                    string2Listbox(EXAMPLERULES, lstRules)
                End If
            End If
            With lstRules
                .SelectedIndex = CInt(IIf(.Items.Count = 0, -1, 0))
            End With
            adjustRuleButtons()
        Catch
            errorHandler("Cannot save Registry values: " & _
                         Err.Number & " " & Err.Description, _
                         Name, "form2Registry")
        End Try
    End Function

    ' -----------------------------------------------------------------
    ' Rule to basic
    '
    '
    ' This method converts the displayed format
    '
    '
    '      condition,action: comment
    '
    '
    ' to 
    '
    '
    '      If condition Then 
    '          Print action & " " & ruleNumber
    '          End
    '      End If
    '
    '
    ' when thorough application is off. When thorough application is on, thie method 
    ' produces:
    '
    '
    '      If condition Then 
    '          Print action & " " & ruleNumber
    '          decisionMade = True
    '      End If
    '
    '
    ' Note: a condition of defaultPolicy is converted to True for non-thorough
    ' application and "Not decisionMade" for thorough application.
    '
    '
    Private Function rule2Basic(ByVal strRule As String, _
                                ByVal intRuleNumber As Integer) As String
        Dim objAction As Object
        Dim strCondition As String
        parseRule.parseRule(strRule, strCondition, objAction)
        If UCase(strCondition) = "DEFAULTPOLICY" Then
            strCondition = CStr(IIf(chkRulesThorough.Checked, "Not decisionMade", "True"))
        End If
        If (TypeOf objAction Is String) Then objAction = enquote(CStr(objAction))
        Return ("If " & strCondition & " Then " & _
                vbNewLine & _
                "    Print " & CStr(objAction) & " & "" "" & " & intRuleNumber & _
                vbNewLine & CStr(IIf(chkRulesThorough.Checked, "    decisionMade = True", "    End")) & _
                vbNewLine & _
                "End If")
    End Function

    ' -----------------------------------------------------------------
    ' Translate one rule to an explanation
    '
    ' Note that we're either explaining the conditional effect of all
    ' rules or the actual effect of only the applied rules.
    '
    '
    Private Function rule2Explanation(ByVal strRule As String, _
                                      ByVal booPredict As Boolean) As String
        Dim objAction As Object
        Dim objScanner As qbScanner.qbScanner
        Dim strComment As String
        Dim strCondition As String
        parseRule.parseRule(strRule, strCondition, objAction, strComment, objScanner)
        With objScanner
            Dim intIndex1 As Integer
            Dim strNexus As String = CStr(IIf(booPredict, "will be", "has been"))
            Dim strNext As String
            Dim strOutstring As String = CStr(IIf(booPredict, "If", "Because"))
            For intIndex1 = 1 To .TokenCount
                If .QBToken(intIndex1).TokenType = qbTokenType.qbTokenType.ENUtokenType.tokenTypeComma Then
                    Exit For
                End If
                Select Case .QBToken(intIndex1).TokenType
                    Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeIdentifier
                        strNext = .QBToken(intIndex1).sourceCode(.SourceCode)
                        If strNext = "NOT" _
                           AndAlso _
                           intIndex1 < .TokenCount _
                           AndAlso _
                           .QBToken(intIndex1 + 1).sourceCode(.SourceCode) = "BANKRUPT" Then
                            strNext = "no"
                        Else
                            strNext = id2ExternalName(strNext)
                        End If
                    Case qbTokenType.qbTokenType.ENUtokenType.tokenTypeOperator
                        strNext = op2ExternalName(.QBToken(intIndex1).sourceCode(.SourceCode))
                    Case Else
                        strNext = .QBToken(intIndex1).sourceCode(.SourceCode)
                End Select
                append(strOutstring, " ", strNext)
            Next intIndex1
            If (TypeOf objAction Is String) Then
                strOutstring &= ", the application " & strNexus & " declined"
            Else
                strOutstring &= ", the application " & strNexus & " " & _
                                "accepted with an Annual Percentage Rate of " & _
                                CSng(objAction)
            End If
            If strComment <> "" Then strComment = " (" & strComment & ")"
            Return strOutstring & strComment
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Translate rules to a complete Basic program
    '
    '
    Private Function rules2Basic() As String
        Dim intIndex1 As Integer
        Dim strOutstring As String = createAssignments()
        For intIndex1 = 0 To lstRules.Items.Count - 1
            append(strOutstring, _
                   vbNewLine, _
                   rule2Basic(CStr(lstRules.Items.Item(intIndex1)), _
                                   intIndex1 + 1))
        Next intIndex1
        Return (strOutstring)
    End Function

    ' -----------------------------------------------------------------
    ' Translate all rules, or rules actually used, to an explanation
    '
    '
    Private Function rules2Explanation(ByVal booPredict As Boolean) As String
        Dim intCount As Integer
        Dim intIndex1 As Integer
        Dim strNext As String
        Dim strOutstring As String
        If booPredict Then
            intIndex1 = 0 : intCount = lstRules.Items.Count - 1
        Else
            intIndex1 = 1
            Try
            intCount = COLevalRuleNumbers.Count
            Catch : End Try
        End If
        Dim intRule As Integer
        Dim strRule As String
        For intIndex1 = intIndex1 To intCount
            If booPredict Then
                strNext = CStr(lstRules.Items.Item(intIndex1))
                strRule = CStr(intIndex1 + 1) & ". "
            Else
                intRule = CInt(COLevalRuleNumbers.Item(intIndex1))
                strNext = CStr(lstRules.Items.Item(intRule - 1))
                strRule = "Rule " & intRule & ": "
            End If
            append(strOutstring, _
                   vbNewLine & vbNewLine, _
                   utilities.utilities.soft2HardParagraph _
                   (strRule & rule2Explanation(strNext, booPredict)))
        Next intIndex1
        Return (strOutstring)
    End Function

    ' -----------------------------------------------------------------
    ' Compare two actions for identity
    '
    '
    Private Function sameAction(ByVal objAction1 As Object, _
                                ByVal objAction2 As Object) As Boolean
        If (TypeOf objAction1 Is String) And (TypeOf objAction2 Is String) Then
            Return True
        ElseIf (TypeOf objAction1 Is String) And Not (TypeOf objAction2 Is String) Then
            Return False
        ElseIf Not (TypeOf objAction1 Is String) And (TypeOf objAction2 Is String) Then
            Return False
        Else
            Return CSng(objAction1) = CSng(objAction2)
        End If
    End Function

    ' -----------------------------------------------------------------
    ' Show the About information
    '
    '
    Private Sub scanEventHandler(ByVal objSender As quickBasicEngine.qbQuickBasicEngine, _
                                 ByVal objToken As qbToken.qbToken)
        updateStatusListBox(lstStatus, _
                            "Scanning Basic rules at token " & objToken.ToString)
    End Sub

    ' -----------------------------------------------------------------
    ' Assign all tooltips
    '
    '
    Private Function setToolTips() As Boolean
        Try
            OBJtooltip = New ToolTip
            OBJtooltip.SetToolTip(chkBankrupt, "Undischarged bankruptcy")
            OBJtooltip.SetToolTip(chkRulesThorough, "Apply all the rules to applicant even after a rule has been fulfilled")
            OBJtooltip.SetToolTip(cmdAbout, "Display information about this application")
            OBJtooltip.SetToolTip(cmdClearRegistry, "Clear form selections as saved in the Registry")
            OBJtooltip.SetToolTip(cmdClose, "Save form selections and close")
            OBJtooltip.SetToolTip(cmdCloseNoSave, "Close: does not save form selections")
            OBJtooltip.SetToolTip(cmdEvaluateCredit, "Apply scoring rules to evaluate credit")
            OBJtooltip.SetToolTip(cmdForm2Registry, "Save form selections")
            OBJtooltip.SetToolTip(cmdRegistry2Form, "Restore form selections from the Registry")
            OBJtooltip.SetToolTip(cmdRulesAdd, "Add a new rule")
            OBJtooltip.SetToolTip(cmdRulesAllCode, "View all Basic code")
            OBJtooltip.SetToolTip(cmdRulesClear, "Clear all rules")
            OBJtooltip.SetToolTip(cmdRulesDefaultPolicy, "Define the default rule")
            OBJtooltip.SetToolTip(cmdRulesDelete, "Deletes the selected rule")
            OBJtooltip.SetToolTip(cmdRulesEdit, "Edits the selected rule")
            OBJtooltip.SetToolTip(cmdRulesExplain, "Explain all the rules")
            OBJtooltip.SetToolTip(cmdRulesExplainRule, "Explain the selected rule")
            OBJtooltip.SetToolTip(cmdRulesShowCode, "Show the Basic code for the selected rule")
            OBJtooltip.SetToolTip(lbl30dayPastDue, "Number of bills paid past due 30 days")
            OBJtooltip.SetToolTip(lbl60DayPastDue, "Number of bills paid past due 60 days")
            OBJtooltip.SetToolTip(lblAnnualIncome, "Annual income")
            OBJtooltip.SetToolTip(radHousingOther, "Housing arrangements are other than own or rent")
            OBJtooltip.SetToolTip(radHousingOwns, "Owns home")
            OBJtooltip.SetToolTip(radHousingRents, "Rents home")
            OBJtooltip.SetToolTip(lstRules, "Identifies the rules")
            OBJtooltip.SetToolTip(lstStatus, "Identifies the operations being carried out in this application")
        Catch
            errorHandler("Cannot set tool tip: " & Err.Number & " " & Err.Description, _
                         Me.Name, "setToolTips", _
                         "Returning False", _
                         booMsgbox:=True, _
                         booInfo:=True)
            Return False
        End Try
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Show the About information
    '
    '
    Private Sub showAbout()
        MsgBox(ABOUT)
    End Sub

    ' -----------------------------------------------------------------
    ' Show the Basic code for all rules
    '
    '
    Private Sub showAllRuleCode()
        With lstRules
            txtOutput.Text = "Translating rules to Basic"
            txtOutput.Refresh()
            txtOutput.Text = rules2Basic()
            txtOutput.Refresh()
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Show the default policy screen and return the policy in the rule
    ' form
    '
    '
    Private Function showDefaultPolicyScreen() As String
        Dim frmPolicy As ruleEntry = mkPolicyForm()
        If frmPolicy Is Nothing Then Return ("")
        With frmPolicy
            .DefaultScreen = True
            .ShowDialog()
            If Not .WasClosed Then Return ("")
            Dim intIndex1 As Integer = findRule("defaultPolicy", .RuleResult)
            If intIndex1 >= 0 Then lstRules.Items.RemoveAt(intIndex1)
            addRule(mkRule("defaultPolicy", .RuleResult, "This will be the default policy"))
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Show the default policy screen and return the policy in the rule
    ' form
    '
    '
    Private Function showPolicyScreen() As String
        Dim frmPolicy As ruleEntry = mkPolicyForm()
        If frmPolicy Is Nothing Then Return ("")
        With frmPolicy
            .DefaultScreen = False
            .ShowDialog()
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Show the Basic code for one specific rule
    '
    '
    Private Sub showRuleCode()
        With lstRules
            If .SelectedIndex < 0 Then
                MsgBox("No rule has been selected") : Return
            End If
            txtOutput.Text = "Converting rule to Basic code"
            txtOutput.Refresh()
            txtOutput.Text = rule2Basic(CStr(.Items.Item(.SelectedIndex)), _
                                            .SelectedIndex + 1)
            txtOutput.Refresh()
        End With
    End Sub

#End Region ' General procedures

End Class
