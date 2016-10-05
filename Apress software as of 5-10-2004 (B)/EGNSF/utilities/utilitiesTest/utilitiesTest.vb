Public Class utilitiesTest
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
    Friend WithEvents lblTestReport As System.Windows.Forms.Label
    Friend WithEvents cmdMethodsSelectAll As System.Windows.Forms.Button
    Friend WithEvents lstMethods As System.Windows.Forms.ListBox
    Friend WithEvents lblTestResults As System.Windows.Forms.Label
    Friend WithEvents cmdMethodsClearSelection As System.Windows.Forms.Button
    Friend WithEvents txtTestResults As System.Windows.Forms.TextBox
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents cmdForm2Registry As System.Windows.Forms.Button
    Friend WithEvents cmdRegistry2Form As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblTestReport = New System.Windows.Forms.Label
Me.cmdMethodsSelectAll = New System.Windows.Forms.Button
Me.lstMethods = New System.Windows.Forms.ListBox
Me.lblTestResults = New System.Windows.Forms.Label
Me.cmdMethodsClearSelection = New System.Windows.Forms.Button
Me.txtTestResults = New System.Windows.Forms.TextBox
Me.cmdTest = New System.Windows.Forms.Button
Me.cmdForm2Registry = New System.Windows.Forms.Button
Me.cmdRegistry2Form = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblTestReport
'
Me.lblTestReport.BackColor = System.Drawing.Color.Navy
Me.lblTestReport.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblTestReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblTestReport.ForeColor = System.Drawing.Color.White
Me.lblTestReport.Location = New System.Drawing.Point(8, 8)
Me.lblTestReport.Name = "lblTestReport"
Me.lblTestReport.Size = New System.Drawing.Size(144, 16)
Me.lblTestReport.TabIndex = 0
Me.lblTestReport.Text = "Methods"
Me.lblTestReport.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdMethodsSelectAll
'
Me.cmdMethodsSelectAll.Location = New System.Drawing.Point(8, 24)
Me.cmdMethodsSelectAll.Name = "cmdMethodsSelectAll"
Me.cmdMethodsSelectAll.Size = New System.Drawing.Size(144, 24)
Me.cmdMethodsSelectAll.TabIndex = 1
Me.cmdMethodsSelectAll.Text = "Select All"
'
'lstMethods
'
Me.lstMethods.BackColor = System.Drawing.SystemColors.Control
Me.lstMethods.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lstMethods.ItemHeight = 14
Me.lstMethods.Location = New System.Drawing.Point(8, 72)
Me.lstMethods.Name = "lstMethods"
Me.lstMethods.Size = New System.Drawing.Size(144, 200)
Me.lstMethods.TabIndex = 2
'
'lblTestResults
'
Me.lblTestResults.BackColor = System.Drawing.Color.Navy
Me.lblTestResults.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblTestResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblTestResults.ForeColor = System.Drawing.Color.White
Me.lblTestResults.Location = New System.Drawing.Point(296, 8)
Me.lblTestResults.Name = "lblTestResults"
Me.lblTestResults.Size = New System.Drawing.Size(416, 16)
Me.lblTestResults.TabIndex = 3
Me.lblTestResults.Text = "Test Results"
Me.lblTestResults.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdMethodsClearSelection
'
Me.cmdMethodsClearSelection.Location = New System.Drawing.Point(8, 48)
Me.cmdMethodsClearSelection.Name = "cmdMethodsClearSelection"
Me.cmdMethodsClearSelection.Size = New System.Drawing.Size(144, 24)
Me.cmdMethodsClearSelection.TabIndex = 4
Me.cmdMethodsClearSelection.Text = "Clear Selection"
'
'txtTestResults
'
Me.txtTestResults.BackColor = System.Drawing.SystemColors.Control
Me.txtTestResults.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtTestResults.Location = New System.Drawing.Point(296, 24)
Me.txtTestResults.Multiline = True
Me.txtTestResults.Name = "txtTestResults"
Me.txtTestResults.Size = New System.Drawing.Size(416, 248)
Me.txtTestResults.TabIndex = 5
Me.txtTestResults.Text = ""
'
'cmdTest
'
Me.cmdTest.Location = New System.Drawing.Point(160, 8)
Me.cmdTest.Name = "cmdTest"
Me.cmdTest.Size = New System.Drawing.Size(128, 40)
Me.cmdTest.TabIndex = 6
Me.cmdTest.Text = "Test %TESTID"
'
'cmdForm2Registry
'
Me.cmdForm2Registry.Location = New System.Drawing.Point(160, 56)
Me.cmdForm2Registry.Name = "cmdForm2Registry"
Me.cmdForm2Registry.Size = New System.Drawing.Size(128, 24)
Me.cmdForm2Registry.TabIndex = 7
Me.cmdForm2Registry.Text = "Save Selections"
'
'cmdRegistry2Form
'
Me.cmdRegistry2Form.Location = New System.Drawing.Point(160, 88)
Me.cmdRegistry2Form.Name = "cmdRegistry2Form"
Me.cmdRegistry2Form.Size = New System.Drawing.Size(128, 24)
Me.cmdRegistry2Form.TabIndex = 8
Me.cmdRegistry2Form.Text = "Restore Selections"
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(160, 232)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(128, 40)
Me.cmdClose.TabIndex = 9
Me.cmdClose.Text = "Close"
'
'utilitiesTest
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(718, 283)
Me.ControlBox = False
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdRegistry2Form)
Me.Controls.Add(Me.cmdForm2Registry)
Me.Controls.Add(Me.cmdTest)
Me.Controls.Add(Me.txtTestResults)
Me.Controls.Add(Me.cmdMethodsClearSelection)
Me.Controls.Add(Me.lblTestResults)
Me.Controls.Add(Me.lstMethods)
Me.Controls.Add(Me.cmdMethodsSelectAll)
Me.Controls.Add(Me.lblTestReport)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "utilitiesTest"
Me.Text = "utilitiesTest"
Me.ResumeLayout(False)

    End Sub

#End Region

End Class
