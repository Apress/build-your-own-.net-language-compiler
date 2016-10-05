Option Strict

' *********************************************************************
' *                                                                   *
' * utilitiesInfo     Display utility information                     *
' *                                                                   *
' *********************************************************************

Public Class utilitiesInfo
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
    Friend WithEvents lblAbstract As System.Windows.Forms.Label
    Friend WithEvents txtAbstract As System.Windows.Forms.TextBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblAbstract = New System.Windows.Forms.Label
Me.txtAbstract = New System.Windows.Forms.TextBox
Me.cmdClose = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblAbstract
'
Me.lblAbstract.BackColor = System.Drawing.Color.Navy
Me.lblAbstract.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblAbstract.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblAbstract.ForeColor = System.Drawing.Color.White
Me.lblAbstract.Location = New System.Drawing.Point(8, 8)
Me.lblAbstract.Name = "lblAbstract"
Me.lblAbstract.Size = New System.Drawing.Size(888, 16)
Me.lblAbstract.TabIndex = 0
Me.lblAbstract.Text = "Abstract"
Me.lblAbstract.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'txtAbstract
'
Me.txtAbstract.BackColor = System.Drawing.SystemColors.Control
Me.txtAbstract.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.txtAbstract.Location = New System.Drawing.Point(8, 24)
Me.txtAbstract.Multiline = True
Me.txtAbstract.Name = "txtAbstract"
Me.txtAbstract.ReadOnly = True
Me.txtAbstract.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
Me.txtAbstract.Size = New System.Drawing.Size(888, 440)
Me.txtAbstract.TabIndex = 1
Me.txtAbstract.Text = ""
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(392, 472)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(144, 24)
Me.cmdClose.TabIndex = 2
Me.cmdClose.Text = "Close"
'
'utilitiesInfo
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(912, 501)
Me.ControlBox = False
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.txtAbstract)
Me.Controls.Add(Me.lblAbstract)
Me.Name = "utilitiesInfo"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "utilitiesInfo"
Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Hide
    End Sub

End Class
