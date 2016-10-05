Option Strict On

Imports utilities.utilities
Imports windowsUtilities.windowsUtilities

' *********************************************************************
' *                                                                   *
' * Display and modify variables                                      *
' *                                                                   *
' *                                                                   *
' * This form displays variables (name, location, type and value).    *
' *                                                                   *
' * The rest of this documentation discusses these topics:            *
' *                                                                   *
' *                                                                   *
' *      *  Properties exposed by this form                           *
' *                                                                   *
' *                                                                   *
' * PROPERTIES EXPOSED BY THIS FORM --------------------------------- *
' *                                                                   *
' * This form exposes the following properties:                       *
' *                                                                   *
' *                                                                   *
' *      *  Change: this read-only property returns True when the     *
' *         variable is changed, False otherwise                      *
' *                                                                   *
' *      *  VariableLocation: this read-write property returns and may*
' *         be set to the variable's location: note that if it is not *
' *         set, or set to -1, then no location displays.             *
' *                                                                   *
' *      *  Variable: this read-write property returns and may be set *
' *         to the variable                                           *
' *                                                                   *
' *                                                                   *
' *********************************************************************
Public Class modifyVariable
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        hideLocation()

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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblVariableName As System.Windows.Forms.Label
    Friend WithEvents txtVariableName As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents txtStorageLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblStorageLocation As System.Windows.Forms.Label
    Friend WithEvents txtVariableType As System.Windows.Forms.TextBox
    Friend WithEvents lblVariableType As System.Windows.Forms.Label
    Friend WithEvents txtVariableValue As System.Windows.Forms.TextBox
    Friend WithEvents lblVariableValue As System.Windows.Forms.Label
    Friend WithEvents cmdModify As System.Windows.Forms.Button
    Friend WithEvents chkMatrixView As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.lblVariableName = New System.Windows.Forms.Label
Me.txtVariableName = New System.Windows.Forms.TextBox
Me.txtStorageLocation = New System.Windows.Forms.TextBox
Me.lblStorageLocation = New System.Windows.Forms.Label
Me.txtVariableType = New System.Windows.Forms.TextBox
Me.lblVariableType = New System.Windows.Forms.Label
Me.txtVariableValue = New System.Windows.Forms.TextBox
Me.lblVariableValue = New System.Windows.Forms.Label
Me.cmdCancel = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.cmdModify = New System.Windows.Forms.Button
Me.chkMatrixView = New System.Windows.Forms.CheckBox
Me.SuspendLayout()
'
'lblVariableName
'
Me.lblVariableName.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
Me.lblVariableName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariableName.Location = New System.Drawing.Point(8, 8)
Me.lblVariableName.Name = "lblVariableName"
Me.lblVariableName.Size = New System.Drawing.Size(96, 20)
Me.lblVariableName.TabIndex = 0
Me.lblVariableName.Text = "Variable Name"
Me.lblVariableName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtVariableName
'
Me.txtVariableName.Location = New System.Drawing.Point(104, 8)
Me.txtVariableName.Name = "txtVariableName"
Me.txtVariableName.ReadOnly = True
Me.txtVariableName.Size = New System.Drawing.Size(176, 20)
Me.txtVariableName.TabIndex = 1
Me.txtVariableName.Text = ""
'
'txtStorageLocation
'
Me.txtStorageLocation.Location = New System.Drawing.Point(384, 8)
Me.txtStorageLocation.Name = "txtStorageLocation"
Me.txtStorageLocation.ReadOnly = True
Me.txtStorageLocation.Size = New System.Drawing.Size(56, 20)
Me.txtStorageLocation.TabIndex = 3
Me.txtStorageLocation.Text = ""
'
'lblStorageLocation
'
Me.lblStorageLocation.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
Me.lblStorageLocation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblStorageLocation.Location = New System.Drawing.Point(288, 8)
Me.lblStorageLocation.Name = "lblStorageLocation"
Me.lblStorageLocation.Size = New System.Drawing.Size(96, 20)
Me.lblStorageLocation.TabIndex = 2
Me.lblStorageLocation.Text = "Storage Location"
Me.lblStorageLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtVariableType
'
Me.txtVariableType.Location = New System.Drawing.Point(104, 32)
Me.txtVariableType.Name = "txtVariableType"
Me.txtVariableType.Size = New System.Drawing.Size(336, 20)
Me.txtVariableType.TabIndex = 5
Me.txtVariableType.Text = ""
'
'lblVariableType
'
Me.lblVariableType.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
Me.lblVariableType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariableType.Location = New System.Drawing.Point(8, 32)
Me.lblVariableType.Name = "lblVariableType"
Me.lblVariableType.Size = New System.Drawing.Size(96, 20)
Me.lblVariableType.TabIndex = 4
Me.lblVariableType.Text = "Type"
Me.lblVariableType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
'
'txtVariableValue
'
Me.txtVariableValue.Location = New System.Drawing.Point(8, 80)
Me.txtVariableValue.Multiline = True
Me.txtVariableValue.Name = "txtVariableValue"
Me.txtVariableValue.Size = New System.Drawing.Size(432, 96)
Me.txtVariableValue.TabIndex = 7
Me.txtVariableValue.Text = ""
'
'lblVariableValue
'
Me.lblVariableValue.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
Me.lblVariableValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblVariableValue.Location = New System.Drawing.Point(8, 56)
Me.lblVariableValue.Name = "lblVariableValue"
Me.lblVariableValue.Size = New System.Drawing.Size(432, 20)
Me.lblVariableValue.TabIndex = 6
Me.lblVariableValue.Text = "Value"
Me.lblVariableValue.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'cmdCancel
'
Me.cmdCancel.Location = New System.Drawing.Point(272, 184)
Me.cmdCancel.Name = "cmdCancel"
Me.cmdCancel.Size = New System.Drawing.Size(80, 24)
Me.cmdCancel.TabIndex = 8
Me.cmdCancel.Text = "Cancel"
'
'cmdClose
'
Me.cmdClose.Location = New System.Drawing.Point(360, 184)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(80, 24)
Me.cmdClose.TabIndex = 9
Me.cmdClose.Text = "Close"
'
'cmdModify
'
Me.cmdModify.Location = New System.Drawing.Point(8, 184)
Me.cmdModify.Name = "cmdModify"
Me.cmdModify.Size = New System.Drawing.Size(80, 24)
Me.cmdModify.TabIndex = 10
Me.cmdModify.Text = "Modify"
'
'chkMatrixView
'
Me.chkMatrixView.Location = New System.Drawing.Point(96, 184)
Me.chkMatrixView.Name = "chkMatrixView"
Me.chkMatrixView.Size = New System.Drawing.Size(136, 32)
Me.chkMatrixView.TabIndex = 11
Me.chkMatrixView.Text = "View array as matrix (read-only)"
Me.chkMatrixView.Visible = False
'
'modifyVariable
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(448, 219)
Me.ControlBox = False
Me.Controls.Add(Me.chkMatrixView)
Me.Controls.Add(Me.cmdModify)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdCancel)
Me.Controls.Add(Me.txtVariableValue)
Me.Controls.Add(Me.lblVariableValue)
Me.Controls.Add(Me.txtVariableType)
Me.Controls.Add(Me.lblVariableType)
Me.Controls.Add(Me.txtStorageLocation)
Me.Controls.Add(Me.lblStorageLocation)
Me.Controls.Add(Me.txtVariableName)
Me.Controls.Add(Me.lblVariableName)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "modifyVariable"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "View and modify variable"
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private OBJvariable As qbVariable.qbVariable
    Private OBJvariableSave As qbVariable.qbVariable
    Private BOOchange As Boolean

#End Region ' " Form data "

#Region " Public properties "

    Public ReadOnly Property Change() As Boolean
        Get
            Return BOOchange
        End Get
    End Property

    Public Property VariableLocation() As Integer
        Get
            Return CInt(txtStorageLocation.Text)
        End Get
        Set(ByVal intNewValue As Integer)
            If intNewValue = -1 Then
                hideLocation()
                Return
            End If
            txtStorageLocation.Text = CStr(intNewValue)
            showLocation()
        End Set
    End Property

    Public Property Variable() As qbVariable.qbVariable
        Get
            Return OBJvariable
        End Get
        Set(ByVal objNewValue As qbVariable.qbVariable)
            OBJvariable = objNewValue
            OBJvariableSave = OBJvariable.clone
            OBJvariableSave.VariableName = OBJvariable.VariableName
            BOOchange = False
            BOOchange = False
            With OBJvariable
                txtVariableName.Text = .VariableName
                txtVariableType.Text = .Dope.ToString
                txtVariableValue.Text = getValueString
            End With
        End Set
    End Property

#End Region ' " Public properties "

#Region " Form events "

    Private Sub chkMatrixView_CheckedChanged(ByVal objSender As Object, _
                                             ByVal objEvents As System.EventArgs) _
            Handles chkMatrixView.CheckedChanged
        With txtVariableValue
            If chkMatrixView.Checked Then
                Dim booDeco As Boolean = _
                    OBJvariable.Dope.isArray _
                    AndAlso _
                    OBJvariable.Dope.VarType = qbVariableType.qbVariableType.ENUvarType.vtVariant
                .Text = OBJvariable.toMatrix(booDeco)
                .ReadOnly = True
            Else
                .Text = getValueString()
                .ReadOnly = False
            End If
            .Refresh()
        End With
    End Sub

    Private Sub cmdCancel_Click(ByVal objSender As Object, _
                                ByVal objEvents As System.EventArgs) _
            Handles cmdCancel.Click
        OBJvariable = OBJvariableSave
        Hide()
    End Sub

    Private Sub cmdClose_Click(ByVal objSender As Object, _
                               ByVal objEvents As System.EventArgs) _
            Handles cmdClose.Click
        If mkFromString() <> OBJvariable.ToString Then
            Select Case MsgBox("Note that the variable's modifications will be lost " & _
                               "if you close now. Click Yes to close, click No to " & _
                               "return to the modifier form.", _
                                MsgBoxStyle.YesNo)
                Case MsgBoxResult.Yes
                Case MsgBoxResult.No : Return
            End Select
        End If
        OBJvariableSave.dispose() : OBJvariableSave = Nothing
        Hide()
    End Sub

    Private Sub cmdModify_Click(ByVal objSender As Object, _
                                ByVal objEvents As System.EventArgs) _
            Handles cmdModify.Click
        makeChange()
        matrixOptionVisibility()
    End Sub

    Private Sub modifyVariable_Load(ByVal objSender As Object, _
                                         ByVal objEvents As System.EventArgs) _
            Handles MyBase.Load
        If (OBJvariable Is Nothing) Then
            Try
                OBJvariable = New qbVariable.qbVariable
            Catch
                errorHandler("Cannot create variable: " & Err.Number & Err.Description, _
                             Name, "modifyVariable_Activated")
                Hide()
            End Try
        End If
        matrixOptionVisibility()
    End Sub

#End Region ' " Form events "

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Obtains the value string
    '
    '
    Private Function getValueString() As String
        With OBJvariable
            Dim strFromString As String = .ToString
            Return Mid(strFromString, _
                       InStr(strFromString & ":", ":") + 1)
        End With
    End Function

    ' -----------------------------------------------------------------
    ' Hides the location
    '
    '
    Private Sub hideLocation()
        lblStorageLocation.Visible = False
        txtStorageLocation.Visible = False
        txtVariableName.Width = txtStorageLocation.Right - lblVariableName.Right
        Refresh()
    End Sub

    ' -----------------------------------------------------------------
    ' Change the variable
    '
    '
    Private Function makeChange() As Boolean
        Dim booOK As Boolean
        Dim strSerialized As String = mkFromString()
        Try
            booOK = OBJvariable.fromString(strSerialized)
        Catch
            MsgBox("Cannot change variable: " & Err.Number & " " & Err.Description)
            Exit Function
        End Try
        If Not booOK Then Exit Function
        If Not verifyChange(OBJvariableSave.ToString, _
                            strSerialized) Then
            OBJvariable.dispose()
            Me.Variable = OBJvariableSave
            Exit Function
        End If
        BOOchange = True
        makeChange = True
    End Function

    ' -----------------------------------------------------------------
    ' Matrix option visibility
    '
    '
    Private Sub matrixOptionVisibility()
        With OBJvariable.Dope
            chkMatrixView.Visible = .isArray AndAlso .Dimensions < 3
        End With
    End Sub

    ' -----------------------------------------------------------------
    ' Create serialized variable from screen
    '
    '
    Private Function mkFromString() As String
        Return txtVariableType.Text & ":" & _
               txtVariableValue.Text
    End Function

    ' -----------------------------------------------------------------
    ' Shows the location
    '
    '
    Private Sub showLocation()
        lblStorageLocation.Visible = True
        txtStorageLocation.Visible = True
        txtVariableName.Width = txtStorageLocation.Left - Grid - txtVariableName.Left
        Refresh()
    End Sub

    ' -----------------------------------------------------------------
    ' Prompt the user: verify change
    '
    '
    Private Function verifyChange(ByVal strOldToString As String, _
                                  ByVal strNewToString As String) As Boolean
        Dim strMessage As String
        Select Case MsgBox("Change the variable " & _
                           "from " & _
                           enquote(ellipsis(strOldToString, 32)) & " " & _
                           "to " & _
                           enquote(ellipsis(strNewToString, 32)) & _
                           "?", _
                           MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Yes : Return True
            Case MsgBoxResult.No : Return False
            Case MsgBoxResult.Cancel : Hide()
            Case Else
                errorHandler("Unexpected MsgBox value", _
                             Name, "verifyChange")
                Hide()
        End Select
    End Function

#End Region

End Class
