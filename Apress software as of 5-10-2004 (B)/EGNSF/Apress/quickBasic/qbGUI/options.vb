Option Strict On

Imports utilities.utilities
Imports windowsUtilities.windowsUtilities
Imports quickBasicEngine.qbQuickBasicEngine

' *********************************************************************
' *                                                                   *
' * qbGUI Option Form                                                 *
' *                                                                   *
' *                                                                   *
' * This class exposes the following methods and properties:          *
' *                                                                   *
' *                                                                   *
' *      *  The read-write ArrayValueDisplay property returns and may *
' *         be set to True to display array values inside the Storage *
' *         display visible in the More mode. By default and to save  *
' *         time, array values are not serialized in this display,    *
' *         only types and bounds, while the types and values of other*
' *         variable types are shown in full.                         *
' *                                                                   *
' *         Note that this property is exposed only when the          *
' *         compile-time symbol QBGUI_EXTENSIONS is True.             *
' *                                                                   *
' *      *  The read-write ConstantFolding property corresponds to the*
' *         ConstantFolding property of the Quick Basic engine.       *
' *                                                                   *
' *      *  The read-write DegenerateOpRemoval property corresponds to*
' *         the DegenerateOpRemoval property of the Quick Basic       *
' *         engine.                                                   *
' *                                                                   *
' *      *  The read-write EventLogging property corresponds to       *
' *         the EventLogging property of the Quick Basic engine.      *
' *                                                                   *
' *      *  The read-write GenerateNOPs property returns and may be   *
' *         set to True if the compiler should generate NOP and REM   *
' *         instructions (primarily as documentation), False          *
' *         otherwise.                                                *
' *                                                                   *
' *         Note that this property is exposed only when the          *
' *         compile-time symbol QBGUI_EXTENSIONS is True.             *
' *                                                                   *
' *      *  The getParseDisplay method returns the mode in which      *
' *         parse information is displayed as an enumeration of type  *
' *         ENUparseDisplay:                                          *
' *                                                                   *
' *              + Outline: the parse display is created in an        *
' *                outline format.                                    *
' *                                                                   *
' *              + XML: the parse display is created in the XML       *
' *                format.                                            *
' *                                                                   *
' *              + NoDisplay: no display is created.                  *
' *                                                                   *
' *      *  The read-write Inspection property causes the Quick Basic *
' *         engine to be inspected after scan, compile, assembly and  *
' *         execute.                                                  *
' *                                                                   *
' *      *  The read-write InspectCompilerObjects property causes the *
' *         Quick Basic engine to inspect compiler objects when these *
' *         are disposed. Objects inspected include the scanner,      *
' *         Polish tokens, variables and variable types.              *
' *                                                                   *
' *         The default value of this property is False since it      *
' *         slows the compiler down. Setting it to True creates a test*
' *         of new versions of the compiler.                          *
' *                                                                   *
' *      *  The read-write ObjectTrace property causes an object code *
' *         trace to appear during execution on the main form.        *
' *                                                                   *
' *      *  The read-write ParseTrace property causes a parse         *
' *         trace to appear during execution on the main form; note   *
' *         that this trace will slow parsing, because it shows both  *
' *         failed and successful attempts.                           *
' *                                                                   *
' *      *  The setParseDisplay(m) method changes the mode in which   *
' *         parse information is displayed. m may be a string or an   *
' *         enumerator of type ENUparse display: it may have one of   *
' *         the following values.                                     *
' *                                                                   *
' *              + Outline: the parse display is created in an        *
' *                outline format.                                    *
' *                                                                   *
' *              + XML: the parse display is created in the XML       *
' *                format.                                            *
' *                                                                   *
' *              + NoDisplay: no display is created.                  *
' *                                                                   *
' *      *  The read-write SourceTrace property causes a source code  *
' *         trace to appear during execution on the main form.        *
' *                                                                   *
' *      *  The read-write SourceTraceOpacity property controls the   *
' *         opacity of the source trace screen that overlays the      *
' *         object trace screen when both forms of trace are in effect*
' *                                                                   *
' *      *  The read-write StopButton property controls the           *
' *         presentation of a stop button which can stop an ongoing or*
' *         runaway lexical analysis, assembly or interpretive        *
' *         execution.                                                *
' *                                                                   *
' *      *  The read-only WasClosed property returns True when the    *
' *         form was dismissed using the Close button.                *
' *                                                                   *
' *                                                                   *
' * To add a Public property, supported by a control on this form:    *
' *                                                                   *
' *                                                                   *
' *      1.  Draw and name the control that corresponds to the        *
' *          property: on the form set its default value (or, add code*
' *          to dynamically draw the control and set its default:     *
' *          this has been done for the QBGUI_EXTENSIONS properties). *
' *                                                                   *
' *      2.  Add documentation of the property to the above list.     *
' *                                                                   *
' *      3.  Write the Property procedure in this file.               *
' *                                                                   *
' *      4.  Add code for the control's property to the following     *
' *          procedures in form1.vb, using comments in those          *
' *          procedures as a guide:                                   *
' *                                                                   *
' *          4.1 showOptions                                          *
' *          4.2 setQBEoptions                                        *
' *          4.3 form2Registry                                        *
' *          4.4 registry2Form                                        *
' *                                                                   *
' *********************************************************************

Public Class options
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

#If QBGUI_EXTENSIONS Then
        chkOptimizationAssemblyRemovesCode.Enabled = False
        createExtensions()
#End If

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        disposer()
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
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents gbxOptimization As System.Windows.Forms.GroupBox
    Friend WithEvents chkOptimizationConstantFolding As System.Windows.Forms.CheckBox
    Friend WithEvents chkOptimizationAssemblyRemovesCode As System.Windows.Forms.CheckBox
    Friend WithEvents chkOptimizationDegenerateOpRemoval As System.Windows.Forms.CheckBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents gbxTracing As System.Windows.Forms.GroupBox
    Friend WithEvents chkTracingObject As System.Windows.Forms.CheckBox
    Friend WithEvents chkTracingSource As System.Windows.Forms.CheckBox
    Friend WithEvents chkTracingParse As System.Windows.Forms.CheckBox
    Friend WithEvents chkMiscInspection As System.Windows.Forms.CheckBox
    Friend WithEvents chkMiscEventLogging As System.Windows.Forms.CheckBox
    Friend WithEvents chkMiscStopButton As System.Windows.Forms.CheckBox
    Friend WithEvents gbxMisc As System.Windows.Forms.GroupBox
    Friend WithEvents gbxParseDisplay As System.Windows.Forms.GroupBox
    Friend WithEvents radParseDisplayOutline As System.Windows.Forms.RadioButton
    Friend WithEvents radParseDisplayXML As System.Windows.Forms.RadioButton
    Friend WithEvents radParseDisplayNone As System.Windows.Forms.RadioButton
    Friend WithEvents chkOptimizationInspectCompilerObjects As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.chkMiscInspection = New System.Windows.Forms.CheckBox
Me.cmdCancel = New System.Windows.Forms.Button
Me.cmdClose = New System.Windows.Forms.Button
Me.gbxOptimization = New System.Windows.Forms.GroupBox
Me.chkOptimizationInspectCompilerObjects = New System.Windows.Forms.CheckBox
Me.chkOptimizationDegenerateOpRemoval = New System.Windows.Forms.CheckBox
Me.chkOptimizationAssemblyRemovesCode = New System.Windows.Forms.CheckBox
Me.chkOptimizationConstantFolding = New System.Windows.Forms.CheckBox
Me.gbxTracing = New System.Windows.Forms.GroupBox
Me.chkTracingParse = New System.Windows.Forms.CheckBox
Me.chkTracingObject = New System.Windows.Forms.CheckBox
Me.chkTracingSource = New System.Windows.Forms.CheckBox
Me.chkMiscEventLogging = New System.Windows.Forms.CheckBox
Me.chkMiscStopButton = New System.Windows.Forms.CheckBox
Me.gbxMisc = New System.Windows.Forms.GroupBox
Me.gbxParseDisplay = New System.Windows.Forms.GroupBox
Me.radParseDisplayNone = New System.Windows.Forms.RadioButton
Me.radParseDisplayXML = New System.Windows.Forms.RadioButton
Me.radParseDisplayOutline = New System.Windows.Forms.RadioButton
Me.gbxOptimization.SuspendLayout()
Me.gbxTracing.SuspendLayout()
Me.gbxMisc.SuspendLayout()
Me.gbxParseDisplay.SuspendLayout()
Me.SuspendLayout()
'
'chkMiscInspection
'
Me.chkMiscInspection.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkMiscInspection.Location = New System.Drawing.Point(8, 56)
Me.chkMiscInspection.Name = "chkMiscInspection"
Me.chkMiscInspection.Size = New System.Drawing.Size(200, 24)
Me.chkMiscInspection.TabIndex = 3
Me.chkMiscInspection.Text = "Inspect Quick Basic Engine"
'
'cmdCancel
'
Me.cmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdCancel.Location = New System.Drawing.Point(232, 280)
Me.cmdCancel.Name = "cmdCancel"
Me.cmdCancel.Size = New System.Drawing.Size(128, 32)
Me.cmdCancel.TabIndex = 4
Me.cmdCancel.Text = "Cancel"
'
'cmdClose
'
Me.cmdClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdClose.Location = New System.Drawing.Point(232, 368)
Me.cmdClose.Name = "cmdClose"
Me.cmdClose.Size = New System.Drawing.Size(128, 32)
Me.cmdClose.TabIndex = 5
Me.cmdClose.Text = "Close"
'
'gbxOptimization
'
Me.gbxOptimization.Controls.Add(Me.chkOptimizationInspectCompilerObjects)
Me.gbxOptimization.Controls.Add(Me.chkOptimizationDegenerateOpRemoval)
Me.gbxOptimization.Controls.Add(Me.chkOptimizationAssemblyRemovesCode)
Me.gbxOptimization.Controls.Add(Me.chkOptimizationConstantFolding)
Me.gbxOptimization.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxOptimization.Location = New System.Drawing.Point(10, 9)
Me.gbxOptimization.Name = "gbxOptimization"
Me.gbxOptimization.Size = New System.Drawing.Size(350, 139)
Me.gbxOptimization.TabIndex = 6
Me.gbxOptimization.TabStop = False
Me.gbxOptimization.Text = "Optimization"
'
'chkOptimizationInspectCompilerObjects
'
Me.chkOptimizationInspectCompilerObjects.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkOptimizationInspectCompilerObjects.Location = New System.Drawing.Point(8, 111)
Me.chkOptimizationInspectCompilerObjects.Name = "chkOptimizationInspectCompilerObjects"
Me.chkOptimizationInspectCompilerObjects.Size = New System.Drawing.Size(328, 18)
Me.chkOptimizationInspectCompilerObjects.TabIndex = 3
Me.chkOptimizationInspectCompilerObjects.Text = "Inspect compiler objects"
'
'chkOptimizationDegenerateOpRemoval
'
Me.chkOptimizationDegenerateOpRemoval.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkOptimizationDegenerateOpRemoval.Location = New System.Drawing.Point(8, 83)
Me.chkOptimizationDegenerateOpRemoval.Name = "chkOptimizationDegenerateOpRemoval"
Me.chkOptimizationDegenerateOpRemoval.Size = New System.Drawing.Size(328, 19)
Me.chkOptimizationDegenerateOpRemoval.TabIndex = 2
Me.chkOptimizationDegenerateOpRemoval.Text = "Remove degenerate operations"
'
'chkOptimizationAssemblyRemovesCode
'
Me.chkOptimizationAssemblyRemovesCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkOptimizationAssemblyRemovesCode.Location = New System.Drawing.Point(8, 55)
Me.chkOptimizationAssemblyRemovesCode.Name = "chkOptimizationAssemblyRemovesCode"
Me.chkOptimizationAssemblyRemovesCode.Size = New System.Drawing.Size(328, 19)
Me.chkOptimizationAssemblyRemovesCode.TabIndex = 1
Me.chkOptimizationAssemblyRemovesCode.Text = "Remove comments && labels during assembly"
'
'chkOptimizationConstantFolding
'
Me.chkOptimizationConstantFolding.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkOptimizationConstantFolding.Location = New System.Drawing.Point(8, 28)
Me.chkOptimizationConstantFolding.Name = "chkOptimizationConstantFolding"
Me.chkOptimizationConstantFolding.Size = New System.Drawing.Size(328, 18)
Me.chkOptimizationConstantFolding.TabIndex = 0
Me.chkOptimizationConstantFolding.Text = "Constant Folding"
'
'gbxTracing
'
Me.gbxTracing.Controls.Add(Me.chkTracingParse)
Me.gbxTracing.Controls.Add(Me.chkTracingObject)
Me.gbxTracing.Controls.Add(Me.chkTracingSource)
Me.gbxTracing.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxTracing.Location = New System.Drawing.Point(232, 152)
Me.gbxTracing.Name = "gbxTracing"
Me.gbxTracing.Size = New System.Drawing.Size(128, 120)
Me.gbxTracing.TabIndex = 7
Me.gbxTracing.TabStop = False
Me.gbxTracing.Text = "Tracing"
'
'chkTracingParse
'
Me.chkTracingParse.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkTracingParse.Location = New System.Drawing.Point(8, 88)
Me.chkTracingParse.Name = "chkTracingParse"
Me.chkTracingParse.Size = New System.Drawing.Size(96, 19)
Me.chkTracingParse.TabIndex = 5
Me.chkTracingParse.Text = "Parse trace"
'
'chkTracingObject
'
Me.chkTracingObject.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkTracingObject.Location = New System.Drawing.Point(8, 56)
Me.chkTracingObject.Name = "chkTracingObject"
Me.chkTracingObject.Size = New System.Drawing.Size(104, 19)
Me.chkTracingObject.TabIndex = 4
Me.chkTracingObject.Text = "Object trace"
'
'chkTracingSource
'
Me.chkTracingSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkTracingSource.Location = New System.Drawing.Point(8, 24)
Me.chkTracingSource.Name = "chkTracingSource"
Me.chkTracingSource.Size = New System.Drawing.Size(102, 18)
Me.chkTracingSource.TabIndex = 3
Me.chkTracingSource.Text = "Source trace"
'
'chkMiscEventLogging
'
Me.chkMiscEventLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkMiscEventLogging.Location = New System.Drawing.Point(8, 24)
Me.chkMiscEventLogging.Name = "chkMiscEventLogging"
Me.chkMiscEventLogging.Size = New System.Drawing.Size(200, 24)
Me.chkMiscEventLogging.TabIndex = 8
Me.chkMiscEventLogging.Text = "Event Log"
'
'chkMiscStopButton
'
Me.chkMiscStopButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.chkMiscStopButton.Location = New System.Drawing.Point(8, 88)
Me.chkMiscStopButton.Name = "chkMiscStopButton"
Me.chkMiscStopButton.Size = New System.Drawing.Size(200, 24)
Me.chkMiscStopButton.TabIndex = 9
Me.chkMiscStopButton.Text = "Stop Button"
'
'gbxMisc
'
Me.gbxMisc.Controls.Add(Me.chkMiscEventLogging)
Me.gbxMisc.Controls.Add(Me.chkMiscInspection)
Me.gbxMisc.Controls.Add(Me.chkMiscStopButton)
Me.gbxMisc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxMisc.Location = New System.Drawing.Point(8, 280)
Me.gbxMisc.Name = "gbxMisc"
Me.gbxMisc.Size = New System.Drawing.Size(216, 120)
Me.gbxMisc.TabIndex = 10
Me.gbxMisc.TabStop = False
Me.gbxMisc.Text = "Miscellaneous"
'
'gbxParseDisplay
'
Me.gbxParseDisplay.Controls.Add(Me.radParseDisplayNone)
Me.gbxParseDisplay.Controls.Add(Me.radParseDisplayXML)
Me.gbxParseDisplay.Controls.Add(Me.radParseDisplayOutline)
Me.gbxParseDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.gbxParseDisplay.Location = New System.Drawing.Point(8, 152)
Me.gbxParseDisplay.Name = "gbxParseDisplay"
Me.gbxParseDisplay.Size = New System.Drawing.Size(216, 120)
Me.gbxParseDisplay.TabIndex = 12
Me.gbxParseDisplay.TabStop = False
Me.gbxParseDisplay.Text = "Parse Display"
'
'radParseDisplayNone
'
Me.radParseDisplayNone.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radParseDisplayNone.Location = New System.Drawing.Point(8, 24)
Me.radParseDisplayNone.Name = "radParseDisplayNone"
Me.radParseDisplayNone.Size = New System.Drawing.Size(200, 24)
Me.radParseDisplayNone.TabIndex = 2
Me.radParseDisplayNone.Text = "No parse display"
'
'radParseDisplayXML
'
Me.radParseDisplayXML.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radParseDisplayXML.Location = New System.Drawing.Point(8, 88)
Me.radParseDisplayXML.Name = "radParseDisplayXML"
Me.radParseDisplayXML.Size = New System.Drawing.Size(200, 24)
Me.radParseDisplayXML.TabIndex = 1
Me.radParseDisplayXML.Text = "XML formatted parse display"
'
'radParseDisplayOutline
'
Me.radParseDisplayOutline.Checked = True
Me.radParseDisplayOutline.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.radParseDisplayOutline.Location = New System.Drawing.Point(8, 56)
Me.radParseDisplayOutline.Name = "radParseDisplayOutline"
Me.radParseDisplayOutline.Size = New System.Drawing.Size(200, 24)
Me.radParseDisplayOutline.TabIndex = 0
Me.radParseDisplayOutline.TabStop = True
Me.radParseDisplayOutline.Text = "Outline parse display"
'
'options
'
Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
Me.ClientSize = New System.Drawing.Size(366, 406)
Me.ControlBox = False
Me.Controls.Add(Me.gbxParseDisplay)
Me.Controls.Add(Me.gbxTracing)
Me.Controls.Add(Me.gbxOptimization)
Me.Controls.Add(Me.cmdClose)
Me.Controls.Add(Me.cmdCancel)
Me.Controls.Add(Me.gbxMisc)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "options"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "Quick Basic Options"
Me.gbxOptimization.ResumeLayout(False)
Me.gbxTracing.ResumeLayout(False)
Me.gbxMisc.ResumeLayout(False)
Me.gbxParseDisplay.ResumeLayout(False)
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form data "

    Private BOOclosed As Boolean
    Private FRMpreview As Form
    Public Enum ENUparseDisplay
        NoDisplay
        Outline
        XML
    End Enum
    Private OBJtoolTips As ToolTip
#If QBGUI_EXTENSIONS Then
    Dim WithEvents CHKoptimizationGenerateNOPs As CheckBox
    Dim CHKmiscArrayValueDisplay As CheckBox
#End If

#End Region ' " Form data "

#Region " Public properties and methods "

#If QBGUI_EXTENSIONS Then
    Public Property ArrayValueDisplay() As Boolean
        Get
            Return (CHKmiscArrayValueDisplay.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            CHKmiscArrayValueDisplay.Checked = booNewValue
        End Set
    End Property
#End If

    Public Property AssemblyRemovesCode() As Boolean
        Get
            Return (chkOptimizationAssemblyRemovesCode.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkOptimizationAssemblyRemovesCode.Checked = booNewValue
        End Set
    End Property

    Public Property ConstantFolding() As Boolean
        Get
            Return (chkOptimizationConstantFolding.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkOptimizationConstantFolding.Checked = booNewValue
        End Set
    End Property

    Public Property DegenerateOpRemoval() As Boolean
        Get
            Return (chkOptimizationDegenerateOpRemoval.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkOptimizationDegenerateOpRemoval.Checked = booNewValue
        End Set
    End Property

    Public Property EventLogging() As Boolean
        Get
            Return (chkMiscEventLogging.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkMiscEventLogging.Checked = booNewValue
        End Set
    End Property

#If QBGUI_EXTENSIONS Then

    Public Property GenerateNOPs() As Boolean
        Get
            Return (CHKoptimizationGenerateNOPs.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            CHKoptimizationGenerateNOPs.Checked = booNewValue
        End Set
    End Property

#End If

    Public Function getParseDisplay() As ENUparseDisplay
        If radParseDisplayNone.Checked Then
            Return ENUparseDisplay.NoDisplay
        ElseIf radParseDisplayOutline.Checked Then
            Return ENUparseDisplay.Outline
        ElseIf radParseDisplayXML.Checked Then
            Return ENUparseDisplay.XML
        Else
            radParseDisplayOutline.Checked = True
        End If
    End Function

    Public Property Inspection() As Boolean
        Get
            Return (chkMiscInspection.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkMiscInspection.Checked = booNewValue
        End Set
    End Property

    Public Property InspectCompilerObjects() As Boolean
        Get
            Return (chkOptimizationInspectCompilerObjects.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkOptimizationInspectCompilerObjects.Checked = booNewValue
        End Set
    End Property

    Public Property ObjectTrace() As Boolean
        Get
            Return (chkTracingObject.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkTracingObject.Checked = booNewValue
        End Set
    End Property

    Public Property ParseTrace() As Boolean
        Get
            Return (chkTracingParse.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkTracingParse.Checked = booNewValue
        End Set
    End Property

    Public Overloads Sub setParseDisplay(ByVal strNewValue As String)
        Select Case UCase(Trim(strNewValue))
            Case "OUTLINE"
                setParseDisplay(ENUparseDisplay.Outline)
            Case "XML"
                setParseDisplay(ENUparseDisplay.XML)
            Case "NODISPLAY"
                setParseDisplay(ENUparseDisplay.NoDisplay)
            Case Else
                errorHandler("Invalid display mode " & _
                             enquote(strNewValue), _
                             Name, "setDisplayMode", _
                             "Display mode not changed")
        End Select
    End Sub

    Public Overloads Sub setParseDisplay(ByVal enuNewValue As ENUparseDisplay)
        Select Case enuNewValue
            Case ENUparseDisplay.NoDisplay
                radParseDisplayNone.Checked = True
            Case ENUparseDisplay.Outline
                radParseDisplayOutline.Checked = True
            Case ENUparseDisplay.XML
                radParseDisplayXML.Checked = True
            Case Else
                errorHandler("Unexpected case", _
                             Name, "setDisplayMode", _
                             "Display mode not changed")
        End Select
    End Sub

    Public Property SourceTrace() As Boolean
        Get
            Return (chkTracingSource.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkTracingSource.Checked = booNewValue
        End Set
    End Property

    Public Property StopButton() As Boolean
        Get
            Return (chkMiscStopButton.Checked)
        End Get
        Set(ByVal booNewValue As Boolean)
            chkMiscStopButton.Checked = booNewValue
        End Set
    End Property

    Public ReadOnly Property WasClosed() As Boolean
        Get
            Return (BOOclosed)
        End Get
    End Property

#End Region ' " Public properties and methods "

#Region " Form Events "

    Private Sub cmdCancel_Click(ByVal objSender As System.Object, _
                                ByVal obEventArgs As System.EventArgs) _
                Handles cmdCancel.Click
        Hide()
    End Sub

    Private Sub cmdClose_Click(ByVal objSender As System.Object, _
                               ByVal obEventArgs As System.EventArgs) _
                Handles cmdClose.Click
        BOOclosed = True
        Hide()
    End Sub

    Private Sub CHKoptimizationGenerateNOPs_CheckedChanged(ByVal objSender As System.Object, _
                                                  ByVal obEventArgs As System.EventArgs)
        chkOptimizationAssemblyRemovesCode.Enabled = CHKoptimizationGenerateNOPs.Checked
    End Sub

    Private Sub options_Load(ByVal objSender As System.Object, _
                             ByVal obEventArgs As System.EventArgs) _
                Handles MyBase.Load
        BOOclosed = False
        Try
            OBJtoolTips = New ToolTip
        Catch
            errorHandler("Cannot create ToolTip object: " & _
                         Err.Number & " " & Err.Description, _
                         Name, _
                         "options_Load", _
                         "Continuing: no Tool Tips will be available")
        End Try
        If Not (OBJtoolTips Is Nothing) Then setToolTips()
    End Sub

#End Region ' " Form Events "

#Region " General procedures "

#If QBGUI_EXTENSIONS Then

    ' -------------------------------------------------------------------------
    ' Create the extension controls
    '
    '
    Private Sub createExtensions()
        createExtensions_insertCheckBox(chkMiscStopButton, _
                       "CHKmiscArrayValueDisplay", _
                       "Display array values", _
                       CHKmiscArrayValueDisplay)
        createExtensions_insertCheckBox(chkOptimizationConstantFolding, _
                       "CHKoptimizationGenerateNOPs", _
                       "Generate NOP and REM opcodes", _
                       CHKoptimizationGenerateNOPs)
        AddHandler CHKoptimizationGenerateNOPs.CheckedChanged, _
                   AddressOf CHKoptimizationGenerateNOPs_CheckedChanged
        Try
            Dim gbxInfo As GroupBox = New GroupBox
            With gbxInfo
                Dim ctlLowest As Control
                Dim intIndex1 As Integer
                For intIndex1 = 0 To Me.Controls.Count - 1
                    If (ctlLowest Is Nothing) _
                       OrElse _
                       Controls(intIndex1).Bottom > ctlLowest.Bottom Then
                        ctlLowest = Controls(intIndex1)
                    End If
                Next intIndex1
                .Left = Grid
                .Width = Me.ClientSize.Width - Grid * 2
                .Top = ctlLowest.Bottom + Grid
                .Height *= 2
                .Text = "Compile Status"
                .Font = New Font(.Font, FontStyle.Bold)
            End With
            Dim lblInfo As Label = New Label
            With lblInfo
                .Width = gbxInfo.Width - Grid * 2
                .Left = Grid
                .Top = Grid * 2
                .Height = gbxInfo.Height - Grid * 3
                .Text = "The following settings were in effect for the " & _
                        "generation of the quickBasicEngine" & _
                        vbNewLine & vbNewLine & vbNewLine & _
                        "EXTENSION: " & CStr(ExtensionAvailable) & ": " & _
                        "compiler extensions will " & _
                        CStr(IIf(ExtensionAvailable, "", "not ")) & _
                        "be available" & _
                        vbNewLine & vbNewLine & _
                        "POPCHECK: " & CStr(PopCheckAvailable) & ": " & _
                        "extra stack checks will " & _
                        CStr(IIf(PopCheckAvailable, "", "not ")) & _
                        "be performed"
                .TextAlign = ContentAlignment.MiddleLeft
                gbxInfo.Controls.Add(lblInfo)
                .Font = New Font(.Font, FontStyle.Regular)
            End With
            With Me
                .ClientSize = New Size(.ClientSize.Width, gbxInfo.Bottom + Grid)
                .Controls.Add(gbxInfo)
                .CenterToScreen()
                .Refresh()
            End With
        Catch
            errorHandler("Cannot create status label: " & _
                         Err.Number & " " & Err.Description, _
                         Me.Name, _
                         "createExtensions")
            Return
        End Try
    End Sub

    ' -------------------------------------------------------------------------
    ' Create the extension controls on behalf of createExtensions
    '
    '
    Private Sub createExtensions_insertCheckBox(ByVal ctlPreceding As Control, _
                                                ByVal strName As String, _
                                                ByVal strText As String, _
                                                ByRef chkHandle As CheckBox)
        Try
            chkHandle = New CheckBox
        Catch
            errorHandler("Cannot create extension check box: " & _
                         Err.Number & " " & Err.Description, _
                         Me.Name, _
                         "createExtensions_insertCheckBox")
            Return
        End Try
        Dim intHeightChange As Integer = ctlPreceding.Height + Grid
        With chkHandle
            .Name = strName
            .Text = strText
            .Left = ctlPreceding.Left
            .Width = ctlPreceding.Width
            .Top = ctlPreceding.Top + intHeightChange
            .Font = New Font(ctlPreceding.Font, FontStyle.Regular)
        End With
        ctlPreceding.Parent.Controls.Add(chkHandle)
        createExtensions_shiftControls(chkHandle, intHeightChange)
        createExtensions_shiftControls(chkHandle.Parent, intHeightChange + Grid)
    End Sub

    ' -------------------------------------------------------------------------
    ' Shift controls below control on behalf of createExtensions
    '
    '
    Private Sub createExtensions_shiftControls(ByVal ctlHandle As Control, _
                                               ByVal intHeightChange As Integer)
        With ctlHandle.Parent
            Dim intIndex1 As Integer
            For intIndex1 = 0 To .Controls.Count - 1
                With .Controls(intIndex1)
                    If .Top > ctlHandle.Top Then
                        .Top += intHeightChange
                    End If
                End With
            Next intIndex1
            If (TypeOf ctlHandle.Parent Is Form) Then
                With CType(ctlHandle.Parent, Form)
                    .ClientSize = New Size(.ClientSize.Width, _
                                           .ClientSize.Height + intHeightChange)
                End With
            Else
                .Height += intHeightChange
            End If
        End With
    End Sub

#End If

    ' -------------------------------------------------------------------------
    ' Disposer
    '
    '
    Public Sub disposer()
        If Not (FRMpreview Is Nothing) Then disposePreview()
    End Sub

    ' -------------------------------------------------------------------------
    ' Dispose preview
    '
    '
    Public Sub disposePreview()
        FRMpreview.Visible = False
        FRMpreview.dispose()
        FRMpreview = Nothing
    End Sub

    ' -------------------------------------------------------------------------
    ' Set tool tips
    '
    '
    Private Sub setToolTips()
        With OBJtoolTips
            .SetToolTip(chkMiscEventLogging, _
                        "Enables event logging in the quickBasicEngine")
            .SetToolTip(chkMiscInspection, _
                        "Enables continuous inspection in the quickBasicEngine")
            .SetToolTip(chkMiscStopButton, _
                        "Provides a Stop button for interrupting compilation and execution")
            .SetToolTip(chkOptimizationAssemblyRemovesCode, _
                        "Causes unnecessary code (comments, no-operations, and " & _
                        "labels) to be removed from object code in the assembler")
            .SetToolTip(chkOptimizationConstantFolding, _
                        "Enables the ""folding"" (compile-time calculation " & _
                        "of constant expressions")
            .SetToolTip(chkOptimizationDegenerateOpRemoval, _
                        "Enables the removal of ""degenerate"" (do-nothing) ops")
            .SetToolTip(chkTracingObject, _
                        "Select to enable object code tracing")
            .SetToolTip(chkTracingParse, _
                        "Select to enable parse tracing")
            .SetToolTip(chkTracingSource, _
                        "Select to enable source tracing")
            .SetToolTip(cmdCancel, _
                        "Dismisses this form: does not change options")
            .SetToolTip(cmdClose, _
                        "Dismisses this form: changes options as indicated")
            .SetToolTip(radParseDisplayNone, _
                        "The parse detail is not displayed when this Radio button is selected")
            .SetToolTip(radParseDisplayOutline, _
                        "The parse detail is in outline form when this Radio button is selected")
            .SetToolTip(radParseDisplayXML, _
                        "The parse detail is in XML form when this Radio button is selected")
#If QBGUI_EXTENSIONS Then
            .SetToolTip(CHKmiscArrayValueDisplay, _
                        "Show array values in the storage display")
            .SetToolTip(CHKoptimizationGenerateNOPs, _
                        "Generate NOP and REM instructions in the object code")
#End If
        End With
    End Sub

#End Region ' " General procedures "

End Class
