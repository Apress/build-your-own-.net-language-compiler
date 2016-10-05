Option Strict On

Imports System.Threading
Imports utilities.utilities

' *********************************************************************
' *                                                                   *
' * Run display                                                       *
' *                                                                   *
' *                                                                   *
' * This form displays the thread status and a Stop button while the  *
' * Quick Basic engine is running.  It exposes the following methods  *
' * and properties (beyond those exposed by the form).                *
' *                                                                   *
' *                                                                   *
' *      *  The new constructor has the following overloaded syntax.  *
' *                                                                   *
' *         + new: creates a run display that has 75% opacity, a      *
' *           default size, that is centered with respect to the      *
' *           screen. The Quick Basic Engine should be specified      *
' *           before the form is used.                                *
' *                                                                   *
' *         + new(qbe): creates a run display that has 75% opacity, a *
' *           default size, that is centered with respect to the      *
' *           screen. The Quick Basic Engine to be monitored is       *
' *           specified in qbe.                                       *
' *                                                                   *
' *         + new(qbe, opacity, left, top, width, height): creates a  *
' *           run display that has the indicated opacity and the      *
' *           specified geometry. The run display will monitor the    *
' *           Quick Basic Engine specified in qbe.                    *  
' *                                                                   *
' *      *  The read-write QuickBasicEngine property returns and must *
' *         be set to the Quick Basic engine to be monitored.         *
' *                                                                   *
' *                                                                   *
' * To use this form properly, Show it non-modally and then start a   *
' * Quick Basic engine method. That is because this form will stay up *
' * until the QBE enters a running state and go down when the QBE     *
' * leaves the running state.                                         *
' *                                                                   *
' *********************************************************************

Public Class frmRun
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        Dim intIndex1 As Integer
        For intIndex1 = Controls.Count - 1 To 0 Step -1
            Controls.Item(intIndex1).dispose()
        Next intIndex1
        MyBase.dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblThreadStatus As System.Windows.Forms.Label
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
Me.cmdStop = New System.Windows.Forms.Button
Me.lblThreadStatus = New System.Windows.Forms.Label
Me.SuspendLayout()
'
'cmdStop
'
Me.cmdStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.cmdStop.Location = New System.Drawing.Point(8, 8)
Me.cmdStop.Name = "cmdStop"
Me.cmdStop.Size = New System.Drawing.Size(80, 48)
Me.cmdStop.TabIndex = 0
Me.cmdStop.Text = "Stop"
'
'lblThreadStatus
'
Me.lblThreadStatus.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
Me.lblThreadStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblThreadStatus.Location = New System.Drawing.Point(96, 8)
Me.lblThreadStatus.Name = "lblThreadStatus"
Me.lblThreadStatus.Size = New System.Drawing.Size(192, 48)
Me.lblThreadStatus.TabIndex = 1
Me.lblThreadStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'frmRun
'
Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
Me.ClientSize = New System.Drawing.Size(292, 61)
Me.ControlBox = False
Me.Controls.Add(Me.lblThreadStatus)
Me.Controls.Add(Me.cmdStop)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
Me.Name = "frmRun"
Me.Opacity = 0.75
Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
Me.ResumeLayout(False)

    End Sub

#End Region

#Region " New constructor "

    ' -----------------------------------------------------------------
    ' New constructor
    '
    '
    ' --- Create object with unset geometry and unset object
    Public Sub New()
        MyBase.New()
        new_(Nothing, DEFAULT_OPACITY, Width, Height, False, 0, 0)
    End Sub
    ' --- Create object with unset geometry and a qbe
    Public Sub New(ByVal objQBE As QuickBasicEngine.qbQuickBasicEngine)
        MyBase.New()
        new_(objQBE, DEFAULT_OPACITY, Width, Height, False, 0, 0)
    End Sub
    ' --- Full Monty
    Public Sub New(ByVal objQBE As QuickBasicEngine.qbQuickBasicEngine, _
                   ByVal sglOpacity As Single, _
                   ByVal intWidth As Integer, _
                   ByVal intHeight As Integer, _
                   ByVal intLeft As Integer, _
                   ByVal intTop As Integer)
        MyBase.New()
        new_(objQBE, sglOpacity, intWidth, intHeight, True, intLeft, intTop)
    End Sub
    ' --- Common logic
    Private Sub new_(ByVal objQBE As QuickBasicEngine.qbQuickBasicEngine, _
                     ByVal sglOpacity As Single, _
                     ByVal intWidth As Integer, _
                     ByVal intHeight As Integer, _
                     ByVal booHavePosition As Boolean, _
                     ByVal intLeft As Integer, _
                     ByVal intTop As Integer)
        InitializeComponent()
        With USRstate
            If sglOpacity < 0 OrElse sglOpacity > 1 Then
                errorHandler("Opacity is invalid", Name, "new_", "Object isn't usable")
                Return
            End If
            If intWidth < 0 Then
                errorHandler("Width is invalid", Name, "new_", "Object isn't usable")
                Return
            End If
            If intHeight < 0 Then
                errorHandler("Height is invalid", Name, "new_", "Object isn't usable")
                Return
            End If
            If intLeft < 0 Then
                errorHandler("Left is invalid", Name, "new_", "Object isn't usable")
                Return
            End If
            If intTop < 0 Then
                errorHandler("Top is invalid", Name, "new_", "Object isn't usable")
                Return
            End If
            .objQuickBasicEngine = objQBE
            .sglOpacity = sglOpacity
            .intWidth = intWidth
            .intHeight = intHeight
            If booHavePosition Then
                .intLeft = intLeft
                .intTop = intTop
            Else
                .intLeft = -1 : .intTop = -1
            End If
            .booUsable = True
        End With
    End Sub

#End Region

#Region " Form data "

    ' ***** Object state *****
    Private Structure TYPstate
        Dim objQuickBasicEngine As QuickBasicEngine.qbQuickBasicEngine
        Dim booUsable As Boolean
        Dim booRunning As Boolean
        Dim sglOpacity As Single
        Dim intWidth As Integer
        Dim intHeight As Integer
        Dim intLeft As Integer
        Dim intTop As Integer
        Dim objKeepInFrontThread As Thread
    End Structure
    Private USRstate As TYPstate

    ' ***** Constants *****
    Private Const KEEPINFRONT_SLEEP As Integer = 10
    Private Const DEFAULT_OPACITY As Single = 0.75

#End Region

#Region " Friendly properties "

    Friend Property QuickBasicEngine() As quickBasicEngine.qbQuickBasicEngine
        Get
            If Not checkUsable("QuickBasicEngine Get", "Returning Nothing") Then
                Return (Nothing)
            End If
            Return USRstate.objQuickBasicEngine
        End Get
        Set(ByVal objNewValue As QuickBasicEngine.qbQuickBasicEngine)
            If Not checkUsable("QuickBasicEngine Let", "No change made") Then
                Return
            End If
            USRstate.objQuickBasicEngine = objNewValue
        End Set
    End Property

#End Region ' Friendly properties

#Region " Form events "

    Private Sub cmdStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStop.Click
        USRstate.objQuickBasicEngine.stopQBE()
    End Sub

    Private Sub frmRun_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not checkUsable("frmRun_Load", "Form not loaded") Then
            Hide()
            Return
        End If
        With USRstate
            .booUsable = False
            If (.objQuickBasicEngine Is Nothing) Then
                utilities.utilities.errorHandler("No Quick Basic engine defined", _
                                                Me.Name, _
                                                "frmRun_Load", _
                                                "Form is not usable")
                Return
            End If
            Width = .intWidth : Top = .intTop
            If .intLeft > -1 Then
                Left = .intLeft : Top = .intTop
            Else
                CenterToScreen()
            End If
            .booRunning = False
            AddHandler USRstate.objQuickBasicEngine.threadStatusChangeEvent, _
                       AddressOf threadStatusChangeEventHandler
            Try
                .objKeepInFrontThread = New Thread(AddressOf keepInFront)
                .objKeepInFrontThread.Start()
            Catch : End Try
            .booUsable = True
            Opacity = .sglOpacity
            Show()
            Refresh()
            threadStatusChangeEventHandler(.objQuickBasicEngine)
            lblThreadStatus.Width = .intWidth _
                                    - _
                                    cmdStop.Width _
                                    - _
                                    windowsUtilities.windowsUtilities.Grid * 4
        End With
    End Sub

#End Region

#Region " General procedures "

    ' -----------------------------------------------------------------
    ' Check the form's usability
    '
    '
    Private Function checkUsable(ByVal strProcedure As String, _
                                 ByVal strPolicy As String) As Boolean
        If Not USRstate.booUsable Then
            errorHandler("Form is not usable", _
                         Name, strProcedure, _
                         strPolicy)
            Return False
        End If
        Return True
    End Function

    ' -----------------------------------------------------------------
    ' Make certain form stays in front: executed as an independent
    ' thread
    '
    '
    Private Sub keepInFront()
        Me.BringToFront() : Me.Refresh()
        System.Threading.Thread.CurrentThread.Sleep(KEEPINFRONT_SLEEP)
    End Sub

    ' -----------------------------------------------------------------
    ' Convert thread status to a color scheme
    '
    '
    Private Sub status2Colors(ByVal strThreadStatus As String, _
                              ByRef objBG As Color, _
                              ByRef objFG As Color)
        Select Case UCase(strThreadStatus)
            Case "INITIALIZING"
                objBG = Color.Aquamarine
                objFG = Color.Black
            Case "READY"
                objBG = Color.LightGreen
                objFG = Color.Black
            Case "RUNNING"
                objBG = Color.Green
                objFG = Color.White
            Case "STOPPING"
                objBG = Color.Yellow
                objFG = Color.Black
            Case "STOPPED"
                objBG = Color.Red
                objFG = Color.White
            Case Else
                objBG = Color.Black
                objFG = Color.White
        End Select
    End Sub

#End Region ' General procedures

#Region " Form events "

    Private Sub threadStatusChangeEventHandler(ByVal objQBE As QuickBasicEngine.qbQuickBasicEngine)
        If Not USRstate.booUsable Then Return
        With objQBE
            Dim strStatus As String = .getThreadStatus
            lblThreadStatus.Text = strStatus & ": " & .runningThreads & " threads running"
            Dim objBG As Color
            Dim objFG As Color
            status2Colors(strStatus, objBG, objFG)
            With lblThreadStatus
                .BackColor = objBG
                .ForeColor = objFG
            End With
        End With
    End Sub

#End Region

End Class
