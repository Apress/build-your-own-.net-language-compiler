VERSION 5.00
Begin VB.Form frmErrorHandler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      ToolTipText     =   "Resume normal execution"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      ToolTipText     =   "For some errors will cause the application to be dismissed or the activity canceled"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CheckBox chkDisplay 
      Caption         =   "Display this Form"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Controls whether this form is shown"
      Top             =   6360
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox chkRaise 
      Caption         =   "Raise errors"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Controls whether errors are raised when the errorHandler is called"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtError 
      BackColor       =   &H00D0D0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Displays an error or purely informational text"
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frmErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************
' *                                                                   *
' * errorHandler     Display error, giving the user the choice of     *
' *                  raising and/or displaying subsequent errors      *
' *                                                                   *
' *                                                                   *
' * This form is used by the errorHandler utility in utilities.DLL    *
' * to display errors.  By default it appears which can be problem-   *
' * atic for server-based applications.  However, the user can then   *
' * check off its Form check box and check on its Raise check box     *
' * to identify the policy needed.                                    *
' *                                                                   *
' * For best results and when using this form in an application, have *
' * your Sub Main or initial form display an innocuous message that   *
' * notifies the installer that this form exists, and allows the      *
' * installer to decide on the policy.  In order to allow people to   *
' * change the policy also implement controls on your flagship form,  *
' * which force this form to appear.  Use the errorHandler utility in *
' * utilities.CLS to display errors and messages, initially and by    *
' * default using this form.                                          *
' *                                                                   *
' * This startup code needs to assign the application's name to the   *
' * write-only Application property of this form in most cases.  If   *
' * this is not done then a default location (utilities) will be used.*
' *                                                                   *
' * Here is the suggested wording of your innocuous message: note that*
' * this wording is returned by the InstallationMessage property of   *
' * this form.                                                        *
' *                                                                   *
' *                                                                   *
' *      Errors from this application can be displayed using this     *
' *      form, or they may be Raised, or both, or neither (this isn't *
' *      recommended, for obvious reasons.)                           *
' *                                                                   *
' *      Check Display this Form to always display this message box   *
' *      on an error.  Check Raise an Error to always Raise an error. *
' *      Check Continue to proceed with these settings.               *
' *                                                                   *
' *                                                                   *
' * Here is an example of the code to do the above, for your Sub Main *
' * or startup form.                                                  *
' *                                                                   *
' *                                                                   *
' *     Private Sub errorInstallation()                               *
' *          '                                                        *
' *          ' Error policy set up                                    *
' *          '                                                        *
' *          '                                                        *
' *          Set FRMerrorHandler = OBJutilities.errorHandlerForm      *
' *          Load FRMerrorHandler                                     *
' *          With FRMerrorHandler                                     *
' *              If Not .LoadDefaults(App.EXEName, _                  *
' *                                   FRMerrorHandler.Name) Then      *
' *                  errorInstallation_                               *
' *              End If                                               *
' *          End With                                                 *
' *      End Sub                                                      *
' *      Private Sub errorInstallation_()                             *
' *          '                                                        *
' *          ' Error policy change                                    *
' *          '                                                        *
' *          '                                                        *
' *          With FRMerrorHandler                                     *
' *              .Application = App.EXEName                           *
' *              .Section = FRMerrorHandler.Name                      *
' *              On Error Resume Next                                 *
' *                  OBJutilities.errorHandler 1, _                   *
' *                                            .InstallationMessage, _*
' *                                            "Form_Load", _         *
' *                                            Me.Name, _             *
' *                                            True, False            *
' *              On Error GoTo 0                                      *
' *          End With                                                 *
' *      End Sub                                                      *
' *                                                                   *
' *                                                                   *
' * Use the errorInstallation procedure in the form load to set up    *
' * error handling.  Use the errorInstallation_ procedure to change   *
' * error handling.                                                   *
' *                                                                   *
' *                                                                   *
' *********************************************************************
Private INTdisplayInitial As Integer
Private INTraiseInitial As Integer
Private STRapplication As String
Private STRsection As String
Public Property Get Application() As String
    Application = STRapplication
End Property
Public Property Let Application(ByVal strNewValue As String)
    STRapplication = strNewValue
    registry2Form STRapplication, STRsection
End Property
Public Function loadDefaults(ByVal strApplicationName As String, _
                             ByVal strSectionName As String) As Boolean
    loadDefaults = registry2Form(strApplicationName, strSectionName)
End Function
Public Property Get InstallationMessage() As String
    InstallationMessage = _
        "Errors from this application can be displayed using this " & _
        "form, or they may be Raised, or both, or neither (this isn't " & _
        "recommended, for obvious reasons.)" & _
        vbNewLine & vbNewLine & _
        "Check Display this Form to always display this message box " & _
        "on an error.  Check Raise an Error to always Raise an error.  " & _
        "Click Continue to proceed with these settings."
End Property

Public Property Get Section() As String
    Section = STRsection
End Property
Public Property Let Section(ByVal strNewValue As String)
    STRsection = strNewValue
    registry2Form STRapplication, STRsection
End Property
Public Property Get ShowForm() As Boolean
    ShowForm = (chkDisplay = 1)
End Property
Public Property Get RaiseError() As Boolean
    RaiseError = (chkRaise = 1)
End Property
Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub cmdContinue_Click()
    form2Registry
    Hide
End Sub
Private Sub Form_Activate()
    Caption = Me.Application
End Sub
Private Sub Form_Load()
    INTdisplayInitial = chkDisplay.Value
    INTraiseInitial = chkRaise.Value
    STRapplication = App.EXEName
    STRsection = Me.Name
    registry2Form STRapplication, STRsection
End Sub
Private Sub form2Registry()
    '
    ' Save form selections
    '
    '
    SaveSetting Me.Application, Me.Section, chkRaise.Name, chkRaise.Value
    SaveSetting Me.Application, Me.Section, chkDisplay.Name, chkDisplay.Value
End Sub
Private Function registry2Form(ByVal strApplicationName As String, _
                               ByVal strSectionName As String) As Boolean
    '
    ' Get Registry settings
    '
    '
    chkRaise.Value = INTraiseInitial: chkDisplay.Value = INTdisplayInitial
    On Error GoTo registry2Form_Lbl1_errorHandler
        chkRaise.Value = GetSetting(strApplicationName, _
                                    strSectionName, _
                                    chkRaise.Name)
        chkDisplay.Value = GetSetting(strApplicationName, _
                                      strSectionName, _
                                      chkDisplay.Name)
    On Error GoTo 0
    registry2Form = True
    Exit Function
registry2Form_Lbl1_errorHandler:
    chkRaise.Value = INTraiseInitial: chkDisplay.Value = INTdisplayInitial
End Function
