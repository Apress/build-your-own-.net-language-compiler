VERSION 5.00
Begin VB.Form frmRefManOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reference Manual Options"
   ClientHeight    =   6972
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6972
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFormat 
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   6015
      Begin VB.CheckBox chkFormatXMLendTagComments 
         Caption         =   "Comment endtags with the associate name in the start tag"
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         ToolTipText     =   $"refManOptions.frx":0000
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CheckBox chkFormatXMLinlineBNF 
         Caption         =   "Add BNF source code as an attribute to production tags"
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         ToolTipText     =   $"refManOptions.frx":008D
         Top             =   1800
         Width           =   4215
      End
      Begin VB.CheckBox chkFormatXMLnewlines 
         Caption         =   "Individual tags should appear on separate lines"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         ToolTipText     =   "Deselect for an XML file that is hard to read but takes up significantly less space"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox chkFormatIncludeXMLsource 
         Caption         =   "The XML should include the BNF source as a comment"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         ToolTipText     =   "Places the source code of the BNF in an XML header comment. Not suitable for large files."
         Top             =   1080
         Width           =   4335
      End
      Begin VB.OptionButton optFormatXML 
         Caption         =   "XML format"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   $"refManOptions.frx":011D
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optFormatText 
         Caption         =   "Text format"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Select for a proprietary text view of the language reference manual as an outline"
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CheckBox chkFormatBoxComments 
         Caption         =   "Place sections of the manual in boxed comments (note: not suitable for large files)"
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         ToolTipText     =   "Will transform the reference manual into a boxed comment suitable for inclusion in your source code"
         Top             =   240
         Width           =   4095
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5760
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.Frame fraSyntax 
      Caption         =   "Syntax Reference"
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   6015
      Begin VB.CheckBox chkIndentReferenceManual 
         Caption         =   "Indent the language reference manual"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Value           =   1  'Checked
         Width           =   5775
      End
      Begin VB.CheckBox chkIncludeBNF 
         Caption         =   "Include formal Backus-Naur definitions"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   5775
      End
      Begin VB.CheckBox chkIncludeSyntax 
         Caption         =   "Include syntax"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   5775
      End
      Begin VB.CheckBox chkIncludeWhere 
         Caption         =   "Include lists showing where symbols are used, in addition to their definitions, in the syntax"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Width           =   5655
      End
   End
   Begin VB.CheckBox chkIncludeTerminalIndex 
      Caption         =   "Include terminal index"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Value           =   1  'Checked
      Width           =   6135
   End
   Begin VB.CheckBox chkIncludeNonterminalIndex 
      Caption         =   "Include non-terminal index"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Value           =   1  'Checked
      Width           =   6135
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Always show this form before producing the reference manual"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Value           =   1  'Checked
      Width           =   6135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close (and create manual)"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdRestoreSettings 
      Caption         =   "Restore Settings"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Click to restore the settings you have selected from your Registry"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdForm2Registry 
      Caption         =   "Save Settings"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click to save the settings you have selected to your Registry for future use"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtLanguageName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "anonymousLanguage"
      ToolTipText     =   "Identify the name of the language for inclusion in the report"
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblLanguageName 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Language Name"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRefManOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' *********************************************************************
' *                                                                   *
' * Reference Manual: Options                                         *
' *                                                                   *
' *                                                                   *
' * This form provides options to the creation of the language refer- *
' * ence manual, and, it allows the user to cancel the creation.  This*
' * form exposes the following Friend properties.                     *
' *                                                                   *
' *                                                                   *
' *      *  The read-only property BoxComments returns True when the  *
' *         form's check box (to place the reference manual sections  *
' *         in comment boxes) is selected.                            *
' *                                                                   *
' *      *  The read-only property Cancel returns True unless the     *
' *         form was dismissed using the Close button.                *
' *                                                                   *
' *      *  The read-only property IncludeBNF returns True when the   *
' *         form's check box (to include formal BNF in the manual) is *
' *         selected.                                                 *
' *                                                                   *
' *      *  The read-only property IncludeNonterminals returns True   *
' *         when the form's check box (to include the nonterminal     *
' *         index) is selected.                                       *
' *                                                                   *
' *      *  The read-only property IncludeSyntax returns True         *
' *         when the form's check box (to include the syntax) is      *
' *         selected.                                                 *
' *                                                                   *
' *      *  The read-only property IncludeTerminals returns True      *
' *         when the form's check box (to include the terminal        *
' *         index) is selected.                                       *
' *                                                                   *
' *      *  The read-only property IncludeWhere returns True when the *
' *         form's check box (to include where-used in the manual) is *
' *         selected.                                                 *
' *                                                                   *
' *      *  The read-only property IncludeXMLsource returns True      *
' *         when the form's check box (to include the source as part  *
' *         of XML output) is selected.                               *
' *                                                                   *
' *      *  The read-only property IndentReferenceManual returns True *
' *         when the form's check box (to indent the rules in the     *
' *         skeleton reference manual) is selected.                   *
' *                                                                   *
' *      *  The read-write property LanguageName returns the language *
' *         name specified on the form.  In addition, this property   *
' *         can be preassigned to show a default language name.       *
' *                                                                   *
' *      *  The read-only property ShowForm returns True when the     *
' *         form's check box (to show the form before creating the    *
' *         reference manual) is selected.                            *
' *                                                                   *
' *      *  The read-only property XML returns True when the          *
' *         reference manual needs to be returned in XML format.      *
' *                                                                   *
' *      *  The read-only property XMLendtagComments returns True when*
' *         the end tag for each production should be commented with  *
' *         the grammar symbol name                                   *
' *                                                                   *
' *      *  The read-only property XMLinlineBNF returns True when     *
' *         the start tag for each production should be commented with*
' *         the production source code.                               *
' *                                                                   *
' *      *  The read-only property XMLnewline returns True if         *
' *         individual XML tags are to be on new lines, False         *
' *         otherwise.                                                *
' *                                                                   *
' *                                                                   *
' *********************************************************************

Private BOOcancel As Boolean
Private OBJutilities As clsUtilities
Friend Property Get BoxComments() As Boolean
    BoxComments = (chkFormatBoxComments.Value = 1)
End Property
Friend Property Get Cancel() As Boolean
    Cancel = BOOcancel
End Property
Friend Property Get IncludeBNF() As Boolean
    IncludeBNF = (chkIncludeBNF = 1)
End Property
Friend Property Get IncludeNonterminals() As Boolean
    IncludeNonterminals = (chkIncludeNonterminalIndex = 1)
End Property
Friend Property Get IncludeSyntax() As Boolean
    IncludeSyntax = (chkIncludeSyntax = 1)
End Property
Friend Property Get IncludeTerminals() As Boolean
    IncludeTerminals = (chkIncludeTerminalIndex = 1)
End Property
Friend Property Get IncludeWhere() As Boolean
    IncludeWhere = (chkIncludeWhere = 1)
End Property
Friend Property Get IndentReferenceManual() As Boolean
    IndentReferenceManual = (chkIndentReferenceManual = 1)
End Property
Friend Property Get LanguageName() As String
    LanguageName = txtLanguageName.Text
End Property
Friend Property Let LanguageName(ByVal strNewValue As String)
    txtLanguageName.Text = strNewValue
End Property
Friend Property Get ShowForm() As Boolean
    ShowForm = (chkShow = 1)
End Property
Friend Property Get XML() As Boolean
    If optFormatXML Then
        XML = True
    ElseIf Not optFormatText Then
        ' Repair the display quietly
        optFormatText = True
    End If
End Property
Friend Property Get XMLendtagComments() As Boolean
    XMLendtagComments = chkFormatXMLendTagComments
End Property
Friend Property Get XMLinlineBNF() As Boolean
    XMLinlineBNF = chkFormatXMLinlineBNF
End Property
Friend Property Get XMLnewline() As Boolean
    XMLnewline = (chkFormatXMLnewlines.Value = 1)
End Property
Friend Property Get IncludeXMLsource() As Boolean
    IncludeXMLsource = chkFormatIncludeXMLsource
End Property

Private Sub chkIncludeSyntax_Click()
    Dim booEnable As Boolean
    booEnable = chkIncludeSyntax.Value
    chkIncludeBNF.Enabled = booEnable
    chkIncludeWhere.Enabled = booEnable
End Sub

Private Sub cmdCancel_Click()
    registry2Form
    Hide
End Sub

Private Sub cmdClose_Click()
    form2Registry
    If Not Me.ShowForm Then
        MsgBox "Note: you can restore the option to show this form " & _
               "by selecting ""reference manual options"" on the main form's " & _
               "menu"
    End If
    BOOcancel = False
    Hide
End Sub

Private Sub cmdForm2Registry_Click()
    form2Registry
End Sub

Private Sub cmdRestoreSettings_Click()
    registry2Form
End Sub

Private Sub Form_Activate()
    BOOcancel = True
End Sub
Private Sub Form_Load()
    On Error GoTo Form_Load_Lbl1_utilitiesErrorHandler
        Set OBJutilities = New clsUtilities
    On Error GoTo 0
    registry2Form
    Exit Sub
Form_Load_Lbl1_utilitiesErrorHandler:
    MsgBox "Can't create utilities: " & Err.Number & " " & Err.Description
End Sub
Private Sub registry2Form()
    '
    ' Obtain form selections
    '
    '
    On Error GoTo registry2Form_Lbl1_errorHandler
        With txtLanguageName
            .Text = GetSetting(App.EXEName, Me.Name, .Name, .Text)
        End With
        With chkFormatBoxComments
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkIncludeBNF
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkIncludeNonterminalIndex
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkIncludeSyntax
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkIncludeTerminalIndex
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkIncludeWhere
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkShow
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkIndentReferenceManual
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With optFormatText
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With optFormatXML
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkFormatIncludeXMLsource
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkFormatXMLnewlines
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkFormatXMLendTagComments
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        With chkFormatXMLinlineBNF
            .Value = GetSetting(App.EXEName, Me.Name, .Name, .Value)
        End With
        adjustFormatDisplay
    On Error GoTo 0
    Exit Sub
registry2Form_Lbl1_errorHandler:
    OBJutilities.errorHandler 0, _
                              "Can't restore form selections: " & _
                              Err.Number & " " & Err.Description, _
                              "registry2Form", _
                              Me.Name
End Sub
Private Sub form2Registry()
    '
    ' Restore form selections
    '
    '
    On Error GoTo form2Registry_Lbl1_errorHandler
        With txtLanguageName
            SaveSetting App.EXEName, Me.Name, .Name, .Text
        End With
        With chkFormatBoxComments
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkIncludeBNF
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkIncludeNonterminalIndex
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkIncludeSyntax
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkIncludeTerminalIndex
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkIncludeWhere
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkShow
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkIndentReferenceManual
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With optFormatText
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With optFormatXML
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkFormatIncludeXMLsource
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkFormatXMLnewlines
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkFormatXMLendTagComments
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
        With chkFormatXMLinlineBNF
            SaveSetting App.EXEName, Me.Name, .Name, .Value
        End With
    On Error GoTo 0
    Exit Sub
form2Registry_Lbl1_errorHandler:
    OBJutilities.errorHandler 0, _
                              "Can't save form selections: " & _
                              Err.Number & " " & Err.Description, _
                              "form2Registry", _
                              Me.Name
End Sub
Private Sub optFormatText_Click()
    adjustFormatDisplay
End Sub
Private Sub adjustFormatDisplay()
    '
    ' Adjust the text versus XML format display
    '
    '
    chkFormatBoxComments.Enabled = optFormatText
    chkIncludeBNF.Enabled = optFormatText
    chkIncludeWhere.Enabled = optFormatText
    chkIndentReferenceManual.Enabled = optFormatText
    chkFormatIncludeXMLsource.Enabled = Not optFormatText
    chkFormatXMLnewlines.Enabled = Not optFormatText
    chkFormatXMLendTagComments.Enabled = Not optFormatText
    chkFormatXMLinlineBNF.Enabled = Not optFormatText
End Sub
Private Sub optFormatXML_Click()
    adjustFormatDisplay
End Sub
